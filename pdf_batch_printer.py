#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件旋转打印（右键菜单集成，单文件 exe）
=============================================

核心设计
--------
所有支持格式先转为 PDF；页码范围、奇偶页过滤、180° 旋转和份数展开均在
"生成 PDF"阶段处理。Word、Excel、PowerPoint 使用本机 Office 或 LibreOffice。

处理顺序（严格固定）
-------------------
    1. 按页码范围取页
    2. 按单/双页过滤（以第 1 步过滤后的页序计）
    3. 按份数展开，偶数份按需旋转 180°

命令行
------
    文件旋转打印.exe                       打开图形界面
    文件旋转打印.exe a.pdf b.pdf 文件夹    右键菜单调用（多选/文件夹递归）
    文件旋转打印.exe --register            注册右键菜单（免管理员）
    文件旋转打印.exe --unregister          卸载右键菜单
"""

from __future__ import annotations

import ctypes
import io
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import zlib
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# pypdf 为必需依赖；pywin32 为可选依赖（仅用于枚举打印机，缺失时自动回退）
# ---------------------------------------------------------------------------
try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import DictionaryObject, NameObject
    PYPDF_OK = True
    PYPDF_ERR = ""
except Exception as exc:  # pragma: no cover
    PYPDF_OK = False
    PYPDF_ERR = str(exc)

try:
    import win32print  # type: ignore
    WIN32PRINT_OK = True
except Exception:
    WIN32PRINT_OK = False

try:
    import win32con
    import win32gui
    import win32ui
    NATIVE_PRINT_OK = True
    NATIVE_PRINT_ERR = ""
except Exception as exc:  # pragma: no cover
    NATIVE_PRINT_OK = False
    NATIVE_PRINT_ERR = str(exc)

try:
    import pypdfium2 as pdfium
    from PIL import Image, ImageSequence, ImageTk, ImageWin
    PREVIEW_OK = True
    PREVIEW_ERR = ""
except Exception as exc:  # pragma: no cover
    PREVIEW_OK = False
    PREVIEW_ERR = str(exc)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "文件旋转打印"
APP_AUTHOR = "005542"
APP_VERSION = "3.3"
MENU_KEY_NAME = "PDFBatchPrint"          # 注册表中的右键菜单项键名
MENU_LABEL = "文件旋转打印"                # 右键菜单显示文字
CONFIG_DIR_NAME = "文件旋转打印"
LEGACY_CONFIG_DIR_NAME = "PDF批量打印工具"

ODD_EVEN_CHOICES = ("全部", "仅单页", "仅双页")
DUPLEX_CHOICES = ("单面", "双面（长边翻转）")
BACKEND_CHOICES = ("自动", "SumatraPDF 命令行", "Windows 原生打印")
BACKEND_AUTO, BACKEND_SUMATRA, BACKEND_NATIVE = 0, 1, 2
WORD_EXTENSIONS = frozenset((".doc", ".docx", ".docm", ".dot", ".dotx",
                             ".dotm", ".rtf", ".txt", ".html", ".htm", ".odt"))
EXCEL_EXTENSIONS = frozenset((".xls", ".xlsx", ".xlsm", ".xlsb", ".csv",
                              ".tsv", ".ods"))
POWERPOINT_EXTENSIONS = frozenset((".ppt", ".pptx", ".pptm", ".pps", ".ppsx",
                                   ".ppsm", ".odp"))
IMAGE_EXTENSIONS = frozenset((".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"))
SUPPORTED_EXTENSIONS = frozenset((".pdf",)) | WORD_EXTENSIONS | EXCEL_EXTENSIONS | \
    POWERPOINT_EXTENSIONS | IMAGE_EXTENSIONS

# 份数 × 页数 超过该阈值时给出体积膨胀警告
PAGE_BUDGET_WARN = 3000
COPIES_WARN = 30

SW_HIDE = 0
SEE_MASK_NOCLOSEPROCESS = 0x00000040
SEE_MASK_NOASYNC = 0x00000100
INFINITE = 0xFFFFFFFF
INSTANCE_REQUESTS = queue.Queue()


def _send_instance_request(port: int, paths: list) -> bool:
    try:
        payload = json.dumps({"app": APP_NAME, "paths": paths},
                             ensure_ascii=False).encode("utf-8") + b"\n"
        with socket.create_connection(("127.0.0.1", port), timeout=2) as client:
            client.settimeout(2)
            client.sendall(payload)
            return client.recv(16) == b"OK\n"
    except OSError:
        return False


def _instance_server(server: socket.socket) -> None:
    while True:
        try:
            connection, _address = server.accept()
        except OSError:
            return
        with connection:
            connection.settimeout(5)
            payload = bytearray()
            try:
                while len(payload) <= 8 * 1024 * 1024:
                    chunk = connection.recv(65536)
                    if not chunk:
                        break
                    payload.extend(chunk)
                    if b"\n" in chunk:
                        break
                request = json.loads(bytes(payload).split(b"\n", 1)[0])
                if not isinstance(request, dict):
                    connection.sendall(b"NO\n")
                    continue
                paths = request.get("paths")
                if (request.get("app") != APP_NAME or not isinstance(paths, list)
                        or not all(isinstance(path, str) for path in paths)):
                    connection.sendall(b"NO\n")
                    continue
                INSTANCE_REQUESTS.put(paths)
                connection.sendall(b"OK\n")
            except (OSError, ValueError, TypeError):
                try:
                    connection.sendall(b"NO\n")
                except OSError:
                    pass


def _start_single_instance(paths: list, port_base: int = None):
    if port_base is None:
        user = os.environ.get("USERNAME") or os.path.expanduser("~")
        port_base = 32000 + zlib.crc32((APP_NAME + user).encode("utf-8")) % 20000
    for port in range(port_base, min(port_base + 64, 65536)):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            server.bind(("127.0.0.1", port))
            server.listen(8)
        except OSError:
            server.close()
            if _send_instance_request(port, paths):
                return None, True
            continue
        threading.Thread(target=_instance_server, args=(server,), daemon=True).start()
        return server, False
    return None, False


# ===========================================================================
# 通用工具
# ===========================================================================
def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_command_prefix() -> str:
    """右键菜单 command 的可执行前缀。"""
    if is_frozen():
        return '"%s"' % sys.executable
    return '"%s" "%s"' % (sys.executable, os.path.abspath(__file__))


def config_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, CONFIG_DIR_NAME)
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        path = tempfile.gettempdir()
    return path


def config_path() -> str:
    return os.path.join(config_dir(), "config.json")


def load_config() -> dict:
    legacy_base = os.environ.get("APPDATA") or os.path.expanduser("~")
    paths = (config_path(), os.path.join(legacy_base, LEGACY_CONFIG_DIR_NAME,
                                         "config.json"))
    for path in paths:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                return data
        except Exception:
            continue
    return {}


def save_config(cfg: dict) -> None:
    try:
        with open(config_path(), "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


def safe_name(text: str, max_len: int = 60) -> str:
    text = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", text).strip().strip(".")
    return (text[:max_len] or "file")


def collect_files(paths) -> list:
    """把传入的文件/文件夹展开为去重后的常见可打印文件列表。"""
    found, seen = [], set()
    visited_dirs = set()

    def add_file(p):
        key = os.path.normcase(os.path.abspath(p))
        if key in seen:
            return
        if not os.path.isfile(p):
            return
        if os.path.splitext(p)[1].lower() not in SUPPORTED_EXTENSIONS:
            return
        seen.add(key)
        found.append(os.path.abspath(p))

    def walk(dirpath):
        directory_key = os.path.normcase(os.path.realpath(dirpath))
        if directory_key in visited_dirs:
            return
        visited_dirs.add(directory_key)
        try:
            entries = sorted(os.listdir(dirpath))
        except Exception:
            return
        for name in entries:
            full = os.path.join(dirpath, name)
            if os.path.isdir(full):
                walk(full)
            elif os.path.splitext(name)[1].lower() in SUPPORTED_EXTENSIONS:
                add_file(full)

    for raw in paths:
        if not raw:
            continue
        p = os.path.abspath(raw)
        if os.path.isdir(p):
            walk(p)
        elif os.path.isfile(p):
            add_file(p)
    return found


class FileConversionError(RuntimeError):
    pass


def _office_to_pdf(src_path: str, out_path: str) -> None:
    """Use installed Microsoft Office COM automation to create a PDF copy."""
    import pythoncom
    import win32com.client

    ext = os.path.splitext(src_path)[1].lower()
    app = document = None
    pythoncom.CoInitialize()
    try:
        if ext in WORD_EXTENSIONS:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            app.AutomationSecurity = 3
            document = app.Documents.Open(
                FileName=os.path.abspath(src_path), ConfirmConversions=False,
                ReadOnly=True, AddToRecentFiles=False, NoEncodingDialog=True)
            document.ExportAsFixedFormat(os.path.abspath(out_path), 17)
        elif ext in EXCEL_EXTENSIONS:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.AutomationSecurity = 3
            document = app.Workbooks.Open(
                Filename=os.path.abspath(src_path), UpdateLinks=0,
                ReadOnly=True, AddToMru=False)
            document.ExportAsFixedFormat(0, os.path.abspath(out_path), 0,
                                         True, False)
        elif ext in POWERPOINT_EXTENSIONS:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = False
            app.AutomationSecurity = 3
            document = app.Presentations.Open(
                os.path.abspath(src_path), ReadOnly=True,
                Untitled=False, WithWindow=False)
            document.ExportAsFixedFormat(os.path.abspath(out_path), 2)
        else:
            raise FileConversionError("此文件类型不支持 Office 转换：%s" % ext)
        if not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
            raise FileConversionError("Office 未生成 PDF 文件")
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _libreoffice_to_pdf(src_path: str, out_path: str) -> None:
    soffice = shutil.which("soffice")
    if not soffice:
        raise FileConversionError("未安装 Microsoft Office 或 LibreOffice")
    output_dir = os.path.dirname(out_path)
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        proc = subprocess.run(
            [soffice, "--headless", "-env:UserInstallation=%s" % Path(profile_dir).as_uri(),
             "--convert-to", "pdf", "--outdir", output_dir, os.path.abspath(src_path)],
            capture_output=True, text=True, timeout=300,
            creationflags=0x08000000)
        generated = os.path.join(
            output_dir, os.path.splitext(os.path.basename(src_path))[0] + ".pdf")
        if proc.returncode != 0 or not os.path.isfile(generated):
            detail = (proc.stderr or proc.stdout or "转换器未生成 PDF").strip()[:400]
            raise FileConversionError(detail)
        if os.path.abspath(generated) != os.path.abspath(out_path):
            os.replace(generated, out_path)
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def convert_to_pdf(src_path: str, out_path: str) -> str:
    """Convert supported office/image formats to a temporary PDF."""
    ext = os.path.splitext(src_path)[1].lower()
    if ext == ".pdf":
        return src_path
    if ext in IMAGE_EXTENSIONS:
        if not PREVIEW_OK:
            raise FileConversionError("图片转换需要 Pillow：%s" % PREVIEW_ERR)
        try:
            with Image.open(src_path) as image:
                frames = [frame.convert("RGB") for frame in ImageSequence.Iterator(image)]
                if not frames:
                    raise FileConversionError("图片文件没有可用页面")
                frames[0].save(out_path, "PDF", save_all=True,
                               append_images=frames[1:], resolution=150.0)
        except FileConversionError:
            raise
        except Exception as exc:
            raise FileConversionError("图片转 PDF 失败：%s" % exc)
        return out_path

    try:
        _office_to_pdf(src_path, out_path)
    except Exception as office_exc:
        try:
            _libreoffice_to_pdf(src_path, out_path)
        except Exception as libreoffice_exc:
            raise FileConversionError(
                "%s 转 PDF 失败（Office：%s；LibreOffice：%s）" %
                (os.path.basename(src_path), office_exc, libreoffice_exc))
    return out_path


def count_file_pages(src_path: str) -> int:
    """Count printable pages through the same conversion path used for printing."""
    if os.path.splitext(src_path)[1].lower() == ".pdf":
        return len(PdfReader(src_path).pages)
    temp_dir = tempfile.mkdtemp(prefix="pdfcount_")
    try:
        converted_path = os.path.join(temp_dir, "count.pdf")
        pdf_path = convert_to_pdf(src_path, converted_path)
        return len(PdfReader(pdf_path).pages)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def page_range_presets(total_pages: int) -> dict:
    presets = {"全部": "", "第1页": "1", "最后一页": "last"}
    if total_pages > 0:
        presets["第1页~第%d页" % total_pages] = "1-%d" % total_pages
    return presets


# ===========================================================================
# F4 页码范围解析
# ===========================================================================
class PageRangeError(ValueError):
    pass


def parse_page_range(text, total_pages: int):
    """解析页码范围，返回 0 基页索引列表（保留书写顺序、自动去重）。

    支持：``1-3,5,last``、``1-last``、``all``、空串（= 全部）。
    越界与非法段自动忽略，不抛异常。
    """
    if total_pages <= 0:
        return []
    raw = (text or "").strip()
    if raw == "" or raw.lower() == "all":
        return list(range(total_pages))

    picked, seen = [], set()

    def push(idx0):
        if 0 <= idx0 < total_pages and idx0 not in seen:
            seen.add(idx0)
            picked.append(idx0)

    def page_number(token):
        return total_pages if token.lower() == "last" else int(token)

    for part in raw.replace("，", ",").split(","):
        part = part.strip()
        if not part or part.lower() == "all":
            if part.lower() == "all":
                for i in range(total_pages):
                    push(i)
            continue
        m = re.fullmatch(r"(\d+|last)\s*-\s*(\d+|last)", part, re.IGNORECASE)
        if m:
            lo, hi = page_number(m.group(1)), page_number(m.group(2))
            if lo > hi:
                lo, hi = hi, lo
            for num in range(lo, hi + 1):
                push(num - 1)          # 越界自动忽略
            continue
        if re.fullmatch(r"\d+|last", part, re.IGNORECASE):
            push(page_number(part) - 1)
            continue
        # 非法段：自动忽略（不中断整体解析）
    return picked


def filter_odd_even(indices, mode: str):
    """单/双页过滤。以传入序列的页序计：第 1 页为单页，第 2 页为双页。"""
    if mode == "仅单页":
        return [v for pos, v in enumerate(indices) if pos % 2 == 0]
    if mode == "仅双页":
        return [v for pos, v in enumerate(indices) if pos % 2 == 1]
    return list(indices)


def output_page_total(total_pages: int, page_range: str, odd_even: str,
                      copies: int) -> int:
    picked = parse_page_range(page_range, total_pages)
    picked = filter_odd_even(picked, odd_even)
    return len(picked) * max(1, min(99, int(copies or 1)))


def order_rows_by_page_total(rows, totals):
    """Stable ascending queue order; unknown totals remain at the end."""
    return [row for _idx, row in sorted(
        enumerate(rows),
        key=lambda pair: (totals.get(id(pair[1]), float("inf")), pair[0]))]


# ===========================================================================
# 核心：生成变换后的 PDF
# ===========================================================================
class GenerateResult:
    def __init__(self):
        self.out_path = ""
        self.src_pages = 0
        self.picked_pages = 0
        self.out_pages = 0
        self.skipped_ranges = ""


def build_pdf(src_path: str, out_path: str, *, page_range: str = "",
              odd_even: str = "全部", copies: int = 1,
              alternate_180: bool = False, file_rotate_180: bool = False,
              duplex: str = "单面",
              progress=None) -> GenerateResult:
    """生成变换后的 PDF。

    按文件和份数位置分别应用 180° 旋转；两种交替旋转可叠加。
    """
    res = GenerateResult()

    reader = PdfReader(src_path)
    if getattr(reader, "is_encrypted", False):
        # 注意：pypdf 的 is_encrypted 在解密成功后仍为 True，
        # 因此必须用 decrypt() 的返回值判断（0 = 解密失败），而非该标志位。
        try:
            rc = reader.decrypt("")     # 尝试空口令
        except Exception:
            rc = 0
        if not rc:
            raise PageRangeError("文件已加密，需要打开口令（V1 不支持带口令的 PDF）")
        try:
            _ = len(reader.pages)       # 再验证一次页面确实可访问
        except Exception as exc:
            raise PageRangeError("文件已加密且无法用空口令打开：%s" % exc)

    try:
        total = len(reader.pages)
    except Exception as exc:
        raise PageRangeError("读取页面失败：%s" % exc)
    res.src_pages = total
    if total == 0:
        raise PageRangeError("PDF 无页面（空文件）")

    # ---- 第 1 步：页码范围 ----
    picked = parse_page_range(page_range, total)
    if not picked:
        raise PageRangeError(
            "页码范围「%s」在 %d 页内无有效页" % (page_range or "(空)", total))

    # ---- 第 2 步：单/双页过滤 ----
    picked = filter_odd_even(picked, odd_even)
    if not picked:
        raise PageRangeError(
            "按「%s」过滤后无剩余页面（有效页 %d 页）" % (odd_even, res.src_pages))
    res.picked_pages = len(picked)

    copies = max(1, min(99, int(copies or 1)))
    # ---- 第 3 步：份数展开 + 偶数份可选180°旋转 ----
    writer = PdfWriter()
    out_count = 0
    for copy_idx in range(copies):
        add_180 = (180 if file_rotate_180 else 0) + (
            180 if alternate_180 and copy_idx % 2 == 1 else 0)
        for src_idx in picked:
            try:
                page = writer.add_page(reader.pages[src_idx])
            except Exception as exc:
                raise PageRangeError(
                    "第 %d 页复制失败：%s" % (src_idx + 1, exc))
            try:
                base = int(page.rotation or 0)      # 源页原有旋转，累加不覆盖
            except Exception:
                base = 0
            page.rotation = (base + add_180) % 360
            out_count += 1
        if progress:
            progress(copy_idx + 1, copies)

    if out_count == 0:
        raise PageRangeError("未生成任何页面")

    # 去掉源文件的元数据/大纲，避免无效引用导致阅读器报错
    try:
        writer.remove_links()
    except Exception:
        pass
    try:
        if writer.outline:
            writer._root_object.pop("/Outlines", None)
    except Exception:
        pass

    viewer_preferences = DictionaryObject()
    viewer_preferences[NameObject("/Duplex")] = NameObject(
        "/DuplexFlipLongEdge" if duplex == "双面（长边翻转）" else "/Simplex")
    writer._root_object[NameObject("/ViewerPreferences")] = viewer_preferences

    tmp_path = out_path + ".part"
    with open(tmp_path, "wb") as fh:
        writer.write(fh)
    os.replace(tmp_path, out_path)

    res.out_path = out_path
    res.out_pages = out_count
    return res


# ===========================================================================
# F7 打印机枚举
# ===========================================================================
def list_printers() -> list:
    names = []
    if WIN32PRINT_OK:
        try:
            flags = (win32print.PRINTER_ENUM_LOCAL |
                     win32print.PRINTER_ENUM_CONNECTIONS)
            for item in win32print.EnumPrinters(flags):
                try:
                    names.append(item[2])
                except Exception:
                    continue
            if names:
                return sorted(set(names))
        except Exception:
            pass
    # 回退：PowerShell 枚举
    for cmd in (
        ["powershell", "-NoProfile", "-Command",
         "Get-Printer | Select-Object -ExpandProperty Name"],
        ["wmic", "printer", "get", "name", "/value"],
    ):
        try:
            out = subprocess.run(cmd, capture_output=True, text=True,
                                 timeout=25, creationflags=0x08000000)
            text = out.stdout or ""
            if "name=" in text.lower():
                for line in text.splitlines():
                    line = line.strip()
                    if line.lower().startswith("name="):
                        v = line.split("=", 1)[1].strip()
                        if v:
                            names.append(v)
            else:
                for line in text.splitlines():
                    line = line.strip()
                    if line:
                        names.append(line)
            if names:
                return sorted(set(names))
        except Exception:
            continue
    return names


def default_printer_name() -> str:
    if WIN32PRINT_OK:
        try:
            return win32print.GetDefaultPrinter()
        except Exception:
            pass
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "(Get-CimInstance Win32_Printer -Filter \"Default=True\").Name"],
            capture_output=True, text=True, timeout=25, creationflags=0x08000000)
        return (out.stdout or "").strip()
    except Exception:
        return ""


# ===========================================================================
# F7 打印后端
# ===========================================================================
def find_sumatra() -> str:
    """查找 SumatraPDF.exe，找不到返回空串。"""
    cands = []
    for env_key in ("SUMATRA_PDF", "SUMATRAPDF"):
        v = os.environ.get(env_key)
        if v:
            cands.append(v)
    which = shutil.which("SumatraPDF.exe") or shutil.which("SumatraPDF")
    if which:
        cands.append(which)
    local = os.environ.get("LOCALAPPDATA", "")
    pf = os.environ.get("PROGRAMFILES", r"C:\Program Files")
    pf86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
    cands += [
        os.path.join(local, "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(pf, "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(pf86, "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.path.dirname(sys.executable), "SumatraPDF.exe"),
        os.path.join(os.getcwd(), "SumatraPDF.exe"),
    ]
    for c in cands:
        try:
            if c and os.path.isfile(c):
                return c
        except Exception:
            continue
    return ""


def _native_print_pdf(pdf_path: str, printer: str, duplex: str, log=None):
    """Render PDF pages with PDFium and send them directly to a Windows printer."""
    if not NATIVE_PRINT_OK or not PREVIEW_OK:
        return False, "Windows 原生打印依赖不可用：%s %s" % (
            NATIVE_PRINT_ERR, PREVIEW_ERR)

    printer_handle = dc = document = None
    job_started = False
    try:
        printer_handle = win32print.OpenPrinter(printer)
        devmode = win32print.GetPrinter(printer_handle, 2)["pDevMode"]
        devmode.Duplex = (win32con.DMDUP_VERTICAL
                          if duplex == "双面（长边翻转）"
                          else win32con.DMDUP_SIMPLEX)
        devmode.Fields |= win32con.DM_DUPLEX
        hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
        dc = win32ui.CreateDCFromHandle(hdc)

        printable_width = dc.GetDeviceCaps(win32con.HORZRES)
        printable_height = dc.GetDeviceCaps(win32con.VERTRES)
        offset_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
        offset_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
        dpi_x = max(72, dc.GetDeviceCaps(win32con.LOGPIXELSX))
        dpi_y = max(72, dc.GetDeviceCaps(win32con.LOGPIXELSY))

        document = pdfium.PdfDocument(pdf_path)
        dc.StartDoc(os.path.basename(pdf_path))
        job_started = True
        for page_index in range(len(document)):
            page = document[page_index]
            page_width, page_height = page.get_size()
            width_at_dpi = page_width * dpi_x / 72.0
            height_at_dpi = page_height * dpi_y / 72.0
            fit = min(1.0, printable_width / width_at_dpi,
                      printable_height / height_at_dpi)
            dest_width = max(1, int(width_at_dpi * fit))
            dest_height = max(1, int(height_at_dpi * fit))
            render_scale = min(dpi_x, dpi_y, 300) * fit / 72.0
            image = page.render(scale=render_scale).to_pil().convert("RGB")
            left = offset_x + (printable_width - dest_width) // 2
            top = offset_y + (printable_height - dest_height) // 2

            dc.StartPage()
            ImageWin.Dib(image).draw(
                dc.GetHandleOutput(),
                (left, top, left + dest_width, top + dest_height))
            dc.EndPage()
            image.close()
            page.close()
        dc.EndDoc()
        job_started = False
        return True, ""
    except Exception as exc:
        if job_started and dc is not None:
            try:
                dc.AbortDoc()
            except Exception:
                pass
        return False, "Windows 原生打印失败：%s" % exc
    finally:
        if document is not None:
            try:
                document.close()
            except Exception:
                pass
        if dc is not None:
            try:
                dc.DeleteDC()
            except Exception:
                pass
        if printer_handle is not None:
            try:
                win32print.ClosePrinter(printer_handle)
            except Exception:
                pass


def print_one(pdf_path: str, printer: str, backend: int, duplex: str = "单面",
              log=None):
    """打印单个 PDF。返回 (ok, message, used_backend)。

    注意：不向驱动传递份数与旋转设置，份数已通过页面展开实现，
    驱动的"旋转 180°""逐份打印"必须保持默认，避免互相抵消。
    """
    def say(msg):
        if log:
            log(msg)

    sumatra = find_sumatra()
    if backend == BACKEND_SUMATRA and not sumatra:
        return False, "未找到 SumatraPDF.exe，无法使用 SumatraPDF 后端", "SumatraPDF"
    use_sumatra = backend == BACKEND_SUMATRA or (
        backend == BACKEND_AUTO and bool(sumatra))
    if use_sumatra:
        say("    后端：SumatraPDF（%s）" % sumatra)
    else:
        say("    后端：Windows 原生打印（PDFium）")

    if use_sumatra:
        duplex_setting = ("duplexlong" if duplex == "双面（长边翻转）"
                          else "simplex")
        cmd = [sumatra, "-print-to", printer, "-print-settings",
               duplex_setting, "-silent", pdf_path]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  timeout=300, creationflags=0x08000000)
        except subprocess.TimeoutExpired:
            return False, "SumatraPDF 打印超时（>300 秒）", "SumatraPDF"
        except Exception as exc:
            return False, "调用 SumatraPDF 失败：%s" % exc, "SumatraPDF"
        if proc.returncode not in (0, None):
            detail = (proc.stderr or proc.stdout or "").strip()[:300]
            return False, "SumatraPDF 返回码 %s %s" % (proc.returncode, detail), "SumatraPDF"
        return True, "", "SumatraPDF"

    ok, err = _native_print_pdf(pdf_path, printer, duplex, log=log)
    return (True, "", "Windows 原生打印") if ok else (False, err, "Windows 原生打印")


# ===========================================================================
# F9 右键菜单注册（HKEY_CURRENT_USER，免管理员）
# ===========================================================================
def _menu_targets():
    return [
        (r"Software\Classes\*\shell", "文件右键（支持多选）", "%1"),
        (r"Software\Classes\Directory\shell", "文件夹右键", "%1"),
        (r"Software\Classes\Directory\Background\shell", "文件夹空白处右键", "%V"),
    ]


def register_menu(exe_path: str = None) -> tuple:
    """注册右键菜单到当前用户配置单元，无需管理员权限。"""
    import winreg
    prefix = exe_path or app_command_prefix()
    icon = sys.executable if is_frozen() else None
    msgs, errors = [], []
    for base, desc, placeholder in _menu_targets():
        argument = placeholder if placeholder == "%*" else '"%s"' % placeholder
        key_path = base + "\\" + MENU_KEY_NAME
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, MENU_LABEL)
                winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
                if icon:
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, '"%s",0' % icon)
                with winreg.CreateKey(key, "command") as ckey:
                    winreg.SetValueEx(ckey, "", 0, winreg.REG_SZ,
                                      '%s %s' % (prefix, argument))
            msgs.append("已注册：%s" % desc)
        except Exception as exc:
            errors.append("%s 失败：%s" % (desc, exc))
    return msgs, errors


def unregister_menu() -> tuple:
    import winreg
    msgs, errors = [], []
    for base, desc, _ph in _menu_targets():
        key_path = base + "\\" + MENU_KEY_NAME
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + "\\command")
        except FileNotFoundError:
            pass
        except Exception as exc:
            errors.append("%s(command)：%s" % (desc, exc))
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            msgs.append("已卸载：%s" % desc)
        except FileNotFoundError:
            msgs.append("未安装，跳过：%s" % desc)
        except Exception as exc:
            errors.append("%s 失败：%s" % (desc, exc))
    return msgs, errors


def menu_registered() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Classes\*\shell" + "\\" + MENU_KEY_NAME):
            return True
    except Exception:
        return False


# ===========================================================================
# GUI
# ===========================================================================
HELP_TEXT = """作者：005542
【参数含义】
作者：005542
· 支持 PDF、Word、Excel、PowerPoint、RTF、文本、常见图片及 OpenDocument 文件。
· Word/Excel/PowerPoint 由本机 Microsoft Office 或 LibreOffice 转成 PDF 后打印。
· 页码范围：1-3,5,last / 1-last / all；勾选“覆盖页码范围”可同时应用到列表中的全部文件。
· 单/双页：按页码范围筛选后的页序取奇数页或偶数页。
· 多份旋转：勾选后第 2、4、6…份旋转 180°，第 1、3、5…份保持原方向；单份不旋转。
· 多文件旋转：按实际打印顺序，第 2、4、6…个文件整体旋转 180°；可与多份旋转同时使用。
· 按页数排队：少页文件优先，同一总打印页数的文件连续提交，页数相同时保留添加顺序。
· 打印面：单面或双面长边翻转；SumatraPDF 后端可明确指定该选项。

【文件与打印】
· 添加文件夹会递归扫描全部子文件夹，并只加入已支持的文件类型。
· 自动后端优先使用 SumatraPDF，否则通过 Windows 原生打印，不依赖 PDF 文件关联。
· Office 文档中的宏不会运行；源文件只读打开，不会被修改或覆盖。
· 预览模式只生成 PDF，不提交打印；生成文件保留在临时目录。
"""


class Row:
    """表格中一行（一个待打印文件）的参数。"""

    def __init__(self, path, page_range="", copies=1, odd_even="全部"):
        self.path = path
        self.page_range = page_range
        self.page_count = None
        self.page_count_error = ""
        self.copies = copies
        self.odd_even = odd_even
        self.status = "待处理"
        self.detail = ""


class App(tk.Tk):
    COLS = ("file", "range", "copies", "oddeven", "status")

    def __init__(self, init_paths=None):
        super().__init__()
        self.title("%s v%s" % (APP_NAME, APP_VERSION))
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = max(680, min(1040, screen_width - 32))
        height = max(480, min(720, screen_height - 48))
        self.geometry("%dx%d" % (width, height))
        self.minsize(min(760, width), min(500, height))

        self.cfg = load_config()
        self.rows: list = []
        self.log_q: "queue.Queue" = queue.Queue()
        self.worker = None
        self.busy = False
        self.temp_dir = ""
        self.editor = None
        self._init_paths = init_paths or []

        self._build_style()
        self._build_ui()
        self._load_state()
        self.after(80, self._drain_log)
        self.after(100, self._drain_instance_requests)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        if is_frozen():
            self.after(250, self._refresh_registered_menu)

        if self._init_paths:
            self.after(200, lambda: self.add_paths(self._init_paths))
        if not PYPDF_OK:
            self.log("严重错误：未找到 pypdf 库（%s）" % PYPDF_ERR)
            messagebox.showerror(APP_NAME, "缺少必需依赖 pypdf，无法生成 PDF。\n%s" % PYPDF_ERR)

    def _refresh_registered_menu(self):
        if not menu_registered():
            return
        messages, errors = register_menu()
        if errors:
            self.log("右键菜单自动更新失败：%s" % "；".join(errors), "err")
        elif messages:
            self.log("已更新已注册的右键菜单名称和程序路径。", "ok")

    # ---------------- UI ----------------
    def _build_style(self):
        try:
            style = ttk.Style(self)
            if "vista" in style.theme_names():
                style.theme_use("vista")
            style.configure("Treeview", rowheight=24)
            style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 9, "bold"))
            style.configure("TButton", font=("Microsoft YaHei UI", 9))
            style.configure("TLabel", font=("Microsoft YaHei UI", 9))
            style.configure("TCheckbutton", font=("Microsoft YaHei UI", 9))
            style.configure("TRadiobutton", font=("Microsoft YaHei UI", 9))
            style.configure("Hint.TLabel", font=("Microsoft YaHei UI", 8),
                            foreground="#666666")
            style.configure("Warn.TLabel", font=("Microsoft YaHei UI", 8),
                            foreground="#a15c00")
        except Exception:
            pass

    def _build_ui(self):
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        # ---- 文件列表 ----
        frm_files = ttk.LabelFrame(root, text=" 待打印文件（双击单元格可修改该行参数） ", padding=6)
        frm_files.pack(fill="both", expand=True)

        bar = ttk.Frame(frm_files)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="添加文件…", command=self.pick_files).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="添加文件夹…", command=self.pick_folder).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="移除选中", command=self.remove_selected).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="清空列表", command=self.clear_rows).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="页面预览", command=self.preview_selected).pack(side="left", padx=(0, 6))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        tools = ttk.Menubutton(bar, text="工具")
        tools_menu = tk.Menu(tools, tearoff=False)
        tools_menu.add_command(label="打开临时目录", command=self.open_temp_dir)
        tools_menu.add_command(label="环境自检", command=self.show_selftest)
        tools_menu.add_command(label="使用说明", command=self.show_help)
        tools.configure(menu=tools_menu)
        tools.pack(side="left")
        ttk.Label(bar, text="作者：%s" % APP_AUTHOR,
              style="Hint.TLabel").pack(side="right", padx=(8, 4))

        cols_wrap = ttk.Frame(frm_files)
        cols_wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(cols_wrap, columns=self.COLS, show="headings",
                                 selectmode="extended")
        headers = {
            "file": ("文件", 380),
            "range": ("页码范围", 170),
            "copies": ("份数", 60),
            "oddeven": ("单/双页", 90),
            "status": ("状态", 240),
        }
        for c in self.COLS:
            text, width = headers[c]
            self.tree.heading(c, text=text)
            self.tree.column(c, width=width, anchor="w", stretch=(c in ("file", "status")))
        vsb = ttk.Scrollbar(cols_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._on_tree_double)
        self.tree.tag_configure("ok", foreground="#1c6b2f")
        self.tree.tag_configure("fail", foreground="#a01818")
        self.tree.tag_configure("disabled", foreground="#999999")

        self.lbl_count = ttk.Label(frm_files, style="Hint.TLabel", text="")
        self.lbl_count.pack(anchor="w", pady=(4, 0))

        # ---- 全局参数 ----
        frm_glob = ttk.LabelFrame(root, text=" 全局参数 ", padding=6)
        frm_glob.pack(fill="x", pady=(8, 0))

        self.v_override_range = tk.BooleanVar(value=False)
        self.v_override_copies = tk.BooleanVar(value=False)
        self.v_override_oddeven = tk.BooleanVar(value=False)

        # 第 1 行：可覆盖的三项
        line1 = ttk.Frame(frm_glob)
        line1.pack(fill="x")
        ttk.Checkbutton(line1, text="覆盖页码范围", variable=self.v_override_range,
                        command=self._apply_override_state).pack(side="left")
        self.ent_g_range = ttk.Entry(line1, width=16)
        self.ent_g_range.pack(side="left", padx=(2, 4))
        self.ent_g_range.insert(0, "all")
        self.cmb_global_range_preset = ttk.Combobox(
            line1, values=("自定义", "全部", "第1页", "最后一页"),
            width=9, state="readonly")
        self.cmb_global_range_preset.set("全部")
        self.cmb_global_range_preset.pack(side="left", padx=(0, 6))
        self.cmb_global_range_preset.bind(
            "<<ComboboxSelected>>", self._apply_global_range_preset)
        ttk.Label(line1, style="Hint.TLabel",
                  text="如 1-3,5,last / 1-last").pack(side="left", padx=(0, 16))

        ttk.Checkbutton(line1, text="覆盖份数", variable=self.v_override_copies,
                        command=self._apply_override_state).pack(side="left")
        self.spn_g_copies = ttk.Spinbox(line1, from_=1, to=99, width=5)
        self.spn_g_copies.pack(side="left", padx=(2, 16))
        self.spn_g_copies.set(1)

        ttk.Checkbutton(line1, text="覆盖单/双页", variable=self.v_override_oddeven,
                        command=self._apply_override_state).pack(side="left")
        self.cmb_g_oddeven = ttk.Combobox(line1, values=ODD_EVEN_CHOICES, width=9,
                                          state="readonly")
        self.cmb_g_oddeven.pack(side="left", padx=(2, 0))
        self.cmb_g_oddeven.set("全部")

        # 第 2 行：同文件与跨文件旋转
        line2 = ttk.Frame(frm_glob)
        line2.pack(fill="x", pady=(6, 0))
        self.v_alternate_180 = tk.BooleanVar(value=False)
        ttk.Checkbutton(line2, text="同一文件多份隔份旋转 180°",
                        variable=self.v_alternate_180).pack(side="left", padx=(0, 16))
        self.v_alternate_files_180 = tk.BooleanVar(value=False)
        ttk.Checkbutton(line2, text="不同文件隔份旋转 180°",
                        variable=self.v_alternate_files_180).pack(side="left")

        line2b = ttk.Frame(frm_glob)
        line2b.pack(fill="x", pady=(4, 0))
        ttk.Label(line2b, text="打印面：").pack(side="left")
        self.cmb_duplex = ttk.Combobox(line2b, values=DUPLEX_CHOICES, width=18,
                                       state="readonly")
        self.cmb_duplex.pack(side="left", padx=(2, 12))
        self.cmb_duplex.set(DUPLEX_CHOICES[0])
        ttk.Label(line2b, style="Hint.TLabel",
                  text="按长边翻转；旋转顺序按实际打印队列计算").pack(side="left")

        # 第 3 行：打印机 + 后端
        line3 = ttk.Frame(frm_glob)
        line3.pack(fill="x", pady=(6, 0))
        ttk.Label(line3, text="打印机：").pack(side="left")
        self.cmb_printer = ttk.Combobox(line3, width=38, state="readonly")
        self.cmb_printer.pack(side="left", padx=(2, 6))
        ttk.Button(line3, text="刷新", width=6, command=self.refresh_printers).pack(side="left", padx=(0, 16))
        ttk.Label(line3, text="打印后端：").pack(side="left")
        self.cmb_backend = ttk.Combobox(line3, values=BACKEND_CHOICES, width=18,
                                        state="readonly")
        self.cmb_backend.pack(side="left", padx=(2, 6))
        self.cmb_backend.current(0)

        # 第 4 行：说明与警告
        line4 = ttk.Frame(frm_glob)
        line4.pack(fill="x", pady=(6, 0))
        ttk.Label(line4, style="Hint.TLabel",
                text="· 页码预览按文件参数显示；Office 文件需本机 Office 或 LibreOffice").pack(side="left")
        self.v_queue_by_pages = tk.BooleanVar(value=False)
        ttk.Checkbutton(line4, text="按打印页数排队（少页优先，同页数连续）",
                variable=self.v_queue_by_pages).pack(side="right")
        line5 = ttk.Frame(frm_glob)
        line5.pack(fill="x", pady=(2, 0))
        ttk.Label(line5, style="Warn.TLabel",
                  text="· 份数靠页面展开实现，份数×页数过大时临时 PDF 体积会显著膨胀，"
                       "建议单次不超过 %d 份" % COPIES_WARN).pack(side="left")

        # ---- 执行区 ----
        frm_run = ttk.Frame(root)
        frm_run.pack(fill="x", pady=(8, 0))
        self.v_preview = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_run, text="只生成不打印（预览模式，保留 PDF 供确认）",
                        variable=self.v_preview).pack(side="left")
        self.v_keep = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_run, text="保留临时 PDF", variable=self.v_keep).pack(side="left", padx=(14, 0))

        self.menu_state = "已注册" if menu_registered() else "未注册"
        ttk.Button(frm_run, text="注册右键菜单（%s）" % self.menu_state,
                   command=self.toggle_menu).pack(side="right", padx=(6, 0))
        self.btn_run = ttk.Button(frm_run, text="开始打印", command=self.start_job)
        self.btn_run.pack(side="right")

        # ---- 日志 ----
        frm_log = ttk.LabelFrame(root, text=" 进度日志 ", padding=4)
        frm_log.pack(fill="both", expand=True, pady=(8, 0))
        self.txt_log = tk.Text(frm_log, height=9, wrap="none",
                               font=("Consolas", 9), state="disabled",
                               background="#fbfbfb")
        lsb = ttk.Scrollbar(frm_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=lsb.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        self.txt_log.tag_configure("ok", foreground="#1c6b2f")
        self.txt_log.tag_configure("err", foreground="#a01818")
        self.txt_log.tag_configure("dim", foreground="#777777")

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", pady=(6, 0))

        self.log("就绪。%s" % ("SumatraPDF：%s" % find_sumatra() if find_sumatra()
                              else "未检测到 SumatraPDF，自动档将使用 Windows 原生打印"))
        self._apply_override_state()

    # ---------------- 状态持久化 ----------------
    def _load_state(self):
        c = self.cfg
        try:
            self.v_alternate_180.set(bool(c.get("alternate_180", c.get("rotate_180", False))))
            # 预览模式刻意不记忆：避免用户上次勾过"只生成不打印"，
            # 下次点"开始打印"时静默不出纸。每次启动都回到"会打印"。
            self.v_preview.set(False)
            self.v_keep.set(bool(c.get("keep_temp", False)))
            duplex = c.get("duplex", DUPLEX_CHOICES[0])
            self.cmb_duplex.set(duplex if duplex in DUPLEX_CHOICES else DUPLEX_CHOICES[0])
            self.spn_g_copies.set(int(c.get("global_copies", 1)))
            self.ent_g_range.delete(0, "end")
            global_range = str(c.get("global_range", "all"))
            self.ent_g_range.insert(0, global_range)
            range_presets = {"": "全部", "all": "全部", "1": "第1页",
                             "last": "最后一页"}
            self.cmb_global_range_preset.set(
                range_presets.get(global_range.strip().lower(), "自定义"))
            oe = c.get("global_odd_even", "全部")
            self.cmb_g_oddeven.set(oe if oe in ODD_EVEN_CHOICES else "全部")
            self.v_override_range.set(bool(c.get("ov_range", False)))
            self.v_override_copies.set(bool(c.get("ov_copies", False)))
            self.v_override_oddeven.set(bool(c.get("ov_oddeven", False)))
            self.v_alternate_files_180.set(bool(c.get("alternate_files_180", False)))
            self.v_queue_by_pages.set(bool(c.get("queue_by_pages", False)))
            be = int(c.get("backend", 0))
            self.cmb_backend.current(be if 0 <= be < len(BACKEND_CHOICES) else 0)
        except Exception:
            pass
        self.refresh_printers(select=c.get("printer", ""))

    def _save_state(self):
        save_config({
            "alternate_180": self.v_alternate_180.get(),
            "alternate_files_180": self.v_alternate_files_180.get(),
            "queue_by_pages": self.v_queue_by_pages.get(),
            "duplex": self.cmb_duplex.get(),
            "keep_temp": self.v_keep.get(),
            "global_copies": self._spin_int(self.spn_g_copies, 1),
            "global_range": self.ent_g_range.get().strip(),
            "global_odd_even": self.cmb_g_oddeven.get(),
            "ov_range": self.v_override_range.get(),
            "ov_copies": self.v_override_copies.get(),
            "ov_oddeven": self.v_override_oddeven.get(),
            "backend": self.cmb_backend.current(),
            "printer": self.cmb_printer.get(),
        })

    @staticmethod
    def _spin_int(widget, default=1):
        try:
            return max(1, min(99, int(str(widget.get()).strip())))
        except Exception:
            return default

    # ---------------- 打印机 ----------------
    def refresh_printers(self, select: str = ""):
        keep = select or self.cmb_printer.get()
        names = list_printers()
        self.cmb_printer["values"] = names
        if keep and keep in names:
            self.cmb_printer.set(keep)
        elif names:
            dflt = default_printer_name()
            self.cmb_printer.set(dflt if dflt in names else names[0])
        else:
            self.cmb_printer.set("")
            self.log("未枚举到任何打印机，请检查打印机连接与驱动安装。")

    # ---------------- 表格 ----------------
    def _apply_override_state(self):
        dis = self.busy
        for var, widget in ((self.v_override_range, self.ent_g_range),
                            (self.v_override_copies, self.spn_g_copies),
                            (self.v_override_oddeven, self.cmb_g_oddeven)):
            try:
                widget.config(state=("normal" if not dis else "disabled"))
            except Exception:
                pass
        self._refresh_tree()

    def _effective(self, row: Row):
        """返回该行最终生效的参数（全局覆盖优先）。"""
        pr = (self.ent_g_range.get().strip()
              if self.v_override_range.get() else str(row.page_range or "").strip())
        cp = (self._spin_int(self.spn_g_copies, 1)
              if self.v_override_copies.get() else int(row.copies or 1))
        oe = (self.cmb_g_oddeven.get()
              if self.v_override_oddeven.get() else (row.odd_even or "全部"))
        return pr, max(1, min(99, cp)), (oe if oe in ODD_EVEN_CHOICES else "全部")

    def _apply_global_range_preset(self, _event=None):
        preset = self.cmb_global_range_preset.get()
        values = {"全部": "all", "第1页": "1", "最后一页": "last"}
        if preset in values:
            self.ent_g_range.delete(0, "end")
            self.ent_g_range.insert(0, values[preset])
            self._refresh_tree()
        elif preset == "自定义":
            self.ent_g_range.focus_set()
            self.ent_g_range.selection_range(0, "end")

    def add_paths(self, paths):
        files = collect_files(paths)
        if not files:
            messagebox.showinfo(APP_NAME, "未找到支持的文件格式。")
            return
        existed = {os.path.normcase(r.path) for r in self.rows}
        added_rows = []
        for p in files:
            if os.path.normcase(p) in existed:
                continue
            row = Row(p)
            self.rows.append(row)
            added_rows.append(row)
        self._refresh_tree()
        self.log("已添加 %d 个文件%s" %
                 (len(added_rows), "" if added_rows else "（可能全部重复）"))
        if added_rows:
            threading.Thread(target=self._load_page_counts, args=(added_rows,),
                             daemon=True).start()

    def _load_page_counts(self, rows):
        for row in rows:
            try:
                row.page_count = count_file_pages(row.path)
            except Exception as exc:
                row.page_count_error = str(exc)
            try:
                self.after(0, self._refresh_tree)
            except tk.TclError:
                return

    def pick_files(self):
        patterns = ["*" + ext for ext in sorted(SUPPORTED_EXTENSIONS)]
        paths = filedialog.askopenfilenames(
            title="选择待打印文件",
            filetypes=[("常见可打印文件", patterns), ("所有文件", "*.*")])
        if paths:
            self.add_paths(list(paths))

    def pick_folder(self):
        path = filedialog.askdirectory(title="选择文件夹（递归收集支持的文件）")
        if path:
            self.add_paths([path])

    def preview_selected(self):
        if not PREVIEW_OK:
            messagebox.showerror(APP_NAME, "页面预览组件不可用：%s" % PREVIEW_ERR)
            return
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo(APP_NAME, "请先选择一个文件。")
            return
        idx = int(selection[0][1:])
        if idx >= len(self.rows):
            return
        row = self.rows[idx]
        page_range, _copies, odd_even = self._effective(row)
        temp_dir = tempfile.mkdtemp(prefix="pdfpreview_")
        pdf_path = os.path.join(temp_dir, "preview.pdf")
        try:
            source_pdf = convert_to_pdf(row.path, pdf_path)
            reader = PdfReader(source_pdf)
            total = len(reader.pages)
            pages = filter_odd_even(parse_page_range(page_range, total), odd_even)
            if not pages:
                raise PageRangeError("当前页码范围/单双页设置没有可预览页面")
            document = pdfium.PdfDocument(source_pdf)
        except Exception as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            messagebox.showerror(APP_NAME, "无法预览文件：%s" % exc)
            return

        window = tk.Toplevel(self)
        window.title("页面范围预览 - %s" % os.path.basename(row.path))
        window.geometry("860x760")
        window.minsize(620, 520)
        page_label = ttk.Label(window, anchor="center")
        page_label.pack(fill="x", padx=8, pady=8)
        image_label = ttk.Label(window, anchor="center")
        image_label.pack(fill="both", expand=True, padx=8, pady=4)
        controls = ttk.Frame(window, padding=8)
        controls.pack(fill="x")
        previous = ttk.Button(controls, text="上一页")
        previous.pack(side="left")
        following = ttk.Button(controls, text="下一页")
        following.pack(side="right")
        current = [0]

        def show_page(position):
            current[0] = max(0, min(len(pages) - 1, position))
            page_index = pages[current[0]]
            page = document[page_index]
            width, height = page.get_size()
            scale = min(1.5, 780.0 / max(width, height))
            image = page.render(scale=scale).to_pil()
            image.thumbnail((800, 620), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            image_label.configure(image=photo)
            image_label.image = photo
            suffixes = []
            if self.v_alternate_180.get():
                suffixes.append("同文件偶数份旋转180°")
            if self.v_alternate_files_180.get():
                suffixes.append("队列偶数文件旋转180°")
            suffix = "；" + "、".join(suffixes) if suffixes else ""
            page_label.configure(text="预览页 %d / %d（源文件第 %d / %d 页）%s" %
                                      (current[0] + 1, len(pages), page_index + 1,
                                       total, suffix))
            previous.configure(state=("normal" if current[0] else "disabled"))
            following.configure(state=("normal" if current[0] < len(pages) - 1
                                       else "disabled"))

        previous.configure(command=lambda: show_page(current[0] - 1))
        following.configure(command=lambda: show_page(current[0] + 1))
        window.bind("<Left>", lambda _event: show_page(current[0] - 1))
        window.bind("<Right>", lambda _event: show_page(current[0] + 1))

        def close_preview():
            try:
                document.close()
            except Exception:
                pass
            shutil.rmtree(temp_dir, ignore_errors=True)
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_preview)
        show_page(0)
        window.transient(self)
        window.grab_set()

    def remove_selected(self):
        sel = set(self.tree.selection())
        if not sel:
            return
        keep = []
        for idx, row in enumerate(self.rows):
            if self._iid(idx) not in sel:
                keep.append(row)
        self.rows = keep
        self._close_editor()
        self._refresh_tree()

    def clear_rows(self):
        if not self.rows:
            return
        if messagebox.askyesno(APP_NAME, "确认清空全部文件？"):
            self.rows = []
            self._close_editor()
            self._refresh_tree()

    @staticmethod
    def _iid(idx):
        return "r%d" % idx

    def _refresh_tree(self):
        self._close_editor()
        self.tree.delete(*self.tree.get_children())
        ov_r = self.v_override_range.get()
        ov_c = self.v_override_copies.get()
        ov_o = self.v_override_oddeven.get()
        for idx, row in enumerate(self.rows):
            pr, cp, oe = self._effective(row)
            selected_range = (pr or "全部") if not ov_r else "%s（全局）" % (pr or "全部")
            if row.page_count:
                selected_range += " (1~%d)" % row.page_count
            else:
                selected_range += " (%s)" % (
                    "页数未知" if row.page_count_error else "读取页数…")
            status = row.status
            if row.detail:
                status = "%s：%s" % (status, row.detail)
            tags = []
            if row.status.startswith("成功"):
                tags.append("ok")
            elif row.status.startswith("失败"):
                tags.append("fail")
            if ov_r and ov_c and ov_o:
                tags.append("disabled")
            self.tree.insert("", "end", iid=self._iid(idx), tags=tags, values=(
                os.path.basename(row.path),
                selected_range,
                ("%d（全局）" % cp) if ov_c else cp,
                ("%s（全局）" % oe) if ov_o else oe,
                status,
            ))
        total_copies = 0
        for row in self.rows:
            _pr, cp, _oe = self._effective(row)
            total_copies += cp
        self.lbl_count.config(
            text="共 %d 个文件；合计份数 %d%s" % (
                len(self.rows), total_copies,
                "（超过 %d 份，临时 PDF 体积可能很大）" % COPIES_WARN
                if total_copies > COPIES_WARN else ""))

    def _on_tree_double(self, event):
        if self.busy:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or not col:
            return
        try:
            cidx = int(col.replace("#", "")) - 1
        except Exception:
            return
        if cidx not in (1, 2, 3):
            return
        ridx = int(item[1:])
        if ridx >= len(self.rows):
            return
        # 被全局覆盖的列不允许编辑
        if (cidx == 1 and self.v_override_range.get()) or \
           (cidx == 2 and self.v_override_copies.get()) or \
           (cidx == 3 and self.v_override_oddeven.get()):
            self.log("该列已被全局参数覆盖，取消勾选对应的“覆盖”后才能逐文件修改。")
            return
        self._open_editor(ridx, cidx)

    def _open_editor(self, ridx, cidx):
        self._close_editor()
        col = self.COLS[cidx]
        bbox = self.tree.bbox(self._iid(ridx), col)
        if not bbox:
            return
        x, y, w, h = bbox
        row = self.rows[ridx]
        if cidx == 1:
            presets = page_range_presets(row.page_count or 0)
            current = next((label for label, value in presets.items()
                            if value == str(row.page_range or "").strip()),
                           str(row.page_range or "全部"))
            var = tk.StringVar(value=current)
            ed = ttk.Combobox(self.tree, values=list(presets),
                              textvariable=var, state="normal")
        elif cidx == 2:
            var = tk.IntVar(value=int(row.copies or 1))
            ed = ttk.Spinbox(self.tree, from_=1, to=99, textvariable=var, width=5)
        else:
            var = tk.StringVar(value=row.odd_even or "全部")
            ed = ttk.Combobox(self.tree, values=ODD_EVEN_CHOICES,
                              textvariable=var, state="readonly", width=8)
        ed.place(x=x, y=y, width=w, height=h)
        ed.focus_set()
        try:
            ed.selection_range(0, "end")
        except Exception:
            pass

        def commit(_evt=None):
            try:
                if cidx == 1:
                    selected = var.get().strip()
                    row.page_range = presets.get(selected, selected)
                elif cidx == 2:
                    row.copies = max(1, min(99, int(str(var.get()).strip() or 1)))
                elif cidx == 3:
                    v = var.get()
                    row.odd_even = v if v in ODD_EVEN_CHOICES else "全部"
                row.status, row.detail = "待处理", ""
            except Exception:
                pass
            self._close_editor()
            self._refresh_tree()

        ed.bind("<Return>", commit)
        if cidx == 1:
            def commit_after_focus_out(_evt=None):
                def finish():
                    if not ed.winfo_exists():
                        return
                    try:
                        popdown = ed.tk.call("ttk::combobox::PopdownWindow", ed)
                        if int(ed.tk.call("winfo", "ismapped", popdown)):
                            return
                    except tk.TclError:
                        pass
                    commit()
                self.after_idle(finish)
            ed.bind("<FocusOut>", commit_after_focus_out)
        else:
            ed.bind("<FocusOut>", commit)
        if cidx == 3:
            ed.bind("<<ComboboxSelected>>", commit)
        elif cidx == 1:
            ed.bind("<<ComboboxSelected>>", commit)
        self.editor = (ed, commit)

    def _close_editor(self):
        if self.editor:
            ed, _commit = self.editor
            try:
                ed.destroy()
            except Exception:
                pass
            self.editor = None

    # ---------------- 日志 ----------------
    def log(self, msg, level=""):
        self.log_q.put((msg, level))

    def _drain_log(self):
        try:
            while True:
                msg, level = self.log_q.get_nowait()
                stamp = datetime.now().strftime("%H:%M:%S")
                self.txt_log.config(state="normal")
                tag = None
                if level == "ok" or msg.strip().startswith("成功"):
                    tag = "ok"
                elif level == "err" or "失败" in msg or "错误" in msg:
                    tag = "err"
                elif msg.startswith("    "):
                    tag = "dim"
                self.txt_log.insert("end", "[%s] %s\n" % (stamp, msg), tag)
                self.txt_log.see("end")
                self.txt_log.config(state="disabled")
        except queue.Empty:
            pass
        self.after(90, self._drain_log)

    def _drain_instance_requests(self):
        try:
            while True:
                paths = INSTANCE_REQUESTS.get_nowait()
                if paths:
                    self.add_paths(paths)
                if self.state() == "iconic":
                    self.deiconify()
                self.lift()
                self.focus_force()
        except queue.Empty:
            pass
        self.after(100, self._drain_instance_requests)

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for w in (self.btn_run,):
            w.config(state=state)
        self.btn_run.config(text="打印中…" if busy else "开始打印")
        self._apply_override_state()

    # ---------------- 任务 ----------------
    def start_job(self):
        if self.busy:
            return
        if not PYPDF_OK:
            messagebox.showerror(APP_NAME, "缺少 pypdf 依赖，无法生成 PDF。")
            return
        if not self.rows:
            messagebox.showinfo(APP_NAME, "请先添加文件或文件夹。")
            return
        backend = self.cmb_backend.current()
        preview = self.v_preview.get()
        printer = self.cmb_printer.get().strip()
        if not preview:
            if not printer:
                messagebox.showwarning(APP_NAME, "请先选择打印机。")
                return
            if backend == BACKEND_SUMATRA and not find_sumatra():
                messagebox.showerror(
                    APP_NAME, "已选择 SumatraPDF 后端，但未找到 SumatraPDF.exe。\n"
                              "请安装 SumatraPDF，或将后端改为“自动”/“Windows 原生打印”。")
                return
        # 体积膨胀提示
        warn = []
        for row in self.rows:
            _pr, cp, _oe = self._effective(row)
            if cp > COPIES_WARN:
                warn.append("%s（%d 份）" % (os.path.basename(row.path), cp))
        if warn and not messagebox.askyesno(
                APP_NAME,
                "以下文件份数超过 %d 份，生成的临时 PDF 体积可能非常大：\n%s\n\n仍要继续吗？"
                % (COPIES_WARN, "\n".join(warn[:8]))):
            return

        self._save_state()
        self._set_busy(True)
        self.progress.config(value=0, maximum=len(self.rows))
        for row in self.rows:
            row.status, row.detail = "待处理", ""
        self._refresh_tree()
        self._set_busy(True)

        args = {
            "rows": self.rows,
            "ov_range": self.v_override_range.get(),
            "ov_copies": self.v_override_copies.get(),
            "ov_oddeven": self.v_override_oddeven.get(),
            "g_range": self.ent_g_range.get().strip(),
            "g_copies": self._spin_int(self.spn_g_copies, 1),
            "g_oddeven": self.cmb_g_oddeven.get(),
            "alternate_180": self.v_alternate_180.get(),
            "alternate_files_180": self.v_alternate_files_180.get(),
            "queue_by_pages": self.v_queue_by_pages.get(),
            "duplex": self.cmb_duplex.get(),
            "printer": printer,
            "backend": backend,
            "preview": preview,
            "keep": self.v_keep.get(),
        }
        self.worker = threading.Thread(target=self._job_main, args=(args,), daemon=True)
        self.worker.start()

    def _queue_page_total(self, row, args):
        page_count = row.page_count
        if page_count is None:
            if row.page_count_error:
                return None
            try:
                page_count = count_file_pages(row.path)
                row.page_count = page_count
            except Exception as exc:
                row.page_count_error = str(exc)
                return None
        page_range = (args["g_range"] if args["ov_range"]
                      else str(row.page_range or "").strip())
        odd_even = (args["g_oddeven"] if args["ov_oddeven"]
                    else (row.odd_even or "全部"))
        copies = (args["g_copies"] if args["ov_copies"]
                  else max(1, min(99, int(row.copies or 1))))
        return output_page_total(page_count, page_range, odd_even, copies)

    def _ordered_rows(self, rows, args):
        if not args["queue_by_pages"]:
            return list(rows)
        counts = {}
        for row in rows:
            page_total = self._queue_page_total(row, args)
            counts[id(row)] = float("inf") if page_total is None else page_total
        ordered = order_rows_by_page_total(rows, counts)
        for row in ordered:
            page_total = counts[id(row)]
            label = "页数未知" if page_total == float("inf") else "%d 页" % page_total
            self.log("排队：%s（%s）" % (os.path.basename(row.path), label))
        return ordered

    def _process_one(self, idx, row, a, temp_dir, file_rotate_180=False):
        """处理单个文件：生成变换后的 PDF，按需打印。

        返回 (outcome, out_path)，outcome 为 "ok" / "preview" / "fail"。
        本方法内部的失败以返回值表达，不向上抛异常。
        """
        name = os.path.basename(row.path)
        pr = a["g_range"] if a["ov_range"] else str(row.page_range or "").strip()
        cp = a["g_copies"] if a["ov_copies"] else max(1, min(99, int(row.copies or 1)))
        oe = a["g_oddeven"] if a["ov_oddeven"] else (row.odd_even or "全部")

        out_name = "%03d_%s.pdf" % (idx, safe_name(os.path.splitext(name)[0]))
        out_path = os.path.join(temp_dir, out_name)

        working_path = row.path
        try:
            if os.path.splitext(row.path)[1].lower() != ".pdf":
                self.log("    正在转换为 PDF…")
                converted_path = os.path.join(temp_dir, "source_%03d.pdf" % idx)
                working_path = convert_to_pdf(row.path, converted_path)
            gen = build_pdf(
                working_path, out_path,
                page_range=pr, odd_even=oe, copies=cp,
                alternate_180=a["alternate_180"],
                file_rotate_180=file_rotate_180, duplex=a["duplex"])
        except PageRangeError as exc:
            row.status, row.detail = "失败", str(exc)
            self.log("    失败：%s" % exc, "err")
            return "fail", ""
        except Exception as exc:
            detail = "处理文件失败：%s" % exc
            row.status, row.detail = "失败", detail
            self.log("    %s" % detail, "err")
            self.log("    %s" % traceback.format_exc().strip().splitlines()[-1], "dim")
            return "fail", ""

        rotations = []
        if a["alternate_180"] and cp > 1:
            rotations.append("同文件偶数份180°")
        if file_rotate_180:
            rotations.append("偶数文件180°")
        self.log("    源 %d 页 → 取页 %d 页 → 输出 %d 页（%d 份，%s，%s，%s）"
                 % (gen.src_pages, gen.picked_pages, gen.out_pages, cp, oe,
                    "、".join(rotations) if rotations else "不旋转", a["duplex"]))

        if a["preview"]:
            row.status, row.detail = "已生成", out_path
            self.log("    已生成（未打印）：%s" % out_path, "ok")
            return "preview", out_path

        ok, msg, used = print_one(
            out_path, a["printer"], a["backend"], duplex=a["duplex"], log=self.log)
        if ok:
            row.status, row.detail = "成功", "已提交（%s）" % used
            self.log("    成功：已提交到打印队列（%s）" % used, "ok")
            return "ok", out_path
        row.status, row.detail = "失败", msg
        self.log("    失败：%s" % msg, "err")
        return "fail", ""

    def _job_main(self, a):
        ok_n = fail_n = skip_n = 0
        temp_dir = ""
        try:
            temp_dir = tempfile.mkdtemp(prefix="pdfprint_")
            self.temp_dir = temp_dir
            self.log("临时目录：%s" % temp_dir, "dim")
            if a["preview"]:
                self.log("预览模式：只生成 PDF，不打印，文件将保留在该目录。")
            if not a["preview"]:
                self.log("打印机：%s ｜ 后端：%s" % (a["printer"], BACKEND_CHOICES[a["backend"]]))

            total = len(a["rows"])
            printed_paths = []
            ordered_rows = self._ordered_rows(a["rows"], a)
            for idx, row in enumerate(ordered_rows, start=1):
                name = os.path.basename(row.path)
                self.log("—— [%d/%d] %s" % (idx, total, name))
                # 单个文件的任何异常都只影响该文件，后续文件继续处理
                try:
                    outcome, out_path = self._process_one(
                        idx, row, a, temp_dir,
                        file_rotate_180=(a["alternate_files_180"] and idx % 2 == 0))
                    if outcome == "ok":
                        ok_n += 1
                        if out_path:
                            printed_paths.append(out_path)
                    elif outcome == "preview":
                        skip_n += 1
                    else:
                        fail_n += 1
                except Exception as exc:
                    fail_n += 1
                    detail = "处理异常：%s" % exc
                    row.status, row.detail = "失败", detail
                    self.log("    %s" % detail, "err")
                    self.log("    %s" % traceback.format_exc().strip().splitlines()[-1], "dim")
                finally:
                    self._mark(self.rows.index(row), row)
                    self._progress(idx)

            # 打印全部提交后，给后台缓冲留出时间再清理临时文件
            if not a["preview"] and printed_paths:
                self.log("等待打印缓冲写入完成…", "dim")
                time.sleep(6)

            if a["preview"] or a["keep"]:
                self.log("已保留临时 PDF：%s" % temp_dir, "dim")
            else:
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    self.log("已清理临时目录。", "dim")
                    self.temp_dir = ""
                except Exception as exc:
                    self.log("临时目录清理失败（可手动删除）：%s" % exc, "dim")

            self.log("=" * 60)
            self.log("任务结束：成功 %d ｜ 失败 %d ｜ 仅生成 %d" % (ok_n, fail_n, skip_n),
                     "ok" if fail_n == 0 else "err")
            self._finish(a, ok_n, fail_n, skip_n, temp_dir)
        except Exception:
            self.log("任务异常终止：\n%s" % traceback.format_exc(), "err")
            self._finish(a, ok_n, fail_n, skip_n, temp_dir)

    def _mark(self, idx, row):
        self.after(0, lambda: self._update_row(idx, row))

    def _update_row(self, idx, row):
        if idx < 0 or idx >= len(self.rows):
            return
        iid = self._iid(idx)
        if not self.tree.exists(iid):
            return
        status = row.status + (("：" + row.detail) if row.detail else "")
        self.tree.set(iid, "status", status)
        self.tree.item(iid, tags=("ok" if row.status in ("成功", "已生成")
                                  else "fail" if row.status == "失败" else ""))

    def _progress(self, done):
        self.after(0, lambda: self.progress.config(value=done))

    def _finish(self, a, ok_n, fail_n, skip_n, temp_dir):
        def done():
            self._set_busy(False)
            self._refresh_tree()
            if a["preview"]:
                msg = "已生成 %d 个 PDF（未打印）。\n目录：%s" % (skip_n, temp_dir)
                messagebox.showinfo(APP_NAME, msg)
            elif fail_n == 0:
                messagebox.showinfo(APP_NAME, "打印完成：成功 %d 个文件。" % ok_n)
            else:
                messagebox.showwarning(
                    APP_NAME, "打印结束：成功 %d ｜ 失败 %d。\n详细原因见日志。" % (ok_n, fail_n))
        self.after(0, done)

    # ---------------- 其他 ----------------
    def open_temp_dir(self):
        d = self.temp_dir
        if not d or not os.path.isdir(d):
            d = os.path.join(tempfile.gettempdir(), "")
            found = [p for p in os.listdir(tempfile.gettempdir())
                     if p.startswith("pdfprint_")]
            if found:
                d = os.path.join(tempfile.gettempdir(), sorted(found)[-1])
        if os.path.isdir(d):
            os.startfile(d)  # noqa: S606 - 本地目录，用户主动触发
        else:
            messagebox.showinfo(APP_NAME, "暂无临时目录（可能已自动清理）。")

    def show_selftest(self):
        """在窗口中显示环境自检结果，可另存为报告文件。"""
        try:
            text = run_selftest()
        except Exception as exc:
            text = "自检执行失败：%s\n%s" % (exc, traceback.format_exc())
        for line in text.splitlines():
            self.log(line)

        win = tk.Toplevel(self)
        win.title("环境自检结果")
        win.geometry("720x520")
        txt = tk.Text(win, wrap="word", font=("Consolas", 10), padx=10, pady=10)
        txt.pack(fill="both", expand=True)
        txt.insert("end", text)
        txt.config(state="disabled")

        bar = ttk.Frame(win, padding=(0, 0, 0, 8))
        bar.pack(fill="x")

        def save_report():
            path = filedialog.asksaveasfilename(
                title="保存自检报告",
                defaultextension=".txt",
                initialfile="PDF批量打印自检报告_%s.txt"
                            % datetime.now().strftime("%Y%m%d_%H%M%S"),
                filetypes=[("文本文件", "*.txt")])
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(text + "\n")
                messagebox.showinfo(APP_NAME, "自检报告已保存：\n%s" % path)
            except Exception as exc:
                messagebox.showerror(APP_NAME, "保存失败：%s" % exc)

        ttk.Button(bar, text="另存报告…", command=save_report).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side="right", padx=(0, 8))

    def show_help(self):
        win = tk.Toplevel(self)
        win.title("使用说明")
        win.geometry("760x560")
        txt = tk.Text(win, wrap="word", font=("Microsoft YaHei UI", 10), padx=10, pady=10)
        txt.pack(fill="both", expand=True)
        txt.insert("end", HELP_TEXT)
        txt.config(state="disabled")
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=(0, 8))

    def toggle_menu(self):
        if menu_registered():
            if not messagebox.askyesno(APP_NAME, "确认卸载右键菜单？"):
                return
            msgs, errors = unregister_menu()
            title = "右键菜单已卸载"
        else:
            if not is_frozen():
                if not messagebox.askyesno(
                        APP_NAME,
                        "当前以脚本方式运行，注册后右键菜单将调用 Python 解释器。\n"
                        "建议先打包成 exe 再注册。\n\n仍要继续吗？"):
                    return
            msgs, errors = register_menu()
            title = "右键菜单已注册（当前用户，免管理员）"
        text = "\n".join(msgs + errors)
        self.menu_state = "已注册" if menu_registered() else "未注册"
        self.log("%s\n%s" % (title, text))
        messagebox.showinfo(APP_NAME, title + "\n\n" + text +
                            "\n\n如资源管理器未立即刷新，注销并重新登录或重启资源管理器即可。")

    def _on_close(self):
        self._save_state()
        if self.busy:
            if not messagebox.askyesno(APP_NAME, "任务仍在执行，确认退出？"):
                return
        try:
            self.destroy()
        except Exception:
            pass


# ===========================================================================
# 入口
# ===========================================================================
def _has_console() -> bool:
    try:
        return bool(ctypes.windll.kernel32.GetConsoleWindow())
    except Exception:
        return False


def report_cli(text: str, title: str = APP_NAME, out_file: str = "") -> str:
    """命令行分支的结果反馈。

    exe 以无控制台方式打包，直接 print 用户看不到；因此：
      · 指定了 out_file → 写入该文件（便于留档与脚本判断）；
      · 从 cmd / bat 调用（已有控制台）→ 打印到控制台，不弹窗，避免阻塞脚本；
      · 双击 exe 调用（无控制台）→ 弹出消息框显示结果。
    返回实际写入的文件绝对路径（未写文件时返回空串）。
    """
    written = ""
    if out_file:
        try:
            with open(out_file, "w", encoding="utf-8") as fh:
                fh.write(text + "\n")
            written = os.path.abspath(out_file)
        except Exception as exc:
            text = "写入报告文件失败：%s\n\n%s" % (exc, text)

    printed = False
    try:
        if sys.stdout is not None:
            print(text)
            sys.stdout.flush()
            printed = True
    except Exception:
        printed = False
    # 已写入文件时不再弹窗，避免脚本被模态框阻塞
    if not written and (not printed or not _has_console()):
        try:
            ctypes.windll.user32.MessageBoxW(None, text, title, 0x40)
        except Exception:
            pass
    return written


def run_selftest() -> str:
    """环境自检：验证 pypdf、页面变换算法、打印机枚举与打印后端可用性。

    仅在临时目录生成并清理测试文件，不改动系统设置。
    """
    lines = ["%s v%s —— 环境自检" % (APP_NAME, APP_VERSION), "=" * 46]
    lines.append("[信息] 作者：%s" % APP_AUTHOR)
    ok_all = True

    # 1) pypdf
    if PYPDF_OK:
        try:
            import pypdf
            lines.append("[正常] pypdf 已可用，版本 %s"
                         % getattr(pypdf, "__version__", "未知"))
        except Exception as exc:
            lines.append("[警告] pypdf 已导入但版本查询失败：%s" % exc)
    else:
        ok_all = False
        lines.append("[缺失] pypdf 不可用：%s" % PYPDF_ERR)
        lines.append("       → 无法生成 PDF，程序不可用。")

    # 2) 打包形态
    lines.append("[信息] 运行形态：%s" % ("单文件 exe（已打包）" if is_frozen()
                                          else "Python 脚本（未打包）"))
    lines.append("[信息] 可执行文件：%s" % sys.executable)

    # 3) 核心算法实测（生成 → 校验旋转）
    if PYPDF_OK:
        tmp = ""
        try:
            tmp = tempfile.mkdtemp(prefix="pdfprint_selftest_")
            src = os.path.join(tmp, "selftest_src.pdf")
            out = os.path.join(tmp, "selftest_out.pdf")
            w = PdfWriter()
            for _ in range(3):
                w.add_blank_page(width=595, height=842)
            with open(src, "wb") as fh:
                w.write(fh)
            gen = build_pdf(src, out, page_range="1-3", odd_even="全部",
                            copies=4, alternate_180=True,
                            duplex="双面（长边翻转）")
            rots = [p.rotation for p in PdfReader(out).pages]
            expect = [0] * 3 + [180] * 3 + [0] * 3 + [180] * 3
            if gen.out_pages == 12 and rots == expect:
                lines.append("[正常] 多份隔份180°旋转实测通过（4 份 → 12 页）")
            else:
                ok_all = False
                lines.append("[异常] 隔份旋转结果不符：页数 %d，旋转 %s"
                             % (gen.out_pages, rots))
            out_single = os.path.join(tmp, "selftest_single_copy.pdf")
            build_pdf(src, out_single, copies=1, alternate_180=True)
            if [p.rotation for p in PdfReader(out_single).pages] == [0, 0, 0]:
                lines.append("[正常] 单份不旋转实测通过")
            else:
                ok_all = False
                lines.append("[异常] 单份启用隔份旋转时不应旋转")
            out_file_rotation = os.path.join(tmp, "selftest_file_rotation.pdf")
            build_pdf(src, out_file_rotation, file_rotate_180=True)
            if [p.rotation for p in PdfReader(out_file_rotation).pages] == [180] * 3:
                lines.append("[正常] 跨文件偶数位旋转实测通过")
            else:
                ok_all = False
                lines.append("[异常] 跨文件旋转结果不符")
            out_combined_rotation = os.path.join(tmp, "selftest_combined_rotation.pdf")
            build_pdf(src, out_combined_rotation, copies=2, alternate_180=True,
                      file_rotate_180=True)
            if [p.rotation for p in PdfReader(out_combined_rotation).pages] == \
               [180] * 3 + [0] * 3:
                lines.append("[正常] 两种隔份旋转叠加实测通过")
            else:
                ok_all = False
                lines.append("[异常] 两种旋转叠加结果不符")
            prefs = PdfReader(out).trailer["/Root"].get("/ViewerPreferences", {})
            if prefs.get("/Duplex") == "/DuplexFlipLongEdge":
                lines.append("[正常] 双面长边翻转设置已写入 PDF")
            else:
                ok_all = False
                lines.append("[异常] 双面长边翻转设置缺失")
            out2 = os.path.join(tmp, "selftest_simplex.pdf")
            build_pdf(src, out2, duplex="单面")
            prefs2 = PdfReader(out2).trailer["/Root"].get("/ViewerPreferences", {})
            if prefs2.get("/Duplex") == "/Simplex":
                lines.append("[正常] 单面打印设置已写入 PDF")
            else:
                ok_all = False
                lines.append("[异常] 单面打印设置缺失")

            nested = os.path.join(tmp, "nested")
            os.makedirs(nested)
            for filename in ("a.pdf", "b.docx", "c.xlsx", "d.png", "ignore.bin"):
                with open(os.path.join(nested, filename), "wb") as fh:
                    fh.write(b"test")
            if len(collect_files([nested])) == 4:
                lines.append("[正常] 常见格式递归扫描实测通过")
            else:
                ok_all = False
                lines.append("[异常] 常见格式递归扫描结果不符")
            if parse_page_range("1,3,4,7", 100) == [0, 2, 3, 6]:
                lines.append("[正常] 自定义离散页码范围实测通过")
            else:
                ok_all = False
                lines.append("[异常] 自定义离散页码范围解析结果不符")
            if parse_page_range("last", 3) == [2] and \
               parse_page_range("1-last", 3) == [0, 1, 2]:
                lines.append("[正常] 最后一页范围解析实测通过")
            else:
                ok_all = False
                lines.append("[异常] 最后一页范围解析结果不符")
            queue_rows = [object(), object(), object(), object()]
            queue_totals = {id(queue_rows[0]): 4, id(queue_rows[1]): 2,
                            id(queue_rows[2]): 4, id(queue_rows[3]): 1}
            queue_order = order_rows_by_page_total(queue_rows, queue_totals)
            if queue_order == [queue_rows[3], queue_rows[1],
                               queue_rows[0], queue_rows[2]]:
                lines.append("[正常] 少页优先且同页数稳定分组实测通过")
            else:
                ok_all = False
                lines.append("[异常] 打印队列顺序结果不符")
            if output_page_total(5, "last", "全部", 2) == 2 and \
               output_page_total(5, "1-last", "全部", 1) == 5:
                lines.append("[正常] 范围覆盖与排队页数计算实测通过")
            else:
                ok_all = False
                lines.append("[异常] 排队页数计算结果不符")
            out_range = os.path.join(tmp, "selftest_range.pdf")
            ranged = build_pdf(src, out_range, page_range="1,3")
            if ranged.picked_pages == 2 and len(PdfReader(out_range).pages) == 2:
                lines.append("[正常] 自定义范围输出页数实测通过（1,3 → 2 页）")
            else:
                ok_all = False
                lines.append("[异常] 自定义范围输出页数不符")
        except Exception as exc:
            ok_all = False
            lines.append("[异常] 核心算法自检失败：%s" % exc)
            lines.append("       %s" % traceback.format_exc().strip().splitlines()[-1])
        finally:
            if tmp:
                shutil.rmtree(tmp, ignore_errors=True)

    # 4) 打印机
    try:
        printers = list_printers()
        dflt = default_printer_name()
        if printers:
            lines.append("[正常] 枚举到 %d 台打印机%s"
                         % (len(printers), "（win32print）" if WIN32PRINT_OK
                            else "（PowerShell 回退）"))
            for p in printers[:6]:
                lines.append("       %s%s" % ("* " if p == dflt else "  ", p))
            if len(printers) > 6:
                lines.append("       … 其余 %d 台省略" % (len(printers) - 6))
            if dflt:
                lines.append("       （* 为系统默认打印机：%s）" % dflt)
        else:
            ok_all = False
            lines.append("[警告] 未枚举到任何打印机 → 打印会失败，请先安装打印机驱动。")
    except Exception as exc:
        ok_all = False
        lines.append("[异常] 打印机枚举失败：%s" % exc)

    # 5) 打印后端
    sumatra = find_sumatra()
    if NATIVE_PRINT_OK and PREVIEW_OK:
        lines.append("[正常] Windows 原生打印组件可用（PDFium + GDI）")
    else:
        lines.append("[缺失] Windows 原生打印组件不可用：%s %s"
                     % (NATIVE_PRINT_ERR, PREVIEW_ERR))
    if sumatra:
        lines.append("[正常] 已找到 SumatraPDF：%s" % sumatra)
        lines.append("       → 自动档将优先使用 SumatraPDF 命令行打印（推荐）")
    else:
        lines.append("[提示] 未找到 SumatraPDF → 自动档将使用 Windows 原生打印")

    original_find_sumatra = find_sumatra
    original_native_print = _native_print_pdf
    native_calls = []
    try:
        globals()["find_sumatra"] = lambda: ""
        globals()["_native_print_pdf"] = lambda *args, **kwargs: (
            native_calls.append(args[1]) or True, "")
        route_ok, _route_error, used_backend = print_one(
            "selftest.pdf", "自检打印机", BACKEND_AUTO, "单面")
        if route_ok and used_backend == "Windows 原生打印" and native_calls:
            lines.append("[正常] 无 SumatraPDF 时自动走 Windows 原生打印")
        else:
            ok_all = False
            lines.append("[异常] 自动打印后端未路由到 Windows 原生打印")
    except Exception as exc:
        ok_all = False
        lines.append("[异常] 原生打印路由自检失败：%s" % exc)
    finally:
        globals()["find_sumatra"] = original_find_sumatra
        globals()["_native_print_pdf"] = original_native_print

    # 6) 右键菜单
    try:
        lines.append("[信息] 右键菜单：%s" % ("已注册" if menu_registered() else "未注册"))
        lines.append("       注册：\"%s\" --register" % (
            sys.executable if is_frozen() else "文件旋转打印.exe"))
        lines.append("       卸载：--unregister")
    except Exception as exc:
        lines.append("[警告] 右键菜单状态查询失败：%s" % exc)

    lines.append("=" * 46)
    lines.append("自检结论：%s" % ("全部正常，可以投入使用" if ok_all else
                                   "存在上述问题，请先处理后再批量打印"))
    lines.append("配置文件：%s" % config_path())
    return "\n".join(lines)


def main(argv):
    lower = [a.lower() for a in argv]
    if "--selftest" in lower or "/selftest" in lower:
        # 可选：紧跟开关之后的非开关参数作为自检报告的输出文件
        out_file = ""
        try:
            i = (lower.index("--selftest") if "--selftest" in lower
                 else lower.index("/selftest"))
        except ValueError:
            i = -1
        if i >= 0 and i + 1 < len(argv):
            nxt = argv[i + 1]
            if nxt and not nxt.startswith("-") and not nxt.startswith("/"):
                out_file = nxt
        report_cli(run_selftest(), "%s —— 环境自检" % APP_NAME, out_file=out_file)
        return 0
    if "--register" in lower or "/register" in lower:
        msgs, errors = register_menu()
        text = "\n".join(msgs + errors)
        if errors:
            text += "\n\n注意：以上失败项需要检查权限或注册表状态。"
        else:
            text += ("\n\n如资源管理器未立即显示菜单，注销并重新登录，"
                     "或重启资源管理器即可。")
        report_cli(text)
        return 0 if not errors else 1
    if "--unregister" in lower or "/unregister" in lower:
        msgs, errors = unregister_menu()
        report_cli("\n".join(msgs + errors))
        return 0 if not errors else 1
    if "--help" in lower or "-h" in lower:
        report_cli(__doc__)
        return 0

    paths = [a for a in argv if not a.startswith("-")]
    instance_server, forwarded = _start_single_instance(paths)
    if forwarded:
        return 0
    if instance_server is None:
        report_cli("无法建立单实例通信端口，请关闭重复运行的程序后重试。")
        return 1
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    app = App(init_paths=paths)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
