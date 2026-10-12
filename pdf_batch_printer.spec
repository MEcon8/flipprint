# -*- mode: python ; coding: utf-8 -*-
"""文件旋转打印 —— PyInstaller 单文件打包配置

用法：
    pyinstaller pdf_batch_printer.spec --clean --noconfirm
或直接用同目录下的「打包exe.bat」。
"""

block_cipher = None

a = Analysis(
    ['pdf_batch_printer.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # pypdf 的部分子模块为动态导入，显式声明避免打包遗漏
        'pypdf',
        'pypdf.generic',
        'pypdf._page',
        'pypdf._reader',
        'pypdf._writer',
        # pypdf 6.x 的加密支持位于以下两个模块（旧版的 pypdf.crypto 已不存在）
        'pypdf._encryption',
        'pypdf._crypt_providers',
        # 页面缩略预览
        'pypdfium2',
        'PIL',
        'PIL.ImageTk',
        'PIL.ImageWin',
        # Office COM 转换
        'pythoncom',
        'win32com.client',
        'win32con',
        'win32gui',
        'win32ui',
        # pywin32 为可选依赖，存在则一并打入（用于枚举打印机）
        'win32print',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除无关的大型库，显著减小体积、加快启动
        'matplotlib', 'numpy', 'pandas', 'scipy', 'PyQt5', 'PyQt6',
        'PySide2', 'PySide6', 'IPython', 'jupyter', 'notebook', 'pytest',
        'setuptools', 'pip', 'tkinter.test', 'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='文件旋转打印',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    # console=False：不弹黑色命令行窗口
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,          # 如需图标，放入 icon.ico 后改为 icon='icon.ico'
    version=None,
)
