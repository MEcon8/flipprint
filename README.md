# 文件旋转打印 v3.3

Windows 批量打印程序，支持 PDF、Office 文档和常见图片，可设置页码范围、单双页、份数和旋转，并提供预览及打印队列排序。

## 文件说明

- `文件旋转打印.exe`：Windows 单文件程序。
- `pdf_batch_printer.py`：Python 源码。
- `启动排队打印程序.bat`：检查 Windows 打印服务及打印机队列；缺少 SumatraPDF 时尝试通过 winget 安装，然后启动程序。
- `pdf_batch_printer.spec`：PyInstaller 配置。
- `打包exe.bat`：从源码构建 exe。

## 多文件右键打印

右键动词使用 Windows Shell 的 `%1` 文件参数替换符，并注册为 `Player` 多选模式。程序收到文件路径后，会将后续启动的路径转发到当前窗口并加入待打印列表。

如此前已注册右键菜单，启动 v3.3 后会自动刷新当前用户的菜单命令。若旧版程序仍在运行，请先关闭旧窗口，再直接启动一次 v3.3。

## 打印组件

原生打印使用 Windows Print Spooler 和目标打印机对应的 Windows 驱动；厂商驱动需按打印机型号安装。SumatraPDF 是可选打印应用，不是打印机驱动。自动后端会优先使用 SumatraPDF，否则使用 Windows 原生打印。

## 从源码构建

Windows 上安装 Python 3.8 或更高版本后，运行 `打包exe.bat`，构建结果位于 `dist\\文件旋转打印.exe`。

## 版本

`3.3`
