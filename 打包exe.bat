@echo off
cd /d "%~dp0"
chcp 65001 >nul
echo ======================================
echo   文件旋转打印 - PyInstaller 打包
echo ======================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8 或更高版本
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查 pyinstaller 是否安装
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo [提示] PyInstaller 未安装，正在安装...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
)

:: 检查 pypdf 是否安装
python -c "import pypdf" >nul 2>&1
if errorlevel 1 (
    echo [提示] pypdf 未安装，正在安装...
    pip install pypdf
)

:: 检查 win32print 是否安装
python -c "import win32print" >nul 2>&1
if errorlevel 1 (
    echo [提示] pywin32 未安装，正在安装...
    pip install pywin32
)

:: 检查 PDF 页面预览依赖是否安装
python -c "import pypdfium2; from PIL import Image, ImageTk" >nul 2>&1
if errorlevel 1 (
    echo [提示] 页面预览组件未安装，正在安装...
    pip install pypdfium2 Pillow
    if errorlevel 1 (
        echo [错误] 页面预览组件安装失败
        pause
        exit /b 1
    )
)

echo [信息] 开始打包...
echo.

:: 强制将输出写到当前脚本目录下
pyinstaller "%~dp0pdf_batch_printer.spec" --clean --noconfirm --distpath "%~dp0dist" --workpath "%~dp0build"

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！
    pause
    exit /b 1
)

echo.
echo ======================================
echo   打包完成！
echo ======================================
echo.
echo 可执行文件位置：
echo   %~dp0dist\文件旋转打印.exe
echo.
echo 按任意键打开输出目录...
pause >nul
explorer "%~dp0dist"
