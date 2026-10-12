@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "APP_EXE=%~dp0文件旋转打印.exe"
if not exist "%APP_EXE%" (
    echo [错误] 未找到同目录下的“文件旋转打印.exe”
    pause
    exit /b 1
)

call :check_windows_print
call :ensure_sumatra
start "" "%APP_EXE%" %*
exit /b 0

:check_windows_print
sc query Spooler | findstr /I "RUNNING" >nul
if not errorlevel 1 goto :check_printers
echo [提示] 正在启动 Windows 打印后台服务...
net start Spooler >nul 2>&1
if errorlevel 1 echo [提示] 无法启动打印后台服务；可能需要管理员权限。

:check_printers
powershell.exe -NoProfile -Command "if (Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Select-Object -First 1) { exit 0 } else { exit 1 }" >nul 2>&1
if errorlevel 1 (
    echo [提示] 尚未检测到已安装的打印机驱动。请按打印机型号安装 Windows 驱动。
)
exit /b 0

:ensure_sumatra
call :find_sumatra
if not errorlevel 1 (
    set "SUMATRA_PDF=%SUMATRA_EXE%"
    exit /b 0
)
where.exe winget.exe >nul 2>&1
if errorlevel 1 (
    echo [提示] 未检测到 SumatraPDF 或 winget，程序仍会启动；自动后端可使用 Windows 原生打印。
    exit /b 0
)
echo [信息] 正在安装 SumatraPDF 打印组件...
winget install --id SumatraPDF.SumatraPDF --exact --silent --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
    echo [提示] SumatraPDF 自动安装失败，程序仍会启动；自动后端可使用 Windows 原生打印。
    exit /b 0
)
call :find_sumatra
if errorlevel 1 (
    echo [提示] 安装后仍未找到 SumatraPDF.exe，程序仍会启动；自动后端可使用 Windows 原生打印。
    exit /b 0
)
set "SUMATRA_PDF=%SUMATRA_EXE%"
exit /b 0

:find_sumatra
set "SUMATRA_EXE="
if defined SUMATRA_PDF if exist "%SUMATRA_PDF%" set "SUMATRA_EXE=%SUMATRA_PDF%"
if defined SUMATRA_EXE exit /b 0
if defined SUMATRAPDF if exist "%SUMATRAPDF%" set "SUMATRA_EXE=%SUMATRAPDF%"
if defined SUMATRA_EXE exit /b 0
if defined LOCALAPPDATA if exist "%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe" set "SUMATRA_EXE=%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"
if defined SUMATRA_EXE exit /b 0
if defined ProgramFiles if exist "%ProgramFiles%\SumatraPDF\SumatraPDF.exe" set "SUMATRA_EXE=%ProgramFiles%\SumatraPDF\SumatraPDF.exe"
if defined SUMATRA_EXE exit /b 0
if exist "%ProgramFiles(x86)%\SumatraPDF\SumatraPDF.exe" set "SUMATRA_EXE=%ProgramFiles(x86)%\SumatraPDF\SumatraPDF.exe"
if defined SUMATRA_EXE exit /b 0
if exist "%~dp0SumatraPDF.exe" set "SUMATRA_EXE=%~dp0SumatraPDF.exe"
if defined SUMATRA_EXE exit /b 0
for /f "delims=" %%P in ('where.exe SumatraPDF.exe 2^>nul') do if not defined SUMATRA_EXE set "SUMATRA_EXE=%%P"
if defined SUMATRA_EXE exit /b 0
exit /b 1
