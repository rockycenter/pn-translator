@echo off
chcp 65001 >nul
title 料号翻译工具 - 一键打包

echo.
echo   ============================================
echo     料号翻译工具 - 全自动打包
echo     首次运行约 3-5 分钟，请耐心等待
echo   ============================================
echo.

cd /d "%~dp0"

:: ── 1. 检查/下载 Python ──────────────────────────────────
set "PYTHON_DIR=%~dp0_python"
set "PYTHON_EXE=%PYTHON_DIR%\python.exe"

if exist "%PYTHON_EXE%" (
    echo [1/4] 使用本地 Python...
) else (
    echo [1/4] 下载 Python 便携版（约 12MB）...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.5/python-3.12.5-embed-amd64.zip' -OutFile '%TEMP%\python-embed.zip'}"
    
    if not exist "%TEMP%\python-embed.zip" (
        echo [错误] 下载失败，请检查网络连接
        pause
        exit /b 1
    )
    
    echo 解压中...
    powershell -Command "Expand-Archive -Path '%TEMP%\python-embed.zip' -DestinationPath '%PYTHON_DIR%' -Force"
    del "%TEMP%\python-embed.zip"
    
    :: 启用 pip
    echo import site>>"%PYTHON_DIR%\python312._pth"
)

:: ── 2. 安装 pip ─────────────────────────────────────────
echo [2/4] 安装 pip...
if not exist "%PYTHON_DIR%\Scripts\pip.exe" (
    "%PYTHON_EXE%" -c "import urllib.request; exec(urllib.request.urlopen('https://bootstrap.pypa.io/get-pip.py').read())" --no-warn-script-location
)

:: ── 3. 安装依赖 ─────────────────────────────────────────
echo [3/4] 安装依赖...
"%PYTHON_DIR%\Scripts\pip.exe" install pandas openpyxl pyinstaller -q --no-warn-script-location

:: ── 4. 打包 exe ─────────────────────────────────────────
echo [4/4] 正在打包 exe（约 1-2 分钟）...
"%PYTHON_DIR%\Scripts\pyinstaller.exe" --onefile --windowed --name "料号翻译工具" --add-data "mapping.xlsx;." gui.py --clean --noconfirm 2>nul

:: ── 完成 ────────────────────────────────────────────────
if exist "dist\料号翻译工具.exe" (
    echo.
    echo   ============================================
    echo     打包成功！
    echo     exe 位置: dist\料号翻译工具.exe
    echo   ============================================
    echo.
    explorer /select,"%~dp0dist\料号翻译工具.exe"
) else (
    echo.
    echo   [错误] 打包失败，请尝试：
    echo   1. 关闭杀毒软件后重试
    echo   2. 手动安装 Python 后运行 build_exe.bat
)
pause
