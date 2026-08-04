@echo off
chcp 65001 >nul
echo ========================================
echo   料号翻译工具 - 打包为 EXE
echo ========================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
pip install pandas openpyxl pyinstaller -q

echo [2/3] 打包中（约 1-2 分钟）...
pyinstaller --onefile --windowed --icon "icons/icon.ico" --name "料号翻译工具" --add-data "mapping.xlsx;." gui.py

echo [3/3] 完成！
echo.
echo 生成的文件在 dist\ 目录下：
echo   dist\料号翻译工具.exe
echo.
echo 将 mapping.xlsx 放在 exe 同目录下即可运行！
pause
