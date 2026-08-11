@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================
echo   Lsy的Excel-Tool - 一键打包
echo ============================================

echo.
echo [1/3] 检查 Python 与依赖...
where python >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 python，请先安装 Python 3.10 及以上版本
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装 openpyxl...
    python -m pip install openpyxl==3.1.5
)

python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    python -m pip install pyinstaller==6.22.0
)

echo.
echo [2/3] 开始打包（约需 1-3 分钟）...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Lsy的Excel-Tool" ^
    --exclude-module unittest --exclude-module pydoc_data --exclude-module test ^
    --exclude-module numpy --exclude-module PIL --exclude-module lxml ^
    --exclude-module scipy --exclude-module pandas --exclude-module matplotlib --exclude-module sklearn app.py
if errorlevel 1 (
    echo.
    echo 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

echo.
echo [3/3] 打包完成！
echo.
echo 输出文件: %~dp0dist\Lsy的Excel-Tool.exe
echo 把这个 exe 拷贝到其他 Windows 电脑上双击即可运行（无需安装 Python/Excel）
echo.
pause
