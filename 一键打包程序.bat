@echo off
chcp 65001 >nul 2>&1
title Scanner Pack Tool

echo ============================================================
echo            Document Scanner - Pack Tool
echo ============================================================
echo.

REM Check Python
echo [1/7] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found, please install Python 3.8+
    pause
    exit /b 1
)
python --version
echo [OK] Python ready
echo.

REM Check and install PyInstaller
echo [2/7] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing PyInstaller...
    pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo [ERROR] PyInstaller install failed
        pause
        exit /b 1
    )
)
echo [OK] PyInstaller ready
echo.

REM Install dependencies
echo [3/7] Installing dependencies...
pip install pymupdf -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install pillow -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install python-pptx -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install ebooklib -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple -q
pip install qrcode[pil] -i https://pypi.tuna.tsinghua.edu.cn/simple -q
echo [OK] Dependencies ready
echo.

REM Check icon
echo [4/7] Checking icon...
if not exist "icon.ico" (
    echo [WARN] icon.ico not found, using default
    set ICON_ARG=
) else (
    echo [OK] icon.ico found
    set ICON_ARG=--icon="icon.ico"
)
echo.

REM Check scanner.py and data
echo [5/7] Checking files...
if not exist "scanner.py" (
    echo [ERROR] scanner.py not found!
    pause
    exit /b 1
)
echo [OK] scanner.py found

if exist "data" (
    echo [OK] data folder found
    set DATA_ARG=--add-data="data;data"
) else (
    echo [WARN] data folder not found
    set DATA_ARG=
)
echo.

REM Clean old files
echo [6/7] Cleaning old files...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "*.spec" del /q *.spec >nul 2>&1
echo [OK] Clean done
echo.

REM Start packaging
echo [7/7] Packaging...
echo ============================================================
echo Script: scanner.py
echo Output: dist\Scanner.exe
echo ============================================================
echo.

pyinstaller --onefile --windowed --name="Scanner" %ICON_ARG% %DATA_ARG% --hidden-import=tkinter --hidden-import=tkinter.ttk --hidden-import=fitz --hidden-import=PIL --hidden-import=PIL.Image --hidden-import=PIL.ImageDraw --hidden-import=PIL.ImageFont --hidden-import=docx --hidden-import=pptx --hidden-import=ebooklib --hidden-import=win32com --hidden-import=win32com.client --hidden-import=qrcode --hidden-import=http.server --hidden-import=socketserver --hidden-import=webbrowser --hidden-import=pathlib --hidden-import=json --hidden-import=hashlib --hidden-import=threading --hidden-import=shutil --hidden-import=ctypes --exclude-module=matplotlib --exclude-module=numpy --exclude-module=pandas --exclude-module=scipy --exclude-module=tensorflow --exclude-module=torch --clean --noconfirm scanner.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo [FAILED] Packaging failed!
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo ============================================================
echo [SUCCESS] Packaging completed!
echo Output: dist\Scanner.exe
echo ============================================================
echo.
echo Usage:
echo   GUI mode: double click Scanner.exe
echo   CMD mode: Scanner.exe --root "folder path"
echo   Options: --root, --out, --ext, --pages, --force
echo.

start explorer dist
pause