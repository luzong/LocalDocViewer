@echo off
title Document Scanner Pack

echo ============================================
echo       Document Scanner Pack Tool
echo ============================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found
    pause
    exit /b 1
)
echo [OK] Python ready
echo.

pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
echo [OK] PyInstaller ready
echo.

pip install pymupdf pillow python-docx python-pptx ebooklib pywin32 qrcode[pil] -i https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
echo [OK] Dependencies installed
echo.

if not exist "scanner.py" (
    echo [ERROR] scanner.py not found
    pause
    exit /b 1
)

set ICON_CMD=
if exist "icon.ico" (
    set ICON_CMD=--icon=icon.ico --add-data "icon.ico;."
    echo [OK] Icon found
) else (
    echo [WARN] No icon
)

set DATA_CMD=
if exist "data" (
    set DATA_CMD=--add-data "data;data"
    echo [OK] Data folder found
) else (
    echo [WARN] No data folder
)
echo.

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo.

echo ============================================
echo       Building Scanner.exe ...
echo ============================================
echo.

pyinstaller --onefile --windowed --name="Scanner" %ICON_CMD% %DATA_CMD% --clean --noconfirm --hidden-import=tkinter --hidden-import=fitz --hidden-import=PIL --hidden-import=docx --hidden-import=pptx --hidden-import=ebooklib --hidden-import=win32com --hidden-import=qrcode --exclude-module=matplotlib --exclude-module=numpy --exclude-module=pandas scanner.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo ============================================
echo       BUILD SUCCESS!
echo ============================================
echo Output: dist\Scanner.exe
echo.

set /p sc="Create desktop shortcut? (Y/N): "
if /i "%sc%"=="Y" (
    echo Set WshShell = WScript.CreateObject("WScript.Shell") > %temp%\sc.vbs
    echo Set lnk = WshShell.CreateShortcut("%USERPROFILE%\Desktop\Scanner.lnk") >> %temp%\sc.vbs
    echo lnk.TargetPath = "%~dp0dist\Scanner.exe" >> %temp%\sc.vbs
    echo lnk.IconLocation = "%~dp0icon.ico" >> %temp%\sc.vbs
    echo lnk.Save >> %temp%\sc.vbs
    cscript //nologo %temp%\sc.vbs
    del %temp%\sc.vbs
    echo [OK] Shortcut created
)

set /p open="Open output folder? (Y/N): "
if /i "%open%"=="Y" explorer "dist"

echo.
pause