@echo off
REM ==========================================
REM   PDF Splitter - Setup (run this once!)
REM ==========================================

echo.
echo  ========================================
echo    PDF Splitter - One-Time Setup
echo  ========================================
echo.
echo  This will:
echo    1. Create a folder for the tool
echo    2. Put a "Split PDF" shortcut on your Desktop
echo.
echo  Press any key to continue (or close this window to cancel)...
pause >nul

REM --- Create the install folder ---
set "INSTALL_DIR=%LOCALAPPDATA%\PDFSplitter"
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM --- Copy the PowerShell script next to this setup file ---
copy /Y "%~dp0Split-PDF.ps1" "%INSTALL_DIR%\Split-PDF.ps1" >nul 2>&1

if not exist "%INSTALL_DIR%\Split-PDF.ps1" (
    echo.
    echo  ERROR: Could not find Split-PDF.ps1
    echo  Make sure Split-PDF.ps1 is in the same folder as this setup file.
    echo.
    pause
    exit /b 1
)

REM --- Create a .bat launcher in the install folder ---
(
echo @echo off
echo powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%%LOCALAPPDATA%%\PDFSplitter\Split-PDF.ps1" %%1
) > "%INSTALL_DIR%\Split-PDF.bat"

REM --- Create a desktop shortcut using PowerShell ---
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $sc = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Split PDF.lnk')); ^
   $sc.TargetPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'PDFSplitter', 'Split-PDF.bat'); ^
   $sc.WorkingDirectory = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'PDFSplitter'); ^
   $sc.Description = 'Split a PDF into individual pages'; ^
   $sc.IconLocation = 'shell32.dll,21'; ^
   $sc.Save()"

echo.
echo  ========================================
echo    All done!
echo  ========================================
echo.
echo  You now have a "Split PDF" shortcut on your Desktop.
echo.
echo  To split a PDF:
echo    - Double-click "Split PDF" on your Desktop
echo    - Pick your PDF file
echo    - The individual pages will appear in a folder
echo      right next to the original file
echo.
echo  You can also drag-and-drop a PDF onto the shortcut!
echo.
echo  (You can delete this setup file now if you like.)
echo.
pause
