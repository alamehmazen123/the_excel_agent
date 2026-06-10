@echo off
REM One-click build of the robust installer.
REM Produces: Output\ExcelIntelligenceAgent-Setup.exe  (give THIS to colleagues)

echo ============================================
echo  Building Excel Intelligence Agent (installer)
echo ============================================

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

REM Regenerate the app icon (Excel-AI themed).
python tools\make_icon.py
if errorlevel 1 goto :error

REM Stamp today's date into buildinfo.py (shown in the app footer).
python tools\stamp_build.py
if errorlevel 1 goto :error

REM 1) PyInstaller -> onedir folder  dist\ExcelIntelligenceAgent\
python -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 goto :error

REM 2) Read APP_VERSION from config.py so the installer version matches.
for /f "usebackq delims=" %%v in (`python -c "import config; print(config.APP_VERSION)"`) do set APPVER=%%v

REM 3) Inno Setup -> Output\ExcelIntelligenceAgent-Setup.exe
set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo.
  echo Inno Setup not found. Install it once with:
  echo    winget install JRSoftware.InnoSetup
  echo The onedir build is still available in dist\ExcelIntelligenceAgent\
  goto :error
)
"%ISCC%" /DAppVersion=%APPVER% installer.iss
if errorlevel 1 goto :error

echo.
echo ============================================
echo  DONE. Distribute this single file:
echo    Output\ExcelIntelligenceAgent-Setup.exe
echo ============================================
pause
exit /b 0

:error
echo.
echo Build FAILED. See messages above.
pause
exit /b 1
