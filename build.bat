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
echo  Built: Output\ExcelIntelligenceAgent-Setup.exe  (v%APPVER%)
echo ============================================

REM 4) Publish to GitHub Releases so installed apps auto-update on next launch.
REM Find gh even if it isn't on PATH (e.g. installed in this same session).
set "GHEXE="
where gh >nul 2>&1 && set "GHEXE=gh"
if not defined GHEXE if exist "%ProgramFiles%\GitHub CLI\gh.exe" set "GHEXE=%ProgramFiles%\GitHub CLI\gh.exe"
if not defined GHEXE if exist "%LOCALAPPDATA%\Programs\GitHub CLI\gh.exe" set "GHEXE=%LOCALAPPDATA%\Programs\GitHub CLI\gh.exe"
if not defined GHEXE (
  echo.
  echo  gh CLI not found -- skipping auto-publish.
  echo  Install once:  winget install GitHub.cli   then   gh auth login
  echo  Or upload Output\ExcelIntelligenceAgent-Setup.exe to a new GitHub Release manually.
  goto :done
)
echo Publishing release v%APPVER% to GitHub...
"%GHEXE%" release create v%APPVER% "Output\ExcelIntelligenceAgent-Setup.exe" --title "v%APPVER%" --notes "Update to v%APPVER%" --repo alamehmazen123/the_excel_agent
if errorlevel 1 (
  echo.
  echo  Publish failed -- most likely you are not logged in yet.
  echo  Run this once:   "%GHEXE%" auth login
  echo  then re-run build.bat. Or upload the Setup.exe to GitHub Releases manually.
  goto :done
)
echo.
echo  Published. Colleagues' apps will auto-update to v%APPVER% on next launch.

:done
echo.
echo  DONE.
pause
exit /b 0

:error
echo.
echo Build FAILED. See messages above.
pause
exit /b 1
