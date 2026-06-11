; Inno Setup script -> ExcelIntelligenceAgent-Setup.exe
; Wraps the PyInstaller onedir build (dist\ExcelIntelligenceAgent) into a single,
; robust, per-user installer. No admin rights required.
;
; Compile with:  "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" installer.iss
; (build.bat does this automatically after PyInstaller.)

#define AppName "Excel Intelligence Agent"
#define AppExe  "ExcelIntelligenceAgent.exe"
#define AppPublisher "Sahel General Hospital"
; AppVersion is passed in by build.bat via /DAppVersion=...; default for manual runs:
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{8B2F3C61-5E2A-4C9D-9F3A-EXCELAGENT001}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Per-user install -> no UAC / admin prompt, works for every colleague.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={localappdata}\Programs\ExcelIntelligenceAgent
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=ExcelIntelligenceAgent-Setup
SetupIconFile=ui\resources\app.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; Close the app automatically if a file is in use during an update. The app is
; relaunched by the [Run] postinstall entry below (in both silent and
; interactive installs), so we do NOT also use Inno's RestartApplications -- that
; would launch a second instance.
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
; The entire onedir folder produced by PyInstaller.
Source: "dist\ExcelIntelligenceAgent\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
; Launch only after an INTERACTIVE (fresh) install. Silent background updates
; happen when the user closes the app, so they must NOT relaunch it.
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
