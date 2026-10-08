; The installer contains application files only. User data is deliberately not
; placed under {app}, so an upgrade cannot overwrite it.
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "CNC程序单自动生成工具"
#define MyAppExeName "CNCProgramSheet.exe"

[Setup]
AppId={{912FB6C9-AECD-49B5-811C-557A3231CF33}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=CNC Program Sheet
DefaultDirName={autopf}\CNC Program Sheet
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=build\installer
OutputBaseFilename={#MyAppName}_{#MyAppVersion}_Setup_x64
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
CloseApplications=yes
RestartApplications=no
UninstallDisplayName={#MyAppName}

[Files]
Source: "build\dist\CNCProgramSheet\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Dirs]
; This directory is owned by the user. The app creates settings, templates,
; logs and update downloads here. It is never in [Files] or [UninstallDelete].
Name: "{localappdata}\CNCProgramSheet"; Flags: uninsneveruninstall

[Run]
; The updater uses /VERYSILENT. Launch the new version at the end so a user
; who started with a portable Desktop EXE is switched to the installed build.
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait
