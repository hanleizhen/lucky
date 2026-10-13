; The installer contains application files only. User data is deliberately not
; placed under {app}, so an upgrade cannot overwrite it.
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "CNC 智能程序单"
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
SetupIconFile=assets\cnc_smart_sheet_icon.ico
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
CloseApplications=yes
RestartApplications=no
UninstallDisplayName={#MyAppName}

[Files]
Source: "build\dist\CNCProgramSheet\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Put a versioned standalone icon beside the EXE.  Using a new icon path for
; every release prevents the Windows Shell shortcut cache from retaining an
; old EXE icon after an in-place upgrade.
Source: "assets\cnc_smart_sheet_icon.ico"; DestDir: "{app}"; DestName: "CNCProgramSheet_icon_{#MyAppVersion}.ico"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\CNCProgramSheet_icon_{#MyAppVersion}.ico"; IconIndex: 0; AppUserModelID: "CNCProgramSheet.CNCSmartSheet"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\CNCProgramSheet_icon_{#MyAppVersion}.ico"; IconIndex: 0; AppUserModelID: "CNCProgramSheet.CNCSmartSheet"

[Registry]
; Per-user Explorer verbs require no administrator privilege.  ``Player``
; passes all selected files to one background process, and ``%*`` preserves
; each selected path as a quoted command-line argument.
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M06"; ValueType: string; ValueName: "MUIVerb"; ValueData: "CNC 智能程序单 - M06 修改"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M06"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\CNCProgramSheet_icon_{#MyAppVersion}.ico,0"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M06"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M06\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" --shell-transform M06 ""%*"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M304"; ValueType: string; ValueName: "MUIVerb"; ValueData: "CNC 智能程序单 - M304 修改"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M304"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\CNCProgramSheet_icon_{#MyAppVersion}.ico,0"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M304"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.M304\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" --shell-transform M304 ""%*"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.Copper"; ValueType: string; ValueName: "MUIVerb"; ValueData: "CNC 智能程序单 - 铜工修改"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.Copper"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\CNCProgramSheet_icon_{#MyAppVersion}.ico,0"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.Copper"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.nc\shell\CNCProgramSheet.Copper\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" --shell-transform TG ""%*"""

[InstallDelete]
; Recreate the current shortcuts so their IconFilename is refreshed.  The app
; id, executable, installation folder and user-data directory remain
; unchanged, so updates stay seamless.
Type: files; Name: "{autoprograms}\CNC 智能程序单.lnk"
Type: files; Name: "{autodesktop}\CNC 智能程序单.lnk"
; Rename only the old app shortcuts.
Type: files; Name: "{autoprograms}\CNC程序单自动生成工具.lnk"
Type: files; Name: "{autodesktop}\CNC程序单自动生成工具.lnk"

[Dirs]
; This directory is owned by the user. The app creates settings, templates,
; logs and update downloads here. It is never in [Files] or [UninstallDelete].
Name: "{localappdata}\CNCProgramSheet"; Flags: uninsneveruninstall

[Run]
; The updater uses /VERYSILENT. Launch the new version at the end so a user
; who started with a portable Desktop EXE is switched to the installed build.
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait
