#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif

#define MyAppName "Auto Tools"
#define MyAppId "{{E53CE2D7-2E64-482E-8F5E-A4F41BC1C570}"
#define PowerShellPath "{sys}\WindowsPowerShell\v1.0\powershell.exe"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppName}
DefaultDirName={localappdata}\Programs\Auto Tools
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=AutoTools-Setup-{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
VersionInfoVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
UninstallDisplayName={#MyAppName}

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\Launcher.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\AutoTools.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\HUONG_DAN_SETUP_AUTO_TOOLS.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\launcher-config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Config\autotools-paths.json"; DestDir: "{app}\Config"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Config\suppliers.csv"; DestDir: "{app}\Config"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Input\Process-CDs\cds_input.txt"; DestDir: "{app}\Input\Process-CDs"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Input\Process-PO\input.txt"; DestDir: "{app}\Input\Process-PO"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\Scripts\*"; DestDir: "{app}\Scripts"; Excludes: "\Temp\*,\Process-CDs\Output\*,\Process-CDs\Older CDs\*,\Process-PO\Output\*,\Process-PO\Older Output PO\*"; Flags: ignoreversion recursesubdirs
Source: "AutoTools.iss"; DestDir: "{app}\Installer"; Flags: ignoreversion

[Dirs]
Name: "{app}\Input"; Flags: uninsneveruninstall
Name: "{app}\Input\ASCP"; Flags: uninsneveruninstall
Name: "{app}\Input\Data Job Short - Not Delete"; Flags: uninsneveruninstall
Name: "{app}\Input\Job Short"; Flags: uninsneveruninstall
Name: "{app}\Input\MR-Outlook"; Flags: uninsneveruninstall
Name: "{app}\Input\Process-CDs"; Flags: uninsneveruninstall
Name: "{app}\Input\Process-PO"; Flags: uninsneveruninstall
Name: "{app}\Input\Supplier Commitment\TC5"; Flags: uninsneveruninstall
Name: "{app}\Input\Supplier Commitment\TN5"; Flags: uninsneveruninstall
Name: "{app}\Output"; Flags: uninsneveruninstall
Name: "{app}\Output\Process-CDs"; Flags: uninsneveruninstall
Name: "{app}\Output\Process-PO"; Flags: uninsneveruninstall
Name: "{app}\Output\Supplier Commitment"; Flags: uninsneveruninstall
Name: "{app}\Logs"; Flags: uninsneveruninstall
Name: "{app}\Temp"; Flags: uninsneveruninstall
Name: "{app}\MR_Out"; Flags: uninsneveruninstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{#PowerShellPath}"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Launcher.ps1"""; WorkingDir: "{app}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{#PowerShellPath}"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Launcher.ps1"""; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{#PowerShellPath}"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\Launcher.ps1"""; WorkingDir: "{app}"; Description: "Launch {#MyAppName}"; Flags: postinstall nowait skipifsilent
