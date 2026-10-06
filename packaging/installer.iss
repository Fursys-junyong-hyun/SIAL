#define MyAppName "유상사급 타처보관 확인서"
#define MyAppVersion "1.04"
#define MyAppPublisher "FURSYS"
#define MyAppExeName "outsourced_inventory_confirmation.exe"
#define MyAppManualName "유상사급타처보관_사용자매뉴얼.pdf"

[Setup]
AppId={{1A30143C-9824-49C8-B7FD-8BAFB4F1B2E1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist_installer
OutputBaseFilename=outsourced_inventory_confirmation_setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면 바로가기 생성"; GroupDescription: "추가 작업:"; Flags: unchecked
Name: "manualshortcut"; Description: "사용 매뉴얼 PDF 바로가기 생성"; GroupDescription: "추가 작업:"

[Files]
Source: "..\dist\outsourced_inventory_confirmation\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\output\doc\{#MyAppManualName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autoprograms}\{#MyAppName} 사용 매뉴얼"; Filename: "{app}\{#MyAppManualName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{autodesktop}\{#MyAppName} 사용 매뉴얼"; Filename: "{app}\{#MyAppManualName}"; Tasks: manualshortcut

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} 실행"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppManualName}"; Description: "사용 매뉴얼 열기"; Flags: shellexec nowait postinstall skipifsilent unchecked
