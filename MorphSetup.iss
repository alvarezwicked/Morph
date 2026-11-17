
#define MyAppName "Morph"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "vincente.soft"
#define MyAppExeName "Morph.exe"

[Setup]
; =====================================================
; APPLICATION INFORMATION
; =====================================================

; Original AppId from Morph 1.0.0.
; Do not change: required for upgrading.
DisableWelcomePage=no
AppId={{957DD601-51FB-45BA-AC60-7DC81D684CD9}

AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}

; =====================================================
; INSTALLATION DIRECTORY
; =====================================================

DefaultDirName={autopf}\Morph
DefaultGroupName=Morph

; Remember the previous installation location.
UsePreviousAppDir=yes

; Always show the destination folder selection.
DisableDirPage=no

; Allow choosing the Start Menu folder.
DisableProgramGroupPage=no

; =====================================================
; INSTALLER OUTPUT
; =====================================================

OutputDir=installer-output
OutputBaseFilename=Morph-Setup-1.0.1

Compression=lzma2
SolidCompression=yes

; =====================================================
; BRANDING AND APPEARANCE
; =====================================================

WizardStyle=modern

; Application icon for the installer EXE.
SetupIconFile=Morph\Assets\Morph.ico

; Large branding image on the welcome/finish pages.
WizardImageFile=Morph\Assets\branding\wizard-large.bmp

; Small branding image on other wizard pages.
WizardSmallImageFile=Morph\Assets\branding\Morph_AppIcon_512.png

; Scale images to fit the wizard.
WizardImageStretch=yes

; =====================================================
; WINDOWS ARCHITECTURE
; =====================================================

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; Required for installing into Program Files.
PrivilegesRequired=admin

; =====================================================
; UPGRADE BEHAVIOR
; =====================================================

; Detect applications using files being updated.
CloseApplications=yes

; Do not automatically restart closed applications.
RestartApplications=no

; =====================================================
; UNINSTALL INFORMATION
; =====================================================

UninstallDisplayName=Morph
UninstallDisplayIcon={app}\Morph.exe

; =====================================================
; OTHER SETTINGS
; =====================================================

AllowNoIcons=yes
ChangesAssociations=no
ShowLanguageDialog=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; =====================================================
; OPTIONAL INSTALLATION TASKS
; =====================================================

; Create a desktop shortcut.
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

; Launch Morph automatically when Windows starts.
Name: "autostart"; Description: "Launch Morph when Windows starts"; GroupDescription: "Additional options:"; Flags: unchecked

[Files]
; =====================================================
; APPLICATION FILES
; =====================================================

; Include the entire self-contained .NET publish directory.
Source: "publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; =====================================================
; SHORTCUTS
; =====================================================

; Start Menu shortcut.
Name: "{group}\Morph"; Filename: "{app}\Morph.exe"; WorkingDir: "{app}"

; Optional desktop shortcut.
Name: "{autodesktop}\Morph"; Filename: "{app}\Morph.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
; =====================================================
; WINDOWS STARTUP
; =====================================================

; Add Morph to startup when the user selects the option.
; Remove the registry value when Morph is uninstalled.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "Morph"; ValueData: """{app}\Morph.exe"""; Flags: uninsdeletevalue; Tasks: autostart

[Run]
; =====================================================
; LAUNCH AFTER INSTALLATION
; =====================================================

Filename: "{app}\Morph.exe"; Description: "Launch Morph"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent
