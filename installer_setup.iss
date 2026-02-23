; Script Inno Setup per Analizzatore Schede Taratura
; Genera un installer Windows professionale

[Setup]
AppId={{8A2B3C4D-5E6F-4A7B-8C9D-0E1F2A3B4C5D}
AppName=Analizzatore Schede Taratura
AppVersion=8.1
AppPublisher=Coemi S.r.l.
DefaultDirName={autopf}\AnalizzatoreSchede
DefaultGroupName=Analizzatore Schede Taratura
AllowNoIcons=yes
; Rimuovi la riga sotto se non hai un file icona .ico
; SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
OutputDir=installer_output
OutputBaseFilename=AnalizzatoreSchede_Setup_v8.1

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\AnalizzatoreSchede.exe"; DestDir: "{app}"; Flags: ignoreversion
; Aggiungi qui eventuali file template o documentazione
; Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Analizzatore Schede Taratura"; Filename: "{app}\AnalizzatoreSchede.exe"
Name: "{group}\{cm:UninstallProgram,Analizzatore Schede Taratura}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Analizzatore Schede Taratura"; Filename: "{app}\AnalizzatoreSchede.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\AnalizzatoreSchede.exe"; Description: "{cm:LaunchProgram,Analizzatore Schede Taratura}"; Flags: nowait postinstall skipifsilent
