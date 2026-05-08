[Setup]
AppName=DocGit
AppVersion=1.1
AppPublisher=Arfan
DefaultDirName={pf}\DocGit
DefaultGroupName=DocGit
OutputDir=Output
OutputBaseFilename=DocGit_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
ChangesEnvironment=yes
PrivilegesRequired=admin

[Files]
; Core executables
Source: "dist\docgit.exe";         DestDir: "{app}"; Flags: ignoreversion
Source: "dist\docgit_server.exe";  DestDir: "{app}"; Flags: ignoreversion
; Add-in manifest
Source: "manifest.xml";            DestDir: "{app}"; Flags: ignoreversion
; mkcert binary for SSL cert generation on user machine
Source: "mkcert.exe";              DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Add {app} to PATH
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; \
    Check: NeedsAddPath(ExpandConstant('{app}'))

; Register trusted catalog for Word sideloading
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEBF\TrustedCatalogs\{{2b339463-b1d6-444d-b94f-4d943b171c7b}"; \
    ValueType: string; ValueName: "Url"; ValueData: "\\localhost\DocGitAddin"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEBF\TrustedCatalogs\{{2b339463-b1d6-444d-b94f-4d943b171c7b}"; \
    ValueType: dword;  ValueName: "Flags"; ValueData: "1"; Flags: uninsdeletekey

; Auto-start server on Windows login (hidden, no console window)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "DocGitServer"; \
    ValueData: """{app}\docgit_server.exe"""; Flags: uninsdeletevalue

[Icons]
Name: "{group}\DocGit Server"; Filename: "{app}\docgit_server.exe"
Name: "{group}\Uninstall DocGit"; Filename: "{uninstallexe}"

[Run]
; Step 1: Install mkcert local CA into Windows trust store
Filename: "{app}\mkcert.exe"; \
    Parameters: "-install"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Installing trusted certificate authority..."

; Step 2: Generate localhost SSL certificate
Filename: "{app}\mkcert.exe"; \
    Parameters: "-cert-file ""{app}\localhost+1.pem"" -key-file ""{app}\localhost+1-key.pem"" localhost 127.0.0.1"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Generating SSL certificate..."

; Step 3a: Remove any existing share (from old install) before recreating
Filename: "cmd.exe"; \
    Parameters: "/c net share DocGitAddin /delete /y"; \
    Flags: runhidden waituntilterminated

; Step 3b: Grant NTFS read permissions so Everyone can access the folder
Filename: "icacls.exe"; \
    Parameters: """{app}"" /grant ""Everyone:(OI)(CI)R"" /T"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Setting folder permissions..."

; Step 3c: Create network share for Word sideloading
Filename: "cmd.exe"; \
    Parameters: "/c net share DocGitAddin=""{app}"" /grant:Everyone,READ"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Configuring Word add-in catalog..."

; Step 4: Launch server immediately (optional, user can skip)
Filename: "{app}\docgit_server.exe"; \
    Flags: nowait postinstall skipifsilent; \
    Description: "Start DocGit Server now"

[UninstallRun]
Filename: "cmd.exe"; Parameters: "/c net share DocGitAddin /delete"; Flags: runhidden

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;
