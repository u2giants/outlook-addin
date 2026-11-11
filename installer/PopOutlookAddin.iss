; Per-user installer for the POP Tools Outlook add-in. No admin rights needed.
; Silent install:  PopOutlookAddin-Setup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART

#ifndef AppVersion
  #define AppVersion "1.0.0.0"
#endif
#define ClsId "{{AB496E0D-CC70-4F80-854B-0EA3E4C4FBE9}"
#define ProgId "PopOutlookAddin.Connect"

[Setup]
AppId={{6A1C2B47-3E0F-4D7B-9E7F-0C1D5B0F7A11}
AppName=POP Tools for Outlook
AppVersion={#AppVersion}
VersionInfoVersion={#AppVersion}
AppPublisher=POP Creations
DefaultDirName={localappdata}\PopOutlookAddin
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=PopOutlookAddin-Setup
Compression=lzma2
SolidCompression=yes
CloseApplications=force
CloseApplicationsFilter=*.dll
UninstallDisplayName=POP Tools for Outlook

[Files]
Source: "..\src\bin\Release\net48\PopOutlookAddin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\bin\Release\net48\Microsoft.Office.Interop.Outlook.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; COM registration (per user), loaded through the .NET runtime.
Root: HKCU; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsId}}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\ProgId"; ValueType: string; ValueData: "{#ProgId}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "PopOutlookAddin.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "PopOutlookAddin, Version={#AppVersion}, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#ClsId}}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}\PopOutlookAddin.dll"
; Tell Outlook to load it at startup.
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "POP Tools"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Retention column and attachment tools"
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
; Keep Outlook from auto-disabling it as "slow".
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "{#ProgId}"; ValueData: 1; Flags: uninsdeletevalue

[Code]
// Retire the old VBA macros so they don't run twice. Kept as a backup, never deleted.
procedure CurStepChanged(CurStep: TSetupStep);
var Otm, Bak: String;
begin
  if CurStep = ssPostInstall then
  begin
    Otm := ExpandConstant('{userappdata}\Microsoft\Outlook\VbaProject.OTM');
    Bak := Otm + '.pre-addin.bak';
    if FileExists(Otm) and not FileExists(Bak) then
      RenameFile(Otm, Bak);
  end;
end;
