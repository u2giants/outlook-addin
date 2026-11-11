# POP Tools for Outlook - per-user install (no admin). Run from the folder holding PopOutlookAddin.dll.
#   Install.ps1            -> install/update, then shows a message
#   Install.ps1 -Silent    -> no message (used by auto-update)
#   Install.ps1 -Uninstall -> remove
param([switch]$Silent, [switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$ProgId = 'PopOutlookAddin.Connect'
$ClsId  = '{AB496E0D-CC70-4F80-854B-0EA3E4C4FBE9}'
$Dest   = Join-Path $env:LOCALAPPDATA 'PopOutlookAddin'
$Src    = Split-Path -Parent $MyInvocation.MyCommand.Path
$AddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
$ClsKey   = "HKCU:\Software\Classes\CLSID\$ClsId"

function Say($m) { if (-not $Silent) { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show($m, 'POP Tools') } }

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
  if ($Silent) { Wait-Process -Name OUTLOOK } else { Say 'Please close Outlook, then click OK.'; Wait-Process -Name OUTLOOK -ErrorAction SilentlyContinue }
}

if ($Uninstall) {
  Remove-Item $AddinKey, $ClsKey, "HKCU:\Software\Classes\$ProgId" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item $Dest -Recurse -Force -ErrorAction SilentlyContinue
  Say 'POP Tools removed.'; exit 0
}

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
Copy-Item (Join-Path $Src 'PopOutlookAddin.dll'), (Join-Path $Src 'Microsoft.Office.Interop.Outlook.dll') $Dest -Force
$dll = Join-Path $Dest 'PopOutlookAddin.dll'
$asm = [System.Reflection.AssemblyName]::GetAssemblyName($dll).FullName

New-Item -Force "HKCU:\Software\Classes\$ProgId\CLSID" -Value $ClsId | Out-Null
Set-Item "HKCU:\Software\Classes\$ProgId" -Value $ProgId
New-Item -Force "$ClsKey\ProgId" -Value $ProgId | Out-Null
Set-Item $ClsKey -Value $ProgId
$ip = "$ClsKey\InprocServer32"
New-Item -Force $ip -Value 'mscoree.dll' | Out-Null
Set-ItemProperty $ip ThreadingModel 'Both'
Set-ItemProperty $ip Class 'PopOutlookAddin.Connect'
Set-ItemProperty $ip Assembly $asm
Set-ItemProperty $ip RuntimeVersion 'v4.0.30319'
Set-ItemProperty $ip CodeBase ('file:///' + $dll.Replace('\', '/'))

New-Item -Force $AddinKey | Out-Null
Set-ItemProperty $AddinKey FriendlyName 'POP Tools'
Set-ItemProperty $AddinKey Description 'Retention column and attachment tools'
New-ItemProperty $AddinKey LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
$res = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList'
if (-not (Test-Path $res)) { New-Item -Force $res | Out-Null }
New-ItemProperty $res $ProgId -Value 1 -PropertyType DWord -Force | Out-Null

# Retire the old VBA macros so they don't run twice (kept as a backup, never deleted).
$otm = Join-Path $env:APPDATA 'Microsoft\Outlook\VbaProject.OTM'
if ((Test-Path $otm) -and -not (Test-Path "$otm.pre-addin.bak")) { Rename-Item $otm "$otm.pre-addin.bak" }

Say "POP Tools installed ($asm). Open Outlook and look for the 'POP Tools' tab."
