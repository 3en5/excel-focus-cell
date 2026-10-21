# Installs the add-in for EVERY user of this computer (terminal servers, shared PCs). Run elevated.
# Windows PowerShell 5.1:  powershell -ExecutionPolicy Bypass -File Install-Machine.ps1 [-Dll <path>]
# Copies the DLL to Program Files, registers the COM class and the Excel add-in under HKLM, reads everything back.
# Nothing is restarted: the add-in loads the next time each user starts Excel. Remove with Uninstall.ps1 -Machine.
param([string]$Dll = (Join-Path $PSScriptRoot '..\bin\Natanor.FocusOverlay.dll'))
$ErrorActionPreference = 'Stop'
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script elevated (as administrator)' }
if (-not (Test-Path $Dll)) { throw "DLL not found: $Dll (run build.ps1 or pass -Dll)" }
$dir = Join-Path $env:ProgramFiles 'Natanor Focus Overlay'
$target = Join-Path $dir 'Natanor.FocusOverlay.dll'
$clsid = '{6E1D7B0A-3C52-4F8E-9A41-2B7C5D0F1E93}'
$progId = 'NatanorFocusOverlay.Connect'

New-Item -ItemType Directory -Force $dir | Out-Null
Copy-Item $Dll $target -Force
Unblock-File $target                     # a downloaded copy carries the internet zone mark, which .NET may refuse to load

$asm = [Reflection.AssemblyName]::GetAssemblyName($target).FullName
$cls = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
New-Item "$cls\InprocServer32" -Force | Out-Null
Set-ItemProperty "$cls\InprocServer32" '(default)' 'mscoree.dll'
Set-ItemProperty "$cls\InprocServer32" 'ThreadingModel' 'Both'
Set-ItemProperty "$cls\InprocServer32" 'Class' 'Natanor.FocusOverlay.Connect'
Set-ItemProperty "$cls\InprocServer32" 'Assembly' $asm
Set-ItemProperty "$cls\InprocServer32" 'RuntimeVersion' 'v4.0.30319'
Set-ItemProperty "$cls\InprocServer32" 'CodeBase' ([Uri]$target).AbsoluteUri
New-Item "$cls\ProgId" -Force | Set-ItemProperty -Name '(default)' -Value $progId
New-Item "HKLM:\SOFTWARE\Classes\$progId\CLSID" -Force | Set-ItemProperty -Name '(default)' -Value $clsid
$addin = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\$progId"
New-Item $addin -Force | Out-Null
Set-ItemProperty $addin 'FriendlyName' 'Natanor Focus Overlay'
Set-ItemProperty $addin 'Description' 'Highlights the row and column of the active cell'
New-ItemProperty $addin 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

"installed: $target"
"codebase:  $((Get-ItemProperty "$cls\InprocServer32").CodeBase)"
"LoadBehavior: $((Get-ItemProperty $addin).LoadBehavior)"
try { [void][Activator]::CreateInstance([Type]::GetTypeFromProgID($progId)); 'COM activation: OK' } catch { "COM activation FAILED: $($_.Exception.InnerException.Message)" }
'Start (or restart) Excel and look for the "Focus" group on the View tab.'
