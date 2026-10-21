# Installs the add-in for the CURRENT USER only. No administrator rights needed.
# Windows PowerShell 5.1:  powershell -ExecutionPolicy Bypass -File Install-User.ps1 [-Dll <path>]
# Limitation: a per-user COM registration is ignored by an ELEVATED Excel (e.g. the built-in Administrator account,
# which runs everything elevated by default). For such accounts use Install-Machine.ps1.
param([string]$Dll = (Join-Path $PSScriptRoot '..\bin\Natanor.FocusOverlay.dll'))
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Dll)) { throw "DLL not found: $Dll (run build.ps1 or pass -Dll)" }
$dir = Join-Path $env:LOCALAPPDATA 'Programs\NatanorFocusOverlay'
$target = Join-Path $dir 'Natanor.FocusOverlay.dll'
$clsid = '{6E1D7B0A-3C52-4F8E-9A41-2B7C5D0F1E93}'
$progId = 'NatanorFocusOverlay.Connect'

if (Get-Process EXCEL -ErrorAction SilentlyContinue) { 'Note: Excel is running - close it before the first use; the DLL may be locked on reinstall.' }
New-Item -ItemType Directory -Force $dir | Out-Null
Copy-Item $Dll $target -Force
Unblock-File $target

$asm = [Reflection.AssemblyName]::GetAssemblyName($target).FullName
$cls = "HKCU:\Software\Classes\CLSID\$clsid"
New-Item "$cls\InprocServer32" -Force | Out-Null
Set-ItemProperty "$cls\InprocServer32" '(default)' 'mscoree.dll'
Set-ItemProperty "$cls\InprocServer32" 'ThreadingModel' 'Both'
Set-ItemProperty "$cls\InprocServer32" 'Class' 'Natanor.FocusOverlay.Connect'
Set-ItemProperty "$cls\InprocServer32" 'Assembly' $asm
Set-ItemProperty "$cls\InprocServer32" 'RuntimeVersion' 'v4.0.30319'
Set-ItemProperty "$cls\InprocServer32" 'CodeBase' ([Uri]$target).AbsoluteUri
New-Item "$cls\ProgId" -Force | Set-ItemProperty -Name '(default)' -Value $progId
New-Item "HKCU:\Software\Classes\$progId\CLSID" -Force | Set-ItemProperty -Name '(default)' -Value $clsid
$addin = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progId"
New-Item $addin -Force | Out-Null
Set-ItemProperty $addin 'FriendlyName' 'Natanor Focus Overlay'
Set-ItemProperty $addin 'Description' 'Highlights the row and column of the active cell'
New-ItemProperty $addin 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

"installed: $target"
"LoadBehavior: $((Get-ItemProperty $addin).LoadBehavior)"
'Start (or restart) Excel and look for the "Focus" group on the View tab.'
