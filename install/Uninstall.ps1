# Removes the add-in.  -Machine removes the all-users install (run elevated); without it, the current user's install.
# Your saved choice (Always / Search / Off) under HKCU\Software\Natanor\FocusOverlay is kept unless -RemoveSettings.
param([switch]$Machine, [switch]$RemoveSettings)
$ErrorActionPreference = 'Stop'
$clsid = '{6E1D7B0A-3C52-4F8E-9A41-2B7C5D0F1E93}'
$progId = 'NatanorFocusOverlay.Connect'
if ($Machine) {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Run elevated to remove the all-users install' }
    $root = 'HKLM:\SOFTWARE'; $dir = Join-Path $env:ProgramFiles 'Natanor Focus Overlay'
} else {
    $root = 'HKCU:\Software'; $dir = Join-Path $env:LOCALAPPDATA 'Programs\NatanorFocusOverlay'
}
Remove-Item "$root\Microsoft\Office\Excel\Addins\$progId" -Recurse -ErrorAction SilentlyContinue
Remove-Item "$root\Classes\CLSID\$clsid" -Recurse -ErrorAction SilentlyContinue
Remove-Item "$root\Classes\$progId" -Recurse -ErrorAction SilentlyContinue
try { Remove-Item $dir -Recurse -Force -ErrorAction Stop; "removed $dir" } catch { "could not remove $dir (close Excel and run again): $($_.Exception.Message)" }
if ($RemoveSettings) { Remove-Item 'HKCU:\Software\Natanor\FocusOverlay' -Recurse -ErrorAction SilentlyContinue }
"add-in key present: $(Test-Path "$root\Microsoft\Office\Excel\Addins\$progId")"
