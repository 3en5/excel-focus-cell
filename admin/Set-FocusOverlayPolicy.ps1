# Central switch for Natanor Focus Overlay (administrators). Windows PowerShell 5.1, run elevated.
#
#   -Disable   turn the highlight OFF for every user on this machine. Users cannot turn it back on.
#   -Allow     leave the choice to each user (the ribbon button works again).
#   -Clear     remove the policy value (same effect as -Allow, and nothing is left behind).
#   (no switch) show the current state only.
#
# Writes only HKLM\SOFTWARE\Policies\Natanor\FocusOverlay\Enabled. Running Excel sessions pick the change up within
# about 5 seconds; nothing is restarted and no one is signed out. The value is read back after every change.
param([switch]$Disable, [switch]$Allow, [switch]$Clear)
$ErrorActionPreference = 'Stop'
$key = 'HKLM:\SOFTWARE\Policies\Natanor\FocusOverlay'

function Show-State {
    $v = (Get-ItemProperty $key -Name Enabled -ErrorAction SilentlyContinue).Enabled
    $state = if ($null -eq $v) { 'NOT SET (each user decides)' } elseif ($v -eq 0) { 'DISABLED for all users' } else { 'ALLOWED (each user decides)' }
    "policy: $state   [$key\Enabled = $(if ($null -eq $v) { '<none>' } else { $v })]"
}

if (($Disable, $Allow, $Clear | Where-Object { $_ }).Count -gt 1) { throw 'Use only one of -Disable, -Allow, -Clear' }
if ($Disable -or $Allow -or $Clear) {
    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $admin) { throw 'Run this script elevated (as administrator)' }
    if ($Clear) {
        Remove-ItemProperty $key -Name Enabled -ErrorAction SilentlyContinue
    } else {
        New-Item $key -Force | Out-Null
        New-ItemProperty $key -Name Enabled -Value $(if ($Disable) { 0 } else { 1 }) -PropertyType DWord -Force | Out-Null
    }
}
Show-State
