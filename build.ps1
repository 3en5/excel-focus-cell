# Builds Natanor.FocusOverlay.dll with the C# compiler that ships with .NET Framework 4.x (no SDK, no Visual Studio).
# Needs the Office interop assemblies (Microsoft.Office.Interop.Excel + office.dll, version 15) - they are in the GAC
# wherever Visual Studio's Office tools or the Office PIA redistributable are installed; or pass -InteropDir.
# Usage: powershell -ExecutionPolicy Bypass -File build.ps1 [-InteropDir <folder with both DLLs>]
param([string]$InteropDir)
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'bin'
$dll = Join-Path $out 'Natanor.FocusOverlay.dll'
if ($InteropDir) {
    $excel = Join-Path $InteropDir 'Microsoft.Office.Interop.Excel.dll'; $office = Join-Path $InteropDir 'office.dll'
} else {
    $gac = "$env:WINDIR\assembly\GAC_MSIL"
    $excel = "$gac\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll"
    $office = "$gac\office\15.0.0.0__71e9bce111e9429c\office.dll"
}
foreach ($r in $excel, $office) { if (-not (Test-Path $r)) { throw "Office interop assembly not found: $r (install the Office PIAs or use -InteropDir)" } }
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
New-Item -ItemType Directory -Force $out | Out-Null
$files = Get-ChildItem (Join-Path $PSScriptRoot 'src\NatanorFocusOverlay') -Filter *.cs | ForEach-Object { $_.FullName }
# /link embeds only the interop types actually used, so the DLL needs no PIA at run time
& $csc /nologo /target:library /platform:anycpu /optimize+ /unsafe /codepage:65001 /out:$dll `
    /link:"$excel" /link:"$office" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:Microsoft.CSharp.dll $files
if ($LASTEXITCODE -ne 0) { throw "csc failed: $LASTEXITCODE" }
"BUILT $dll  sha256 $((Get-FileHash $dll).Hash)"
