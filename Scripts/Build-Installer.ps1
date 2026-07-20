#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidatePattern('^\d+\.\d+\.\d+(?:\.\d+)?$')]
    [string]$Version = '1.1.0',

    [string]$OutputDirectory = 'dist',

    [switch]$SkipTests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$manifest = Join-Path $root 'Installer\AutoTools.iss'
$testScript = Join-Path $root 'Tests\Test-Installer.ps1'
$launcherScript = Join-Path $root 'Launcher.ps1'
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    [System.IO.Path]::GetFullPath($OutputDirectory)
}
else {
    [System.IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
}

if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
    throw "Khong tim thay Inno Setup manifest: $manifest"
}

$launcherVersion = [regex]::Match(
    (Get-Content -LiteralPath $launcherScript -Raw),
    '(?m)^\$AppVersion\s*=\s*''([^'']+)''$'
)

if (-not $launcherVersion.Success) {
    throw 'Khong doc duoc $AppVersion trong Launcher.ps1.'
}

if ($Version -ne $launcherVersion.Groups[1].Value) {
    throw "Build version $Version khong khop Launcher version $($launcherVersion.Groups[1].Value)."
}

if (-not $SkipTests -and (Test-Path -LiteralPath $testScript -PathType Leaf)) {
    & $testScript
}

$isccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
$isccCandidates = @($env:INNO_SETUP_ISCC)

if ($null -ne $isccCommand) {
    $isccCandidates += $isccCommand.Source
}

if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    $isccCandidates += Join-Path `
        $env:LOCALAPPDATA `
        'Programs\Inno Setup 6\ISCC.exe'
}

foreach ($programFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if (-not [string]::IsNullOrWhiteSpace($programFiles)) {
        $isccCandidates += Join-Path $programFiles 'Inno Setup 6\ISCC.exe'
    }
}

$iscc = $isccCandidates |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_) -and
        (Test-Path -LiteralPath $_ -PathType Leaf)
    } |
    Select-Object -First 1

if ([string]::IsNullOrWhiteSpace($iscc)) {
    throw @'
Khong tim thay ISCC.exe. Hay cai Inno Setup 6 tu:
https://jrsoftware.org/isdl.php
'@
}

$null = New-Item -Path $outputPath -ItemType Directory -Force

& $iscc @(
    '/Qp'
    "/DMyAppVersion=$Version"
    "/O$outputPath"
    $manifest
)

if ($LASTEXITCODE -ne 0) {
    throw "ISCC.exe failed with exit code $LASTEXITCODE."
}

$installer = Join-Path $outputPath "AutoTools-Setup-$Version.exe"
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
    throw "Khong tao duoc bo cai: $installer"
}

Write-Host "Done: $installer"
