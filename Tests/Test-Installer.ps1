#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$installerScript = Join-Path $root 'Installer\AutoTools.iss'
$buildScript = Join-Path $root 'Scripts\Build-Installer.ps1'
$launcherScript = Join-Path $root 'Launcher.ps1'

function Assert-True {
    param(
        [Parameter(Mandatory)]
        [bool]$Condition,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Contains {
    param(
        [Parameter(Mandatory)]
        [string]$Content,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [Parameter(Mandatory)]
        [string]$Message
    )

    Assert-True -Condition ($Content -match $Pattern) -Message $Message
}

Assert-True `
    -Condition (Test-Path -LiteralPath $installerScript -PathType Leaf) `
    -Message 'Missing Installer\AutoTools.iss.'

Assert-True `
    -Condition (Test-Path -LiteralPath $buildScript -PathType Leaf) `
    -Message 'Missing Scripts\Build-Installer.ps1.'

Assert-True `
    -Condition (Test-Path -LiteralPath $launcherScript -PathType Leaf) `
    -Message 'Missing Launcher.ps1.'

$manifest = Get-Content -LiteralPath $installerScript -Raw
$builder = Get-Content -LiteralPath $buildScript -Raw
$launcher = Get-Content -LiteralPath $launcherScript -Raw

Assert-Contains $manifest '(?m)^PrivilegesRequired=lowest$' `
    'Installer must run without administrator elevation.'
Assert-Contains $manifest '(?m)^DefaultDirName=\{localappdata\}\\Programs\\Auto Tools$' `
    'Installer must use a per-user writable directory.'
Assert-Contains $manifest '(?m)^\[Files\]$' 'Installer is missing [Files].'
Assert-Contains $manifest '(?m)^\[Dirs\]$' 'Installer is missing [Dirs].'
Assert-Contains $manifest '(?m)^\[Icons\]$' 'Installer is missing [Icons].'
Assert-Contains $manifest '(?m)^\[Run\]$' 'Installer is missing [Run].'
Assert-Contains $manifest 'launcher-config\.json.*onlyifdoesntexist.*uninsneveruninstall' `
    'User launcher configuration must survive upgrades and uninstall.'
Assert-Contains $manifest 'Config\\suppliers\.csv.*onlyifdoesntexist.*uninsneveruninstall' `
    'Supplier configuration must survive upgrades and uninstall.'
Assert-Contains $manifest '\{sys\}\\WindowsPowerShell\\v1\.0\\powershell\.exe' `
    'Shortcuts must launch Auto Tools with Windows PowerShell.'

Assert-Contains $builder 'ISCC\.exe' 'Build script must locate ISCC.exe.'
Assert-Contains $builder 'Installer\\AutoTools\.iss' `
    'Build script must compile Installer\AutoTools.iss.'

$launcherVersion = [regex]::Match(
    $launcher,
    '(?m)^\$AppVersion\s*=\s*''([^'']+)''$'
)
$manifestVersion = [regex]::Match(
    $manifest,
    '(?m)^\s*#define MyAppVersion "([^"]+)"$'
)
$builderVersion = [regex]::Match(
    $builder,
    '(?m)^\s*\[string\]\$Version\s*=\s*''([^'']+)'''
)

Assert-True $launcherVersion.Success 'Launcher is missing $AppVersion.'
Assert-True $manifestVersion.Success 'Manifest is missing its default version.'
Assert-True $builderVersion.Success 'Build script is missing its default version.'
Assert-True ($launcherVersion.Groups[1].Value -eq '1.1.0') `
    'Launcher version must be 1.1.0.'
Assert-True `
    ($launcherVersion.Groups[1].Value -eq $manifestVersion.Groups[1].Value) `
    'Launcher and Inno Setup versions do not match.'
Assert-True `
    ($launcherVersion.Groups[1].Value -eq $builderVersion.Groups[1].Value) `
    'Launcher and build script versions do not match.'
Assert-Contains $launcher 'Developed by Hugo Le Chi Quoc Hung' `
    'Launcher is missing the developer credit.'
Assert-Contains `
    $launcher `
    '(?s)\$footerLabel\.Text\s*=.*?Developed by Hugo Le Chi Quoc Hung.*?\$footerLabel\.TextAlign\s*=\s*''MiddleRight''' `
    'Launcher footer must be right aligned.'

Write-Host 'Installer manifest checks passed.'
