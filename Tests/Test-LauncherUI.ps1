#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$launcherPath = Join-Path $root 'Launcher.ps1'
$configPath = Join-Path $root 'launcher-config.json'

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

$tokens = $null
$parseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    $launcherPath,
    [ref]$tokens,
    [ref]$parseErrors
)

Assert-True `
    -Condition ($parseErrors.Count -eq 0) `
    -Message "Launcher.ps1 has parse errors: $($parseErrors -join '; ')"

$launcher = Get-Content -LiteralPath $launcherPath -Raw
$configData = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
$programs = @($configData)

Assert-True ($programs.Count -gt 0) 'launcher-config.json must contain at least one tool.'
foreach ($program in $programs) {
    Assert-True `
        -Condition (-not [string]::IsNullOrWhiteSpace([string]$program.Name)) `
        -Message 'Every configured tool must have a Name.'
    Assert-True `
        -Condition (-not [string]::IsNullOrWhiteSpace([string]$program.Path)) `
        -Message 'Every configured tool must have a Path.'

    $resolvedPath = if ([System.IO.Path]::IsPathRooted([string]$program.Path)) {
        [System.IO.Path]::GetFullPath([string]$program.Path)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path $root ([string]$program.Path)))
    }

    Assert-True `
        -Condition (Test-Path -LiteralPath $resolvedPath -PathType Leaf) `
        -Message "Configured tool does not exist: $($program.Path)"
}

Assert-Contains $launcher '\$form\.AutoScaleMode\s*=\s*\[System\.Windows\.Forms\.AutoScaleMode\]::Dpi' `
    'Launcher must scale with Windows display DPI.'
Assert-Contains $launcher 'New-Object System\.Windows\.Forms\.TableLayoutPanel' `
    'Launcher must use layout containers instead of fixed-position page layout.'
Assert-Contains $launcher 'New-Object System\.Windows\.Forms\.FlowLayoutPanel' `
    'Launcher must use a responsive scrolling tool list.'
Assert-Contains $launcher "'READY'" `
    'Launcher must show a text readiness state for available tools.'
Assert-Contains $launcher "'MISSING'" `
    'Launcher must show a text readiness state for missing tools.'
Assert-Contains $launcher '\$runButton\.AccessibleName\s*=' `
    'Run actions must expose an accessible name.'
Assert-Contains $launcher '\$deleteButton\.AccessibleName\s*=' `
    'Remove actions must expose an accessible name.'
Assert-Contains $launcher '\$btnAdd\.AccessibleName\s*=' `
    'Add Tool must expose an accessible name.'
Assert-Contains $launcher '\$btnEdit\.AccessibleName\s*=' `
    'Open Config must expose an accessible name.'
Assert-Contains $launcher 'Developed by Hugo Le Chi Quoc Hung' `
    'Launcher must retain the developer credit.'

Write-Host 'PASS: Launcher UI and configuration contract checks.' -ForegroundColor Green
