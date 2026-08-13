#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scriptPath = Join-Path $root 'Scripts\AutoImport.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)

if ($parseErrors.Count -gt 0) {
    throw "AutoImport.ps1 has PowerShell parse errors: $($parseErrors -join '; ')"
}

$definition = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Move-JobShortSourceToArchive'
}, $true)

if ($null -eq $definition) {
    throw "Function 'Move-JobShortSourceToArchive' was not found in AutoImport.ps1."
}

. ([scriptblock]::Create($definition.Extent.Text))

$content = Get-Content -LiteralPath $scriptPath -Raw
if ($content -match '[A-Za-z]:\\Users\\[^\\]+\\') {
    throw 'AutoImport.ps1 must not contain a user-specific absolute path.'
}

$saveIndex = $content.IndexOf('$wb.Save()')
$archiveIndex = $content.IndexOf('$archivedPath = Move-JobShortSourceToArchive')
if ($saveIndex -lt 0 -or $archiveIndex -lt 0 -or $archiveIndex -lt $saveIndex) {
    throw 'The source folder must be archived only after the workbook is saved successfully.'
}

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoTools_AutoImport_{0}" -f [Guid]::NewGuid())
try {
    $sourceFolder = New-Item -ItemType Directory -Path (Join-Path $testRoot '13.08') -Force
    Set-Content -LiteralPath (Join-Path $sourceFolder.FullName 'sample.txt') -Value 'sample'

    $archivedPath = Move-JobShortSourceToArchive -SourceFolder $sourceFolder -RootFolder $testRoot

    if (Test-Path -LiteralPath $sourceFolder.FullName) {
        throw 'The source folder must be removed from the active Job Short root.'
    }
    if (-not (Test-Path -LiteralPath $archivedPath -PathType Container)) {
        throw 'The source folder must be moved into Older.'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $archivedPath 'sample.txt') -PathType Leaf)) {
        throw 'Archiving must preserve the source folder contents.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

Write-Host 'PASS: AutoImport archive behavior and portability checks.' -ForegroundColor Green
