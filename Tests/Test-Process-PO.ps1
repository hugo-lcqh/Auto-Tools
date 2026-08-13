#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$ProjectRoot = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
}

$scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Process-PO\Process-PO.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)

if ($parseErrors.Count -gt 0) {
    throw "Process-PO.ps1 has PowerShell parse errors: $($parseErrors -join '; ')"
}

foreach ($functionName in @(
    'Select-POGroupingMode',
    'Group-PORecords',
    'Get-POOutputRelativePath'
)) {
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq $functionName
    }, $true)

    if ($null -eq $definition) {
        throw "Function '$functionName' was not found in Process-PO.ps1."
    }

    . ([scriptblock]::Create($definition.Extent.Text))
}

function Assert-Equal($actual, $expected, $message) {
    if ($actual -ne $expected) {
        throw "$message Expected '$expected', got '$actual'."
    }
}

$records = @(
    [PSCustomObject]@{ BU = 'TC5'; VENDOR_SITE = 'TVC-100-B'; ITEM_NO = '10001' },
    [PSCustomObject]@{ BU = 'TC5'; VENDOR_SITE = 'TVC-200-B'; ITEM_NO = '20001' },
    [PSCustomObject]@{ BU = 'TN5'; VENDOR_SITE = 'TXV-100-B'; ITEM_NO = '30001' }
)

$supplierGroups = @(Group-PORecords -Records $records -GroupingMode 'Supplier')
Assert-Equal $supplierGroups.Count 3 'Supplier mode must create one group per BU and supplier site.'

$siteGroups = @(Group-PORecords -Records $records -GroupingMode 'Site')
Assert-Equal $siteGroups.Count 2 'Site mode must create one group per BU only.'
$tc5Group = $siteGroups | Where-Object { $_.Group[0].BU -eq 'TC5' } | Select-Object -First 1
Assert-Equal $tc5Group.Count 2 'Site mode must keep different TC5 suppliers in the same group.'

$supplierPath = Get-POOutputRelativePath `
    -GroupingMode 'Supplier' `
    -BU 'TC5' `
    -VendorSite 'TVC-100-B' `
    -VendorFolderName 'Vendor A'
Assert-Equal $supplierPath 'TC5\Vendor A\TVC-100-B.txt' `
    'Supplier mode must preserve the existing supplier folder and file naming.'

$sitePath = Get-POOutputRelativePath `
    -GroupingMode 'Site' `
    -BU 'TC5' `
    -VendorSite 'TVC-100-B' `
    -VendorFolderName 'Vendor A'
Assert-Equal $sitePath 'TC5\TC5.txt' `
    'Site mode must create one file directly inside the BU folder.'

Assert-Equal (Select-POGroupingMode -RequestedMode 'Supplier') 'Supplier' `
    'The Supplier parameter value must select the current behavior.'
Assert-Equal (Select-POGroupingMode -RequestedMode 'Site') 'Site' `
    'The Site parameter value must select the new behavior.'

$script:promptAnswers = New-Object 'System.Collections.Generic.Queue[string]'
$script:promptAnswers.Enqueue('invalid')
$script:promptAnswers.Enqueue('2')
$script:promptCount = 0
function Read-Host {
    param([string]$Prompt)
    $script:promptCount++
    return $script:promptAnswers.Dequeue()
}

Assert-Equal (Select-POGroupingMode -RequestedMode '') 'Site' `
    'Interactive choice 2 must select site-only grouping.'
Assert-Equal $script:promptCount 2 `
    'An invalid interactive choice must be rejected and prompted again.'

Write-Host 'PASS: Process-PO grouping mode tests.' -ForegroundColor Green
