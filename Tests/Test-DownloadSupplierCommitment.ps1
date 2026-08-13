$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Scripts\Download-SupplierCommitment.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)

if ($parseErrors.Count -gt 0) {
    throw "Download-SupplierCommitment.ps1 has PowerShell parse errors: $($parseErrors -join '; ')"
}

foreach ($functionName in @(
    'Parse-SupplierCommitmentSubject',
    'New-SupplierLookup',
    'Get-OutlookFolderTree',
    'Normalize-SupplierKey',
    'ConvertTo-OutlookRecipientList',
    'Get-SupplierReminderCandidates',
    'Get-SupplierReminderSubject',
    'New-SupplierReminderHtml',
    'Write-Log',
    'New-SupplierReminderDrafts'
)) {
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq $functionName
    }, $true)

    if ($null -eq $definition) {
        throw "Function '$functionName' was not found in Download-SupplierCommitment.ps1."
    }

    . ([scriptblock]::Create($definition.Extent.Text))
}

function Assert-Equal($actual, $expected, $message) {
    if ($actual -ne $expected) {
        throw "$message Expected '$expected', got '$actual'."
    }
}

function Assert-Null($actual, $message) {
    if ($null -ne $actual) {
        throw "$message Expected null, got '$actual'."
    }
}

$tc5 = Parse-SupplierCommitmentSubject 'TTI ISPIH Alert -- MRP Supplier Commitment(SF_TC5--07/26/2026 14:09:15)(12258)'
Assert-Equal $tc5.BU 'TC5' 'TC5 subject must map to BU TC5.'
Assert-Equal $tc5.VendorCode '12258' 'TC5 subject must expose the supplier code from the final parentheses.'

$tn5 = Parse-SupplierCommitmentSubject 'tti ispih alert -- mrp supplier commitment(SF_TN5--07/27/2026 08:01:02)(10603)'
Assert-Equal $tn5.BU 'TN5' 'Subject parsing must be case-insensitive.'
Assert-Equal $tn5.VendorCode '10603' 'TN5 subject must expose the supplier code.'

Assert-Null (Parse-SupplierCommitmentSubject 'RE: Supplier Commitment TC5 (12258)') 'Unrelated subjects must not match.'
Assert-Null (Parse-SupplierCommitmentSubject 'TTI ISPIH Alert -- MRP Supplier Commitment(SF_TC5--07/26/2026 14:09:15)') 'A subject without the final supplier code must not match.'

$supplierLookup = New-SupplierLookup @(
    [PSCustomObject]@{ Keyword = 'ABC'; VendorCode = ' 12258 '; VendorName = ' Vendor A ' },
    [PSCustomObject]@{ Keyword = 'XYZ'; VendorCode = '10603'; VendorName = 'Vendor B' },
    [PSCustomObject]@{ Keyword = 'EMPTY'; VendorCode = ''; VendorName = 'Ignored Vendor' }
)

Assert-Equal $supplierLookup['12258'].Name 'Vendor A' 'Supplier lookup must map column B VendorCode to column C VendorName.'
Assert-Equal $supplierLookup['10603'].Display '10603 - Vendor B' 'Supplier lookup must build the download display name.'
Assert-Equal $supplierLookup.ContainsKey('') $false 'Rows without a supplier code must be ignored.'

$supplierConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Config\suppliers.csv'
$configuredSupplierLookup = New-SupplierLookup @(Import-Csv -LiteralPath $supplierConfigPath)
Assert-Equal $configuredSupplierLookup[$tc5.VendorCode].Name 'NGHIA LONG METAL PRODUCTS COMPANY LIMITED' 'The sample code 12258 must resolve through Config\suppliers.csv.'

$leafFolder = [PSCustomObject]@{ Name = 'SF renamed'; Folders = @() }
$buFolder = [PSCustomObject]@{ Name = 'Custom BU folder'; Folders = @($leafFolder) }
$rootFolder = [PSCustomObject]@{ Name = 'Supplier Commitment'; Folders = @($buFolder) }
$folderTree = @(Get-OutlookFolderTree $rootFolder)

Assert-Equal $folderTree.Count 3 'Fallback folder scan must include the Supplier Commitment root and all descendants.'
Assert-Equal (($folderTree.Name | Sort-Object) -join '|') 'Custom BU folder|SF renamed|Supplier Commitment' 'Fallback folder scan must find renamed nested folders.'

$nbsp = [char]0x00A0
$recipientList = ConvertTo-OutlookRecipientList " first@example.com;SECOND@example.com; first@example.com;${nbsp}third@example.com "
Assert-Equal $recipientList 'first@example.com; SECOND@example.com; third@example.com' 'Recipient lists must be trimmed, de-duplicated, and Outlook-compatible.'

$missingRows = @(
    [PSCustomObject]@{ BU = 'TN5'; VendorCode = '100'; VendorName = 'A & B SUPPLIER' },
    [PSCustomObject]@{ BU = 'TC5'; VendorCode = '100'; VendorName = 'A & B SUPPLIER' },
    [PSCustomObject]@{ BU = 'TC5'; VendorCode = '200'; VendorName = 'NO EMAIL SUPPLIER' }
)
$supplierRows = @(
    [PSCustomObject]@{ VendorCode = '100'; VendorName = 'A & B SUPPLIER'; 'Email To' = 'buyer@example.com; buyer@example.com'; 'Email CC' = 'planner@example.com' },
    [PSCustomObject]@{ VendorCode = '200'; VendorName = 'NO EMAIL SUPPLIER'; 'Email To' = ''; 'Email CC' = 'planner@example.com' }
)
$reminderCandidates = @(Get-SupplierReminderCandidates -MissingRows $missingRows -Suppliers $supplierRows)
Assert-Equal $reminderCandidates.Count 2 'Missing rows must be grouped into one reminder candidate per supplier.'

$supplier100 = $reminderCandidates | Where-Object { $_.VendorCode -eq '100' } | Select-Object -First 1
Assert-Equal $supplier100.MissingBUs 'TC5, TN5' 'A reminder must list every missing BU for the supplier.'
Assert-Equal $supplier100.To 'buyer@example.com' 'Duplicate supplier recipients must be removed.'
Assert-Equal $supplier100.CanCreateDraft $true 'A supplier with Email To must be eligible for draft creation.'

$supplier200 = $reminderCandidates | Where-Object { $_.VendorCode -eq '200' } | Select-Object -First 1
Assert-Equal $supplier200.CanCreateDraft $false 'A supplier without Email To must not be eligible for draft creation.'

$weekStart = [datetime]'2026-07-26'
$weekEnd = [datetime]'2026-08-01'
$subject = Get-SupplierReminderSubject -Candidate $supplier100 -WeekStart $weekStart -WeekEnd $weekEnd
Assert-Equal $subject 'Reminder: Supplier Commitment - TC5, TN5 - 26/07/2026-01/08/2026' 'Reminder subject must identify the missing BUs and week.'

$html = New-SupplierReminderHtml -Candidate $supplier100 -WeekStart $weekStart -WeekEnd $weekEnd
if ($html -notmatch 'A &amp; B SUPPLIER') {
    throw 'Reminder HTML must encode the supplier name.'
}
if ($html -notmatch '<li>TC5</li>' -or $html -notmatch '<li>TN5</li>') {
    throw 'Reminder HTML must list every missing BU.'
}

$script:fakeDisplayCount = 0
$script:fakeSaveCount = 0
$fakeMail = [PSCustomObject]@{
    To       = ''
    CC       = ''
    Subject  = ''
    HTMLBody = '<p>Default Outlook signature</p>'
}
$fakeMail | Add-Member -MemberType ScriptMethod -Name Display -Value { $script:fakeDisplayCount++ }
$fakeMail | Add-Member -MemberType ScriptMethod -Name Save -Value { $script:fakeSaveCount++ }
$fakeOutlook = [PSCustomObject]@{ Mail = $fakeMail }
$fakeOutlook | Add-Member -MemberType ScriptMethod -Name CreateItem -Value { param($itemType) return $this.Mail }

$draftCount = New-SupplierReminderDrafts `
    -Outlook $fakeOutlook `
    -Candidates @($supplier100) `
    -WeekStart $weekStart `
    -WeekEnd $weekEnd `
    -SignatureDelayMilliseconds 0
Assert-Equal $draftCount 1 'Exactly one draft must be created for one selected supplier.'
Assert-Equal $fakeMail.To 'buyer@example.com' 'Draft To recipients must come from suppliers.csv.'
Assert-Equal $fakeMail.CC 'planner@example.com' 'Draft CC recipients must come from suppliers.csv.'
Assert-Equal $script:fakeSaveCount 1 'The reminder email must be saved as a draft.'
Assert-Equal $script:fakeDisplayCount 2 'The reminder draft must remain displayed for user review.'
if ($fakeMail.HTMLBody -notmatch 'Default Outlook signature') {
    throw 'Reminder draft must preserve the default Outlook signature.'
}

Write-Host 'PASS: Download-SupplierCommitment parsing, lookup, and reminder tests.' -ForegroundColor Green
