# =====================================================================
#  Check MR Pull-In
#  - Input: Part, MR Qty, Requested Date, ASCP updated Yes/No
#  - ASCP not updated: remaining FG after Supply Reply vs MR Qty
#  - ASCP updated: previous-week Supply Reply vs Plan Order, with FG consumed by Supply Reply
# =====================================================================

[CmdletBinding()]
param(
    [string]$PartNum,
    [double]$MRQty,
    [string]$RequestedDate,
    [string]$ASCPUpdated,
    [string]$Site,
    [string]$VendorCode,
    [string]$Supplier,
    [switch]$NoPause
)

. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")

$ErrorActionPreference = "Stop"
$AutoToolsPaths = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

$SupplierCommitmentRoot = $AutoToolsConfig.SupplierCommitmentOutput
$VendorStockRoot = Join-Path -Path $AutoToolsPaths.Output -ChildPath "Vendor Stock"
$ReportRoot = Join-Path -Path $AutoToolsPaths.Output -ChildPath "MR Check"

if (-not (Test-Path -LiteralPath $ReportRoot -PathType Container)) {
    $null = New-Item -Path $ReportRoot -ItemType Directory -Force
}

$Script:HtmlTableCache = @{}
$Script:VendorStockFileCache = @{}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Color = "White"
    )

    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message) -ForegroundColor $Color
}

function ConvertTo-Num {
    param($Value)

    if ($null -eq $Value) { return 0.0 }
    if ($Value -is [array]) {
        if ($Value.Length -eq 0) { return 0.0 }
        $Value = $Value[0]
        if ($null -eq $Value) { return 0.0 }
    }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        return [double]$Value
    }

    $text = ([string]$Value).Trim()
    if ($text -eq "") { return 0.0 }
    $text = $text -replace ',', ''

    $number = 0.0
    if ([double]::TryParse(
        $text,
        [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )) {
        return $number
    }

    return 0.0
}

function ConvertTo-Str {
    param($Value)

    if ($null -eq $Value) { return "" }
    if ($Value -is [array]) {
        if ($Value.Length -eq 0) { return "" }
        $Value = $Value[0]
        if ($null -eq $Value) { return "" }
    }

    return ([string]$Value).Trim()
}

function ConvertTo-PartKey {
    param([string]$Value)

    return (ConvertTo-Str $Value).ToUpperInvariant()
}

function Read-RequiredInput {
    param(
        [string]$Label,
        [string]$CurrentValue
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return $CurrentValue.Trim()
    }

    do {
        $answer = Read-Host $Label
    } while ([string]::IsNullOrWhiteSpace($answer))

    return $answer.Trim()
}

function Read-OptionalInput {
    param(
        [string]$Label,
        [string]$CurrentValue
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return $CurrentValue.Trim()
    }

    return (Read-Host $Label).Trim()
}

function Read-YesNo {
    param(
        [string]$Label,
        [bool]$CurrentValue,
        [bool]$HasCurrentValue
    )

    if ($HasCurrentValue) {
        return $CurrentValue
    }

    do {
        $answer = (Read-Host "$Label (Y/N)").Trim()
    } while ($answer -notmatch '^(?i)y(es)?|n(o)?$')

    return ($answer -match '^(?i)y')
}

function ConvertTo-BoolInput {
    param([string]$Value)

    $text = (ConvertTo-Str $Value).ToUpperInvariant()
    if ($text -in @("Y", "YES", "TRUE", "1")) {
        return $true
    }
    if ($text -in @("N", "NO", "FALSE", "0")) {
        return $false
    }

    throw "Gia tri ASCPUpdated khong hop le. Dung Yes/No hoac True/False."
}

function ConvertTo-UserDate {
    param($Value)

    if ($Value -is [datetime] -and $Value -ne [datetime]::MinValue) {
        return ([datetime]$Value).Date
    }

    $currentValue = ""
    if ($null -ne $Value -and -not [string]::IsNullOrWhiteSpace([string]$Value)) {
        $currentValue = [string]$Value
    }

    $text = (Read-RequiredInput -Label "Requested pull-in date (yyyy-mm-dd hoac dd/mm/yyyy)" -CurrentValue $currentValue).Trim()

    if ($text -match '^(?<year>\d{4})-(?<month>\d{1,2})-(?<day>\d{1,2})$') {
        return [datetime]::new([int]$matches.year, [int]$matches.month, [int]$matches.day)
    }

    if ($text -match '^(?<day>\d{1,2})[\/.](?<month>\d{1,2})[\/.](?<year>\d{4})$') {
        return [datetime]::new([int]$matches.year, [int]$matches.month, [int]$matches.day)
    }

    $date = [datetime]::MinValue
    if ([datetime]::TryParse(
        $text,
        [System.Globalization.CultureInfo]::GetCultureInfo("vi-VN"),
        [System.Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        return $date.Date
    }

    throw "Ngay requested khong hop le: $text. Vui long nhap theo yyyy-mm-dd hoac dd/mm/yyyy, vi du 2026-07-22 hoac 22/07/2026."
}

function Get-LatestVendorStockFolder {
    if (-not (Test-Path -LiteralPath $VendorStockRoot -PathType Container)) {
        return $null
    }

    $folders = @(
        Get-ChildItem -LiteralPath $VendorStockRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object {
                @(Get-ChildItem -LiteralPath $_.FullName -File -Filter "*.xlsx" -ErrorAction SilentlyContinue).Count -gt 0 -or
                @(Get-ChildItem -LiteralPath $_.FullName -File -Filter "*.xls" -ErrorAction SilentlyContinue).Count -gt 0
            } |
            Sort-Object LastWriteTime -Descending
    )

    if ($folders.Count -eq 0) {
        return $null
    }

    return $folders[0].FullName
}

function Test-SiteMatch {
    param(
        [string]$CellSite,
        [string]$RequiredSite
    )

    if ([string]::IsNullOrWhiteSpace($RequiredSite)) {
        return $true
    }

    $target = $RequiredSite.Trim().ToUpperInvariant()
    $parts = (ConvertTo-Str $CellSite).ToUpperInvariant() -split '[;,/ ]+' |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    return ($parts -contains $target)
}

function Get-ExcelCell {
    param(
        [object]$Data,
        [int]$Row,
        [int]$Column
    )

    try {
        return $Data.GetValue($Row, $Column)
    }
    catch {
        return $null
    }
}

function Find-HeaderColumn {
    param(
        [object]$Data,
        [int]$HeaderRow,
        [int]$ColumnCount,
        [string[]]$Names
    )

    for ($c = 1; $c -le $ColumnCount; $c++) {
        $header = (ConvertTo-Str (Get-ExcelCell -Data $Data -Row $HeaderRow -Column $c)).ToUpperInvariant()
        foreach ($name in $Names) {
            if ($header -eq $name.ToUpperInvariant()) {
                return $c
            }
        }
    }

    return 0
}

function Find-HeaderIndex {
    param(
        [object[]]$Headers,
        [string[]]$Names
    )

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $header = (ConvertTo-Str $Headers[$i]).ToUpperInvariant()
        foreach ($name in $Names) {
            if ($header -eq $name.ToUpperInvariant()) {
                return $i
            }
        }
    }

    return -1
}

function Get-ArrayCell {
    param(
        [object[]]$Row,
        [int]$Index
    )

    if ($Index -lt 0 -or $Index -ge $Row.Count) {
        return $null
    }

    return $Row[$Index]
}

function Test-HtmlExcelFile {
    param([string]$Path)

    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $prefixLength = [Math]::Min($stream.Length, 32)
            $bytes = New-Object byte[] $prefixLength
            $null = $stream.Read($bytes, 0, $prefixLength)
        }
        finally {
            $stream.Close()
        }
        $prefixLength = [Math]::Min($bytes.Length, 32)
        $prefix = [System.Text.Encoding]::UTF8.GetString($bytes, 0, $prefixLength)
        $prefix = $prefix.TrimStart([char]0xFEFF, [char]0x200B, [char]0x00A0, [char]9, [char]10, [char]13, [char]32)
        return $prefix.StartsWith("<", [System.StringComparison]::Ordinal)
    }
    catch {
        return $false
    }
}

function Read-HtmlTableRows {
    param([string]$Path)

    $cacheKey = [System.IO.Path]::GetFullPath($Path).ToUpperInvariant()
    if ($Script:HtmlTableCache.ContainsKey($cacheKey)) {
        return $Script:HtmlTableCache[$cacheKey]
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            $content = $reader.ReadToEnd()
        }
        finally {
            $reader.Close()
        }
    }
    finally {
        if ($null -ne $stream) {
            $stream.Close()
        }
    }
    $rows = @()
    $rowMatches = [regex]::Matches($content, '(?is)<tr\b[^>]*>(.*?)</tr>')

    foreach ($rowMatch in $rowMatches) {
        $rowHtml = $rowMatch.Groups[1].Value
        $cellMatches = [regex]::Matches($rowHtml, '(?is)<td\b[^>]*>(.*?)(?=<td\b|</tr>|$)')
        $cells = @()

        foreach ($cellMatch in $cellMatches) {
            $cellHtml = $cellMatch.Groups[1].Value
            $text = [regex]::Replace($cellHtml, '(?is)<[^>]+>', '')
            $text = [System.Net.WebUtility]::HtmlDecode($text)
            $cells += (ConvertTo-Str $text)
        }

        if ($cells.Count -gt 0) {
            $rows += [PSCustomObject]@{
                Cells = [object[]]$cells
            }
        }
    }

    $Script:HtmlTableCache[$cacheKey] = @($rows)
    return @($rows)
}

function ConvertTo-CommitmentDate {
    param([string]$HeaderText)

    $text = (ConvertTo-Str $HeaderText).ToUpperInvariant()
    if ($text -eq "PAST DUE") {
        return [datetime]::MinValue
    }

    if ($text -match '^(?<day>\d{1,2})-(?<month>[A-Z]{3})-(?<year>\d{2,4})$') {
        $monthMap = @{
            JAN = 1; FEB = 2; MAR = 3; APR = 4; MAY = 5; JUN = 6
            JUL = 7; AUG = 8; SEP = 9; OCT = 10; NOV = 11; DEC = 12
        }

        $monthText = $matches.month
        if (-not $monthMap.ContainsKey($monthText)) {
            return $null
        }

        $year = [int]$matches.year
        if ($year -lt 100) {
            if ($year -lt 80) {
                $year += 2000
            }
            else {
                $year += 1900
            }
        }

        return [datetime]::new($year, $monthMap[$monthText], [int]$matches.day)
    }

    return $null
}

function Read-SupplierCommitmentHtml {
    param(
        [string]$Path,
        [string]$FileSite,
        [string]$RequiredPart,
        [datetime]$RequiredDate
    )

    $tableRows = @(Read-HtmlTableRows -Path $Path)
    if ($tableRows.Count -eq 0) {
        return @()
    }

    $headers = @($tableRows[0].Cells)
    $colPart = Find-HeaderIndex -Headers $headers -Names @("Part Num", "Part_No", "Part No")
    $colType = Find-HeaderIndex -Headers $headers -Names @("Type")
    $colVendorCode = Find-HeaderIndex -Headers $headers -Names @("Vendor Code")
    $colVendorName = Find-HeaderIndex -Headers $headers -Names @("Vendor Name")
    $colFG = Find-HeaderIndex -Headers $headers -Names @("FG")
    $colPastDue = Find-HeaderIndex -Headers $headers -Names @("Past Due")

    if ($colPart -lt 0 -or $colType -lt 0 -or $colPastDue -lt 0) {
        Write-Log "  Bo qua file do thieu cot Part Num/Type/Past Due." "Yellow"
        return @()
    }

    $dateColumns = @()
    for ($i = $colPastDue; $i -lt $headers.Count; $i++) {
        $headerText = ConvertTo-Str $headers[$i]
        if ($headerText -match '^(?i)total demand$') {
            break
        }

        $bucketDate = ConvertTo-CommitmentDate -HeaderText $headerText
        if ($null -ne $bucketDate -and $bucketDate -le $RequiredDate) {
            $dateColumns += [PSCustomObject]@{
                Index = $i
                Name = $headerText
                Date = $bucketDate
            }
        }
    }

    if ($dateColumns.Count -eq 0) {
        return @()
    }

    $matchedRows = @()
    $partKey = ConvertTo-PartKey $RequiredPart
    for ($r = 1; $r -lt $tableRows.Count; $r++) {
        $row = @($tableRows[$r].Cells)
        $part = ConvertTo-PartKey (Get-ArrayCell -Row $row -Index $colPart)
        if ($part -ne $partKey) {
            continue
        }

        $type = ConvertTo-Str (Get-ArrayCell -Row $row -Index $colType)
        if ($type -notin @("Plan Order", "Supply Reply")) {
            continue
        }

        $valuesByBucket = @{}
        $sum = 0.0
        foreach ($bucket in $dateColumns) {
            $value = ConvertTo-Num (Get-ArrayCell -Row $row -Index $bucket.Index)
            $valuesByBucket[$bucket.Index] = $value
            $sum += $value
        }

        $matchedRows += [PSCustomObject]@{
            Type = $type
            Quantity = $sum
            FG = if ($colFG -ge 0) { ConvertTo-Num (Get-ArrayCell -Row $row -Index $colFG) } else { 0.0 }
            ValuesByBucket = $valuesByBucket
            VendorCode = if ($colVendorCode -ge 0) { ConvertTo-Str (Get-ArrayCell -Row $row -Index $colVendorCode) } else { "" }
            VendorName = if ($colVendorName -ge 0) { ConvertTo-Str (Get-ArrayCell -Row $row -Index $colVendorName) } else { "" }
        }
    }

    if ($matchedRows.Count -eq 0) {
        return @()
    }

    $vendorCodeFromRow = ConvertTo-Str (@($matchedRows | Where-Object { $_.VendorCode } | Select-Object -First 1).VendorCode)
    $vendorNameFromRow = ConvertTo-Str (@($matchedRows | Where-Object { $_.VendorName } | Select-Object -First 1).VendorName)
    $planRows = @($matchedRows | Where-Object { $_.Type -eq "Plan Order" })
    $supplyRows = @($matchedRows | Where-Object { $_.Type -eq "Supply Reply" })
    $commitmentFG = Measure-PropertySum -Rows $supplyRows -Property "FG"
    $planTotal = 0.0
    $usableSupplyReply = 0.0
    $remainingFG = $commitmentFG
    $minRemainingFG = $remainingFG
    $minReplyGap = $null
    $prevSupply = 0.0
    $timeline = @()

    foreach ($bucket in $dateColumns) {
        $plan = 0.0
        foreach ($row in $planRows) {
            if ($row.ValuesByBucket.ContainsKey($bucket.Index)) {
                $plan += [double]$row.ValuesByBucket[$bucket.Index]
            }
        }

        $supply = 0.0
        foreach ($row in $supplyRows) {
            if ($row.ValuesByBucket.ContainsKey($bucket.Index)) {
                $supply += [double]$row.ValuesByBucket[$bucket.Index]
            }
        }

        $replyGap = $prevSupply - $plan
        if ($null -eq $minReplyGap -or $replyGap -lt $minReplyGap) {
            $minReplyGap = $replyGap
        }

        $remainingFG = $remainingFG - $prevSupply
        if ($remainingFG -lt $minRemainingFG) {
            $minRemainingFG = $remainingFG
        }

        $planTotal += $plan
        $usableSupplyReply += $prevSupply
        $timeline += [PSCustomObject]@{
            Week = $bucket.Name
            Plan = [math]::Round($plan, 4)
            PrevSupply = [math]::Round($prevSupply, 4)
            EnteredSupply = [math]::Round($supply, 4)
            ReplyGap = [math]::Round($replyGap, 4)
            RemainingFG = [math]::Round($remainingFG, 4)
        }
        $prevSupply = $supply
    }

    return @([PSCustomObject]@{
        File = $Path
        Site = $FileSite
        VendorCode = $vendorCodeFromRow
        VendorName = $vendorNameFromRow
        CommitmentFG = [math]::Round($commitmentFG, 4)
        PlanOrder = [math]::Round($planTotal, 4)
        SupplyReply = [math]::Round($usableSupplyReply, 4)
        EndingGap = [math]::Round($remainingFG, 4)
        MinGap = [math]::Round($minRemainingFG, 4)
        MinReplyGap = [math]::Round($(if ($null -eq $minReplyGap) { 0.0 } else { $minReplyGap }), 4)
        Buckets = ($dateColumns | ForEach-Object { $_.Name }) -join ", "
        SupplyBuckets = "Supply Reply uses previous bucket to cover current Plan Order"
        Timeline = @($timeline)
    })
}

function Get-CommitmentFiles {
    param(
        [string]$RequiredSite,
        [string]$RequiredVendorCode
    )

    if (-not (Test-Path -LiteralPath $SupplierCommitmentRoot -PathType Container)) {
        throw "Khong tim thay folder Supplier Commitment: $SupplierCommitmentRoot"
    }

    $siteFolders = @("TC5", "TN5")
    if (-not [string]::IsNullOrWhiteSpace($RequiredSite)) {
        $siteFolders = @($RequiredSite.Trim().ToUpperInvariant())
    }

    $files = @()
    foreach ($siteFolder in $siteFolders) {
        $folder = Join-Path -Path $SupplierCommitmentRoot -ChildPath $siteFolder
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            continue
        }

        $folderFiles = @(
            Get-ChildItem -LiteralPath $folder -File -Filter "*.xls" -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath $folder -File -Filter "*.xlsx" -ErrorAction SilentlyContinue
        ) | Where-Object { -not $_.Name.StartsWith("~$") }

        if (-not [string]::IsNullOrWhiteSpace($RequiredVendorCode)) {
            $folderFiles = @($folderFiles | Where-Object { $_.Name -like "*$RequiredVendorCode*" })
        }

        foreach ($file in $folderFiles) {
            $files += [PSCustomObject]@{
                Path = $file.FullName
                Site = $siteFolder
            }
        }
    }

    return @($files)
}

function Read-SupplierCommitment {
    param(
        [object]$Excel,
        [string]$RequiredPart,
        [string]$RequiredSite,
        [string]$RequiredVendorCode,
        [datetime]$RequiredDate
    )

    $files = @(Get-CommitmentFiles -RequiredSite $RequiredSite -RequiredVendorCode $RequiredVendorCode)
    $results = @()
    $partKey = ConvertTo-PartKey $RequiredPart

    foreach ($file in $files) {
        $wb = $null
        try {
            Write-Log "Doc Supply Commitment: $([System.IO.Path]::GetFileName($file.Path))" "Cyan"
            if (Test-HtmlExcelFile -Path $file.Path) {
                $results += @(Read-SupplierCommitmentHtml `
                    -Path $file.Path `
                    -FileSite $file.Site `
                    -RequiredPart $RequiredPart `
                    -RequiredDate $RequiredDate)
                continue
            }

            if ($null -eq $Excel) {
                throw "Can Excel COM de doc file Supplier Commitment khong phai HTML: $($file.Path)"
            }

            $wb = $Excel.Workbooks.Open($file.Path)
            $ws = $wb.Sheets.Item(1)
            $used = $ws.UsedRange
            $rows = $used.Rows.Count
            $cols = $used.Columns.Count
            $data = $used.Value2

            $colPart = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("Part Num", "Part_No", "Part No")
            $colType = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("Type")
            $colVendorCode = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("Vendor Code")
            $colVendorName = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("Vendor Name")
            $colFG = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("FG")
            $colPastDue = Find-HeaderColumn -Data $data -HeaderRow 1 -ColumnCount $cols -Names @("Past Due")

            if ($colPart -eq 0 -or $colType -eq 0 -or $colPastDue -eq 0) {
                Write-Log "  Bo qua file do thieu cot Part Num/Type/Past Due." "Yellow"
                continue
            }

            $dateColumns = @()
            for ($c = $colPastDue; $c -le $cols; $c++) {
                $headerText = ConvertTo-Str $ws.Cells.Item(1, $c).Text
                if ($headerText -match '^(?i)total demand$') {
                    break
                }

                $bucketDate = ConvertTo-CommitmentDate -HeaderText $headerText
                if ($null -ne $bucketDate -and $bucketDate -le $RequiredDate) {
                    $dateColumns += [PSCustomObject]@{
                        Column = $c
                        Name = $headerText
                        Date = $bucketDate
                    }
                }
            }

            if ($dateColumns.Count -eq 0) {
                continue
            }

            $matchedRows = @()
            for ($r = 2; $r -le $rows; $r++) {
                $part = ConvertTo-PartKey (Get-ExcelCell -Data $data -Row $r -Column $colPart)
                if ($part -ne $partKey) {
                    continue
                }

                $type = ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colType)
                if ($type -notin @("Plan Order", "Supply Reply")) {
                    continue
                }

                $valuesByBucket = @{}
                $sum = 0.0
                foreach ($bucket in $dateColumns) {
                    $value = ConvertTo-Num (Get-ExcelCell -Data $data -Row $r -Column $bucket.Column)
                    $valuesByBucket[$bucket.Column] = $value
                    $sum += $value
                }

                $matchedRows += [PSCustomObject]@{
                    Type = $type
                    Quantity = $sum
                    FG = if ($colFG -gt 0) { ConvertTo-Num (Get-ExcelCell -Data $data -Row $r -Column $colFG) } else { 0.0 }
                    ValuesByBucket = $valuesByBucket
                    VendorCode = if ($colVendorCode -gt 0) { ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colVendorCode) } else { "" }
                    VendorName = if ($colVendorName -gt 0) { ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colVendorName) } else { "" }
                }
            }

            if ($matchedRows.Count -gt 0) {
                $vendorCodeFromRow = ConvertTo-Str (@($matchedRows | Where-Object { $_.VendorCode } | Select-Object -First 1).VendorCode)
                $vendorNameFromRow = ConvertTo-Str (@($matchedRows | Where-Object { $_.VendorName } | Select-Object -First 1).VendorName)
                $planRows = @($matchedRows | Where-Object { $_.Type -eq "Plan Order" })
                $supplyRows = @($matchedRows | Where-Object { $_.Type -eq "Supply Reply" })
                $commitmentFG = Measure-PropertySum -Rows $supplyRows -Property "FG"
                $planTotal = 0.0
                $usableSupplyReply = 0.0
                $remainingFG = $commitmentFG
                $minRemainingFG = $remainingFG
                $minReplyGap = $null
                $prevSupply = 0.0
                $timeline = @()

                foreach ($bucket in $dateColumns) {
                    $plan = 0.0
                    foreach ($row in $planRows) {
                        if ($row.ValuesByBucket.ContainsKey($bucket.Column)) {
                            $plan += [double]$row.ValuesByBucket[$bucket.Column]
                        }
                    }

                    $supply = 0.0
                    foreach ($row in $supplyRows) {
                        if ($row.ValuesByBucket.ContainsKey($bucket.Column)) {
                            $supply += [double]$row.ValuesByBucket[$bucket.Column]
                        }
                    }

                    $replyGap = $prevSupply - $plan
                    if ($null -eq $minReplyGap -or $replyGap -lt $minReplyGap) {
                        $minReplyGap = $replyGap
                    }

                    $remainingFG = $remainingFG - $prevSupply
                    if ($remainingFG -lt $minRemainingFG) {
                        $minRemainingFG = $remainingFG
                    }

                    $planTotal += $plan
                    $usableSupplyReply += $prevSupply
                    $timeline += [PSCustomObject]@{
                        Week = $bucket.Name
                        Plan = [math]::Round($plan, 4)
                        PrevSupply = [math]::Round($prevSupply, 4)
                        EnteredSupply = [math]::Round($supply, 4)
                        ReplyGap = [math]::Round($replyGap, 4)
                        RemainingFG = [math]::Round($remainingFG, 4)
                    }
                    $prevSupply = $supply
                }

                $results += [PSCustomObject]@{
                    File = $file.Path
                    Site = $file.Site
                    VendorCode = $vendorCodeFromRow
                    VendorName = $vendorNameFromRow
                    CommitmentFG = [math]::Round($commitmentFG, 4)
                    PlanOrder = [math]::Round($planTotal, 4)
                    SupplyReply = [math]::Round($usableSupplyReply, 4)
                    EndingGap = [math]::Round($remainingFG, 4)
                    MinGap = [math]::Round($minRemainingFG, 4)
                    MinReplyGap = [math]::Round($(if ($null -eq $minReplyGap) { 0.0 } else { $minReplyGap }), 4)
                    Buckets = ($dateColumns | ForEach-Object { $_.Name }) -join ", "
                    SupplyBuckets = "Supply Reply uses previous bucket to cover current Plan Order"
                    Timeline = @($timeline)
                }
            }
        }
        catch {
            $lineNumber = if ($_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "unknown" }
            Write-Log "  Loi doc file dong ${lineNumber}: $($_.Exception.Message)" "Red"
        }
        finally {
            if ($null -ne $wb) {
                try { $wb.Close($false) } catch {}
            }
        }
    }

    return @($results)
}

function Get-VendorStockFileRows {
    param(
        [object]$Excel,
        [object]$File
    )

    $cacheKey = [System.IO.Path]::GetFullPath($File.FullName).ToUpperInvariant()
    if ($Script:VendorStockFileCache.ContainsKey($cacheKey)) {
        return @($Script:VendorStockFileCache[$cacheKey])
    }

    $wb = $null
    $rowsOut = @()
    try {
        Write-Log "Doc Vendor Stock: $($File.Name)" "Cyan"
        $wb = $Excel.Workbooks.Open($File.FullName)
        $ws = $wb.Sheets.Item(1)
        $used = $ws.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count
        $data = $used.Value2

        $headerRow = 0
        $colPart = 0
        $colSite = 0
        $colVendorCode = 0
        $colVendor = 0
        $colFG = 0

        for ($r = 1; $r -le [Math]::Min($rows, 12); $r++) {
            $tryPart = Find-HeaderColumn -Data $data -HeaderRow $r -ColumnCount $cols -Names @("Part_No", "Part Num", "Part No")
            $tryFG = Find-HeaderColumn -Data $data -HeaderRow $r -ColumnCount $cols -Names @("FG")
            if ($tryPart -gt 0 -and $tryFG -gt 0) {
                $headerRow = $r
                $colPart = $tryPart
                $colFG = $tryFG
                $colSite = Find-HeaderColumn -Data $data -HeaderRow $r -ColumnCount $cols -Names @("Site")
                $colVendorCode = Find-HeaderColumn -Data $data -HeaderRow $r -ColumnCount $cols -Names @("Vendor code", "Vendor Code")
                $colVendor = Find-HeaderColumn -Data $data -HeaderRow $r -ColumnCount $cols -Names @("Vendor", "Vendor Name")
                break
            }
        }

        if ($headerRow -eq 0) {
            Write-Log "  Bo qua file do khong tim thay header Part_No/FG." "Yellow"
            $Script:VendorStockFileCache[$cacheKey] = @()
            return @()
        }

        for ($r = $headerRow + 1; $r -le $rows; $r++) {
            $part = ConvertTo-PartKey (Get-ExcelCell -Data $data -Row $r -Column $colPart)
            if ([string]::IsNullOrWhiteSpace($part)) {
                continue
            }

            $rowSite = if ($colSite -gt 0) { ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colSite) } else { "" }
            $rowVendorCode = if ($colVendorCode -gt 0) { ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colVendorCode) } else { "" }
            $fg = [math]::Round((ConvertTo-Num (Get-ExcelCell -Data $data -Row $r -Column $colFG)), 4)

            $rowsOut += [PSCustomObject]@{
                File = $File.FullName
                PartKey = $part
                Site = $rowSite
                VendorCode = $rowVendorCode
                VendorName = if ($colVendor -gt 0) { ConvertTo-Str (Get-ExcelCell -Data $data -Row $r -Column $colVendor) } else { "" }
                FG = $fg
            }
        }
    }
    catch {
        Write-Log "  Loi doc file: $($_.Exception.Message)" "Red"
    }
    finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
        }
    }

    $Script:VendorStockFileCache[$cacheKey] = @($rowsOut)
    return @($rowsOut)
}

function Get-VendorStockFiles {
    param([string]$RequiredVendorCode)

    $folder = Get-LatestVendorStockFolder
    if ($null -eq $folder) {
        Write-Log "Khong tim thay folder Vendor Stock co file Excel." "Yellow"
        return @()
    }

    $files = @(
        Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue |
            Where-Object {
                -not $_.Name.StartsWith("~$") -and
                $_.Extension -in @(".xlsx", ".xls")
            }
    )

    if ([string]::IsNullOrWhiteSpace($RequiredVendorCode)) {
        return @($files)
    }

    $candidateFiles = @()
    if (Test-Path -LiteralPath $AutoToolsConfig.SupplierMaster -PathType Leaf) {
        $supplierRows = @(
            Import-Csv -LiteralPath $AutoToolsConfig.SupplierMaster |
                Where-Object {
                    [string]::Equals(
                        (ConvertTo-Str $_.VendorCode),
                        $RequiredVendorCode.Trim(),
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        )

        $tokens = @()
        foreach ($supplier in $supplierRows) {
            $tokens += ((ConvertTo-Str $supplier.Keyword) -split ';')
            $tokens += (ConvertTo-Str $supplier.VendorName)
        }

        $tokens = @(
            $tokens |
                ForEach-Object { Normalize-MatchText $_ } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )

        if ($tokens.Count -gt 0) {
            $candidateFiles = @(
                $files |
                    Where-Object {
                        $baseName = Normalize-MatchText $_.BaseName
                        $matched = $false
                        foreach ($token in $tokens) {
                            if ($baseName.Contains($token)) {
                                $matched = $true
                                break
                            }
                        }
                        $matched
                    }
            )
        }
    }

    if ($candidateFiles.Count -gt 0) {
        if ($candidateFiles.Count -gt 1) {
            $selectedFile = @($candidateFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
            Write-Log ("Tim thay {0} file Vendor Stock khop supplier, dung file moi nhat: {1}" -f $candidateFiles.Count, $selectedFile[0].Name) "Yellow"
            return @($selectedFile)
        }

        return @($candidateFiles)
    }

    Write-Log "Khong tim thay file Vendor Stock khop vendor code $RequiredVendorCode. Bo qua Vendor Stock de tranh quet tat ca supplier." "Yellow"
    return @()
}

function Read-VendorStock {
    param(
        [object]$Excel,
        [string]$RequiredPart,
        [string]$RequiredSite,
        [string]$RequiredVendorCode
    )

    $results = @()
    $seenRows = @{}
    $partKey = ConvertTo-PartKey $RequiredPart
    $files = @(Get-VendorStockFiles -RequiredVendorCode $RequiredVendorCode)

    foreach ($file in $files) {
        $fileRows = @(Get-VendorStockFileRows -Excel $Excel -File $file)
        foreach ($row in $fileRows) {
            if ($row.PartKey -ne $partKey) {
                continue
            }

            if (-not (Test-SiteMatch -CellSite $row.Site -RequiredSite $RequiredSite)) {
                continue
            }

            if (
                -not [string]::IsNullOrWhiteSpace($RequiredVendorCode) -and
                -not [string]::Equals($row.VendorCode, $RequiredVendorCode.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
            ) {
                continue
            }

            $dedupeKey = "{0}|{1}|{2}|{3}" -f $row.Site.ToUpperInvariant(), $row.VendorCode.ToUpperInvariant(), $row.PartKey, $row.FG
            if ($seenRows.ContainsKey($dedupeKey)) {
                continue
            }
            $seenRows[$dedupeKey] = $true

            $results += [PSCustomObject]@{
                File = $row.File
                Site = $row.Site
                VendorCode = $row.VendorCode
                VendorName = $row.VendorName
                FG = $row.FG
            }
        }
    }

    return @($results)
}

function Format-Qty {
    param([double]$Value)

    return $Value.ToString("#,##0.####", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-FileNameSafe {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return "N/A"
    }

    return [System.IO.Path]::GetFileName($Path)
}

function Measure-PropertySum {
    param(
        [object[]]$Rows,
        [string]$Property
    )

    $sum = 0.0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $propertyInfo = $row.PSObject.Properties[$Property]
        if ($null -eq $propertyInfo -or $null -eq $propertyInfo.Value) { continue }
        $sum += [double]$propertyInfo.Value
    }

    return $sum
}

function Measure-PropertyMinimum {
    param(
        [object[]]$Rows,
        [string]$Property,
        [double]$Default = 0.0
    )

    $hasValue = $false
    $minimum = 0.0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $propertyInfo = $row.PSObject.Properties[$Property]
        if ($null -eq $propertyInfo -or $null -eq $propertyInfo.Value) { continue }

        $value = [double]$propertyInfo.Value
        if (-not $hasValue -or $value -lt $minimum) {
            $minimum = $value
            $hasValue = $true
        }
    }

    if ($hasValue) { return $minimum }
    return $Default
}

function New-KeyValueTable {
    param([object[]]$Rows)

    if ($Rows.Count -eq 0) {
        return @()
    }

    $metricWidth = 7
    $valueWidth = 6
    foreach ($row in $Rows) {
        $metricWidth = [Math]::Max($metricWidth, (ConvertTo-Str $row.Metric).Length)
        $valueWidth = [Math]::Max($valueWidth, (ConvertTo-Str $row.Value).Length)
    }

    $border = "+-{0}-+-{1}-+" -f ("-" * $metricWidth), ("-" * $valueWidth)
    $lines = @()
    $lines += $border
    $lines += "| {0} | {1} |" -f ("Thong so".PadRight($metricWidth)), ("Gia tri".PadRight($valueWidth))
    $lines += $border
    foreach ($row in $Rows) {
        $lines += "| {0} | {1} |" -f ((ConvertTo-Str $row.Metric).PadRight($metricWidth)), ((ConvertTo-Str $row.Value).PadRight($valueWidth))
    }
    $lines += $border

    return @($lines)
}

function Add-ReportLines {
    param(
        [System.Collections.Generic.List[string]]$Lines,
        [string[]]$NewLines
    )

    foreach ($line in $NewLines) {
        $Lines.Add($line)
    }
}

function New-TimelineTable {
    param([object[]]$Rows)

    if ($Rows.Count -eq 0) {
        return @()
    }

    $tableRows = @(
        $Rows | ForEach-Object {
            [PSCustomObject]@{
                Tuan = ConvertTo-Str $_.Week
                Plan = Format-Qty ([double]$_.Plan)
                SupplyTruoc = Format-Qty ([double]$_.PrevSupply)
                SupplyNhap = Format-Qty ([double]$_.EnteredSupply)
                ReplyGap = Format-Qty ([double]$_.ReplyGap)
                RemainingFG = Format-Qty ([double]$_.RemainingFG)
            }
        }
    )

    $metricRows = @(
        [PSCustomObject]@{ Metric = "Plan Order"; Values = @($tableRows | ForEach-Object { $_.Plan }) }
        [PSCustomObject]@{ Metric = "Supply truoc"; Values = @($tableRows | ForEach-Object { $_.SupplyTruoc }) }
        [PSCustomObject]@{ Metric = "Supply nhap"; Values = @($tableRows | ForEach-Object { $_.SupplyNhap }) }
        [PSCustomObject]@{ Metric = "Chenh Reply"; Values = @($tableRows | ForEach-Object { $_.ReplyGap }) }
        [PSCustomObject]@{ Metric = "FG con lai"; Values = @($tableRows | ForEach-Object { $_.RemainingFG }) }
    )

    $widths = @(13)
    foreach ($metric in $metricRows) {
        $widths[0] = [Math]::Max($widths[0], (ConvertTo-Str $metric.Metric).Length)
    }
    for ($i = 0; $i -lt $tableRows.Count; $i++) {
        $width = (ConvertTo-Str $tableRows[$i].Tuan).Length
        foreach ($metric in $metricRows) {
            $width = [Math]::Max($width, (ConvertTo-Str $metric.Values[$i]).Length)
        }
        $widths += $width
    }

    $border = "+-" + (($widths | ForEach-Object { "-" * $_ }) -join "-+-") + "-+"

    $lines = @()
    $lines += $border

    $headerCells = @("Thong so".PadRight($widths[0]))
    for ($i = 0; $i -lt $tableRows.Count; $i++) {
        $headerCells += (ConvertTo-Str $tableRows[$i].Tuan).PadLeft($widths[$i + 1])
    }
    $lines += "| " + ($headerCells -join " | ") + " |"
    $lines += $border

    foreach ($metric in $metricRows) {
        $cells = @((ConvertTo-Str $metric.Metric).PadRight($widths[0]))
        for ($i = 0; $i -lt $tableRows.Count; $i++) {
            $cells += (ConvertTo-Str $metric.Values[$i]).PadLeft($widths[$i + 1])
        }
        $lines += "| " + ($cells -join " | ") + " |"
    }

    $lines += $border
    return @($lines)
}

function New-MRReport {
    param(
        [string]$RequiredPart,
        [double]$RequiredQty,
        [datetime]$RequiredDate,
        [bool]$Updated,
        [string]$RequiredSite,
        [string]$RequiredVendorCode,
        [object[]]$Commitments,
        [object[]]$Stocks
    )

    $planOrder = Measure-PropertySum -Rows $Commitments -Property "PlanOrder"
    $usableSupplyReply = Measure-PropertySum -Rows $Commitments -Property "SupplyReply"
    $commitmentFG = Measure-PropertySum -Rows $Commitments -Property "CommitmentFG"
    $fgStock = Measure-PropertySum -Rows $Stocks -Property "FG"
    $fgForCalc = [Math]::Max($fgStock, $commitmentFG)
    $fgUsedForDisplay = if ($Updated) { $commitmentFG } else { $fgForCalc }
    $remainingSource = $fgUsedForDisplay - $usableSupplyReply
    $minCommitmentGap = 0.0
    $minReplyGap = 0.0
    if ($Commitments.Count -gt 0) {
        $minCommitmentGap = Measure-PropertyMinimum -Rows $Commitments -Property "MinGap"
        $minReplyGap = Measure-PropertyMinimum -Rows $Commitments -Property "MinReplyGap"
    }
    $demandNeeded = if ($Updated) { $planOrder } else { $RequiredQty }
    $replySurplus = $usableSupplyReply - $planOrder
    $mrSurplus = $remainingSource - $RequiredQty
    $displaySurplus = if ($Updated) { $replySurplus } else { $mrSurplus }
    $hasCommitment = ($Commitments.Count -gt 0)
    $hasStock = ($Stocks.Count -gt 0)
    $needsManualCheck = $false
    $dataWarnings = New-Object System.Collections.Generic.List[string]

    if ($Updated) {
        if (-not $hasCommitment) {
            $needsManualCheck = $true
            $dataWarnings.Add("Khong tim thay ma vat tu trong Supply Commitment theo site/vendor/ngay da chon.")
        }
        elseif ($planOrder -le 0) {
            $needsManualCheck = $true
            $dataWarnings.Add("Plan Order bang 0 truoc ngay requested; chua xac nhan duoc demand pull-in trong file commitment.")
        }
    }
    else {
        if (-not $hasCommitment) {
            $dataWarnings.Add("Khong tim thay ma vat tu trong Supply Commitment; Supply Reply duoc tinh la 0.")
        }
        if (-not $hasStock) {
            if ($commitmentFG -gt 0) {
                $dataWarnings.Add("Khong tim thay ma vat tu trong Vendor Stock; dung FG trong Supply Commitment de tinh.")
            }
            else {
                $dataWarnings.Add("Khong tim thay ma vat tu trong Vendor Stock; FG duoc tinh la 0.")
            }
        }
        if (-not $hasCommitment -and -not $hasStock) {
            $needsManualCheck = $true
            $dataWarnings.Add("Khong tim thay ma vat tu trong ca Supply Commitment va Vendor Stock.")
        }
    }

    if ($needsManualCheck) {
        $recommendation = "Need manual check"
    }
    elseif ($Updated) {
        if ($minReplyGap -lt -0.0001 -or $remainingSource -lt -0.0001 -or $minCommitmentGap -lt -0.0001) {
            $recommendation = "Can not meet MR"
        }
        else {
            $recommendation = "Can meet"
        }
    }
    elseif ($remainingSource + 0.0001 -ge $RequiredQty) {
        $recommendation = "Can meet"
    }
    else {
        $recommendation = "Can not meet MR"
    }

    $recommendationVi = switch ($recommendation) {
        "Can meet" { "Co the dap ung" }
        "Can not meet MR" { "Khong dap ung MR" }
        "Need manual check" { "Can kiem tra thu cong" }
        default { $recommendation }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("===== KET QUA CHECK MR PULL-IN =====")
    $lines.Add(("Ma vat tu: {0}" -f $RequiredPart))
    $lines.Add(("So luong MR nhap vao: {0}" -f (Format-Qty $RequiredQty)))
    $lines.Add(("Ngay can pull-in: {0}" -f $RequiredDate.ToString("yyyy-MM-dd")))
    $lines.Add(("Site: {0}" -f $(if ([string]::IsNullOrWhiteSpace($RequiredSite)) { "All" } else { $RequiredSite.ToUpperInvariant() })))
    $lines.Add(("Vendor Code: {0}" -f $(if ([string]::IsNullOrWhiteSpace($RequiredVendorCode)) { "All" } else { $RequiredVendorCode })))
    $lines.Add(("Da update ASCP: {0}" -f $(if ($Updated) { "Co" } else { "Chua" })))
    $lines.Add("")
    $lines.Add("----- Tinh trang du lieu -----")
    $lines.Add(("Dong Supply Commitment tim thay: {0}" -f $Commitments.Count))
    if (-not $Updated) {
        $lines.Add(("Dong Vendor Stock tim thay: {0}" -f $Stocks.Count))
    }
    if ($dataWarnings.Count -eq 0) {
        $lines.Add("Khong co canh bao du lieu.")
    }
    else {
        foreach ($warning in $dataWarnings) {
            $lines.Add(("Canh bao: {0}" -f $warning))
        }
    }
    $lines.Add("")
    $lines.Add("----- Cach tinh -----")
    if ($Updated) {
        $lines.Add("Da update ASCP = Co")
        $lines.Add("Nhu cau can dap ung = tong Plan Order den tuan requested")
        $lines.Add("Supply Reply kha dung = Supply Reply cua tuan truoc dung de cover Plan Order tuan hien tai")
        $lines.Add("Chenh Reply = Supply Reply kha dung - Plan Order")
        $lines.Add("FG con lai = FG trong Supply Commitment - Supply Reply kha dung")
        $lines.Add("Ket luan = Chenh Reply khong am va FG con lai khong am")
    }
    else {
        $lines.Add("Da update ASCP = Chua")
        $lines.Add("FG dung de tinh = max(FG Vendor Stock, FG trong Supply Commitment)")
        $lines.Add("Supply Reply kha dung = Supply Reply cua tuan truoc dung de cover timing hien tai")
        $lines.Add("FG con lai = FG dung de tinh - Supply Reply kha dung")
        $lines.Add("Ket luan = FG con lai >= So luong MR nhap vao")
    }
    $lines.Add("")

    $summaryRows = @(
        [PSCustomObject]@{ Metric = "Ket luan"; Value = $recommendationVi }
        [PSCustomObject]@{ Metric = "So luong MR nhap vao"; Value = Format-Qty $RequiredQty }
        [PSCustomObject]@{ Metric = "Nhu cau Plan Order"; Value = Format-Qty $planOrder }
        [PSCustomObject]@{ Metric = "FG Vendor Stock"; Value = Format-Qty $fgStock }
        [PSCustomObject]@{ Metric = "FG Supply Commitment"; Value = Format-Qty $commitmentFG }
        [PSCustomObject]@{ Metric = "FG dung de tinh"; Value = Format-Qty $fgUsedForDisplay }
        [PSCustomObject]@{ Metric = "Supply Reply kha dung"; Value = Format-Qty $usableSupplyReply }
        [PSCustomObject]@{ Metric = "FG con lai sau Supply Reply"; Value = Format-Qty $remainingSource }
        [PSCustomObject]@{ Metric = $(if ($Updated) { "Chenh Reply so voi Plan Order" } else { "Du / Thieu so voi MR" }); Value = Format-Qty $displaySurplus }
        [PSCustomObject]@{ Metric = "Chenh Reply thap nhat"; Value = Format-Qty $minReplyGap }
        [PSCustomObject]@{ Metric = "FG con lai thap nhat"; Value = Format-Qty $minCommitmentGap }
    )

    $lines.Add("----- Bang tom tat -----")
    Add-ReportLines -Lines $lines -NewLines (New-KeyValueTable -Rows $summaryRows)
    $lines.Add("")

    $lines.Add(("Ket luan: {0}" -f $recommendationVi))
    if ($recommendation -eq "Need manual check") {
        if ($dataWarnings.Count -gt 0) {
            $lines.Add(("Ly do: {0}" -f ($dataWarnings[$dataWarnings.Count - 1])))
        }
        else {
            $lines.Add("Ly do: Thieu du lieu nguon hoac du lieu chua du de ket luan.")
        }
    }
    elseif ($recommendation -eq "Can meet") {
        if ($Updated) {
            $lines.Add(("Ly do: Supply Reply cover du Plan Order, chenh thap nhat {0}; FG con lai {1}." -f (Format-Qty $minReplyGap), (Format-Qty $remainingSource)))
        }
        else {
            $lines.Add(("Ly do: FG con lai sau khi tru Supply Reply du so voi MR, con du {0}." -f (Format-Qty $mrSurplus)))
        }
    }
    else {
        if ($Updated -and $minReplyGap -lt -0.0001) {
            $lines.Add(("Ly do: Supply Reply tuan truoc khong cover du Plan Order. Chenh Reply thap nhat: {0}." -f (Format-Qty $minReplyGap)))
        }
        elseif ($remainingSource -lt -0.0001 -or $minCommitmentGap -lt -0.0001) {
            $lines.Add(("Ly do: FG con lai bi am sau khi tru Supply Reply. FG con lai thap nhat: {0}." -f (Format-Qty $minCommitmentGap)))
        }
        else {
            $lines.Add(("Ly do: FG con lai sau khi tru Supply Reply thieu {0} so voi MR." -f (Format-Qty ([Math]::Max((-1 * $mrSurplus), 0)))))
        }
    }

    $lines.Add("")
    $lines.Add("----- Chi tiet Supply Commitment -----")
    if ($Commitments.Count -eq 0) {
        $lines.Add("Khong tim thay dong Supply Commitment phu hop.")
    }
    else {
        foreach ($row in $Commitments) {
            $lines.Add(("Site: {0}" -f $row.Site))
            $lines.Add(("Supplier: {0} - {1}" -f $(if ($row.VendorCode) { $row.VendorCode } else { "N/A" }), $(if ($row.VendorName) { $row.VendorName } else { "N/A" })))
            $lines.Add(("File: {0}" -f (Get-FileNameSafe -Path $row.File)))
            $lines.Add(("Bucket da tinh: {0}" -f $(if ($row.Buckets) { $row.Buckets } else { "N/A" })))
            $detailRows = @(
                [PSCustomObject]@{ Metric = "FG Supply Commitment"; Value = Format-Qty ([double]$row.CommitmentFG) }
                [PSCustomObject]@{ Metric = "Plan Order"; Value = Format-Qty ([double]$row.PlanOrder) }
                [PSCustomObject]@{ Metric = "Supply Reply kha dung"; Value = Format-Qty ([double]$row.SupplyReply) }
                [PSCustomObject]@{ Metric = "Chenh Reply thap nhat"; Value = Format-Qty ([double]$row.MinReplyGap) }
                [PSCustomObject]@{ Metric = "FG con lai cuoi ky"; Value = Format-Qty ([double]$row.EndingGap) }
                [PSCustomObject]@{ Metric = "FG con lai thap nhat"; Value = Format-Qty ([double]$row.MinGap) }
            )
            Add-ReportLines -Lines $lines -NewLines (New-KeyValueTable -Rows $detailRows)
            if ($row.Timeline) {
                $lines.Add("Bang rolling theo tuan:")
                Add-ReportLines -Lines $lines -NewLines (New-TimelineTable -Rows @($row.Timeline))
            }
        }
    }

    if (-not $Updated) {
        $lines.Add("")
        $lines.Add("----- Chi tiet Vendor Stock (chi lay FG) -----")
        if ($Stocks.Count -eq 0) {
            $lines.Add("Khong tim thay dong Vendor Stock phu hop.")
        }
        else {
            foreach ($row in $Stocks) {
                $lines.Add(("{0} | Vendor {1} | {2} | FG {3} | File: {4}" -f
                    $(if ($row.Site) { $row.Site } else { "N/A" }),
                    $(if ($row.VendorCode) { $row.VendorCode } else { "N/A" }),
                    $(if ($row.VendorName) { $row.VendorName } else { [System.IO.Path]::GetFileNameWithoutExtension($row.File) }),
                    (Format-Qty ([double]$row.FG)),
                    (Get-FileNameSafe -Path $row.File)
                ))
            }
        }
    }

    return ($lines -join [Environment]::NewLine)
}

function Normalize-MatchText {
    param([string]$Value)

    return ((ConvertTo-Str $Value).ToUpperInvariant() -replace '[^A-Z0-9]+', '')
}

function Get-SupplierRows {
    if (-not (Test-Path -LiteralPath $AutoToolsConfig.SupplierMaster -PathType Leaf)) {
        return @()
    }

    return @(Import-Csv -LiteralPath $AutoToolsConfig.SupplierMaster)
}

function Resolve-SupplierInput {
    param(
        [string]$InputValue,
        [bool]$AllowPrompt
    )

    $inputText = ConvertTo-Str $InputValue
    if ([string]::IsNullOrWhiteSpace($inputText)) {
        return [PSCustomObject]@{
            VendorCode = ""
            SupplierLabel = "All"
        }
    }

    $supplierRows = @(Get-SupplierRows)
    if ($supplierRows.Count -eq 0) {
        if ($inputText -match '^\d+$') {
            return [PSCustomObject]@{
                VendorCode = $inputText
                SupplierLabel = $inputText
            }
        }

        throw "Khong tim thay Config\suppliers.csv de map ten supplier sang vendor code."
    }

    if ($inputText -match '^\d+$') {
        $exactCode = @(
            $supplierRows |
                Where-Object {
                    [string]::Equals(
                        (ConvertTo-Str $_.VendorCode),
                        $inputText,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        )

        if ($exactCode.Count -gt 0) {
            return [PSCustomObject]@{
                VendorCode = ConvertTo-Str $exactCode[0].VendorCode
                SupplierLabel = ConvertTo-Str $exactCode[0].VendorName
            }
        }

        return [PSCustomObject]@{
            VendorCode = $inputText
            SupplierLabel = $inputText
        }
    }

    $inputNorm = Normalize-MatchText $inputText
    $matches = @(
        $supplierRows |
            Where-Object {
                $vendorName = ConvertTo-Str $_.VendorName
                $keyword = ConvertTo-Str $_.Keyword
                $searchValues = @($vendorName) + ($keyword -split ';')

                $isMatch = $false
                foreach ($value in $searchValues) {
                    $candidateNorm = Normalize-MatchText $value
                    if (
                        -not [string]::IsNullOrWhiteSpace($candidateNorm) -and
                        ($candidateNorm.Contains($inputNorm) -or $inputNorm.Contains($candidateNorm))
                    ) {
                        $isMatch = $true
                        break
                    }
                }

                $isMatch
            }
    )

    if ($matches.Count -eq 0) {
        throw "Khong tim thay supplier khop voi input: $inputText"
    }

    if ($matches.Count -eq 1) {
        return [PSCustomObject]@{
            VendorCode = ConvertTo-Str $matches[0].VendorCode
            SupplierLabel = ConvertTo-Str $matches[0].VendorName
        }
    }

    if (-not $AllowPrompt) {
        $choices = ($matches | ForEach-Object { "{0} - {1}" -f $_.VendorCode, $_.VendorName }) -join "; "
        throw "Supplier input '$inputText' khop nhieu supplier: $choices"
    }

    Write-Host ""
    Write-Host "Tim thay nhieu supplier khop '$inputText':" -ForegroundColor Yellow
    for ($i = 0; $i -lt $matches.Count; $i++) {
        Write-Host ("  {0}. {1} - {2}" -f ($i + 1), $matches[$i].VendorCode, $matches[$i].VendorName)
    }

    do {
        $choiceText = Read-Host "Chon so thu tu supplier"
        $choice = 0
        $ok = [int]::TryParse($choiceText, [ref]$choice)
    } while (-not $ok -or $choice -lt 1 -or $choice -gt $matches.Count)

    $selected = $matches[$choice - 1]
    return [PSCustomObject]@{
        VendorCode = ConvertTo-Str $selected.VendorCode
        SupplierLabel = ConvertTo-Str $selected.VendorName
    }
}

function ConvertTo-PartRequests {
    param(
        [string]$InputText,
        [double]$DefaultQty
    )

    $tokens = @(
        (ConvertTo-Str $InputText) -split '[,;\r\n\t ]+' |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($tokens.Count -eq 0) {
        throw "Chua co part item hop le."
    }

    $requests = @()
    foreach ($token in $tokens) {
        $part = $token
        $qty = $DefaultQty

        if ($token -match '^(?<part>[^:=]+)[:=](?<qty>.+)$') {
            $part = $matches.part
            $qty = ConvertTo-Num $matches.qty
        }

        $part = (ConvertTo-Str $part)
        if ([string]::IsNullOrWhiteSpace($part)) {
            continue
        }
        if ($qty -le 0) {
            throw "MR Qty cua part $part phai lon hon 0."
        }

        $requests += [PSCustomObject]@{
            PartNum = $part
            MRQty = [double]$qty
        }
    }

    if ($requests.Count -eq 0) {
        throw "Chua co part item hop le."
    }

    return @($requests)
}

function Test-PartInputNeedsDefaultQty {
    param([string]$InputText)

    $tokens = @(
        (ConvertTo-Str $InputText) -split '[,;\r\n\t ]+' |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($tokens.Count -eq 0) {
        return $true
    }

    foreach ($token in $tokens) {
        if ($token -notmatch '^[^:=]+[:=].+$') {
            return $true
        }
    }

    return $false
}

function Invoke-MRCheckBatch {
    param(
        [object[]]$PartRequests,
        [datetime]$RequiredDate,
        [string]$RequiredSite,
        [string]$RequiredVendorCode,
        [bool]$Updated,
        [string]$SupplierLabel
    )

    Write-Host ""
    Write-Log "Dang doc du lieu..." "Cyan"

    $excel = $null
    $reports = @()
    try {
        if (-not $Updated) {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
        }

        foreach ($request in $PartRequests) {
            Write-Host ""
            Write-Host ("===== CHECK MA VAT TU {0} =====" -f $request.PartNum) -ForegroundColor Cyan

            $commitments = @(Read-SupplierCommitment `
                -Excel $excel `
                -RequiredPart $request.PartNum `
                -RequiredSite $RequiredSite `
                -RequiredVendorCode $RequiredVendorCode `
                -RequiredDate $RequiredDate)

            $stocks = @()
            if (-not $Updated) {
                $stocks = @(Read-VendorStock `
                    -Excel $excel `
                    -RequiredPart $request.PartNum `
                    -RequiredSite $RequiredSite `
                    -RequiredVendorCode $RequiredVendorCode)
            }

            $report = New-MRReport `
                -RequiredPart $request.PartNum `
                -RequiredQty $request.MRQty `
                -RequiredDate $RequiredDate `
                -Updated $Updated `
                -RequiredSite $RequiredSite `
                -RequiredVendorCode $RequiredVendorCode `
                -Commitments $commitments `
                -Stocks $stocks

            $reports += $report
            Write-Host ""
            Write-Host $report
        }
    }
    finally {
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        }
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    if ($PartRequests.Count -eq 1) {
        $safePart = ($PartRequests[0].PartNum -replace '[^A-Za-z0-9._-]', '_')
        $reportPath = Join-Path -Path $ReportRoot -ChildPath ("MR_Check_{0}_{1}.txt" -f $safePart, $stamp)
    }
    else {
        $reportPath = Join-Path -Path $ReportRoot -ChildPath ("MR_Check_Batch_{0}.txt" -f $stamp)
    }

    $batchHeader = @(
        "===== THONG TIN BATCH CHECK MR ====="
        ("Supplier: {0}" -f $(if ([string]::IsNullOrWhiteSpace($SupplierLabel)) { "All" } else { $SupplierLabel }))
        ("Vendor Code: {0}" -f $(if ([string]::IsNullOrWhiteSpace($RequiredVendorCode)) { "All" } else { $RequiredVendorCode }))
        ("Site: {0}" -f $(if ([string]::IsNullOrWhiteSpace($RequiredSite)) { "All" } else { $RequiredSite }))
        ("Ngay can pull-in: {0}" -f $RequiredDate.ToString("yyyy-MM-dd"))
        ("Da update ASCP: {0}" -f $(if ($Updated) { "Co" } else { "Chua" }))
        ""
    )

    Set-Content -LiteralPath $reportPath -Value (($batchHeader + ($reports -join ([Environment]::NewLine + [Environment]::NewLine))) -join [Environment]::NewLine) -Encoding UTF8
    Write-Host ""
    Write-Log "Da luu report: $reportPath" "Green"
}

try {
    Write-Host ""
    Write-Host "===== CHECK MR PULL-IN =====" -ForegroundColor Cyan
    Write-Host "Cach tinh:"
    Write-Host "  Chua update ASCP = FG con lai sau khi tru Supply Reply kha dung so voi MR Qty"
    Write-Host "  Da update ASCP   = Supply Reply tuan truoc cover Plan Order, dong thoi FG con lai khong am"
    Write-Host ""
    Write-Host "Nhap nhieu part bang dau phay, vi du: 685937001,525476023"
    Write-Host "Neu moi part co qty rieng, dung dang: 685937001:500,525476023:1000"
    Write-Host ""

    $runAgain = $true
    $firstRun = $true
    while ($runAgain) {
        $partInput = Read-RequiredInput -Label "Part item(s)" -CurrentValue $(if ($firstRun) { $PartNum } else { "" })

        $qtyValue = $MRQty
        $needsDefaultQty = Test-PartInputNeedsDefaultQty -InputText $partInput
        if ($needsDefaultQty -and ($qtyValue -le 0 -or -not $firstRun)) {
            $qtyText = Read-RequiredInput -Label "MR Qty mac dinh" -CurrentValue ""
            $qtyValue = ConvertTo-Num $qtyText
        }
        if ($needsDefaultQty -and $qtyValue -le 0) {
            throw "MR Qty phai lon hon 0."
        }

        $partRequests = @(ConvertTo-PartRequests -InputText $partInput -DefaultQty $qtyValue)
        $requestDateValue = ConvertTo-UserDate $(if ($firstRun) { $RequestedDate } else { "" })
        $siteValue = Read-OptionalInput -Label "Site (TC5/TN5, bo trong neu check tat ca)" -CurrentValue $(if ($firstRun) { $Site } else { "" })

        $supplierInput = ""
        if ($firstRun -and -not [string]::IsNullOrWhiteSpace($VendorCode)) {
            $supplierInput = $VendorCode
        }
        elseif ($firstRun -and -not [string]::IsNullOrWhiteSpace($Supplier)) {
            $supplierInput = $Supplier
        }
        else {
            $supplierInput = Read-OptionalInput -Label "Supplier name/code (vi du YAO-I, WONDERWARD, 9781; bo trong neu check tat ca)" -CurrentValue ""
        }

        $supplierResolved = Resolve-SupplierInput -InputValue $supplierInput -AllowPrompt (-not $NoPause)
        if (-not [string]::IsNullOrWhiteSpace($supplierResolved.VendorCode)) {
            Write-Host ("Supplier da chon: {0} - {1}" -f $supplierResolved.VendorCode, $supplierResolved.SupplierLabel) -ForegroundColor Cyan
        }

        $hasASCPParam = $firstRun -and $PSBoundParameters.ContainsKey("ASCPUpdated") -and -not [string]::IsNullOrWhiteSpace($ASCPUpdated)
        $ascpCurrentValue = if ($hasASCPParam) { ConvertTo-BoolInput -Value $ASCPUpdated } else { $false }
        $ascpUpdatedValue = Read-YesNo -Label "Da update ASCP chua" -CurrentValue $ascpCurrentValue -HasCurrentValue $hasASCPParam

        Invoke-MRCheckBatch `
            -PartRequests $partRequests `
            -RequiredDate $requestDateValue `
            -RequiredSite $siteValue `
            -RequiredVendorCode $supplierResolved.VendorCode `
            -Updated $ascpUpdatedValue `
            -SupplierLabel $supplierResolved.SupplierLabel

        if ($NoPause) {
            $runAgain = $false
        }
        else {
            $runAgain = Read-YesNo -Label "Tiep tuc check ma khac khong" -CurrentValue $false -HasCurrentValue $false
            Write-Host ""
        }

        $firstRun = $false
        $PartNum = ""
        $MRQty = 0
        $RequestedDate = ""
        $Site = ""
        $VendorCode = ""
        $Supplier = ""
        $ASCPUpdated = ""
    }
}
catch {
    Write-Host ""
    Write-Log "Loi: $($_.Exception.Message)" "Red"
    exit 1
}
finally {
    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Nhan Enter de thoat"
    }
}
