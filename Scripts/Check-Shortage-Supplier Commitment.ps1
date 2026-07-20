# =====================================================================
#  Check Shortage - Kiem tra ma vat tu thieu hang (Gap am)
#  - Tu dong quet TOAN BO file .xls trong TC5 va TN5 cua folder tuan hien tai
#  - Tinh Gap cong don theo tung tuan cho moi Part Num
#  - Xuat output dang TIMELINE: moi dong 1 ma, moi cot 1 tuan
#    O thieu hang -> hien Gap am + to mau do. O du -> de trong.
#  - Tien copy vao mail highlight cho supplier.
#
#  Cong thuc Gap[t] = Gap[t-1] + SupplyReply[t-1] - PlanOrder[t]
#  Moi Part Num gom 5 dong: Firm / Forecast / Plan Order / Supply Reply / Gap
# =====================================================================

# ----------------- CAU HINH (chinh o day) ----------------------------
. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")
$AutoToolsPaths  = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

$OutputFolder = $AutoToolsConfig.SupplierCommitmentOutput
$runDate = Get-Date
$weekStart = $runDate.Date.AddDays(-1 * [int]$runDate.DayOfWeek)
$weekEnd = $weekStart.AddDays(6)
$weekFolderName = "{0}-{1}" -f $weekStart.ToString("dd.MM"), $weekEnd.ToString("dd.MM.yyyy")
$RootFolder = Join-Path -Path $OutputFolder -ChildPath $weekFolderName
$BUFolders  = @("TC5", "TN5")
$Extensions = @("*.xls", "*.xlsx")

# So tuan lich can kiem tra (bo qua Past Due). 12 tuan ~ 3 thang.
$WeeksToCheck = 12
# ---------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
$ErrorActionPreference = "Stop"

function Write-Log($msg, $color = "White") {
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg) -ForegroundColor $color
}

function ConvertTo-Num($v) {
    if ($null -eq $v) { return 0.0 }
    if ($v -is [array]) {
        if ($v.Length -eq 0) { return 0.0 }
        $v = $v[0]; if ($null -eq $v) { return 0.0 }
    }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) { return [double]$v }
    $s = ([string]$v).Trim()
    if ($s -eq "") { return 0.0 }
    $s = $s -replace ',', ''
    $out = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
    return 0.0
}

function ConvertTo-Str($v) {
    if ($null -eq $v) { return "" }
    if ($v -is [array]) {
        if ($v.Length -eq 0) { return "" }
        $v = $v[0]; if ($null -eq $v) { return "" }
    }
    return ([string]$v).Trim()
}

# --- Kiem tra folder goc ---
if (-not (Test-Path $RootFolder)) {
    [System.Windows.Forms.MessageBox]::Show("Khong tim thay folder:`n$RootFolder", "Loi", "OK", "Error") | Out-Null
    Write-Log "Khong tim thay folder goc: $RootFolder" "Red"
    Read-Host "Enter de thoat"; exit 1
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# --- Thu thap danh sach file ---
$fileList = @()
foreach ($bu in $BUFolders) {
    $buPath = Join-Path $RootFolder $bu
    if (-not (Test-Path $buPath)) { Write-Log "Bo qua: khong co folder '$bu'" "Yellow"; continue }
    foreach ($ext in $Extensions) {
        Get-ChildItem -Path $buPath -Filter $ext -File -ErrorAction SilentlyContinue | ForEach-Object {
            $fileList += @{ Path = $_.FullName; BU = $bu }
        }
    }
}
if ($fileList.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("Khong tim thay file .xls/.xlsx nao trong TC5 va TN5.", "Ket qua", "OK", "Information") | Out-Null
    Write-Log "Khong co file de quet." "Yellow"; Read-Host "Enter de thoat"; exit 0
}
Write-Log "Tim thay $($fileList.Count) file de quet." "Cyan"

# --- Mo Excel ---
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
} catch {
    Write-Log "Khong khoi dong duoc Excel. Can cai Microsoft Excel." "Red"
    Read-Host "Enter de thoat"; exit 1
}

$allResults = @()
$calWeekNames = $null   # ten cac tuan lich (cho header timeline) - lay tu file dau tien

foreach ($item in $fileList) {
    $file = $item.Path
    $bu   = $item.BU
    Write-Log "[$bu] $([System.IO.Path]::GetFileName($file))" "Cyan"
    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($file)
        $ws = $wb.Sheets.Item(1)
        $used = $ws.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count
        $data = $used.Value2

        $getCell = {
            param($row, $col)
            try { return $data.GetValue($row, $col) } catch { return $null }
        }

        $colPart = 2
        $colType = 10
        $colFirstWeek = 11   # Past Due

        $colLastWeek = $cols
        for ($c = $colFirstWeek; $c -le $cols; $c++) {
            $h = ConvertTo-Str (& $getCell 1 $c)
            if ($h -match '(?i)total demand') { $colLastWeek = $c - 1; break }
        }

        $colCalFirst = $colFirstWeek + 1
        $colCheckEnd = $colCalFirst + $WeeksToCheck - 1
        if ($colCheckEnd -gt $colLastWeek) { $colCheckEnd = $colLastWeek }

        # Ten cac tuan (doc .Text de ra ngay thang dung)
        $weekNames = @()
        for ($c = $colFirstWeek; $c -le $colLastWeek; $c++) {
            $txt = ""
            try { $txt = [string]$ws.Cells.Item(1, $c).Text } catch {}
            if ([string]::IsNullOrWhiteSpace($txt)) { $txt = ConvertTo-Str (& $getCell 1 $c) }
            $weekNames += $txt.Trim()
        }

        # Ten cac tuan lich trong pham vi kiem tra (cho header timeline)
        if ($null -eq $calWeekNames) {
            $calWeekNames = @()
            for ($c = $colCalFirst; $c -le $colCheckEnd; $c++) {
                $calWeekNames += $weekNames[$c - $colFirstWeek]
            }
        }

        $r = 2
        while ($r + 4 -le $rows) {
            $t2 = ConvertTo-Str (& $getCell ($r + 2) $colType)
            $t3 = ConvertTo-Str (& $getCell ($r + 3) $colType)
            $t4 = ConvertTo-Str (& $getCell ($r + 4) $colType)

            if ($t2 -eq "Plan Order" -and $t3 -eq "Supply Reply" -and $t4 -eq "Gap") {
                $part = ConvertTo-Str (& $getCell $r $colPart)

                $gap = 0.0; $prevGap = 0.0; $prevSupply = 0.0
                $negCount = 0; $totalShortage = 0.0; $minGap = 0.0
                $firstNegWeek = $null; $lastNegWeek = $null
                $gapByWeek = @{}   # ten tuan lich -> gap am (chi luu tuan thieu)
                $wi = 0
                for ($c = $colFirstWeek; $c -le $colLastWeek; $c++) {
                    $plan   = ConvertTo-Num (& $getCell ($r + 2) $c)
                    $supply = ConvertTo-Num (& $getCell ($r + 3) $c)
                    $gap = $prevGap + $prevSupply - $plan

                    $inRange = ($c -ge $colCalFirst -and $c -le $colCheckEnd)
                    if ($inRange -and $gap -lt -0.001) {
                        $negCount++
                        $wname = $weekNames[$wi]
                        if ($null -eq $firstNegWeek) { $firstNegWeek = $wname }
                        $lastNegWeek = $wname
                        $totalShortage += $gap
                        if ($gap -lt $minGap) { $minGap = $gap }
                        $gapByWeek[$wname] = [math]::Round($gap, 0)
                    }
                    $prevGap = $gap
                    $prevSupply = $supply
                    $wi++
                }

                if ($negCount -gt 0) {
                    $allResults += [PSCustomObject]@{
                        BU            = $bu
                        Supplier      = [System.IO.Path]::GetFileNameWithoutExtension($file)
                        PartNum       = $part
                        SoTuanThieu   = $negCount
                        TuanThieuDau  = $firstNegWeek
                        TuanThieuCuoi = $lastNegWeek
                        GapAmThapNhat = [math]::Round($minGap, 0)
                        TongThieu     = [math]::Round($totalShortage, 0)
                        GapByWeek     = $gapByWeek
                    }
                }
            }
            $r += 5
        }
        $wb.Close($false)
    } catch {
        Write-Log "  Loi khi doc file: $_" "Red"
        if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    }
}

Write-Log "Tong so ma thieu hang: $($allResults.Count)" "Yellow"

if ($allResults.Count -eq 0) {
    $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.Windows.Forms.MessageBox]::Show("Khong co ma vat tu nao bi thieu hang (Gap am).", "Ket qua", "OK", "Information") | Out-Null
    Read-Host "Enter de thoat"; exit 0
}

$sorted = $allResults | Sort-Object BU, Supplier, PartNum

# =====================================================================
#  XUAT FILE TIMELINE (.xlsx) bang Excel COM, to mau o thieu
# =====================================================================
Write-Log "Dang tao file timeline..." "Cyan"

$nWeeks = $calWeekNames.Count
$wbOut = $excel.Workbooks.Add()
$wsOut = $wbOut.Sheets.Item(1)
$wsOut.Name = "Shortage Timeline"

# --- Header ---
# Cot: 1=BU, 2=Supplier, 3=PartNum, 4..(3+nWeeks)=tuan, +1=TongThieu, +2=SoTuanThieu
$fixedCols = 3
$wsOut.Cells.Item(1,1).Value2 = "BU"
$wsOut.Cells.Item(1,2).Value2 = "Supplier"
$wsOut.Cells.Item(1,3).Value2 = "Part Num"
for ($i = 0; $i -lt $nWeeks; $i++) {
    $wsOut.Cells.Item(1, $fixedCols + 1 + $i).Value2 = $calWeekNames[$i]
}
$colTotal = $fixedCols + 1 + $nWeeks
$colCount = $colTotal + 1
$wsOut.Cells.Item(1, $colTotal).Value2 = "Tong Thieu"
$wsOut.Cells.Item(1, $colCount).Value2 = "So Tuan Thieu"

# Dinh dang header
$lastCol = $colCount
$headerRange = $wsOut.Range($wsOut.Cells.Item(1,1), $wsOut.Cells.Item(1, $lastCol))
$headerRange.Font.Bold = $true
$headerRange.Font.Color = 16777215   # trang
$headerRange.Interior.Color = 6299648 # xanh dam (BGR)
$headerRange.HorizontalAlignment = -4108  # center
$headerRange.VerticalAlignment = -4108

# --- Du lieu ---
$rowOut = 2
foreach ($res in $sorted) {
    $wsOut.Cells.Item($rowOut, 1).Value2 = $res.BU
    $wsOut.Cells.Item($rowOut, 2).Value2 = $res.Supplier
    $wsOut.Cells.Item($rowOut, 3).Value2 = $res.PartNum

    for ($i = 0; $i -lt $nWeeks; $i++) {
        $wname = $calWeekNames[$i]
        $cell = $wsOut.Cells.Item($rowOut, $fixedCols + 1 + $i)
        if ($res.GapByWeek.ContainsKey($wname)) {
            $cell.Value2 = [string]($res.GapByWeek[$wname])
            $cell.Interior.Color = 13551615   # do nhat (BGR ~ #FFCCCC)
            $cell.Font.Color = 192            # do dam
            $cell.Font.Bold = $true
            $cell.HorizontalAlignment = -4108
        }
    }
    $wsOut.Cells.Item($rowOut, $colTotal).Value2 = [string]$res.TongThieu
    $wsOut.Cells.Item($rowOut, $colTotal).Font.Bold = $true
    $wsOut.Cells.Item($rowOut, $colCount).Value2 = [string]$res.SoTuanThieu
    $wsOut.Cells.Item($rowOut, $colCount).HorizontalAlignment = -4108
    $rowOut++
}

# --- Dinh dang chung ---
$allRange = $wsOut.Range($wsOut.Cells.Item(1,1), $wsOut.Cells.Item($rowOut-1, $lastCol))
$allRange.Borders.LineStyle = 1
$allRange.Borders.Weight = 2
$wsOut.Columns.Item(2).ColumnWidth = 28   # Supplier
$wsOut.Columns.Item(3).ColumnWidth = 14   # Part Num
for ($i = 0; $i -lt $nWeeks; $i++) { $wsOut.Columns.Item($fixedCols + 1 + $i).ColumnWidth = 11 }
$wsOut.Columns.Item($colTotal).ColumnWidth = 12
$wsOut.Columns.Item($colCount).ColumnWidth = 8

# Dong bang header + 3 cot dau de cuon van thay
$wsOut.Application.ActiveWindow.SplitColumn = 3
$wsOut.Application.ActiveWindow.SplitRow = 1
$wsOut.Application.ActiveWindow.FreezePanes = $true

# Auto-filter
$headerRange.AutoFilter() | Out-Null

# --- Luu file ---
$stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$outPath = Join-Path $OutputFolder "Shortage_Timeline_$stamp.xlsx"
$wbOut.SaveAs($outPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
$wbOut.Close($false)

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Log "Da luu timeline: $outPath" "Green"

# Tom tat tren man hinh
Write-Host ""
Write-Host "===== TOM TAT THIEU HANG (trong $WeeksToCheck tuan toi) =====" -ForegroundColor Cyan
$sorted | Format-Table BU, PartNum, SoTuanThieu, TuanThieuDau, TuanThieuCuoi, TongThieu -AutoSize

Start-Process $outPath
Read-Host "Nhan Enter de thoat"
