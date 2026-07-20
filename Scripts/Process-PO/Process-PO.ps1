<#
==============================================================================
 Process-PO.ps1
 ------------------------------------------------------------------------------
 Đọc dữ liệu PO (copy từ Excel, tab-delimited) -> nhóm theo vendor_site ->
 dán vào sheet "Upload PO_NoSER" của file PO Template -> lưu thành file
 Text (Tab delimited) (.txt), sắp xếp theo cấu trúc folder:
       <Output>\<dd.MM-dd.MM.yyyy>\<BU>\<vendor_site>\<vendor_site>.txt
==============================================================================
 CÁCH DÙNG:
   1. Đặt 3 thứ trong cùng 1 thư mục (bất kỳ đâu cũng được):
        - Process-PO.ps1   (script này)
        - PO_Template_5.xlsx  (file mẫu)
        - input.txt        (dữ liệu copy từ Excel, dán vào file này)
   2. Chuột phải Process-PO.ps1 -> Run with PowerShell
      (script tự tìm template/input nằm CÙNG THƯ MỤC với nó, không phụ thuộc
       thư mục hiện tại của PowerShell / Launcher / shortcut.)

   Tuỳ chọn ghi đè tham số:
      .\Process-PO.ps1 -TemplatePath "D:\...\PO_Template_5.xlsx" `
                       -InputPath "D:\...\input.txt" `
                       -OutputRoot "D:\...\Output"
==============================================================================
#>

param(
    [string]$TemplatePath = "",
    [string]$InputPath    = "",
    [string]$OutputRoot   = "",
    [string[]]$MoqPath    = @(),
    [string]$SheetName    = "Upload PO_NoSER"
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Xác định THƯ MỤC CHỨA SCRIPT làm gốc cho mọi đường dẫn mặc định.
# Nhờ vậy script luôn tìm đúng template/input nằm cạnh nó, bất kể đang chạy
# từ đâu (double-click, shortcut, Launcher, hay PowerShell ở thư mục khác).
# ============================================================================
if ($PSScriptRoot -and $PSScriptRoot.Trim().Length -gt 0) {
    $ScriptDir = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = (Get-Location).Path
}

$CommonScript = Join-Path -Path (Split-Path -Path $ScriptDir -Parent) -ChildPath "AutoTools.Common.ps1"
. $CommonScript
$AutoToolsPaths  = Initialize-AutoToolsPaths -StartPath $ScriptDir
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

# Neu nguoi dung khong truyen tham so -> dung cau truc Auto Tools moi.
# Template/input van fallback ve file nam canh script de khong dut workflow cu.
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = $AutoToolsConfig.ProcessPOTemplate
    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        $TemplatePath = Join-Path $ScriptDir "PO_Template_5.xlsx"
    }
}
if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $InputPath = $AutoToolsConfig.ProcessPOInput
    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        $InputPath = Join-Path $ScriptDir "input.txt"
    }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = $AutoToolsConfig.ProcessPOOutput
}

# Tạo folder tuần làm việc theo Chủ Nhật -> Thứ Bảy, ví dụ 05.07-11.07.2026.
$runDate = Get-Date
$weekStart = $runDate.Date.AddDays(-1 * [int]$runDate.DayOfWeek)
$weekEnd = $weekStart.AddDays(6)
$weekFolderName = "{0}-{1}" -f $weekStart.ToString("dd.MM"), $weekEnd.ToString("dd.MM.yyyy")
$OutputRoot = Join-Path -Path $OutputRoot -ChildPath $weekFolderName
foreach ($buFolder in @("TC5", "TN5")) {
    New-Item -ItemType Directory -Path (Join-Path $OutputRoot $buFolder) -Force | Out-Null
}
Write-Host ("Folder tuần làm việc: {0}" -f $OutputRoot)


# ---- Tên cột trên file template (đúng thứ tự A..J) ----
# A:BU  B:VENDOR_CODE  C:VENDOR SITE CODE  D:ITEM_NO  E:ITEM_revision
# F:NEW DOCK DATE  G:NEW ORDER QUANTITY  H:currency  I:EXCESS  J:ship_to

# ---- Bảng tra cứu: vendor id -> tên công ty ----
# Lưu ý: id 10861 có 2 tên (WEIDA Switch / Stamping) -> dùng tên chung.
$VendorMap = @{
    '9826'  = 'CTY TNHH KY THUAT VA CONG NGHE LONG BAO VIET NAM'
    '9781'  = 'YAO-I VIETNAM COMPANY LIMITED'
    '12225' = 'TALWAY VIET NAM COMPANY LIMITED'
    '9340'  = 'GREEN (VIET NAM) CO., LTD'
    '10861' = 'WEIDA (VIETNAM) MANUFACTURING CO.,LTD'
    '11126' = 'HAIXING TECHNOLOGY (VIET NAM) COMPANY LIMITED'
    '10603' = 'XIAN HUA VIET NAM INDUSTRIAL COMPANY LIMITED'
    '11500' = 'KHGEARS VIETNAM COMPANY LIMITED'
    '12259' = 'YU XIN VIET NAM COMPANY LIMITED'
    '13313' = 'NEW SUN VIET NAM TECHNOLOGY COMPANY LIMITED'
    '13553' = 'CONG TY TNHH MAGNET JC (VIET NAM)'
    '10534' = 'MINH THU ACCURATELY MECHANIC COMPANY LIMITED'
    '12504' = 'WONDERWARD TECHNOLOGY (HONG KONG) LIMITED'
    '13316' = 'VIETNAM SHANGHU ELECTRONICS COMPANY LIMITED'
    '12258' = 'NGHIA LONG METAL PRODUCTS COMPANY LIMITED'
    '12728' = 'JUN JAM METAL PRODUCTS TECHNOLOGY LIMITED'
    '12448' = 'CHANTING INTELLIGENT TECHNOLOGY CO., LTD'
    '10981' = 'MINH ANH ELECTRONICS SERVICE-TRADING COMPANY LIMITED'
    '10058' = 'CONG TY TNHH KY THUAT NOVATEK'
}

function Resolve-FullPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    $scriptCandidate = Join-Path $ScriptDir $p
    if (Test-Path -LiteralPath $scriptCandidate) { return $scriptCandidate }
    return (Join-Path $AutoToolsPaths.Root $p)
}

# Tách số id vendor từ vendor_site, vd 'TVC-9781-B' -> '9781'
function Get-VendorId([string]$site) {
    $m = [regex]::Match($site, '\d+')
    if ($m.Success) { return $m.Value }
    return $null
}

# Loại bỏ ký tự cấm trong tên folder Windows
function Sanitize-FolderName([string]$name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[{0}]" -f [regex]::Escape($invalid)
    $clean = ($name -replace $pattern, ' ').Trim().TrimEnd('.')
    return $clean
}

# Trả về tên folder con dựa trên vendor_site (tên công ty nếu tra được)
function Get-VendorFolderName([string]$site) {
    $id = Get-VendorId $site
    if ($id -and $VendorMap.ContainsKey($id)) {
        return (Sanitize-FolderName $VendorMap[$id])
    }
    Write-Host ("   [Cảnh báo] Không tìm thấy tên vendor cho '{0}' (id={1}) -> dùng mã site." -f $site, $id) -ForegroundColor Yellow
    return (Sanitize-FolderName $site)
}

function Split-ConfigList([string]$value) {
    return @(Split-AutoToolsConfigList -Value $value)
}

function Select-MoqFiles {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title  = "Chọn một hoặc nhiều file MOQ (Full carton qty)"
    $dlg.Filter = "Excel/CSV/Text (*.xlsx;*.xls;*.csv;*.txt)|*.xlsx;*.xls;*.csv;*.txt|Tất cả (*.*)|*.*"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return @($dlg.FileNames) }
    return @()
}

function Parse-Qty([string]$s) {
    return (ConvertTo-AutoToolsIntQuantity -Value $s)
}

function Load-MoqTable([string]$path) {
    $table = @{}
    $ext = [System.IO.Path]::GetExtension($path).ToLower()

    if ($ext -eq '.xlsx' -or $ext -eq '.xls') {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        try {
            $wb = $xl.Workbooks.Open($path)
            $ws = $wb.Worksheets.Item(1)
            $used = $ws.UsedRange
            $rowsN = $used.Rows.Count
            for ($r = 1; $r -le $rowsN; $r++) {
                $itemRaw = $ws.Cells.Item($r, 1).Text
                $qtyRaw  = $ws.Cells.Item($r, 2).Text
                if ($itemRaw -notmatch '^\s*\d{4,}\s*$') { continue }
                $qty = Parse-Qty $qtyRaw
                if ($null -ne $qty -and $qty -gt 0) { $table[$itemRaw.Trim()] = $qty }
            }
            $wb.Close($false)
        }
        finally {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
    else {
        foreach ($ln in (Get-Content -Path $path -Encoding UTF8)) {
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            if ($ln.Contains("`t")) { $cols = $ln -split "`t" }
            else                    { $cols = $ln -split ',', 3 }
            if ($cols.Count -lt 2) { continue }
            $itemRaw = $cols[0].Trim()
            if ($itemRaw -notmatch '^\d{4,}$') { continue }
            $qty = Parse-Qty $cols[1]
            if ($null -ne $qty -and $qty -gt 0) { $table[$itemRaw] = $qty }
        }
    }

    return $table
}

function Load-MoqTables([string[]]$paths) {
    $merged = @{}
    $uniquePaths = @(
        $paths |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { Resolve-FullPath $_ } |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
            Select-Object -Unique
    )

    foreach ($path in $uniquePaths) {
        Write-Host ("Nap MOQ PO: {0}" -f $path) -ForegroundColor Gray
        $table = Load-MoqTable $path
        foreach ($key in $table.Keys) {
            if ($merged.ContainsKey($key)) {
                Write-Host ("  [Canh bao] Item {0} bi trung MOQ, dung gia tri moi tu {1}" -f $key, [System.IO.Path]::GetFileName($path)) -ForegroundColor Yellow
            }
            $merged[$key] = [PSCustomObject]@{
                Qty    = [int]$table[$key]
                Source = $path
            }
        }
    }

    if ($uniquePaths.Count -gt 0) {
        Write-Host ("Da nap {0} ma MOQ PO." -f $merged.Count) -ForegroundColor Gray
    }

    return $merged
}

function RoundUp-ToMoq([int]$qty, [int]$moq) {
    if ($moq -le 0) { return $qty }
    return ([math]::Ceiling($qty / [double]$moq)) * $moq
}

$TemplatePath = Resolve-FullPath $TemplatePath
$InputPath    = Resolve-FullPath $InputPath
$OutputRoot   = Resolve-FullPath $OutputRoot

Write-Host "Template : $TemplatePath"
Write-Host "Input    : $InputPath"
Write-Host "Output   : $OutputRoot"
Write-Host ""

if (-not (Test-Path $TemplatePath)) { throw "Không tìm thấy file template: $TemplatePath" }
if (-not (Test-Path $InputPath))    { throw "Không tìm thấy file input: $InputPath" }

# ============================================================================
# BƯỚC 0: Đọc & phân tích input
# ============================================================================
$rawLines = Get-Content -Path $InputPath -Encoding UTF8 | Where-Object { $_.Trim().Length -gt 0 }
if ($rawLines.Count -lt 2) { throw "File input không có dữ liệu." }

# Bỏ dòng header nếu dòng đầu bắt đầu bằng 'BU' (không phân biệt hoa thường)
$startIdx = 0
if ($rawLines[0] -match '^\s*BU\b') { $startIdx = 1 }

$records = @()
for ($i = $startIdx; $i -lt $rawLines.Count; $i++) {
    $cols = $rawLines[$i] -split "`t"
    if ($cols.Count -lt 8) { continue }   # bỏ dòng thiếu cột

    # Chuẩn hoá ngày -> yyyy/MM/dd
    $rawDate = $cols[5].Trim()
    $dateOut = $rawDate
    $parsed  = [datetime]::MinValue
    if ([datetime]::TryParse($rawDate, [ref]$parsed)) {
        $dateOut = $parsed.ToString('yyyy/MM/dd')
    } else {
        $dateOut = ($rawDate -replace '-', '/')   # phòng trường hợp 2026-07-12
    }

    $records += [PSCustomObject]@{
        BU          = $cols[0].Trim()
        VENDOR_CODE = $cols[1].Trim()
        VENDOR_SITE = $cols[2].Trim()
        ITEM_NO     = $cols[3].Trim()
        ITEM_REV    = $cols[4].Trim()
        DOCK_DATE   = $dateOut
        QTY         = $cols[6].Trim()
        CURRENCY    = $cols[7].Trim()
        EXCESS      = if ($cols.Count -gt 8) { $cols[8].Trim() } else { '-' }
        SHIP_TO     = if ($cols.Count -gt 9) { $cols[9].Trim() } else { '' }
    }
}

if ($records.Count -eq 0) { throw "Không phân tích được dòng dữ liệu nào." }
Write-Host ("Đã đọc {0} dòng dữ liệu." -f $records.Count)

# Nhóm theo BU + vendor_site
$groups = $records | Group-Object -Property BU, VENDOR_SITE
Write-Host ("Số nhóm (BU/vendor_site): {0}" -f $groups.Count)
Write-Host ""

# ============================================================================
# TUY CHON MOQ: chi lam tron cho supplier nguoi dung chon
# ============================================================================
$MoqVendorIds = @{}
$MoqTable = @{}

$applyMoq = Read-Host "Co muon lam tron PO theo MOQ cho supplier nao khong? (Y/N, Enter = bo qua)"
if ($applyMoq -eq 'Y' -or $applyMoq -eq 'y') {
    $vendorChoices = @()
    $index = 1
    foreach ($g in $groups) {
        $site = $g.Group[0].VENDOR_SITE
        $vendorId = Get-VendorId $site
        if ([string]::IsNullOrWhiteSpace($vendorId)) { continue }

        $vendorName = Get-VendorFolderName $site
        if ($vendorChoices | Where-Object { $_.VendorId -eq $vendorId }) { continue }

        $vendorChoices += [PSCustomObject]@{
            Index    = $index
            VendorId = $vendorId
            Name     = $vendorName
            Site     = $site
        }
        $index++
    }

    if ($vendorChoices.Count -eq 0) {
        Write-Host "[Canh bao] Khong tim thay supplier hop le de ap dung MOQ." -ForegroundColor Yellow
    }
    else {
        Write-Host ""
        Write-Host "Danh sach supplier trong input:" -ForegroundColor Cyan
        foreach ($choice in $vendorChoices) {
            Write-Host ("  {0}. {1} ({2}) - vi du site: {3}" -f $choice.Index, $choice.Name, $choice.VendorId, $choice.Site)
        }

        $selectedText = Read-Host "Nhap so supplier can lam tron MOQ, cach nhau bang dau phay, hoac ALL"
        $selectedVendorIds = @()
        if ($selectedText -match '(?i)^ALL$') {
            $selectedVendorIds = @($vendorChoices | ForEach-Object { $_.VendorId })
        }
        else {
            $selectedIndexes = @(
                $selectedText -split '[,; ]+' |
                    Where-Object { $_ -match '^\d+$' } |
                    ForEach-Object { [int]$_ }
            )

            foreach ($selectedIndex in $selectedIndexes) {
                $selected = $vendorChoices | Where-Object { $_.Index -eq $selectedIndex } | Select-Object -First 1
                if ($null -ne $selected) {
                    $selectedVendorIds += $selected.VendorId
                }
            }
        }

        foreach ($vendorId in @($selectedVendorIds | Select-Object -Unique)) {
            $MoqVendorIds[$vendorId] = $true
        }

        if ($MoqVendorIds.Count -eq 0) {
            Write-Host "[Canh bao] Chua chon supplier nao -> bo qua MOQ." -ForegroundColor Yellow
        }
        else {
            $moqPaths = @()
            foreach ($pathValue in @($MoqPath)) {
                foreach ($path in (Split-ConfigList $pathValue)) {
                    $moqPaths += $path
                }
            }

            if (Test-Path -LiteralPath $AutoToolsConfig.ProcessPOMoq -PathType Leaf) {
                $moqPaths += $AutoToolsConfig.ProcessPOMoq
            }

            if ($moqPaths.Count -eq 0) {
                $moqPaths = @(Select-MoqFiles)
            }

            if ($moqPaths.Count -eq 0) {
                Write-Host "[Canh bao] Khong chon file MOQ -> bo qua MOQ." -ForegroundColor Yellow
                $MoqVendorIds.Clear()
            }
            else {
                $MoqTable = Load-MoqTables -paths $moqPaths
                if ($MoqTable.Count -eq 0) {
                    Write-Host "[Canh bao] Khong nap duoc ma MOQ nao -> bo qua MOQ." -ForegroundColor Yellow
                    $MoqVendorIds.Clear()
                }
            }
        }
    }

    Write-Host ""
}

# ============================================================================
# Khởi tạo Excel COM
# ============================================================================
$excel = $null
$moqMissing = @()
try {
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw "Không khởi tạo được Microsoft Excel. Máy này cần cài Excel để chạy script. Chi tiết: $($_.Exception.Message)"
    }
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    $xlTextWindows = 20   # định dạng "Text (Tab delimited)"

    foreach ($g in $groups) {
        $bu   = $g.Group[0].BU
        $site = $g.Group[0].VENDOR_SITE
        $vendorId = Get-VendorId $site
        $rows = $g.Group

        Write-Host ("-> {0} / {1}  ({2} dòng)" -f $bu, $site, $rows.Count)

        # --- BƯỚC 2: tạo folder  Output\<tuần>\BU\<tên vendor> ---
        $vendorName = Get-VendorFolderName $site
        $siteFolder = Join-Path (Join-Path $OutputRoot $bu) $vendorName
        if (-not (Test-Path $siteFolder)) {
            New-Item -ItemType Directory -Path $siteFolder -Force | Out-Null
        }

        # --- Mở 1 bản sao mới của template cho mỗi nhóm ---
        $wb = $excel.Workbooks.Open($TemplatePath)
        $ws = $wb.Worksheets.Item($SheetName)

        # --- BƯỚC 1a: xoá dữ liệu cũ (từ dòng 2 trở xuống), giữ header dòng 1 ---
        $used = $ws.UsedRange
        $lastRow = $used.Row + $used.Rows.Count - 1
        if ($lastRow -ge 2) {
            $clearRange = $ws.Range("A2:J$lastRow")
            $clearRange.ClearContents() | Out-Null
        }

        # --- BƯỚC 1b: dán dữ liệu mới bắt đầu từ dòng 2 ---
        # Build mảng 2 chiều để gán 1 lần (nhanh)
        $n = $rows.Count
        $arr = New-Object 'object[,]' $n, 10
        for ($r = 0; $r -lt $n; $r++) {
            $rec = $rows[$r]
            $qtyOut = $rec.QTY

            if (
                $vendorId -and
                $MoqVendorIds.ContainsKey($vendorId) -and
                $MoqTable.Count -gt 0
            ) {
                if ($MoqTable.ContainsKey($rec.ITEM_NO)) {
                    $currentQty = Parse-Qty $rec.QTY
                    $moqInfo = $MoqTable[$rec.ITEM_NO]
                    if ($null -ne $currentQty) {
                        $newQty = [int](RoundUp-ToMoq $currentQty ([int]$moqInfo.Qty))
                        if ($newQty -ne $currentQty) {
                            Write-Host ("     MOQ PO: {0}  {1} -> {2}  (MOQ={3}, file={4})" -f $rec.ITEM_NO, $currentQty, $newQty, $moqInfo.Qty, [System.IO.Path]::GetFileName($moqInfo.Source))
                        }
                        $qtyOut = [string]$newQty
                    }
                }
                else {
                    $moqMissing += ("{0}/{1}/{2}" -f $bu, $site, $rec.ITEM_NO)
                }
            }

            $arr[$r,0] = $rec.BU
            $arr[$r,1] = $rec.VENDOR_CODE
            $arr[$r,2] = $rec.VENDOR_SITE
            $arr[$r,3] = $rec.ITEM_NO
            $arr[$r,4] = $rec.ITEM_REV
            $arr[$r,5] = $rec.DOCK_DATE     # text yyyy/MM/dd
            $arr[$r,6] = $qtyOut
            $arr[$r,7] = $rec.CURRENCY
            $arr[$r,8] = $rec.EXCESS
            $arr[$r,9] = $rec.SHIP_TO
        }

        $endRow = 1 + $n
        $target = $ws.Range("A2:J$endRow")
        # Ép cột ngày (F) & item_no/qty về Text để không bị Excel tự đổi định dạng
        $ws.Range("F2:F$endRow").NumberFormat = "@"
        $ws.Range("D2:D$endRow").NumberFormat = "@"
        $target.Value2 = $arr

        # --- BƯỚC 3: lưu thành Text (Tab delimited) .txt ---
        $txtPath = Join-Path $siteFolder ("{0}.txt" -f $site)
        # SaveAs txt chỉ lưu sheet đang active -> kích hoạt sheet đích trước
        $ws.Activate() | Out-Null
        $wb.SaveAs($txtPath, $xlTextWindows)

        # Đóng workbook (không lưu lại .xlsx gốc)
        $wb.Close($false)

        Write-Host ("   Đã lưu: {0}" -f $txtPath)
    }

    if ($moqMissing.Count -gt 0) {
        Write-Host ""
        Write-Host ("CANH BAO MOQ PO: {0} item khong co trong file MOQ (da giu nguyen so luong):" -f $moqMissing.Count) -ForegroundColor Yellow
        foreach ($missing in @($moqMissing | Select-Object -Unique)) {
            Write-Host ("   - {0}" -f $missing) -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "HOÀN TẤT. Tất cả file .txt nằm trong: $OutputRoot"
}
catch {
    Write-Host ""
    Write-Host "==================== CÓ LỖI XẢY RA ====================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.InvocationInfo) {
        Write-Host ("Tại dòng: {0}" -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Red
    }
    Write-Host "=======================================================" -ForegroundColor Red
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    Write-Host ""
    Read-Host "Nhấn Enter để đóng cửa sổ"
}
