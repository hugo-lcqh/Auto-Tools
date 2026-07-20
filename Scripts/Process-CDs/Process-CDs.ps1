<#
==============================================================================
 Process-CDs.ps1
 ------------------------------------------------------------------------------
 Đọc dữ liệu Pre-CD (báo cáo Thiếu CDS) -> nhóm theo Supplier (vendor) ->
 điền vào template SpreadsheetML (Template_PreCD.xml) -> lưu mỗi vendor thành
 1 file .xml mở được bằng Excel, sắp theo cấu trúc:
       <Output>\<dd.MM-dd.MM.yyyy>\<BU>\<Vendor name>\<Vendor name>.xml

 Quy tắc điền template (sheet "Sheet1"):
   - Cột A  Vendor Code  = mã vendor (vd YAO-I -> 9781)
   - Cột B  Vendor name  = tên đầy đủ (vd YAO-I VIETNAM COMPANY LIMITED)
   - Cột F  Part_No      = mã item
   - Cột H  Date type    = "ETA"
   - Ô  K1               = ngày tạo CDs (ngày chạy script)
   - Ô  K2 trở xuống     = số lượng cần tạo CDs (= cột ThieuCDS của input)
   - Bỏ qua các dòng có ThieuCDS = 0
==============================================================================
 CÁCH DÙNG:
   Đặt cùng thư mục: Process-CDs.ps1 , Template_PreCD.xml , cds_input.txt
   Rồi chạy:  powershell -ExecutionPolicy Bypass -File .\Process-CDs.ps1
==============================================================================
#>

param(
    [string]$TemplatePath = "",
    [string]$InputPath    = "",
    [string]$OutputRoot   = "",
    [string[]]$MoqPath    = @()         # mot hoac nhieu file MOQ; de trong thi co the chon hoac bo qua
)

# Đảm bảo console hiển thị tiếng Việt đúng
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

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
# Template/MOQ van fallback ve file nam canh script de khong dut workflow cu.
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = $AutoToolsConfig.ProcessCDsTemplate
    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        $TemplatePath = Join-Path $ScriptDir "Template_PreCD.xml"
    }
}
if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $InputPath = $AutoToolsConfig.ProcessCDsInput
    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        $InputPath = Join-Path $ScriptDir "cds_input.txt"
    }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = $AutoToolsConfig.ProcessCDsOutput
}

function Resolve-FullPath([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $p }
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    $scriptCandidate = Join-Path $ScriptDir $p
    if (Test-Path -LiteralPath $scriptCandidate) { return $scriptCandidate }
    return (Join-Path $AutoToolsPaths.Root $p)
}

function ConvertTo-Bool([string]$value, [bool]$defaultValue = $false) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $defaultValue }

    switch -Regex ($value.Trim()) {
        '^(?i:true|yes|y|1|x)$'  { return $true }
        '^(?i:false|no|n|0)$'    { return $false }
        default                  { return $defaultValue }
    }
}

function Split-ConfigList([string]$value) {
    return @(Split-AutoToolsConfigList -Value $value)
}

function New-DefaultSupplierMaster([string]$path) {
    $csv = @'
Keyword,VendorCode,VendorName,ProcessCDs,RoundMOQ,MOQFile,Combine,ItemPrefix
YAO-I,9781,YAO-I VIETNAM COMPANY LIMITED,TRUE,TRUE,,FALSE,
XIAN HUA,10603,XIAN HUA VIET NAM INDUSTRIAL COMPANY LIMITED,TRUE,FALSE,,TRUE,3;6
YU XIN,12259,YU XIN VIET NAM COMPANY LIMITED,TRUE,FALSE,,FALSE,
NEW SUN,13313,NEW SUN VIET NAM TECHNOLOGY COMPANY LIMITED,TRUE,FALSE,,FALSE,
KHGEAR,11500,KHGEARS VIETNAM COMPANY LIMITED,TRUE,FALSE,,FALSE,
MINH THU,10534,MINH THU ACCURATELY MECHANIC COMPANY LIMITED,TRUE,FALSE,,FALSE,
WONDERWARD,12504,WONDERWARD TECHNOLOGY (HONG KONG) LIMITED,TRUE,FALSE,,FALSE,
SHANGHU,13316,VIETNAM SHANGHU ELECTRONICS COMPANY LIMITED,TRUE,FALSE,,FALSE,
WEIDA,10861,"WEIDA (VIETNAM) MANUFACTURING CO.,LTD",TRUE,FALSE,,FALSE,
TALWAY,12225,TALWAY VIET NAM COMPANY LIMITED,TRUE,FALSE,,FALSE,
GREEN,9340,"GREEN (VIET NAM) CO., LTD",TRUE,FALSE,,FALSE,
HAIXING,11126,HAIXING TECHNOLOGY (VIET NAM) COMPANY LIMITED,TRUE,FALSE,,FALSE,
MAGNET,13553,CONG TY TNHH MAGNET JC (VIET NAM),TRUE,FALSE,,FALSE,
NGHIA LONG,12258,NGHIA LONG METAL PRODUCTS COMPANY LIMITED,TRUE,FALSE,,FALSE,
JUN JAM,12728,JUN JAM METAL PRODUCTS TECHNOLOGY LIMITED,TRUE,FALSE,,FALSE,
CHANTING,12448,"CHANTING INTELLIGENT TECHNOLOGY CO., LTD",TRUE,FALSE,,FALSE,
MINH ANH,10981,MINH ANH ELECTRONICS SERVICE-TRADING COMPANY LIMITED,TRUE,FALSE,,FALSE,
NOVATEK,10058,CONG TY TNHH KY THUAT NOVATEK,TRUE,FALSE,,FALSE,
LONG BAO,9826,CTY TNHH KY THUAT VA CONG NGHE LONG BAO VIET NAM,TRUE,FALSE,,FALSE,
'@

    $parent = Split-Path -Path $path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $csv, $enc)
}

function Load-SupplierMaster([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Write-Host "Khong tim thay supplier master, dang tao file mau: $path" -ForegroundColor Yellow
        New-DefaultSupplierMaster -path $path
    }

    $rows = Import-Csv -LiteralPath $path -Encoding UTF8
    $suppliers = @()

    foreach ($row in $rows) {
        if (
            [string]::IsNullOrWhiteSpace($row.Keyword) -or
            [string]::IsNullOrWhiteSpace($row.VendorCode) -or
            [string]::IsNullOrWhiteSpace($row.VendorName)
        ) {
            continue
        }

        if (-not (ConvertTo-Bool $row.ProcessCDs $true)) {
            continue
        }

        $suppliers += [PSCustomObject]@{
            Key        = $row.Keyword.Trim()
            Code       = [int]$row.VendorCode
            Name       = $row.VendorName.Trim()
            RoundMOQ   = ConvertTo-Bool $row.RoundMOQ $false
            MOQFiles   = Split-ConfigList $row.MOQFile
            Combine    = ConvertTo-Bool $row.Combine $false
            ItemPrefix = Split-ConfigList $row.ItemPrefix
        }
    }

    if ($suppliers.Count -eq 0) {
        throw "Khong co supplier hop le trong file: $path"
    }

    return @($suppliers)
}

function Sanitize-FolderName([string]$name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[{0}]" -f [regex]::Escape($invalid)
    return (($name -replace $pattern, ' ').Trim().TrimEnd('.'))
}

# Tìm vendor dựa trên chuỗi Supplier (so khớp keyword không phân biệt hoa thường)
function Find-Vendor([string]$supplier) {
    $up = $supplier.ToUpper()
    foreach ($v in $VendorTable) {
        foreach ($keyword in (Split-ConfigList $v.Key)) {
            if ($up.Contains($keyword.ToUpper())) { return $v }
        }
    }
    return $null
}

function Xml-Escape([string]$s) {
    return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;')
}

$TemplatePath = Resolve-FullPath $TemplatePath
$InputPath    = Resolve-FullPath $InputPath
$OutputRoot   = Resolve-FullPath $OutputRoot
$SupplierMasterPath = Resolve-FullPath $AutoToolsConfig.SupplierMaster
$VendorTable = Load-SupplierMaster -path $SupplierMasterPath

Write-Host "Template : $TemplatePath"
Write-Host "Input    : $InputPath"
Write-Host "Output   : $OutputRoot"
Write-Host "Supplier : $SupplierMasterPath"
Write-Host ""

if (-not (Test-Path $TemplatePath)) { throw "Không tìm thấy template: $TemplatePath" }
if (-not (Test-Path $InputPath))    { throw "Không tìm thấy input: $InputPath" }

# ============================================================================
# Nạp bảng MOQ (Full carton qty) — dùng để làm tròn số cho vendor có RoundMOQ
# ============================================================================
# Hộp thoại chọn nhiều file MOQ. Trả về danh sách đường dẫn hoặc mảng rỗng nếu huỷ.
function Select-MoqFiles {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title  = "Chọn một hoặc nhiều file MOQ (Full carton qty)"
    $dlg.Filter = "Excel/CSV/Text (*.xlsx;*.xls;*.csv;*.txt)|*.xlsx;*.xls;*.csv;*.txt|Tất cả (*.*)|*.*"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return @($dlg.FileNames) }
    return @()
}

# Làm sạch chuỗi số: bỏ khoảng trắng, dấu phẩy ngăn nghìn -> [int]
function Parse-Qty([string]$s) {
    return (ConvertTo-AutoToolsIntQuantity -Value $s)
}

# Đọc file MOQ -> hashtable: Item -> MOQ (số nguyên).
# Hỗ trợ: .xlsx/.xls (qua Excel COM) và .csv/.txt (tab hoặc phẩy).
function Load-MoqTable([string]$path) {
    $table = @{}
    $ext = [System.IO.Path]::GetExtension($path).ToLower()

    if ($ext -eq '.xlsx' -or $ext -eq '.xls') {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        try {
            $wb = $xl.Workbooks.Open($path)
            $ws = $wb.Worksheets.Item(1)
            $used = $ws.UsedRange
            $rowsN = $used.Rows.Count
            for ($r = 1; $r -le $rowsN; $r++) {
                $itemRaw = $ws.Cells.Item($r,1).Text
                $qtyRaw  = $ws.Cells.Item($r,2).Text
                if ($itemRaw -notmatch '^\s*\d{4,}\s*$') { continue }  # bỏ header / dòng lạ
                $qty = Parse-Qty $qtyRaw
                if ($null -ne $qty -and $qty -gt 0) { $table[$itemRaw.Trim()] = $qty }
            }
            $wb.Close($false)
        } finally {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    } else {
        # CSV/TXT: tách theo Tab nếu có, ngược lại theo dấu phẩy ĐẦU TIÊN
        foreach ($ln in (Get-Content -Path $path -Encoding UTF8)) {
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            if ($ln.Contains("`t")) { $cols = $ln -split "`t" }
            else                    { $cols = $ln -split ',', 3 }  # giữ qty có thể chứa phẩy? -> dùng tab tốt hơn
            if ($cols.Count -lt 2) { continue }
            $itemRaw = $cols[0].Trim()
            if ($itemRaw -notmatch '^\d{4,}$') { continue }
            $qty = Parse-Qty $cols[1]
            if ($null -ne $qty -and $qty -gt 0) { $table[$itemRaw] = $qty }
        }
    }
    return $table
}

function Load-MoqTables([string[]]$paths, [string]$label) {
    $merged = @{}
    $uniquePaths = @(
        $paths |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { Resolve-FullPath $_ } |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
            Select-Object -Unique
    )

    foreach ($path in $uniquePaths) {
        Write-Host ("Nap MOQ {0}: {1}" -f $label, $path) -ForegroundColor Gray
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
        Write-Host ("Da nap {0} ma MOQ cho {1}." -f $merged.Count, $label) -ForegroundColor Gray
    }

    return $merged
}

function Get-ConfiguredMoqPaths {
    [CmdletBinding()]
    param(
        [object[]]$Suppliers
    )

    $paths = @()

    foreach ($pathValue in @($MoqPath)) {
        foreach ($path in (Split-ConfigList $pathValue)) {
            $paths += $path
        }
    }

    if (Test-Path -LiteralPath $AutoToolsConfig.ProcessCDsMoq -PathType Leaf) {
        $paths += $AutoToolsConfig.ProcessCDsMoq
    }

    foreach ($supplier in $Suppliers) {
        foreach ($path in @($supplier.MOQFiles)) {
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                $paths += $path
            }
        }
    }

    return @($paths | Select-Object -Unique)
}

# Làm tròn LÊN bội số gần nhất của MOQ
function RoundUp-ToMoq([int]$qty, [int]$moq) {
    if ($moq -le 0) { return $qty }
    return ([math]::Ceiling($qty / [double]$moq)) * $moq
}

# Co vendor nao can lam tron MOQ khong?
$needMoq = @($VendorTable | Where-Object { $_.RoundMOQ }).Count -gt 0
$GlobalMoqTable = @{}
$MoqTablesByVendor = @{}

if ($needMoq) {
    $configuredMoqPaths = @(Get-ConfiguredMoqPaths -Suppliers $VendorTable)

    if ($configuredMoqPaths.Count -eq 0) {
        $answer = Read-Host "Co supplier can MOQ. Ban co muon chon file MOQ khong? (Y/N, Enter = bo qua)"
        if ($answer -eq 'Y' -or $answer -eq 'y') {
            $configuredMoqPaths = @(Select-MoqFiles)
        }
    }

    if ($configuredMoqPaths.Count -eq 0) {
        Write-Host "[Canh bao] Khong co file MOQ -> se KHONG lam tron, giu nguyen so luong." -ForegroundColor Yellow
    }
    else {
        $globalPaths = @()
        foreach ($pathValue in @($MoqPath)) {
            foreach ($path in (Split-ConfigList $pathValue)) {
                $globalPaths += $path
            }
        }

        if (Test-Path -LiteralPath $AutoToolsConfig.ProcessCDsMoq -PathType Leaf) {
            $globalPaths += $AutoToolsConfig.ProcessCDsMoq
        }

        if ($globalPaths.Count -eq 0) {
            $globalPaths = @(
                $configuredMoqPaths |
                    Where-Object {
                        $path = $_
                        -not @($VendorTable | ForEach-Object { $_.MOQFiles } | ForEach-Object { $_ }).Contains($path)
                    }
            )
        }

        $GlobalMoqTable = Load-MoqTables -paths $globalPaths -label "chung"

        foreach ($supplier in $VendorTable) {
            if ($supplier.MOQFiles.Count -gt 0) {
                $MoqTablesByVendor[$supplier.Key] = Load-MoqTables -paths $supplier.MOQFiles -label $supplier.Key
            }
        }

        if ($GlobalMoqTable.Count -eq 0 -and $MoqTablesByVendor.Count -eq 0) {
            Write-Host "[Canh bao] Khong nap duoc ma MOQ nao -> se giu nguyen so luong." -ForegroundColor Yellow
        }
    }

    Write-Host ""
}

# ----- Ngày tạo CDs = hôm nay (định dạng SpreadsheetML DateTime) -----
$today    = Get-Date
$todayISO = $today.ToString("yyyy-MM-ddT00:00:00.000")
$weekStart = $today.Date.AddDays(-1 * [int]$today.DayOfWeek)
$weekEnd = $weekStart.AddDays(6)
$weekFolderName = "{0}-{1}" -f $weekStart.ToString("dd.MM"), $weekEnd.ToString("dd.MM.yyyy")
$OutputRoot = Join-Path -Path $OutputRoot -ChildPath $weekFolderName
foreach ($buFolder in @("TC5", "TN5")) {
    New-Item -ItemType Directory -Path (Join-Path $OutputRoot $buFolder) -Force | Out-Null
}
Write-Host ("Folder tuần làm việc: {0}" -f $OutputRoot)
Write-Host ("Ngày tạo CDs (ô K1): {0}" -f $today.ToString('yyyy/MM/dd'))
Write-Host ""

# ============================================================================
# BƯỚC 0: Đọc & phân tích input (cột căn lề bằng khoảng trắng)
# ============================================================================
$lines = Get-Content -Path $InputPath -Encoding UTF8
$hasThieuPOCdsColumn = [bool]($lines | Where-Object { $_ -match '\bThieuPO_CDS\b' } | Select-Object -First 1)
$records = @()
foreach ($ln in $lines) {
    if ([string]::IsNullOrWhiteSpace($ln)) { continue }
    if ($ln -match '^\s*BU\b')              { continue }   # header (cột đầu là BU)
    if ($ln -match '^\s*Supplier\b')        { continue }   # header kiểu cũ (không BU)
    if ($ln -match '^\s*-{3,}')             { continue }   # dòng gạch ----

    # Cấu trúc dòng:  BU  Supplier...  Item  <các số>  ThieuCDS ThieuHang [ThieuPO_CDS]  CanhBao(text)
    # BU = token đầu tiên (không chứa khoảng trắng).
    $trim = $ln.TrimStart()
    $sp   = $trim.IndexOf(' ')
    if ($sp -lt 0) { continue }
    $bu        = $trim.Substring(0, $sp).Trim()
    $afterBU   = $trim.Substring($sp)   # phần còn lại: Supplier... Item <số...> text

    # Item = token toàn số >= 6 chữ số đầu tiên; Supplier = phần giữa BU và Item
    $m = [regex]::Match($afterBU, '\s(\d{6,})\s')
    if (-not $m.Success) { continue }

    $supplier = $afterBU.Substring(0, $m.Index).Trim()
    $afterItem = $afterBU.Substring($m.Index).Trim()
    # Phần số nằm trước cột CanhBao (text). Cắt tại ký tự chữ cái đầu tiên.
    $letterIdx = [regex]::Match($afterItem, '[A-Za-z]').Index
    if ($letterIdx -gt 0) { $numPart = $afterItem.Substring(0, $letterIdx) }
    else                  { $numPart = $afterItem }

    # Lấy toàn bộ số: nums[0]=Item, sau đó các cột số.
    # ThieuCDS = số áp chót với input cũ, hoặc số thứ 3 từ phải khi có thêm ThieuPO_CDS.
    $nums = [regex]::Matches($numPart, '-?\d+') | ForEach-Object { $_.Value }
    if ($nums.Count -lt 3) { continue }   # cần ít nhất Item + ThieuCDS + ThieuHang
    $item     = $nums[0]
    $thieuCDS = [int]$nums[$nums.Count - $(if ($hasThieuPOCdsColumn) { 3 } else { 2 })]

    if ($thieuCDS -le 0) { continue }   # bỏ dòng ThieuCDS = 0

    $records += [PSCustomObject]@{
        BU       = $bu
        Supplier = $supplier
        Item     = $item
        Qty      = $thieuCDS
    }
}

if ($records.Count -eq 0) { throw "Không có dòng dữ liệu hợp lệ (ThieuCDS > 0)." }
Write-Host ("Đã đọc {0} dòng (ThieuCDS > 0)." -f $records.Count)

# Nhóm theo BU + Supplier
$groups = $records | Group-Object -Property BU, Supplier
Write-Host ("Số nhóm (BU/supplier): {0}" -f $groups.Count)
Write-Host ""

# ============================================================================
# Tách template thành 3 phần: HEAD | (header row) | TAIL
# HEAD  = mọi thứ trước <Row> đầu tiên (gồm Styles, Column...)
# TAIL  = từ </Table> của Sheet1 trở đi
# Header row + data rows sẽ được sinh lại.
# ============================================================================
$tpl = Get-Content -Path $TemplatePath -Raw -Encoding UTF8

# Cắt HEAD: tới ngay trước <Row đầu tiên (dòng header)
$firstRowIdx = $tpl.IndexOf("`r`n   <Row")
if ($firstRowIdx -lt 0) { $firstRowIdx = $tpl.IndexOf("<Row") }
$head = $tpl.Substring(0, $firstRowIdx)

# Cắt TAIL: từ vị trí "  </Table>" đầu tiên (đóng Table của Sheet1)
$tableCloseIdx = $tpl.IndexOf("  </Table>")
$tail = $tpl.Substring($tableCloseIdx)   # bắt đầu bằng "  </Table>..."

# Header row (12 cell A..L). K1 = ngày tạo CDs.
function Build-HeaderRow([string]$dateISO) {
    $nl = "`r`n"
    $s = "   <Row ss:AutoFitHeight=`"0`">$nl"
    $s += "    <Cell ss:StyleID=`"s65`"><Data ss:Type=`"String`">Vendor Code</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s66`"><Data ss:Type=`"String`">Vendor name</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s67`"><Data ss:Type=`"String`">location</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s68`"><Data ss:Type=`"String`">line no.</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s69`"><Data ss:Type=`"String`">Project Type</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s70`"><Data ss:Type=`"String`">Part_No</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s71`"><Data ss:Type=`"String`">Rev.</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s65`"><Data ss:Type=`"String`">Date type(ETD/ETA)</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s72`"><Data ss:Type=`"String`">SO</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s72`"><Data ss:Type=`"String`">Remark</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s73`"><Data ss:Type=`"DateTime`">$dateISO</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s74`"/>$nl"
    $s += "   </Row>$nl"
    return $s
}

# Data row: A=code B=name F=part H=ETA K=qty
function Build-DataRow($code, $name, $part, $qty) {
    $nl = "`r`n"
    $nameEsc = Xml-Escape $name
    $s = "   <Row ss:AutoFitHeight=`"0`">$nl"
    $s += "    <Cell ss:StyleID=`"s75`"><Data ss:Type=`"Number`">$code</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s76`"><Data ss:Type=`"String`">$nameEsc</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s77`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s77`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s78`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s77`"><Data ss:Type=`"Number`">$part</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s78`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s77`"><Data ss:Type=`"String`">ETA</Data></Cell>$nl"
    $s += "    <Cell ss:StyleID=`"s78`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s78`"/>$nl"
    $s += "    <Cell ss:StyleID=`"s79`"><Data ss:Type=`"Number`">$qty</Data></Cell>$nl"
    $s += "   </Row>$nl"
    return $s
}

# ============================================================================
# Sinh file cho từng vendor
# ============================================================================
$unknown = @()
$combineNotes = @()   # tích luỹ thông báo cho vendor cần combine (vd XIAN HUA)
$moqMissing  = @()   # mã của vendor RoundMOQ nhưng không có trong file MOQ
foreach ($g in $groups) {
    $bu       = $g.Group[0].BU
    $supplier = $g.Group[0].Supplier
    $rows = $g.Group
    $v = Find-Vendor $supplier

    if ($null -eq $v) {
        Write-Host ("-> [BỎ QUA] Không nhận dạng được vendor: '{0}' (BU={1})" -f $supplier, $bu) -ForegroundColor Yellow
        $unknown += ("{0}/{1}" -f $bu, $supplier)
        continue
    }

    # --- Vendor co item can COMBINE: tach rieng item combine, item con lai van tao file ---
    if ($v.Combine) {
        # Neu ItemPrefix co gia tri, chi combine cac item co dau ma nam trong danh sach.
        # Neu ItemPrefix rong, combine toan bo supplier va khong tao file rieng.
        $prefixes = @($v.ItemPrefix)
        if ($prefixes.Count -gt 0) {
            $combineRows = @($rows | Where-Object {
                $first = $_.Item.Substring(0,1)
                $prefixes -contains $first
            })
            $outputRows = @($rows | Where-Object {
                $first = $_.Item.Substring(0,1)
                -not ($prefixes -contains $first)
            })
        }
        else {
            $combineRows = @($rows)
            $outputRows = @()
        }

        if ($combineRows.Count -gt 0) {
            Write-Host ("-> {0} / {1}  =>  {2}  [COMBINE {3} item]" -f $bu, $supplier, $v.Name, $combineRows.Count) -ForegroundColor Cyan
            $itemList = @($combineRows | ForEach-Object { "{0} (SL: {1})" -f $_.Item, $_.Qty })
            $combineNotes += [PSCustomObject]@{
                BU     = $bu
                Vendor = $v.Name
                Items  = $itemList
            }
            foreach ($it in $itemList) { Write-Host ("     - {0}" -f $it) }
        }
        else {
            Write-Host ("     (Không có mã nào đầu {0})" -f ($prefixes -join '/'))
        }

        if ($outputRows.Count -eq 0) {
            Write-Host ("     Khong tao file rieng vi tat ca item cua supplier nay thuoc nhom combine.") -ForegroundColor Cyan
            continue
        }

        Write-Host ("     Con {0} item khong thuoc nhom combine -> van tao file CDs." -f $outputRows.Count) -ForegroundColor Cyan
        $rows = $outputRows
    }

    Write-Host ("-> {0} / {1}  =>  {2} ({3})  [{4} item]" -f $bu, $supplier, $v.Name, $v.Code, $rows.Count)

    $activeMoqTable = $GlobalMoqTable
    if ($MoqTablesByVendor.ContainsKey($v.Key) -and $MoqTablesByVendor[$v.Key].Count -gt 0) {
        $activeMoqTable = $MoqTablesByVendor[$v.Key]
    }

    $doRound = ($v.RoundMOQ -and $activeMoqTable.Count -gt 0)

    # Build các data row
    $dataRows = ""
    foreach ($r in $rows) {
        $qty = $r.Qty
        if ($doRound) {
            if ($activeMoqTable.ContainsKey($r.Item)) {
                $moqInfo = $activeMoqTable[$r.Item]
                $moq    = [int]$moqInfo.Qty
                $newQty = [int](RoundUp-ToMoq $qty $moq)
                if ($newQty -ne $qty) {
                    Write-Host ("     MOQ: {0}  {1} -> {2}  (MOQ={3}, file={4})" -f $r.Item, $qty, $newQty, $moq, [System.IO.Path]::GetFileName($moqInfo.Source))
                }
                $qty = $newQty
            } else {
                Write-Host ("     [Cảnh báo] {0} không có trong file MOQ -> giữ nguyên {1}" -f $r.Item, $qty) -ForegroundColor Yellow
                $moqMissing += ("{0}/{1}/{2}" -f $bu, $v.Name, $r.Item)
            }
        }
        $dataRows += Build-DataRow $v.Code $v.Name $r.Item $qty
    }

    # Cập nhật ExpandedRowCount = 1 (header) + số data row
    $rowCount = 1 + $rows.Count
    $headFixed = [regex]::Replace($head, 'ss:ExpandedRowCount="\d+"', "ss:ExpandedRowCount=`"$rowCount`"")

    $xml = $headFixed + (Build-HeaderRow $todayISO) + $dataRows + $tail

    # Tạo folder Output\<tuần>\<BU>\<tên vendor> + lưu
    $buFolder     = Sanitize-FolderName $bu
    $folderName   = Sanitize-FolderName $v.Name
    $vendorFolder = Join-Path (Join-Path $OutputRoot $buFolder) $folderName
    if (-not (Test-Path $vendorFolder)) {
        New-Item -ItemType Directory -Path $vendorFolder -Force | Out-Null
    }
    $outPath = Join-Path $vendorFolder ("{0}.xml" -f $folderName)

    # Ghi UTF-8 (không BOM) để Excel mở đúng SpreadsheetML
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($outPath, $xml, $enc)

    Write-Host ("   Đã lưu: {0}" -f $outPath)
}

Write-Host ""
if ($unknown.Count -gt 0) {
    Write-Host ("CẢNH BÁO: {0} supplier chưa map được: {1}" -f $unknown.Count, ($unknown -join ', ')) -ForegroundColor Yellow
    Write-Host "Hay bo sung supplier vao file Config\suppliers.csv." -ForegroundColor Yellow
    Write-Host ""
}

# --- Tổng hợp các mã cần COMBINE (nhắn chị Mai) ---
if ($combineNotes.Count -gt 0) {
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ' Nhắn tin cho chị "Tina Le Thi Mai (VN.OP-SC-PMC)" để nhờ combine khai CDs chung' -ForegroundColor Cyan
    Write-Host " Các mã cần khai CDs (đầu 3 hoặc 6):" -ForegroundColor Cyan
    foreach ($note in $combineNotes) {
        Write-Host ("   [{0}] {1}" -f $note.BU, $note.Vendor) -ForegroundColor Cyan
        foreach ($it in $note.Items) {
            Write-Host ("      - {0}" -f $it) -ForegroundColor Cyan
        }
    }
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ""
}

# --- Cảnh báo các mã (vendor RoundMOQ) không tìm thấy trong file MOQ ---
if ($moqMissing.Count -gt 0) {
    Write-Host ("CẢNH BÁO MOQ: {0} mã không có trong file MOQ (đã giữ nguyên số lượng, không làm tròn):" -f $moqMissing.Count) -ForegroundColor Yellow
    foreach ($mm in $moqMissing) { Write-Host ("   - {0}" -f $mm) -ForegroundColor Yellow }
    Write-Host ""
}

Write-Host "HOÀN TẤT. Các file CDs nằm trong: $OutputRoot"
Write-Host ""
Read-Host "Nhấn Enter để đóng cửa sổ"
