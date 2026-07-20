<#
====================================================================
 CheckCD_Shortage.ps1
 Quet toan bo cac sheet supplier trong file working, kiem tra:
   1. THIEU CDS : demand(N tuan) - (CD Open Qty + Pre CD) > 0
   2. THIEU HANG: demand(N tuan) - Vendor PO QTY > 0
   3. THIEU PO CHO CDS: Vendor PO QTY - CD Open Qty - Pre CD - ThieuCDS < 0
 Xuat ra 1 file report Excel liet ke item/supplier co van de.
 --------------------------------------------------------------------
 Chay doc lap, KHONG ghi de gi vao file working (chi doc).
====================================================================
#>

# ====================== CAU HINH ======================
. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")
$AutoToolsPaths  = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

$RootFolder  = $AutoToolsConfig.JobShortRoot
$WorkingFile = "20240410 Jun JS.xlsb"
$HeaderRow   = 6     # dong chua tieu de cot
$DataStart   = 7     # dong dau tien co data
$SupplierMasterPath = if ([System.IO.Path]::IsPathRooted($AutoToolsConfig.SupplierMaster)) {
    $AutoToolsConfig.SupplierMaster
} else {
    Join-Path $AutoToolsPaths.Root $AutoToolsConfig.SupplierMaster
}

# --- ASCP: dung de double-check item co BU_Category = "Shared" ---
$ASCP_RootFolder = $AutoToolsConfig.ASCPRoot
$ASCP_File       = "OPVN_ASCP.xlsb"
$ASCP_Sheet      = "ASCP"
$ASCP_COL_PART   = 2     # B  - Part_No (ma item)
$ASCP_COL_BUYER  = 14    # N  - Buyer (nguoi phu trach)
$ASCP_DataStart  = 11    # dong bat dau du lieu
$TargetBuyer     = "TVN634465, Le Chi Quoc Hung (VN.OP-SUC)"
$SharedTag       = "Shared"   # gia tri cot A can double-check

# Vi tri cot (1-based, theo Excel). Da xac dinh tu cau truc file:
$COL_BU      = 1     # A  - BU_Category
$COL_PLANNER = 2     # B  - Planner_Code
$COL_ITEM    = 4     # D  - Item
$COL_ORGCODE = 3     # C  - ORG_CODE
$COL_VENDORCD= 6     # F  - Vendor code (dung de tra lead time)
$COL_VENDOR  = 7     # G  - Vendor Name
$COL_PO      = 10    # J  - Vendor PO QTY
$COL_CDOPEN  = 11    # K  - CD Open Qty
$COL_PRECD   = 15    # O  - Pre CD
# Cot ngay (demand) bat dau ngay sau cot Pre CD, la cac cot co header dang ngay (serial > 40000)

# Cac sheet KHONG phai supplier (bo qua khi quet)
$SkipSheets  = @("DELIVERY (Confirmed)","OVERVIEW","JS_11365","CD",
                 "Outstanding_11801","JIT_JOB","RAW",
                 "6. HAIXING","15. LONG BAO")
# ======================================================

$ErrorActionPreference = "Stop"
function Say($m,$c="White"){ Write-Host $m -ForegroundColor $c }
function ToNum($v){ $d=0.0; if($v -ne $null -and [double]::TryParse("$v",[ref]$d)){return $d}; return 0.0 }
# Truy cap an toan o (r,c) tu ket qua Range.Value2 (mang 2 chieu hoac gia tri don)
function Get2D($raw,$isArray,$r,$c){
    if ($isArray) { try { return $raw.GetValue($r,$c) } catch { return $null } }
    else { if ($r -eq 1 -and $c -eq 1) { return $raw } else { return $null } }
}
function Load-LeadTimeTable($path) {
    $table = @{}
    foreach ($row in (Import-Csv -LiteralPath $path -Encoding UTF8)) {
        $code = ("" + $row.VendorCode).Trim()
        $rule = ("" + $row.PreCDLeadTime).Trim()
        if ($code -ne "" -and $rule -ne "") { $table[$code] = $rule }
    }
    return $table
}
function Get-LeadTime($table,$vendorCode,$item) {
    if (-not $table.ContainsKey($vendorCode)) { return $null }
    $rule = $table[$vendorCode]
    $lt = 0
    if ([int]::TryParse($rule, [ref]$lt)) { return $lt }

    foreach ($part in (Split-AutoToolsConfigList $rule)) {
        if ($part -match '^([^:]+):(\d+)$' -and $item.StartsWith($matches[1])) {
            return [int]$matches[2]
        }
    }
    return $null
}

# ---------- Bat dau ----------
Say "`n=== CHECK CD & SHORTAGE ===" Cyan
Say "Khung thoi gian check = PreCD Lead Time rieng cua tung vendor.`n" Gray
$LeadTime = Load-LeadTimeTable $SupplierMasterPath
Say "Da nap Lead Time tu: $SupplierMasterPath" Gray

# Tim folder ngay moi nhat (dang dd.MM) trong thu muc goc, lay file working trong do.
# Neu file nam thang trong thu muc goc thi dung luon.
$workingPath = $null
$directPath  = Join-Path $RootFolder $WorkingFile
if (Test-Path $directPath) {
    $workingPath = $directPath
} else {
    $dateFolders = Get-ChildItem -Path $RootFolder -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{2}\.\d{2}$' }
    if ($dateFolders) {
        $latest = $dateFolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $candidate = Join-Path $latest.FullName $WorkingFile
        if (Test-Path $candidate) {
            $workingPath = $candidate
            Say "Dung file working trong folder ngay: $($latest.Name)" Gray
        }
    }
}
if (-not $workingPath) {
    throw "Khong tim thay file '$WorkingFile' trong $RootFolder hoac trong folder ngay (dd.MM) nao."
}

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    # ---------- NAP BANG TRA ASCP (cho item Shared) ----------
    # Tim file ASCP trong folder moi nhat, doc cot Part_No + Buyer vao hashtable.
    # $ascpBuyers[ma_item] = danh sach buyer (1 ma co the co nhieu dong).
    $ascpBuyers = @{}
    $ascpLoaded = $false
    $ascpPath = $null
    $ascpDirect = Join-Path $ASCP_RootFolder $ASCP_File
    if (Test-Path $ascpDirect) {
        $ascpPath = $ascpDirect
    } else {
        $ascpFolders = Get-ChildItem -Path $ASCP_RootFolder -Directory -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending
        foreach ($f in $ascpFolders) {
            $cand = Join-Path (Join-Path $f.FullName "Import") $ASCP_File
            if (Test-Path $cand) { $ascpPath = $cand; break }
        }
    }

    if (-not $ascpPath) {
        Say "[!] CANH BAO: Khong tim thay $ASCP_File trong $ASCP_RootFolder (folder moi nhat\Import)." Yellow
        Say "    Cac item 'Shared' se khong duoc double-check va se bi bo qua." Yellow
    } else {
        Say "Dang nap bang tra ASCP (file lon, co the mat 10-30 giay)..." Gray
        $wbA = $excel.Workbooks.Open($ascpPath, $false, $true)  # ReadOnly
        $shA = $wbA.Worksheets.Item($ASCP_Sheet)
        $usedA = $shA.UsedRange
        $lastRowA = $usedA.Row + $usedA.Rows.Count - 1
        # Doc 2 cot (Part_No, Buyer) trong 1 lan
        $rngPart  = $shA.Range($shA.Cells.Item($ASCP_DataStart,$ASCP_COL_PART),  $shA.Cells.Item($lastRowA,$ASCP_COL_PART)).Value2
        $rngBuyer = $shA.Range($shA.Cells.Item($ASCP_DataStart,$ASCP_COL_BUYER), $shA.Cells.Item($lastRowA,$ASCP_COL_BUYER)).Value2
        $nA = $lastRowA - $ASCP_DataStart + 1
        for ($i = 1; $i -le $nA; $i++) {
            $partRaw = if ($nA -eq 1) { $rngPart } else { $rngPart.GetValue($i,1) }
            if ($null -eq $partRaw -or "$partRaw".Trim() -eq "") { continue }
            $key = ("" + $partRaw).Trim()
            if ($key -match '^\d+\.0+$') { $key = $key -replace '\.0+$','' }   # 658251001.0 -> 658251001
            $buyerRaw = if ($nA -eq 1) { $rngBuyer } else { $rngBuyer.GetValue($i,1) }
            $buyer = ("" + $buyerRaw).Trim()
            if (-not $ascpBuyers.ContainsKey($key)) {
                $ascpBuyers[$key] = New-Object System.Collections.Generic.HashSet[string]
            }
            [void]$ascpBuyers[$key].Add($buyer)
        }
        $wbA.Close($false)
        $ascpLoaded = $true
        Say "Da nap ASCP: $($ascpBuyers.Count) ma item." Gray
    }

    Say "Dang mo file working (co the mat vai giay)..." Gray
    $wb = $excel.Workbooks.Open($workingPath, $false, $true)  # ReadOnly = true

    # Thu thap ket qua vao mang de in ra console (khong tao file Excel)
    $results  = New-Object System.Collections.Generic.List[object]
    $totalCDS = 0; $totalShort = 0; $totalPOForCDS = 0
    $skippedCPT = 0
    $warnLeadTime      = New-Object System.Collections.Generic.List[string]
    $warnSharedNotMine = New-Object System.Collections.Generic.List[string]
    $warnSharedNotFound= New-Object System.Collections.Generic.List[string]

    foreach ($sheet in $wb.Worksheets) {
        $name = $sheet.Name
        if ($SkipSheets -contains $name) { continue }

        $used = $sheet.UsedRange
        $lastRow = $used.Row + $used.Rows.Count - 1
        $lastCol = $used.Column + $used.Columns.Count - 1
        if ($lastRow -lt $DataStart) { continue }

        # --- Tim cac cot ngay (demand): header la so serial ngay (>40000), nam sau cot Pre CD ---
        $dayCols = @()
        for ($c = $COL_PRECD + 1; $c -le $lastCol; $c++) {
            $hvNum = ToNum($sheet.Cells.Item($HeaderRow, $c).Value2)
            if ($hvNum -gt 40000) {
                $dayCols += $c
            } elseif ($dayCols.Count -gt 0) {
                break   # da het block ngay dau tien
            }
        }
        if ($dayCols.Count -eq 0) { continue }
        # Giu toan bo cot ngay; so ngay lay theo lead time cua tung item ben duoi.

        # --- Doc nhanh toan bo vung du lieu 1 lan (nhanh hon doc tung cell) ---
        $nRows = $lastRow - $DataStart + 1
        $dataRange = $sheet.Range($sheet.Cells.Item($DataStart,1), $sheet.Cells.Item($lastRow,$lastCol))
        $raw = $dataRange.Value2

        # Excel COM tra ve kieu khac nhau tuy kich thuoc vung:
        #  - Nhieu o      -> mang 2 chieu [1..nRows, 1..nCols]
        #  - 1 o duy nhat -> gia tri don (khong phai mang)
        $isArray = $raw -is [array]

        for ($r = 1; $r -le $nRows; $r++) {
            $item = Get2D $raw $isArray $r $COL_ITEM
            if ($null -eq $item -or "$item".Trim() -eq "") { continue }

            $bu = "" + (Get2D $raw $isArray $r $COL_BU)
            $buTrim = $bu.Trim()
            $planner = ("" + (Get2D $raw $isArray $r $COL_PLANNER)).Trim().ToUpper()
            if ($planner -match 'CPT' -or $planner -notmatch 'OP') { $skippedCPT++; continue }

            $itemStr   = ("" + $item).Trim()
            if ($itemStr -match '^\d+\.0+$') { $itemStr = $itemStr -replace '\.0+$','' }  # 658251001.0 -> 658251001

            # --- Double-check item "Shared" voi file ASCP ---
            if ($buTrim -eq $SharedTag) {
                if (-not $ascpLoaded) {
                    # Khong nap duoc ASCP -> khong the double-check -> bo qua
                    $warnSharedNotFound.Add("$name | item $itemStr (chua nap duoc ASCP)")
                    continue
                }
                if (-not $ascpBuyers.ContainsKey($itemStr)) {
                    # Khong tim thay ma trong ASCP -> thong bao + bo qua
                    $warnSharedNotFound.Add("$name | item $itemStr")
                    continue
                }
                # Tim thay: kiem tra buyer co dung la minh khong
                if (-not $ascpBuyers[$itemStr].Contains($TargetBuyer)) {
                    # Buyer khac -> khong phai cua minh -> thong bao + bo qua
                    $otherBuyer = ($ascpBuyers[$itemStr] | Select-Object -First 1)
                    $warnSharedNotMine.Add("$name | item $itemStr | buyer: $otherBuyer")
                    continue
                }
                # Buyer dung la minh -> tinh binh thuong (di tiep)
            }

            $po  = ToNum(Get2D $raw $isArray $r $COL_PO)
            $cd  = ToNum(Get2D $raw $isArray $r $COL_CDOPEN)
            $pre = ToNum(Get2D $raw $isArray $r $COL_PRECD)
            if ($po -eq 0 -and $cd -eq 0) { continue }

            # --- Xac dinh lead time (so ngay) cho item theo Vendor Code ---
            $vcRaw  = Get2D $raw $isArray $r $COL_VENDORCD
            $vcode  = ("" + $vcRaw).Trim()
            if ($vcode -match '^\d+\.0+$') { $vcode = $vcode -replace '\.0+$','' }  # 9781.0 -> 9781

            $lt = Get-LeadTime $LeadTime $vcode $itemStr
            if ($null -eq $lt) {
                $warnLeadTime.Add("$name | vendor code '$vcode' | item $itemStr")
                continue
            }

            # Lay dung so cot ngay = lead time (tu dau dai ngay)
            $windowCols = $dayCols | Select-Object -First $lt

            $demand = 0.0
            foreach ($c in $windowCols) {
                $demand += ToNum(Get2D $raw $isArray $r $c)
            }
            if ($demand -le 0) { continue }

            $cdsLack   = $demand - ($cd + $pre)   # thieu CDS
            $shortLack = $demand - $po            # thieu hang
            $thieuCDS  = [math]::Max(0,$cdsLack)
            $poForCdsLack = if ($thieuCDS -gt 0) { [math]::Max(0, -($po - $cd - $pre - $thieuCDS)) } else { 0 }

            if ($cdsLack -gt 0 -or $shortLack -gt 0) {
                $org    = Get2D $raw $isArray $r $COL_ORGCODE
                $alerts = @()
                if ($cdsLack   -gt 0) { $alerts += "Thieu CDS";  $totalCDS++ }
                if ($shortLack -gt 0) { $alerts += "Thieu Hang"; $totalShort++ }
                if ($poForCdsLack -gt 0) { $alerts += "Thieu PO cho CDS"; $totalPOForCDS++ }

                $results.Add([pscustomobject]@{
                    Supplier  = [string]$name
                    OrgCode   = [string]("" + $org)
                    Item      = [string]("" + $item)
                    LT        = [int]$lt
                    PO_QTY    = [double]$po
                    CD_Open   = [double]$cd
                    Pre_CD    = [double]$pre
                    Demand    = [double]$demand
                    ThieuCDS  = [double]$thieuCDS
                    ThieuHang = [double][math]::Max(0,$shortLack)
                    ThieuPO_CDS = [double]$poForCdsLack
                    CanhBao   = [string]($alerts -join " + ")
                })
            }
        }
        Say "  Da quet: $name" Gray
    }

    # Dong file working (chi doc, khong luu)
    $wb.Close($false)
    $excel.Quit()

    # ---------- IN KET QUA RA CONSOLE ----------
    Say "`n========================================" Green
    Say " KET QUA KIEM TRA (theo Lead Time tung vendor)" Green
    Say "========================================" Green
    if ($skippedCPT -gt 0) { Say " (Da bo qua $skippedCPT item do Planner_Code khong phai OP hoac co CPT)" DarkGray }

    if ($results.Count -eq 0) {
        Say "`n Khong co item nao thieu CDS hoac thieu hang. Tot!`n" Green
    } else {
        # Sap xep: theo BU (TC5/TN5), roi thieu hang len truoc, roi theo supplier
        $sorted = $results | Sort-Object OrgCode, @{E={$_.ThieuHang};Descending=$true}, Supplier

        # In dang bang: BU | Supplier | Item | LT | Demand | CD+Pre | PO | ThieuCDS | ThieuHang | ThieuPO_CDS | CanhBao
        $tbl = $sorted | Select-Object `
            @{N='BU';        E={$_.OrgCode}},
            Supplier,
            Item,
            @{N='LT';        E={[int]$_.LT}},
            @{N='Demand';   E={[int]$_.Demand}},
            @{N='CD+Pre';   E={[int]($_.CD_Open + $_.Pre_CD)}},
            @{N='PO';        E={[int]$_.PO_QTY}},
            @{N='ThieuCDS';  E={[int]$_.ThieuCDS}},
            @{N='ThieuHang'; E={[int]$_.ThieuHang}},
            @{N='ThieuPO_CDS'; E={[int]$_.ThieuPO_CDS}},
            CanhBao
        $tbl | Format-Table -AutoSize | Out-String -Width 4096 | Write-Host

        $poCdsMissing = $sorted | Where-Object { $_.ThieuPO_CDS -gt 0 }
        if ($poCdsMissing) {
            Say "[!] Cac ma thieu PO de tao them CDS:" Red
            $poCdsMissing | Select-Object `
                @{N='BU'; E={$_.OrgCode}},
                Supplier,
                Item,
                @{N='ThieuPO_CDS'; E={[int]$_.ThieuPO_CDS}} |
                Format-Table -AutoSize | Out-String -Width 4096 | Write-Host
        }

        Say "----------------------------------------" Gray
        Say " Tong item thieu CDS : $totalCDS" Yellow
        Say " Tong item thieu hang: $totalShort" Red
        Say " Tong item thieu PO de tao CDS: $totalPOForCDS" Red
        Say "========================================`n" Green
    }

    # ---------- IN CANH BAO (neu co) ----------
    if ($warnLeadTime.Count -gt 0) {
        Say "[!] Vendor Code / item KHONG co PreCDLeadTime trong suppliers.csv (da bo qua, can bo sung):" Yellow
        foreach ($w in $warnLeadTime) { Say "    - $w" Yellow }
        Say ""
    }
    if ($warnSharedNotMine.Count -gt 0) {
        Say "[!] Item 'Shared' KHONG phai do ban phu trach (da bo qua):" Yellow
        foreach ($w in $warnSharedNotMine) { Say "    - $w" Yellow }
        Say ""
    }
    if ($warnSharedNotFound.Count -gt 0) {
        Say "[!] Item 'Shared' KHONG tim thay ma trong ASCP (da bo qua):" Yellow
        foreach ($w in $warnSharedNotFound) { Say "    - $w" Yellow }
        Say ""
    }
}
catch {
    Say "`nLOI: $($_.Exception.Message)" Red
    if ($wbA)   { try { $wbA.Close($false) } catch {} }
    if ($wb)    { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
}
finally {
    if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Read-Host "Nhan Enter de dong"
}
