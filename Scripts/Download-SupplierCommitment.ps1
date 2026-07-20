# =====================================================================
#  Download Supplier Commitment (SF - Forecast chuan)
#  - Duyet 2 folder Outlook: TC5 - Cu Chi / TN5 - Dau Giay
#  - Tai file .xls dinh kem cua email tu Thu 2 dau tuan -> hom nay
#  - Doc ten supplier tu CATEGORY: "Vendor_<ma> - <TEN>"
#  - Doi ten file: "<ma> - <ten> - <BU>.xls"
#  - Luu vao folder tuan lam viec, ben trong co TC5 / TN5 (trung ten -> ghi de)
#  - Xuat report supplier/BU chua co data theo Config\suppliers.csv
# =====================================================================

# ----------------- CAU HINH (chinh o day) ----------------------------
. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")
$AutoToolsPaths  = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

# Folder dich trong Auto Tools, se tao folder tuan lam viec va 2 folder con TC5/TN5 ben trong
$DestRoot = $AutoToolsConfig.SupplierCommitmentOutput
$SupplierMasterPath = $AutoToolsConfig.SupplierMaster

# Duong dan folder Outlook (theo cay thu muc trong anh).
# Sua lai ten cho khop chinh xac voi Outlook cua ban neu khac.
# Dinh dang: "<Ten store/account>\Inbox\Supplier Commitment\<BU>\SF (Forecast chuan)"
# De trong $StoreName de tu dong dung tai khoan mac dinh.
$StoreName = ""   # vd: "ten.email@tti.com.hk" ; de trong = mailbox mac dinh

$Folders = @(
    @{ BU = "TC5"; Path = @("Supplier Commitment", "TC5 - Củ Chi",  "SF (Forecast chuẩn)") },
    @{ BU = "TN5"; Path = @("Supplier Commitment", "TN5 - Dầu Giây", "SF (Forecast chuẩn)") }
)

# Chi tai file dinh kem co duoi nay
$AttachmentExt = @(".xls", ".xlsx")
# ---------------------------------------------------------------------

$ErrorActionPreference = "Stop"

function Write-Log($msg, $color = "White") {
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg) -ForegroundColor $color
}

# --- Tinh ngay Thu 2 cua tuan hien tai (00:00) ---
$today = Get-Date
# DayOfWeek: Sunday=0, Monday=1 ... Saturday=6
$dow = [int]$today.DayOfWeek
if ($dow -eq 0) { $dow = 7 }              # coi Chu Nhat la cuoi tuan
$monday = $today.Date.AddDays(-1 * ($dow - 1))
$weekStart = $monday.AddDays(-1)          # folder tuan theo dang Chu Nhat -> Thu Bay, vd 28.06-04.07.2026
$weekEnd = $weekStart.AddDays(6)
$weekFolderName = "{0}-{1}" -f $weekStart.ToString("dd.MM"), $weekEnd.ToString("dd.MM.yyyy")
$DestRoot = Join-Path -Path $DestRoot -ChildPath $weekFolderName
New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null

Write-Log "Tai email tu $($monday.ToString('dd/MM/yyyy')) den $($today.ToString('dd/MM/yyyy'))" "Cyan"
Write-Log "Folder tuan lam viec: $weekFolderName" "Cyan"

# --- Lam sach ten file/folder ---
function Clean-Name($name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $name = $name.Replace($c, ' ') }
    return ($name -replace '\s+', ' ').Trim()
}

# --- Lay ten supplier tu category "Vendor_<ma> - <TEN>" ---
# Tra ve hashtable @{ Code=...; Name=...; Display="ma - ten" } hoac $null
function Parse-Vendor($categories) {
    if ([string]::IsNullOrWhiteSpace($categories)) { return $null }
    # Categories co the co nhieu cai, ngan cach boi dau phay
    foreach ($cat in ($categories -split ',')) {
        $c = $cat.Trim()
        # Khop dang: Vendor_10603 - XIAN HUA VIET NAM
        if ($c -match '(?i)vendor[_\s]*(\d+)\s*-\s*(.+)') {
            return @{
                Code    = $matches[1].Trim()
                Name    = $matches[2].Trim()
                Display = ("{0} - {1}" -f $matches[1].Trim(), $matches[2].Trim())
            }
        }
    }
    return $null
}

# --- Tim folder con theo duong dan ten ---
function Get-SubFolder($parent, $pathArray) {
    $cur = $parent
    foreach ($name in $pathArray) {
        $found = $null
        foreach ($f in $cur.Folders) {
            if ($f.Name -eq $name) { $found = $f; break }
        }
        if ($null -eq $found) {
            # thu khop khong phan biet hoa thuong / khoang trang thua
            foreach ($f in $cur.Folders) {
                if ($f.Name.Trim().ToLower() -eq $name.Trim().ToLower()) { $found = $f; break }
            }
        }
        if ($null -eq $found) {
            throw "Khong tim thay folder con: '$name'"
        }
        $cur = $found
    }
    return $cur
}

# --- Ket noi Outlook ---
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
} catch {
    Write-Log "Khong the ket noi Outlook. Hay mo Outlook truoc khi chay." "Red"
    Read-Host "Nhan Enter de thoat"
    exit 1
}

# --- Tim Inbox goc ---
if ([string]::IsNullOrWhiteSpace($StoreName)) {
    $inbox = $ns.GetDefaultFolder(6)   # 6 = olFolderInbox
} else {
    $store = $null
    foreach ($s in $ns.Stores) {
        if ($s.DisplayName -eq $StoreName -or $s.DisplayName -like "*$StoreName*") { $store = $s; break }
    }
    if ($null -eq $store) {
        Write-Log "Khong tim thay tai khoan '$StoreName'. Dung mailbox mac dinh." "Yellow"
        $inbox = $ns.GetDefaultFolder(6)
    } else {
        $inbox = $store.GetRootFolder().Folders.Item("Inbox")
    }
}

$totalSaved = 0
$totalSkip  = 0
$downloadedRows = New-Object System.Collections.Generic.List[object]

function Normalize-SupplierKey($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    $value = $name.ToUpperInvariant()
    $value = $value -replace '&', ' AND '
    return ($value -replace '[^A-Z0-9]+', '')
}

function Test-SupplierMatched($supplier, $downloadedForBu) {
    $expectedCode = [string]$supplier.VendorCode
    if (-not [string]::IsNullOrWhiteSpace($expectedCode)) {
        foreach ($row in $downloadedForBu) {
            if ([string]$row.VendorCode -eq $expectedCode.Trim()) { return $true }
        }
    }

    $keys = New-Object System.Collections.Generic.List[string]
    $vendorNameKey = Normalize-SupplierKey ([string]$supplier.VendorName)
    if (-not [string]::IsNullOrWhiteSpace($vendorNameKey)) { $keys.Add($vendorNameKey) }

    if (-not [string]::IsNullOrWhiteSpace($supplier.Keyword)) {
        foreach ($keyword in ([string]$supplier.Keyword -split ';')) {
            $keywordKey = Normalize-SupplierKey $keyword
            if (-not [string]::IsNullOrWhiteSpace($keywordKey)) { $keys.Add($keywordKey) }
        }
    }

    foreach ($row in $downloadedForBu) {
        $downloadedKey = Normalize-SupplierKey ([string]$row.VendorName)
        if ([string]::IsNullOrWhiteSpace($downloadedKey)) { continue }

        foreach ($key in $keys) {
            if ($downloadedKey.Contains($key) -or $key.Contains($downloadedKey)) { return $true }
        }
    }

    return $false
}

function Export-MissingSupplierCommitmentReport($masterPath, $outputDirectory, $bus, $downloaded) {
    if (-not (Test-Path -LiteralPath $masterPath -PathType Leaf)) {
        Write-Log "Khong tim thay file master supplier: $masterPath" "Yellow"
        return
    }

    $suppliers = @(Import-Csv -LiteralPath $masterPath)
    if ($suppliers.Count -eq 0) {
        Write-Log "File master supplier khong co data: $masterPath" "Yellow"
        return
    }

    $missing = New-Object System.Collections.Generic.List[object]
    foreach ($bu in $bus) {
        $downloadedForBu = @($downloaded | Where-Object { $_.BU -eq $bu })
        foreach ($supplier in $suppliers) {
            if ([string]::IsNullOrWhiteSpace($supplier.VendorName)) { continue }

            if (-not (Test-SupplierMatched $supplier $downloadedForBu)) {
                $missing.Add([PSCustomObject]@{
                    BU         = $bu
                    VendorCode = $supplier.VendorCode
                    VendorName = $supplier.VendorName
                    Keyword    = $supplier.Keyword
                    Status     = "Missing Supplier Commitment data file"
                })
            }
        }
    }

    $reportPath = Join-Path -Path $outputDirectory -ChildPath ("MissingSupplierCommitment_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $missing | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8

    if ($missing.Count -eq 0) {
        Write-Log "Tat ca supplier trong suppliers.csv deu da co Supplier Commitment data." "Green"
    } else {
        Write-Log "Con $($missing.Count) dong supplier/BU chua co data. Report: $reportPath" "Yellow"
    }
}

foreach ($cfg in $Folders) {
    $bu = $cfg.BU
    Write-Log "----- Xu ly BU: $bu -----" "Yellow"

    # Tao folder dich
    $destDir = Join-Path $DestRoot $bu
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    # Lay folder Outlook
    try {
        $folder = Get-SubFolder $inbox $cfg.Path
    } catch {
        Write-Log "  Bo qua $bu : $_" "Red"
        continue
    }

    # Loc email theo ngay bang Restrict cho nhanh
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    $filter = "[ReceivedTime] >= '" + $monday.ToString("MM/dd/yyyy HH:mm") + "'"
    $mails = $items.Restrict($filter)

    Write-Log "  Tim thay $($mails.Count) email trong khoang thoi gian." "Gray"

    foreach ($mail in $mails) {
        # Chi xu ly email thuc su (MailItem)
        if ($mail.Class -ne 43) { continue }   # 43 = olMail

        $vendor = Parse-Vendor $mail.Categories
        if ($null -eq $vendor) {
            Write-Log "  ! Bo qua 1 email khong co category Vendor (Subject: $($mail.Subject))" "DarkYellow"
            continue
        }

        if ($mail.Attachments.Count -eq 0) { continue }

        foreach ($att in $mail.Attachments) {
            $origName = $att.FileName
            $ext = [System.IO.Path]::GetExtension($origName).ToLower()

            # Mot so attachment (vd file nhung/anh chu ky) khong phai Excel -> bo qua.
            # Neu extension rong nhung ten co chua "commitment"/".xls" thi van coi la Excel.
            if ([string]::IsNullOrWhiteSpace($ext)) {
                if ($origName -match '(?i)\.xlsx?($|\s)') {
                    $ext = if ($origName -match '(?i)\.xlsx') { ".xlsx" } else { ".xls" }
                } else {
                    continue
                }
            }
            if ($AttachmentExt -notcontains $ext) { continue }

            # Tao ten moi: dat $ext rieng de khong bao gio mat duoi file
            $baseName = Clean-Name("$($vendor.Display) - $bu")
            $newName  = $baseName + $ext
            $savePath = Join-Path $destDir $newName

            Write-Log "    (goc: $origName)" "DarkGray"

            try {
                $att.SaveAsFile($savePath)   # ghi de neu trung
                Write-Log "  OK  $newName" "Green"
                $totalSaved++
                $downloadedRows.Add([PSCustomObject]@{
                    BU           = $bu
                    VendorCode   = $vendor.Code
                    VendorName   = $vendor.Name
                    Subject      = $mail.Subject
                    Attachment   = $origName
                    SavedAs      = $savePath
                    ReceivedTime = $mail.ReceivedTime
                })
            } catch {
                Write-Log "  Loi luu $newName : $_" "Red"
            }
        }
    }
}

if ($downloadedRows.Count -gt 0) {
    $logPath = Join-Path -Path $DestRoot -ChildPath ("DownloadLog_SupplierCommitment_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $downloadedRows | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
    Write-Log "Da ghi log: $logPath" "Gray"
}

Export-MissingSupplierCommitmentReport `
    -masterPath $SupplierMasterPath `
    -outputDirectory $DestRoot `
    -bus @($Folders | ForEach-Object { $_.BU }) `
    -downloaded $downloadedRows

Write-Log "===== HOAN TAT: da luu $totalSaved file =====" "Cyan"
Write-Log "Folder dich: $DestRoot" "Cyan"

# Giai phong COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

Read-Host "Nhan Enter de thoat"
