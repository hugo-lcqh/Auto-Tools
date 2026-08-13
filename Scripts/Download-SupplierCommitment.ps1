# =====================================================================
#  Download Supplier Commitment (SF - Forecast chuan)
#  - Uu tien 2 folder Outlook: TC5 - Cu Chi / TN5 - Dau Giay
#  - Neu thieu folder BU, quet Supplier Commitment theo tieu de SF_TC5 / SF_TN5
#  - Tai file .xls dinh kem cua email tu Thu 2 dau tuan -> hom nay
#  - Doc ma supplier tu tieu de va tra ten trong Config\suppliers.csv
#  - Email cu khong dung mau tieu de van doc CATEGORY: "Vendor_<ma> - <TEN>"
#  - Doi ten file: "<ma> - <ten> - <BU>.xls"
#  - Luu vao folder tuan lam viec, ben trong co TC5 / TN5 (trung ten -> ghi de)
#  - Xuat report supplier/BU chua co data theo Config\suppliers.csv
#  - Cho phep chon supplier can nhac (tung supplier hoac tat ca)
#  - Tao Outlook Draft voi Email To/CC tu Config\suppliers.csv, khong tu dong gui
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

# --- Lay BU va ma supplier tu tieu de mail ISPIH ---
function Parse-SupplierCommitmentSubject($subject) {
    if ([string]::IsNullOrWhiteSpace($subject)) { return $null }

    $pattern = '(?i)^TTI\s+ISPIH\s+Alert\s*--\s*MRP\s+Supplier\s+Commitment\(SF_(TC5|TN5)--[^)]*\)\((\d+)\)\s*$'
    if ($subject.Trim() -notmatch $pattern) { return $null }

    return @{
        BU         = $matches[1].ToUpperInvariant()
        VendorCode = $matches[2].Trim()
    }
}

# --- Tao bang ma supplier (cot B) -> ten vendor (cot C) ---
function New-SupplierLookup($suppliers) {
    $lookup = @{}

    foreach ($supplier in $suppliers) {
        $code = ([string]$supplier.VendorCode).Trim()
        $name = ([string]$supplier.VendorName).Trim()
        if ([string]::IsNullOrWhiteSpace($code) -or [string]::IsNullOrWhiteSpace($name)) { continue }

        $lookup[$code] = @{
            Code    = $code
            Name    = $name
            Display = "$code - $name"
        }
    }

    return $lookup
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

# --- Lay folder goc va tat ca folder con de dung khi duong dan BU bi thay doi ---
function Get-OutlookFolderTree($root) {
    $result = New-Object System.Collections.Generic.List[object]
    $pending = New-Object System.Collections.Stack
    $pending.Push($root)

    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        $result.Add($current)

        foreach ($child in $current.Folders) {
            $pending.Push($child)
        }
    }

    return $result.ToArray()
}

$masterSuppliers = @()
if (Test-Path -LiteralPath $SupplierMasterPath -PathType Leaf) {
    try {
        $masterSuppliers = @(Import-Csv -LiteralPath $SupplierMasterPath)
    } catch {
        Write-Log "Khong doc duoc file master supplier: $SupplierMasterPath. $_" "Yellow"
    }
} else {
    Write-Log "Khong tim thay file master supplier: $SupplierMasterPath" "Yellow"
}
$supplierLookup = New-SupplierLookup $masterSuppliers

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
$downloadedRows = New-Object System.Collections.Generic.List[object]

function Normalize-SupplierKey($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    $value = $name.ToUpperInvariant()
    $value = $value -replace '&', ' AND '
    return ($value -replace '[^A-Z0-9]+', '')
}

function ConvertTo-OutlookRecipientList($value) {
    if ([string]::IsNullOrWhiteSpace([string]$value)) { return "" }

    $seen = @{}
    $addresses = New-Object System.Collections.Generic.List[string]
    $cleanValue = ([string]$value).Replace([char]0x00A0, ' ')

    foreach ($address in ($cleanValue -split '[;,]')) {
        $trimmed = $address.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

        $key = $trimmed.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $addresses.Add($trimmed)
        }
    }

    return ($addresses -join '; ')
}

function Get-SupplierReminderCandidates {
    param(
        [object[]]$MissingRows,
        [object[]]$Suppliers
    )

    $supplierByCode = @{}
    $supplierByName = @{}
    foreach ($supplier in $Suppliers) {
        $code = ([string]$supplier.VendorCode).Trim()
        $nameKey = Normalize-SupplierKey ([string]$supplier.VendorName)
        if (-not [string]::IsNullOrWhiteSpace($code)) { $supplierByCode[$code] = $supplier }
        if (-not [string]::IsNullOrWhiteSpace($nameKey)) { $supplierByName[$nameKey] = $supplier }
    }

    $groups = @{}
    foreach ($row in $MissingRows) {
        $code = ([string]$row.VendorCode).Trim()
        $name = ([string]$row.VendorName).Trim()
        $nameKey = Normalize-SupplierKey $name
        $key = if (-not [string]::IsNullOrWhiteSpace($code)) { "CODE:$code" } else { "NAME:$nameKey" }

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [PSCustomObject]@{
                VendorCode = $code
                VendorName = $name
                BUs        = New-Object System.Collections.Generic.List[string]
            }
        }

        $bu = ([string]$row.BU).Trim().ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($bu) -and -not $groups[$key].BUs.Contains($bu)) {
            $groups[$key].BUs.Add($bu)
        }
    }

    $candidates = foreach ($group in $groups.Values) {
        $supplier = $null
        if (-not [string]::IsNullOrWhiteSpace($group.VendorCode) -and $supplierByCode.ContainsKey($group.VendorCode)) {
            $supplier = $supplierByCode[$group.VendorCode]
        } else {
            $groupNameKey = Normalize-SupplierKey $group.VendorName
            if ($supplierByName.ContainsKey($groupNameKey)) { $supplier = $supplierByName[$groupNameKey] }
        }

        $to = if ($null -ne $supplier) { ConvertTo-OutlookRecipientList $supplier.'Email To' } else { "" }
        $cc = if ($null -ne $supplier) { ConvertTo-OutlookRecipientList $supplier.'Email CC' } else { "" }

        [PSCustomObject]@{
            VendorCode    = $group.VendorCode
            VendorName    = $group.VendorName
            MissingBUs    = (($group.BUs | Sort-Object) -join ', ')
            To            = $to
            CC            = $cc
            CanCreateDraft = -not [string]::IsNullOrWhiteSpace($to)
        }
    }

    return @($candidates | Sort-Object VendorName, VendorCode)
}

function Get-SupplierReminderSubject {
    param(
        [Parameter(Mandatory)]$Candidate,
        [Parameter(Mandatory)][datetime]$WeekStart,
        [Parameter(Mandatory)][datetime]$WeekEnd
    )

    return 'Reminder: Supplier Commitment - {0} - {1}-{2}' -f `
        $Candidate.MissingBUs,
        $WeekStart.ToString('dd/MM/yyyy'),
        $WeekEnd.ToString('dd/MM/yyyy')
}

function New-SupplierReminderHtml {
    param(
        [Parameter(Mandatory)]$Candidate,
        [Parameter(Mandatory)][datetime]$WeekStart,
        [Parameter(Mandatory)][datetime]$WeekEnd
    )

    $supplierName = [System.Net.WebUtility]::HtmlEncode([string]$Candidate.VendorName)
    $buItems = foreach ($bu in ([string]$Candidate.MissingBUs -split '\s*,\s*')) {
        if (-not [string]::IsNullOrWhiteSpace($bu)) {
            '<li>{0}</li>' -f [System.Net.WebUtility]::HtmlEncode($bu)
        }
    }
    $weekLabel = '{0} - {1}' -f $WeekStart.ToString('dd/MM/yyyy'), $WeekEnd.ToString('dd/MM/yyyy')

    return @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000">
<p>Dear Supplier,</p>
<p>We have not received the Supplier Commitment file from <b>$supplierName</b> for the following BU(s):</p>
<ul>$($buItems -join '')</ul>
<p>Week: <b>$weekLabel</b></p>
<p>Please send the missing Supplier Commitment file(s) as soon as possible.</p>
<p>Thank you.</p>
</div>
"@
}

function Show-SupplierReminderSelection {
    param([object[]]$Candidates)

    $sendable = @($Candidates | Where-Object { $_.CanCreateDraft })
    if ($sendable.Count -eq 0) { return @() }

    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Supplier Commitment Reminder'
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(980, 500)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = 'FixedDialog'

    $instruction = New-Object System.Windows.Forms.Label
    $instruction.Text = 'Chon supplier can tao email nhac. Supplier khong duoc chon se khong tao draft.'
    $instruction.Location = New-Object System.Drawing.Point(16, 16)
    $instruction.Size = New-Object System.Drawing.Size(940, 24)
    $form.Controls.Add($instruction)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true
    $list.HorizontalScrollbar = $true
    $list.Location = New-Object System.Drawing.Point(16, 48)
    $list.Size = New-Object System.Drawing.Size(948, 340)
    $list.Font = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($candidate in $sendable) {
        $label = '{0} | {1} | Thieu: {2} | To: {3}' -f `
            $candidate.VendorCode, $candidate.VendorName, $candidate.MissingBUs, $candidate.To
        [void]$list.Items.Add($label, $false)
    }
    $form.Controls.Add($list)

    $missingEmailCount = @($Candidates | Where-Object { -not $_.CanCreateDraft }).Count
    $status = New-Object System.Windows.Forms.Label
    $status.Text = if ($missingEmailCount -gt 0) {
        "$missingEmailCount supplier thieu Email To trong suppliers.csv va se duoc bo qua."
    } else {
        'Tat ca supplier trong danh sach deu co Email To.'
    }
    $status.ForeColor = if ($missingEmailCount -gt 0) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DarkGreen }
    $status.Location = New-Object System.Drawing.Point(16, 398)
    $status.Size = New-Object System.Drawing.Size(520, 24)
    $form.Controls.Add($status)

    $selectAll = New-Object System.Windows.Forms.Button
    $selectAll.Text = 'Chon tat ca'
    $selectAll.Location = New-Object System.Drawing.Point(16, 440)
    $selectAll.Size = New-Object System.Drawing.Size(110, 32)
    $selectAll.Add_Click({
        for ($index = 0; $index -lt $list.Items.Count; $index++) {
            $list.SetItemChecked($index, $true)
        }
    })
    $form.Controls.Add($selectAll)

    $clearAll = New-Object System.Windows.Forms.Button
    $clearAll.Text = 'Bo chon tat ca'
    $clearAll.Location = New-Object System.Drawing.Point(134, 440)
    $clearAll.Size = New-Object System.Drawing.Size(120, 32)
    $clearAll.Add_Click({
        for ($index = 0; $index -lt $list.Items.Count; $index++) {
            $list.SetItemChecked($index, $false)
        }
    })
    $form.Controls.Add($clearAll)

    $createDrafts = New-Object System.Windows.Forms.Button
    $createDrafts.Text = 'Tao Draft'
    $createDrafts.Location = New-Object System.Drawing.Point(738, 440)
    $createDrafts.Size = New-Object System.Drawing.Size(110, 32)
    $createDrafts.Add_Click({
        if ($list.CheckedItems.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'Hay chon it nhat mot supplier, hoac bam Huy.',
                'Supplier Commitment Reminder',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($createDrafts)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Huy'
    $cancel.Location = New-Object System.Drawing.Point(856, 440)
    $cancel.Size = New-Object System.Drawing.Size(108, 32)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $createDrafts
    $form.CancelButton = $cancel
    $dialogResult = $form.ShowDialog()

    $selected = New-Object System.Collections.Generic.List[object]
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($index in $list.CheckedIndices) {
            $selected.Add($sendable[[int]$index])
        }
    }

    $form.Dispose()
    return $selected.ToArray()
}

function New-SupplierReminderDrafts {
    param(
        [Parameter(Mandatory)]$Outlook,
        [Parameter(Mandatory)][object[]]$Candidates,
        [Parameter(Mandatory)][datetime]$WeekStart,
        [Parameter(Mandatory)][datetime]$WeekEnd,
        [int]$SignatureDelayMilliseconds = 500
    )

    $created = 0
    foreach ($candidate in $Candidates) {
        $mail = $null
        try {
            $mail = $Outlook.CreateItem(0)
            $mail.Display()
            if ($SignatureDelayMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $SignatureDelayMilliseconds
            }

            $signature = [string]$mail.HTMLBody
            $mail.To = $candidate.To
            $mail.CC = $candidate.CC
            $mail.Subject = Get-SupplierReminderSubject -Candidate $candidate -WeekStart $WeekStart -WeekEnd $WeekEnd
            $mail.HTMLBody = (New-SupplierReminderHtml -Candidate $candidate -WeekStart $WeekStart -WeekEnd $WeekEnd) + $signature
            $mail.Save()
            $mail.Display()

            $created++
            Write-Log "Da tao draft: $($candidate.VendorCode) - $($candidate.VendorName) [$($candidate.MissingBUs)]" 'Green'
        } catch {
            Write-Log "Khong tao duoc draft cho $($candidate.VendorCode) - $($candidate.VendorName): $_" 'Red'
        }
    }

    return $created
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

    return [PSCustomObject]@{
        ReportPath  = $reportPath
        MissingRows = $missing.ToArray()
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

    # Uu tien duong dan Outlook da cau hinh. Neu folder bi doi ten/di chuyen,
    # quet toan bo cay Supplier Commitment va loc theo tieu de mail.
    $subjectOnly = $false
    try {
        $mailFolders = @(Get-SubFolder $inbox $cfg.Path)
    } catch {
        Write-Log "  Khong tim thay duong dan folder cua $bu. Chuyen sang quet theo tieu de mail." "Yellow"
        try {
            $supplierCommitmentRoot = Get-SubFolder $inbox @("Supplier Commitment")
            $mailFolders = @(Get-OutlookFolderTree $supplierCommitmentRoot)
            $subjectOnly = $true
        } catch {
            Write-Log "  Bo qua $bu : $_" "Red"
            continue
        }
    }

    $mailCount = 0
    foreach ($folder in $mailFolders) {
        # Loc email theo ngay bang Restrict cho nhanh
        $items = $folder.Items
        $items.Sort("[ReceivedTime]", $true)
        $filter = "[ReceivedTime] >= '" + $monday.ToString("MM/dd/yyyy HH:mm") + "'"
        $mails = $items.Restrict($filter)
        $mailCount += $mails.Count

        foreach ($mail in $mails) {
            # Chi xu ly email thuc su (MailItem)
            if ($mail.Class -ne 43) { continue }   # 43 = olMail

            $subjectInfo = Parse-SupplierCommitmentSubject $mail.Subject
            if ($subjectOnly -and ($null -eq $subjectInfo -or $subjectInfo.BU -ne $bu)) { continue }

            if ($null -ne $subjectInfo) {
                # Tieu de co BU ro rang: khong de mail cua BU khac roi vao sai folder dich.
                if ($subjectInfo.BU -ne $bu) { continue }

                $vendor = $supplierLookup[$subjectInfo.VendorCode]
                if ($null -eq $vendor) {
                    Write-Log "  ! Khong tim thay ma supplier $($subjectInfo.VendorCode) trong suppliers.csv (Subject: $($mail.Subject))" "DarkYellow"
                    continue
                }
            } else {
                # Giu tuong thich voi email cu trong folder chuan.
                $vendor = Parse-Vendor $mail.Categories
                if ($null -eq $vendor) {
                    Write-Log "  ! Bo qua email khong dung mau tieu de va khong co category Vendor (Subject: $($mail.Subject))" "DarkYellow"
                    continue
                }
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

    Write-Log "  Da quet $mailCount email trong khoang thoi gian." "Gray"
}

if ($downloadedRows.Count -gt 0) {
    $logPath = Join-Path -Path $DestRoot -ChildPath ("DownloadLog_SupplierCommitment_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $downloadedRows | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
    Write-Log "Da ghi log: $logPath" "Gray"
}

$missingReport = Export-MissingSupplierCommitmentReport `
    -masterPath $SupplierMasterPath `
    -outputDirectory $DestRoot `
    -bus @($Folders | ForEach-Object { $_.BU }) `
    -downloaded $downloadedRows

if ($null -ne $missingReport -and @($missingReport.MissingRows).Count -gt 0) {
    $reminderCandidates = @(Get-SupplierReminderCandidates `
        -MissingRows $missingReport.MissingRows `
        -Suppliers $masterSuppliers)

    $missingEmailCandidates = @($reminderCandidates | Where-Object { -not $_.CanCreateDraft })
    foreach ($candidate in $missingEmailCandidates) {
        Write-Log "Khong co Email To, bo qua reminder: $($candidate.VendorCode) - $($candidate.VendorName) [$($candidate.MissingBUs)]" 'Yellow'
    }

    $sendableCandidates = @($reminderCandidates | Where-Object { $_.CanCreateDraft })
    if ($sendableCandidates.Count -gt 0) {
        Write-Log "Mo danh sach chon supplier can tao reminder draft..." 'Cyan'
        $selectedCandidates = @(Show-SupplierReminderSelection -Candidates $reminderCandidates)

        if ($selectedCandidates.Count -gt 0) {
            $draftCount = New-SupplierReminderDrafts `
                -Outlook $outlook `
                -Candidates $selectedCandidates `
                -WeekStart $weekStart `
                -WeekEnd $weekEnd
            Write-Log "Da tao $draftCount Outlook draft. Vui long doc lai truoc khi Send." 'Cyan'
        } else {
            Write-Log "Khong co supplier nao duoc chon. Khong tao reminder draft." 'Gray'
        }
    } else {
        Write-Log "Khong co supplier nao du Email To de tao reminder draft." 'Yellow'
    }
}

Write-Log "===== HOAN TAT: da luu $totalSaved file =====" "Cyan"
Write-Log "Folder dich: $DestRoot" "Cyan"

# Giai phong COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

Read-Host "Nhan Enter de thoat"
