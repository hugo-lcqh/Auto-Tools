# =====================================================================
#  Download Vendor Stock
#  - Duyet folder Outlook: Inbox\Vendor Stock
#  - Mac dinh tai file Excel cua tuan hien tai, tinh tu Chu Nhat -> Thu Bay
#    Vi du ngay 24/06/2026 => 21/06/2026 den 27/06/2026
#  - Tao folder theo tuan: Output\Vendor Stock\21.06-27.06.2026
#  - Dat ten file theo supplier lay tu subject hoac ten file dinh kem
#  - Xuat report supplier chua co file data theo Config\suppliers.csv
# =====================================================================

[CmdletBinding()]
param(
    # De trong = mailbox mac dinh. Co the dien mot phan ten store/email neu can.
    [string]$StoreName = "",

    # Duong dan folder tinh tu Inbox.
    [string[]]$FolderPath = @("Vendor Stock"),

    # De trong = Output\Vendor Stock trong Auto Tools.
    [string]$DestRoot = "",

    # File master supplier dung de check supplier chua phan hoi.
    [string]$SupplierMasterPath = "",

    # Ngay dung de tinh tuan hien tai.
    [datetime]$RunDate = (Get-Date),

    # Mac dinh dung tuan Chu Nhat -> Thu Bay theo vi du 21-27/06/2026.
    [ValidateSet("Sunday", "Monday")]
    [string]$WeekStartsOn = "Sunday",

    # Tuy chon nhap truc tiep ngay bat dau/ket thuc neu muon tai lai tuan cu.
    [datetime]$WeekStart,
    [datetime]$WeekEnd,

    # Neu bat, se quet ca folder con cua Vendor Stock.
    [switch]$IncludeSubfolders,

    # Bo qua mail template ban dau do noi bo gui cho supplier.
    [string[]]$ExcludedSenderAddresses = @("AnhThu.Nguyen1@ttigroup.com.vn"),

    # Dung them sender name de phong Outlook/Exchange khong tra ve SMTP address.
    [string[]]$ExcludedSenderNames = @("Judy Nguyen Anh Thu (VN.OP-SUC)"),

    # Neu bat, file trung ten se ghi de. Mac dinh se them (2), (3) de khong mat file.
    [switch]$Overwrite
)

. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")

$ErrorActionPreference = "Stop"

$AutoToolsPaths = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($DestRoot)) {
    $DestRoot = Join-Path -Path $AutoToolsPaths.Output -ChildPath "Vendor Stock"
}
if ([string]::IsNullOrWhiteSpace($SupplierMasterPath)) {
    $SupplierMasterPath = Join-Path -Path $AutoToolsPaths.Config -ChildPath "suppliers.csv"
}

$AttachmentExt = @(".xls", ".xlsx", ".xlsm", ".xlsb")
$script:Results = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Color = "White"
    )

    Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message) -ForegroundColor $Color
}

function Get-WeekRange {
    param(
        [datetime]$Date,
        [string]$StartsOn,
        [bool]$HasWeekStart,
        $StartOverride,
        [bool]$HasWeekEnd,
        $EndOverride
    )

    if ($HasWeekStart) {
        $start = $StartOverride.Date
    }
    else {
        $startDay = [System.Enum]::Parse([System.DayOfWeek], $StartsOn)
        $daysBack = (([int]$Date.DayOfWeek - [int]$startDay + 7) % 7)
        $start = $Date.Date.AddDays(-1 * $daysBack)
    }

    if ($HasWeekEnd) {
        $end = $EndOverride.Date
    }
    else {
        $end = $start.AddDays(6)
    }

    if ($end -lt $start) {
        throw "WeekEnd phai lon hon hoac bang WeekStart."
    }

    return [PSCustomObject]@{
        Start        = $start
        End          = $end
        EndExclusive = $end.AddDays(1)
        FolderName   = ("{0}-{1}" -f $start.ToString("dd.MM"), $end.ToString("dd.MM.yyyy"))
    }
}

function Clean-FileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $clean = $Name.Trim()
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $clean = $clean.Replace($c, " ")
    }

    $clean = $clean -replace '\s+', ' '
    $clean = $clean.Trim(" .-_")

    return $clean
}

function Normalize-Subject {
    param([string]$Subject)

    if ([string]::IsNullOrWhiteSpace($Subject)) {
        return ""
    }

    $value = $Subject.Trim()
    $previous = $null
    while ($previous -ne $value) {
        $previous = $value
        $value = ($value -replace '^\s*((RE|FW|FWD)\s*:\s*)+', '').Trim()
    }

    return $value
}

function Normalize-SupplierKey {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $value = $Name.ToUpperInvariant()
    $value = $value -replace '&', ' AND '
    $value = $value -replace '[^A-Z0-9]+', ''

    return $value
}

function Test-SupplierMatched {
    param(
        [string]$ExpectedName,
        [string[]]$ExpectedKeywords,
        [System.Collections.Generic.HashSet[string]]$DownloadedKeys
    )

    $expectedKey = Normalize-SupplierKey -Name $ExpectedName
    if ([string]::IsNullOrWhiteSpace($expectedKey)) {
        return $false
    }

    if ($DownloadedKeys.Contains($expectedKey)) {
        return $true
    }

    foreach ($downloadedKey in $DownloadedKeys) {
        if ([string]::IsNullOrWhiteSpace($downloadedKey)) {
            continue
        }

        if ($downloadedKey.Contains($expectedKey) -or $expectedKey.Contains($downloadedKey)) {
            return $true
        }
    }

    foreach ($keyword in $ExpectedKeywords) {
        $keywordKey = Normalize-SupplierKey -Name $keyword
        if ([string]::IsNullOrWhiteSpace($keywordKey)) {
            continue
        }

        foreach ($downloadedKey in $DownloadedKeys) {
            if ($downloadedKey.Contains($keywordKey) -or $keywordKey.Contains($downloadedKey)) {
                return $true
            }
        }
    }

    return $false
}

function Export-MissingSupplierReport {
    param(
        [Parameter(Mandatory)]
        [string]$MasterPath,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [Parameter(Mandatory)]
        [datetime]$WeekStartDate,

        [Parameter(Mandatory)]
        [datetime]$WeekEndDate,

        [Parameter(Mandatory)]
        [object[]]$DownloadedRows
    )

    if (-not (Test-Path -LiteralPath $MasterPath -PathType Leaf)) {
        Write-Log ("Khong tim thay file master supplier: {0}" -f $MasterPath) "Yellow"
        return
    }

    $suppliers = Import-Csv -LiteralPath $MasterPath
    if ($null -eq $suppliers -or $suppliers.Count -eq 0) {
        Write-Log ("File master supplier khong co data: {0}" -f $MasterPath) "Yellow"
        return
    }

    $downloadedKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in $DownloadedRows) {
        $key = Normalize-SupplierKey -Name $row.Supplier
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            [void]$downloadedKeys.Add($key)
        }
    }

    $missing = New-Object System.Collections.Generic.List[object]
    foreach ($supplier in $suppliers) {
        $vendorName = [string]$supplier.VendorName
        if ([string]::IsNullOrWhiteSpace($vendorName)) {
            continue
        }

        $keywords = @()
        if (-not [string]::IsNullOrWhiteSpace($supplier.Keyword)) {
            $keywords = $supplier.Keyword -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }

        if (-not (Test-SupplierMatched -ExpectedName $vendorName -ExpectedKeywords $keywords -DownloadedKeys $downloadedKeys)) {
            $missing.Add([PSCustomObject]@{
                VendorCode = $supplier.VendorCode
                VendorName = $vendorName
                Keyword    = $supplier.Keyword
                WeekStart  = $WeekStartDate.ToString("dd/MM/yyyy")
                WeekEnd    = $WeekEndDate.ToString("dd/MM/yyyy")
                Status     = "Missing Vendor Stock reply/data file"
            })
        }
    }

    $reportPath = Join-Path -Path $OutputDirectory -ChildPath ("MissingVendorStock_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $missing | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8

    if ($missing.Count -eq 0) {
        Write-Log "Tat ca supplier trong suppliers.csv deu da co file data." "Green"
    }
    else {
        Write-Log ("Con {0} supplier chua co file data. Report: {1}" -f $missing.Count, $reportPath) "Yellow"
    }
}

function Get-MailSenderAddress {
    param(
        [Parameter(Mandatory)]
        $Mail
    )

    try {
        $smtpAddress = $Mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
        if (-not [string]::IsNullOrWhiteSpace($smtpAddress)) {
            return $smtpAddress.Trim().ToLowerInvariant()
        }
    }
    catch {
    }

    try {
        if ($Mail.SenderEmailType -eq "EX" -and $null -ne $Mail.Sender) {
            $exchangeUser = $Mail.Sender.GetExchangeUser()
            if ($null -ne $exchangeUser -and -not [string]::IsNullOrWhiteSpace($exchangeUser.PrimarySmtpAddress)) {
                return $exchangeUser.PrimarySmtpAddress.Trim().ToLowerInvariant()
            }
        }
    }
    catch {
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($Mail.SenderEmailAddress)) {
            return $Mail.SenderEmailAddress.Trim().ToLowerInvariant()
        }
    }
    catch {
    }

    return ""
}

function Test-ExcludedSender {
    param(
        [string]$SenderAddress,
        [string]$SenderName,
        [string[]]$BlockedAddresses,
        [string[]]$BlockedNames
    )

    foreach ($address in $BlockedAddresses) {
        if ([string]::IsNullOrWhiteSpace($address)) {
            continue
        }

        if ($SenderAddress -eq $address.Trim().ToLowerInvariant()) {
            return $true
        }
    }

    foreach ($name in $BlockedNames) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($SenderName -eq $name.Trim() -or $SenderName -like ("*" + $name.Trim() + "*")) {
            return $true
        }
    }

    return $false
}

function Get-SupplierName {
    param(
        [Parameter(Mandatory)]
        $Mail,

        [string]$AttachmentName
    )

    $subject = Normalize-Subject -Subject $Mail.Subject

    if ($subject -match '(?i)vendor\s*stock\s*[-_:]\s*(.+)$') {
        $name = Clean-FileName $matches[1]
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            return $name
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($AttachmentName)) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($AttachmentName)
        if ($base -match '(?i)vendor\s*stock\s*[-_\s]+(.+)$') {
            $name = Clean-FileName $matches[1]
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                return $name
            }
        }
    }

    return $null
}

function Get-SubFolder {
    param(
        [Parameter(Mandatory)]
        $Parent,

        [Parameter(Mandatory)]
        [string[]]$PathArray
    )

    $current = $Parent
    foreach ($name in $PathArray) {
        $found = $null
        foreach ($folder in $current.Folders) {
            if ($folder.Name -eq $name) {
                $found = $folder
                break
            }
        }

        if ($null -eq $found) {
            foreach ($folder in $current.Folders) {
                if ($folder.Name.Trim().ToLowerInvariant() -eq $name.Trim().ToLowerInvariant()) {
                    $found = $folder
                    break
                }
            }
        }

        if ($null -eq $found) {
            throw "Khong tim thay folder con: '$name'"
        }

        $current = $found
    }

    return $current
}

function Get-MailFolders {
    param(
        [Parameter(Mandatory)]
        $Folder,

        [bool]$Recursive
    )

    $folders = New-Object System.Collections.Generic.List[object]
    $folders.Add($Folder)

    if ($Recursive) {
        foreach ($child in $Folder.Folders) {
            foreach ($nested in (Get-MailFolders -Folder $child -Recursive $true)) {
                $folders.Add($nested)
            }
        }
    }

    return $folders
}

function Get-UniqueSavePath {
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$BaseName,

        [Parameter(Mandatory)]
        [string]$Extension,

        [bool]$AllowOverwrite
    )

    $path = Join-Path -Path $Directory -ChildPath ($BaseName + $Extension)
    if ($AllowOverwrite -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return $path
    }

    for ($i = 2; $i -lt 10000; $i++) {
        $candidate = Join-Path -Path $Directory -ChildPath ("{0} ({1}){2}" -f $BaseName, $i, $Extension)
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }

    throw "Qua nhieu file trung ten trong folder dich: $BaseName"
}

$range = Get-WeekRange `
    -Date $RunDate `
    -StartsOn $WeekStartsOn `
    -HasWeekStart $PSBoundParameters.ContainsKey("WeekStart") `
    -StartOverride $WeekStart `
    -HasWeekEnd $PSBoundParameters.ContainsKey("WeekEnd") `
    -EndOverride $WeekEnd

$destDir = Join-Path -Path $DestRoot -ChildPath $range.FolderName
if (-not (Test-Path -LiteralPath $destDir -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $destDir -Force
}

Write-Log ("Tai file tu {0} den {1}" -f $range.Start.ToString("dd/MM/yyyy"), $range.End.ToString("dd/MM/yyyy")) "Cyan"
Write-Log ("Folder dich: {0}" -f $destDir) "Cyan"

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
}
catch {
    Write-Log "Khong the ket noi Outlook. Hay mo Outlook truoc khi chay script." "Red"
    Read-Host "Nhan Enter de thoat"
    exit 1
}

try {
    if ([string]::IsNullOrWhiteSpace($StoreName)) {
        $inbox = $namespace.GetDefaultFolder(6)
    }
    else {
        $store = $null
        foreach ($candidate in $namespace.Stores) {
            if ($candidate.DisplayName -eq $StoreName -or $candidate.DisplayName -like "*$StoreName*") {
                $store = $candidate
                break
            }
        }

        if ($null -eq $store) {
            throw "Khong tim thay mailbox/store: $StoreName"
        }

        $inbox = $store.GetRootFolder().Folders.Item("Inbox")
    }

    $vendorStockFolder = Get-SubFolder -Parent $inbox -PathArray $FolderPath
    $mailFolders = Get-MailFolders -Folder $vendorStockFolder -Recursive $IncludeSubfolders.IsPresent

    $totalMails = 0
    $totalSaved = 0
    $totalSkipped = 0
    $totalExcludedMails = 0

    $filter = "[ReceivedTime] >= '" + $range.Start.ToString("MM/dd/yyyy HH:mm") + "' AND [ReceivedTime] < '" + $range.EndExclusive.ToString("MM/dd/yyyy HH:mm") + "'"

    foreach ($folder in $mailFolders) {
        Write-Log ("----- Xu ly folder Outlook: {0} -----" -f $folder.FolderPath) "Yellow"

        $items = $folder.Items
        $items.Sort("[ReceivedTime]", $true)
        $mails = $items.Restrict($filter)
        $totalMails += $mails.Count

        Write-Log ("Tim thay {0} email trong khoang ngay." -f $mails.Count) "Gray"

        foreach ($mail in $mails) {
            if ($mail.Class -ne 43) {
                continue
            }

            $senderAddress = Get-MailSenderAddress -Mail $mail
            $senderName = [string]$mail.SenderName

            if (Test-ExcludedSender -SenderAddress $senderAddress -SenderName $senderName -BlockedAddresses $ExcludedSenderAddresses -BlockedNames $ExcludedSenderNames) {
                $totalExcludedMails++
                Write-Log ("Bo qua mail template tu {0} <{1}>: {2}" -f $senderName, $senderAddress, $mail.Subject) "DarkYellow"
                continue
            }

            if ($mail.Attachments.Count -eq 0) {
                continue
            }

            foreach ($attachment in $mail.Attachments) {
                $originalName = $attachment.FileName
                $extension = [System.IO.Path]::GetExtension($originalName).ToLowerInvariant()

                if ($AttachmentExt -notcontains $extension) {
                    continue
                }

                if ($originalName -like '~$*') {
                    continue
                }

                $supplier = Get-SupplierName -Mail $mail -AttachmentName $originalName
                if ([string]::IsNullOrWhiteSpace($supplier)) {
                    $supplier = Clean-FileName ([System.IO.Path]::GetFileNameWithoutExtension($originalName))
                    Write-Log ("Khong lay duoc supplier tu subject, dung ten file goc: {0}" -f $supplier) "DarkYellow"
                }

                if ([string]::IsNullOrWhiteSpace($supplier)) {
                    $supplier = "UNKNOWN_SUPPLIER"
                }

                $baseName = Clean-FileName $supplier
                $savePath = Get-UniqueSavePath -Directory $destDir -BaseName $baseName -Extension $extension -AllowOverwrite $Overwrite.IsPresent

                try {
                    $attachment.SaveAsFile($savePath)
                    $totalSaved++
                    Write-Log ("OK  {0}" -f [System.IO.Path]::GetFileName($savePath)) "Green"

                    $script:Results.Add([PSCustomObject]@{
                        ReceivedTime = $mail.ReceivedTime
                        Supplier     = $supplier
                        Subject      = $mail.Subject
                        Attachment   = $originalName
                        SavedAs      = $savePath
                        Sender       = $mail.SenderName
                        SenderEmail  = $senderAddress
                    })
                }
                catch {
                    $totalSkipped++
                    Write-Log ("Loi luu file {0}: {1}" -f $originalName, $_.Exception.Message) "Red"
                }
            }
        }
    }

    if ($script:Results.Count -gt 0) {
        $logPath = Join-Path -Path $destDir -ChildPath ("DownloadLog_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        $script:Results | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
        Write-Log ("Da ghi log: {0}" -f $logPath) "Gray"
    }

    Export-MissingSupplierReport `
        -MasterPath $SupplierMasterPath `
        -OutputDirectory $destDir `
        -WeekStartDate $range.Start `
        -WeekEndDate $range.End `
        -DownloadedRows $script:Results.ToArray()

    Write-Log ("===== HOAN TAT: quet {0} email, bo qua {1} mail template, luu {2} file Excel =====" -f $totalMails, $totalExcludedMails, $totalSaved) "Cyan"
    if ($totalSkipped -gt 0) {
        Write-Log ("Co {0} file bi loi khi luu." -f $totalSkipped) "Yellow"
    }
}
catch {
    Write-Log $_.Exception.Message "Red"
}
finally {
    if ($null -ne $namespace) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace)
    }
    if ($null -ne $outlook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
    }
}

Read-Host "Nhan Enter de thoat"
