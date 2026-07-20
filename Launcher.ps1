#requires -Version 5.1

<#
.SYNOPSIS
    PowerShell Program Launcher.

.DESCRIPTION
    Doc danh sach chuong trinh tu launcher-config.json.

    Ho tro:
    - Duong dan tuong doi theo thu muc AutoTools.
    - Duong dan tuyet doi cho script ben ngoai AutoTools.
    - Tu dong tao cac thu muc he thong.
    - Ghi log vao Logs\Launcher.log.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$AppVersion = '1.1.0'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# ============================================================
# THU MUC HE THONG
# ============================================================

# Launcher.ps1 nam truc tiep trong AutoTools
$RootPath = [System.IO.Path]::GetFullPath($PSScriptRoot)

$Paths = [ordered]@{
    Root    = $RootPath
    Scripts = Join-Path -Path $RootPath -ChildPath 'Scripts'
    Config  = Join-Path -Path $RootPath -ChildPath 'Config'
    Input   = Join-Path -Path $RootPath -ChildPath 'Input'
    Output  = Join-Path -Path $RootPath -ChildPath 'Output'
    Logs    = Join-Path -Path $RootPath -ChildPath 'Logs'
    Temp    = Join-Path -Path $RootPath -ChildPath 'Temp'
}

$ConfigPath = Join-Path -Path $RootPath -ChildPath 'launcher-config.json'
$SupplierMasterPath = Join-Path -Path $Paths.Config -ChildPath 'suppliers.csv'
$LogPath    = Join-Path -Path $Paths.Logs -ChildPath 'Launcher.log'

# Tao cac thu muc neu chua ton tai
foreach ($directoryPath in $Paths.Values) {
    if (-not (Test-Path -LiteralPath $directoryPath -PathType Container)) {
        $null = New-Item `
            -Path $directoryPath `
            -ItemType Directory `
            -Force
    }
}

# ============================================================
# BANG MAU GIAO DIEN
# ============================================================

$ColorBg       = [System.Drawing.Color]::FromArgb(22, 27, 31)
$ColorPanel    = [System.Drawing.Color]::FromArgb(31, 38, 44)
$ColorBtn      = [System.Drawing.Color]::FromArgb(35, 111, 119)
$ColorBtnHover = [System.Drawing.Color]::FromArgb(43, 135, 145)
$ColorBtnDel   = [System.Drawing.Color]::FromArgb(151, 70, 76)
$ColorAccent   = [System.Drawing.Color]::FromArgb(239, 181, 76)
$ColorText     = [System.Drawing.Color]::FromArgb(242, 245, 247)
$ColorSubText  = [System.Drawing.Color]::FromArgb(163, 174, 184)
$ColorCard     = [System.Drawing.Color]::FromArgb(42, 50, 58)
$ColorGreen    = [System.Drawing.Color]::FromArgb(69, 154, 105)
$ColorWarning  = [System.Drawing.Color]::FromArgb(226, 136, 66)

# ============================================================
# CAC HAM HO TRO
# ============================================================

function Get-LauncherUserName {
    [CmdletBinding()]
    param()

    $rawName = if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
        [string]$env:USERNAME
    }
    else {
        [System.Security.Principal.WindowsIdentity]::GetCurrent().Name -replace '^.*\\', ''
    }

    if ([string]::IsNullOrWhiteSpace($rawName)) {
        return 'Material Control'
    }

    return ($rawName -replace '[._-]+', ' ').Trim()
}

function Get-LauncherDateTimeText {
    [CmdletBinding()]
    param()

    return (Get-Date -Format 'dddd, MMM dd yyyy  HH:mm:ss')
}

function Get-WorkdayStatus {
    [CmdletBinding()]
    param()

    $now = Get-Date
    $start = $now.Date.AddHours(7).AddMinutes(30)
    $end = $now.Date.AddHours(16)
    $totalSeconds = ($end - $start).TotalSeconds

    if ($now -lt $start) {
        $percent = 100
        $text = 'Workday 07:30-16:00 | Chua bat dau'
    }
    elseif ($now -ge $end) {
        $percent = 0
        $text = 'Workday 07:30-16:00 | Da ket thuc'
    }
    else {
        $remaining = $end - $now
        $percent = [Math]::Round(($remaining.TotalSeconds / $totalSeconds) * 100)
        $text = 'Workday 07:30-16:00 | Con lai {0}h {1}m ({2}%)' -f `
            [Math]::Floor($remaining.TotalHours),
            $remaining.Minutes,
            $percent
    }

    return [PSCustomObject]@{
        Percent = [Math]::Max(0, [Math]::Min(100, [int]$percent))
        Text    = $text
    }
}

function Write-LauncherLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logLine = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message

        Add-Content `
            -LiteralPath $LogPath `
            -Value $logLine `
            -Encoding UTF8
    }
    catch {
        # Khong lam dung launcher neu ghi log bi loi.
    }
}

function Write-FileNoBom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Content
    )

    $parentDirectory = Split-Path -Path $Path -Parent

    if (
        -not [string]::IsNullOrWhiteSpace($parentDirectory) -and
        -not (Test-Path -LiteralPath $parentDirectory -PathType Container)
    ) {
        $null = New-Item `
            -Path $parentDirectory `
            -ItemType Directory `
            -Force
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    [System.IO.File]::WriteAllText(
        $Path,
        $Content,
        $utf8NoBom
    )
}

function Resolve-ProgramPath {
    <#
        Chuyen duong dan trong JSON thanh duong dan tuyet doi.

        Vi du:
        Scripts\AutoImport.ps1
        =>
        C:\...\AutoTools\Scripts\AutoImport.ps1
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $expandedPath = [Environment]::ExpandEnvironmentVariables(
        $Path.Trim().Trim('"')
    )

    if ([System.IO.Path]::IsPathRooted($expandedPath)) {
        return [System.IO.Path]::GetFullPath($expandedPath)
    }

    return [System.IO.Path]::GetFullPath(
        (Join-Path -Path $RootPath -ChildPath $expandedPath)
    )
}

function ConvertTo-PortablePath {
    <#
        Neu script nam trong AutoTools, luu duong dan tuong doi.

        Vi du:
        C:\Tools\AutoTools\Scripts\AutoImport.ps1
        =>
        Scripts\AutoImport.ps1

        Neu script nam ben ngoai AutoTools, giu duong dan tuyet doi.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    $normalizedRoot = $RootPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    ) + [System.IO.Path]::DirectorySeparatorChar

    if (
        $fullPath.StartsWith(
            $normalizedRoot,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        return $fullPath.Substring($normalizedRoot.Length)
    }

    return $fullPath
}

function Get-PowerShellExecutable {
    [CmdletBinding()]
    param()

    # Neu Launcher dang chay bang PowerShell 7
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $pwshPath = Join-Path -Path $PSHOME -ChildPath 'pwsh.exe'

        if (Test-Path -LiteralPath $pwshPath -PathType Leaf) {
            return $pwshPath
        }
    }

    # Windows PowerShell 5.1
    $windowsPowerShellPath = Join-Path `
        -Path $PSHOME `
        -ChildPath 'powershell.exe'

    if (
        Test-Path `
            -LiteralPath $windowsPowerShellPath `
            -PathType Leaf
    ) {
        return $windowsPowerShellPath
    }

    $command = Get-Command `
        -Name 'powershell.exe' `
        -ErrorAction SilentlyContinue

    if ($null -ne $command) {
        return $command.Source
    }

    throw 'Khong tim thay powershell.exe hoac pwsh.exe.'
}

function Show-LauncherMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [string]$Title = 'Program Launcher',

        [ValidateSet(
            'None',
            'Information',
            'Warning',
            'Error',
            'Question'
        )]
        [string]$Icon = 'Information'
    )

    [void][System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::$Icon
    )
}

function New-DefaultConfiguration {
    [CmdletBinding()]
    param()

    $defaultPrograms = @(
        [PSCustomObject]@{
            Name = 'Auto Import'
            Path = 'Scripts\AutoImport.ps1'
        }
        [PSCustomObject]@{
            Name = 'Check CD Shortage'
            Path = 'Scripts\CheckCD_Shortage.ps1'
        }
        [PSCustomObject]@{
            Name = 'Check Shortage Supplier Commitment'
            Path = 'Scripts\Check-Shortage-Supplier Commitment.ps1'
        }
        [PSCustomObject]@{
            Name = 'Download Supplier Commitment'
            Path = 'Scripts\Download-SupplierCommitment.ps1'
        }
        [PSCustomObject]@{
            Name = 'Process CDs'
            Path = 'Scripts\Process-CDs\Process-CDs.ps1'
        }
        [PSCustomObject]@{
            Name = 'Process PO'
            Path = 'Scripts\Process-PO\Process-PO.ps1'
        }
    )

    $json = ConvertTo-Json `
        -InputObject @($defaultPrograms) `
        -Depth 5

    Write-FileNoBom `
        -Path $ConfigPath `
        -Content $json

    Write-LauncherLog -Message "Da tao file cau hinh mac dinh: $ConfigPath"
}

function New-DefaultSupplierMaster {
    [CmdletBinding()]
    param()

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

    Write-FileNoBom `
        -Path $SupplierMasterPath `
        -Content $csv

    Write-LauncherLog -Message "Da tao file supplier mac dinh: $SupplierMasterPath"
}

function Get-DefaultSupplierInputText {
    [CmdletBinding()]
    param()

    return @'
WEIDA (VIETNAM) MANUFACTURING CO.,LTD	10861
MINH ANH ELECTRONICS SERVICE-TRADING COMPANY LIMITED	10981
TALWAY VIET NAM COMPANY LIMITED	12225
GREEN (VIET NAM) CO., LTD	9340
YAO-I VIETNAM COMPANY LIMITED	9781
XIAN HUA VIET NAM INDUSTRIAL COMPANY LIMITED	10603
VIETNAM SHANGHU ELECTRONICS COMPANY LIMITED	13316
NGHIA LONG METAL PRODUCTS COMPANY LIMITED	12258
CHANTING INTELLIGENT TECHNOLOGY CO., LTD	12448
JUN JAM METAL PRODUCTS TECHNOLOGY LIMITED	12728
NEW SUN VIET NAM TECHNOLOGY COMPANY LIMITED	13313
YU XIN VIET NAM COMPANY LIMITED	12259
KHGEARS VIETNAM COMPANY LIMITED	11500
CONG TY TNHH MAGNET JC (VIET NAM)	13553
MINH THU ACCURATELY MECHANIC COMPANY LIMITED	10534
WONDERWARD TECHNOLOGY (HONG KONG) LIMITED	12504
'@
}

function Get-SupplierKeyword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$VendorName
    )

    $name = $VendorName.ToUpperInvariant()

    $knownKeywords = [ordered]@{
        'YAO-I'       = 'YAO-I'
        'XIAN HUA'    = 'XIAN HUA'
        'YU XIN'      = 'YU XIN'
        'NEW SUN'     = 'NEW SUN'
        'KHGEARS'     = 'KHGEAR;KHGEARS'
        'KHGEAR'      = 'KHGEAR;KHGEARS'
        'MINH THU'    = 'MINH THU'
        'WONDERWARD'  = 'WONDERWARD'
        'SHANGHU'     = 'SHANGHU;VN SHANGHU;VIETNAM SHANGHU'
        'WEIDA'       = 'WEIDA'
        'TALWAY'      = 'TALWAY'
        'GREEN'       = 'GREEN'
        'HAIXING'     = 'HAIXING'
        'MAGNET'      = 'MAGNET'
        'NGHIA LONG'  = 'NGHIA LONG'
        'JUN JAM'     = 'JUN JAM'
        'CHANTING'    = 'CHANTING'
        'MINH ANH'    = 'MINH ANH'
        'NOVATEK'     = 'NOVATEK'
        'LONG BAO'    = 'LONG BAO'
    }

    foreach ($key in $knownKeywords.Keys) {
        if ($name.Contains($key)) {
            return $knownKeywords[$key]
        }
    }

    $clean = $name
    $removePatterns = @(
        '\([^)]*\)',
        '\bCONG TY\b',
        '\bCTY\b',
        '\bTNHH\b',
        '\bVIET NAM\b',
        '\bVIETNAM\b',
        '\bCOMPANY\b',
        '\bLIMITED\b',
        '\bCO\b',
        '\bLTD\b',
        '\bTECHNOLOGY\b',
        '\bELECTRONICS\b',
        '\bINDUSTRIAL\b',
        '\bMANUFACTURING\b',
        '\bSERVICE\b',
        '\bTRADING\b',
        '\bACCURATELY\b',
        '\bMECHANIC\b',
        '\bPRODUCTS\b'
    )

    foreach ($pattern in $removePatterns) {
        $clean = [regex]::Replace($clean, $pattern, ' ')
    }

    $clean = ($clean -replace '[^A-Z0-9\- ]', ' ' -replace '\s+', ' ').Trim()
    if (-not [string]::IsNullOrWhiteSpace($clean)) {
        return $clean
    }

    return $VendorName.Trim()
}

function ConvertTo-CsvField {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -match '[,"\r\n]') {
        return '"' + $Value.Replace('"', '""') + '"'
    }

    return $Value
}

function Convert-SupplierInputToCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputText
    )

    $rows = @()
    $seenCodes = @{}

    foreach ($line in ($InputText -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $trimmedLine = $line.Trim()
        if ($trimmedLine -match '(?i)^(vendor\s*name|supplier|processcds|roundmoq)') {
            continue
        }

        if ($trimmedLine.Contains("`t")) {
            $columns = @($trimmedLine -split "`t")

            if ($columns.Count -lt 2) {
                throw "Dong supplier khong hop le: $line"
            }

            $vendorCode = $columns[$columns.Count - 1].Trim()
            $vendorName = (($columns[0..($columns.Count - 2)] -join ' ').Trim())
        }
        elseif ($trimmedLine -match '^(?<name>.+?)\s+(?<code>\d+)\s*$') {
            $vendorName = $matches.name.Trim()
            $vendorCode = $matches.code.Trim()
        }
        else {
            throw "Dong supplier khong hop le: $line"
        }

        if ([string]::IsNullOrWhiteSpace($vendorName) -or $vendorCode -notmatch '^\d+$') {
            throw "Dong supplier khong hop le: $line"
        }

        if ($seenCodes.ContainsKey($vendorCode)) {
            continue
        }

        $seenCodes[$vendorCode] = $true

        $rows += [PSCustomObject]@{
            Keyword    = Get-SupplierKeyword -VendorName $vendorName
            VendorCode = $vendorCode
            VendorName = $vendorName
            ProcessCDs = 'TRUE'
            RoundMOQ   = 'FALSE'
            MOQFile    = ''
            Combine    = 'FALSE'
            ItemPrefix = ''
        }
    }

    if ($rows.Count -eq 0) {
        throw 'Chua co supplier hop le de tao Config\suppliers.csv.'
    }

    $lines = @('Keyword,VendorCode,VendorName,ProcessCDs,RoundMOQ,MOQFile,Combine,ItemPrefix')
    foreach ($row in $rows) {
        $fields = @(
            ConvertTo-CsvField $row.Keyword
            ConvertTo-CsvField $row.VendorCode
            ConvertTo-CsvField $row.VendorName
            ConvertTo-CsvField $row.ProcessCDs
            ConvertTo-CsvField $row.RoundMOQ
            ConvertTo-CsvField $row.MOQFile
            ConvertTo-CsvField $row.Combine
            ConvertTo-CsvField $row.ItemPrefix
        )

        $lines += ($fields -join ',')
    }

    return ($lines -join "`r`n")
}

function Show-SupplierSetupDialog {
    [CmdletBinding()]
    param()

    $setupForm = New-Object System.Windows.Forms.Form
    $setupForm.Text = 'Setup Supplier Master'
    $setupForm.Size = New-Object System.Drawing.Size(760, 560)
    $setupForm.StartPosition = 'CenterScreen'
    $setupForm.MinimizeBox = $false
    $setupForm.MaximizeBox = $false
    $setupForm.FormBorderStyle = 'FixedDialog'
    $setupForm.BackColor = $ColorBg
    $setupForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = (
        "Chua co Config\suppliers.csv.`r`n" +
        "Paste danh sach supplier theo mau: Vendor Name<Tab>Vendor Code.`r`n" +
        "Cac cot ProcessCDs/RoundMOQ/MOQFile/Combine/ItemPrefix se duoc tao mac dinh."
    )
    $label.Location = New-Object System.Drawing.Point(16, 14)
    $label.Size = New-Object System.Drawing.Size(710, 58)
    $label.ForeColor = $ColorText
    $setupForm.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(16, 78)
    $textBox.Size = New-Object System.Drawing.Size(710, 360)
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $true
    $textBox.WordWrap = $false
    $textBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $textBox.Text = Get-DefaultSupplierInputText
    $setupForm.Controls.Add($textBox)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(16, 448)
    $statusLabel.Size = New-Object System.Drawing.Size(710, 22)
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(240, 200, 110)
    $setupForm.Controls.Add($statusLabel)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Tao suppliers.csv'
    $btnSave.Location = New-Object System.Drawing.Point(466, 478)
    $btnSave.Size = New-Object System.Drawing.Size(126, 32)
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(50, 130, 90)
    $btnSave.ForeColor = $ColorText
    $btnSave.FlatStyle = 'Flat'
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::None
    $setupForm.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Huy'
    $btnCancel.Location = New-Object System.Drawing.Point(606, 478)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 32)
    $btnCancel.BackColor = $ColorPanel
    $btnCancel.ForeColor = $ColorText
    $btnCancel.FlatStyle = 'Flat'
    $setupForm.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        try {
            $csv = Convert-SupplierInputToCsv -InputText $textBox.Text
            Write-FileNoBom -Path $SupplierMasterPath -Content $csv
            Write-LauncherLog -Message "Da tao supplier master tu input nguoi dung: $SupplierMasterPath"
            $setupForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $setupForm.Close()
        }
        catch {
            $statusLabel.Text = $_.Exception.Message
            Write-LauncherLog -Level ERROR -Message "Loi tao supplier master: $($_.Exception.Message)"
        }
    })

    $btnCancel.Add_Click({
        $setupForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $setupForm.Close()
    })

    return $setupForm.ShowDialog()
}

function Load-Programs {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
            return @()
        }

        $json = Get-Content `
            -LiteralPath $ConfigPath `
            -Raw `
            -Encoding UTF8

        if ([string]::IsNullOrWhiteSpace($json)) {
            return @()
        }

        $data = $json | ConvertFrom-Json -ErrorAction Stop
        $validPrograms = @()

        foreach ($program in @($data)) {
            if ($null -eq $program) {
                continue
            }

            if (
                -not ($program.PSObject.Properties.Name -contains 'Name') -or
                -not ($program.PSObject.Properties.Name -contains 'Path')
            ) {
                Write-LauncherLog `
                    -Level WARNING `
                    -Message 'Bo qua mot muc cau hinh vi thieu Name hoac Path.'

                continue
            }

            $name = [string]$program.Name
            $path = [string]$program.Path

            if (
                [string]::IsNullOrWhiteSpace($name) -or
                [string]::IsNullOrWhiteSpace($path)
            ) {
                Write-LauncherLog `
                    -Level WARNING `
                    -Message 'Bo qua mot muc cau hinh co Name hoac Path rong.'

                continue
            }

            $programObject = [PSCustomObject]@{
                Name = $name.Trim()
                Path = $path.Trim()
            }
            if ($program.PSObject.Properties.Name -contains 'Arguments') {
                $programObject | Add-Member -NotePropertyName Arguments -NotePropertyValue ([string]$program.Arguments).Trim()
            }
            $validPrograms += $programObject
        }

        return @($validPrograms)
    }
    catch {
        $errorMessage = $_.Exception.Message
        $lineNumber = if ($_.InvocationInfo) {
            $_.InvocationInfo.ScriptLineNumber
        }
        else {
            'unknown'
        }

        Write-LauncherLog `
            -Level ERROR `
            -Message "Loi doc file cau hinh tai dong ${lineNumber}: $errorMessage"

        Show-LauncherMessage `
            -Title 'Loi cau hinh' `
            -Icon Error `
            -Message (
                "Khong the doc launcher-config.json.`r`n`r`n" +
                "Chi tiet: $errorMessage`r`n`r`n" +
                "File: $ConfigPath"
            )

        return @()
    }
}

function Save-Programs {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowEmptyCollection()]
        [object[]]$Programs = @()
    )

    try {
        $programArray = @($Programs)

        if ($programArray.Count -eq 0) {
            $json = '[]'
        }
        else {
            $json = ConvertTo-Json `
                -InputObject $programArray `
                -Depth 5

            # Windows PowerShell 5.1 co the bo dau [] khi chi co 1 object.
            if ($programArray.Count -eq 1) {
                $trimmedJson = $json.Trim()

                if (-not $trimmedJson.StartsWith('[')) {
                    $json = "[`r`n$json`r`n]"
                }
            }
        }

        Write-FileNoBom `
            -Path $ConfigPath `
            -Content $json

        Write-LauncherLog `
            -Message "Da luu $($programArray.Count) chuong trinh vao cau hinh."
    }
    catch {
        $errorMessage = $_.Exception.Message

        Write-LauncherLog `
            -Level ERROR `
            -Message "Loi luu file cau hinh: $errorMessage"

        Show-LauncherMessage `
            -Title 'Loi luu cau hinh' `
            -Icon Error `
            -Message (
                "Khong the luu launcher-config.json.`r`n`r`n" +
                "Chi tiet: $errorMessage"
            )
    }
}

function Start-LauncherProgram {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Program
    )

    try {
        $resolvedPath = Resolve-ProgramPath -Path ([string]$Program.Path)

        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            Write-LauncherLog `
                -Level WARNING `
                -Message "Khong tim thay script: $resolvedPath"

            Show-LauncherMessage `
                -Title 'Khong tim thay file' `
                -Icon Warning `
                -Message (
                    "Khong tim thay file PowerShell:`r`n`r`n" +
                    "$resolvedPath`r`n`r`n" +
                    "Duong dan trong cau hinh:`r`n" +
                    "$($Program.Path)"
                )

            return
        }

        if (
            -not [string]::Equals(
                [System.IO.Path]::GetExtension($resolvedPath),
                '.ps1',
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            Show-LauncherMessage `
                -Title 'File khong hop le' `
                -Icon Warning `
                -Message "File duoc chon khong phai file .ps1:`r`n$resolvedPath"

            return
        }

        $powerShellExecutable = Get-PowerShellExecutable
        $workingDirectory = Split-Path -Path $resolvedPath -Parent

        # Truyen RootPath cho script con thong qua bien moi truong.
        # Script con co the doc bang:
        # $env:AUTOTOOLS_ROOT
        $previousRootEnvironment = $env:AUTOTOOLS_ROOT
        $env:AUTOTOOLS_ROOT = $RootPath

        try {
            $extraArguments = ''
            if ($Program.PSObject.Properties['Arguments']) {
                $extraArguments = ' ' + [string]$Program.Arguments
            }

            $argumentList = (
                '-NoProfile ' +
                '-NoExit ' +
                '-ExecutionPolicy Bypass ' +
                '-File "{0}"{1}' -f $resolvedPath.Replace('"', '\"'), $extraArguments
            )

            Start-Process `
                -FilePath $powerShellExecutable `
                -ArgumentList $argumentList `
                -WorkingDirectory $workingDirectory `
                -ErrorAction Stop

            Write-LauncherLog `
                -Message "Da chay '$($Program.Name)': $resolvedPath"
        }
        finally {
            if ($null -eq $previousRootEnvironment) {
                Remove-Item `
                    -Path 'Env:\AUTOTOOLS_ROOT' `
                    -ErrorAction SilentlyContinue
            }
            else {
                $env:AUTOTOOLS_ROOT = $previousRootEnvironment
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message

        Write-LauncherLog `
            -Level ERROR `
            -Message "Khong the chay '$($Program.Name)': $errorMessage"

        Show-LauncherMessage `
            -Title 'Loi chay chuong trinh' `
            -Icon Error `
            -Message (
                "Khong the chay chuong trinh:`r`n" +
                "$($Program.Name)`r`n`r`n" +
                "Chi tiet: $errorMessage"
            )
    }
}

# ============================================================
# TAO CAU HINH MAC DINH
# ============================================================

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    try {
        New-DefaultConfiguration
    }
    catch {
        Show-LauncherMessage `
            -Title 'Loi khoi tao' `
            -Icon Error `
            -Message (
                "Khong the tao launcher-config.json.`r`n`r`n" +
                $_.Exception.Message
            )

        throw
    }
}

if (-not (Test-Path -LiteralPath $SupplierMasterPath -PathType Leaf)) {
    try {
        $setupResult = Show-SupplierSetupDialog
        if ($setupResult -ne [System.Windows.Forms.DialogResult]::OK) {
            throw 'Ban can tao Config\suppliers.csv truoc khi su dung Launcher.'
        }
    }
    catch {
        Show-LauncherMessage `
            -Title 'Loi khoi tao supplier' `
            -Icon Error `
            -Message (
                "Khong the tao Config\suppliers.csv.`r`n`r`n" +
                $_.Exception.Message
            )

        throw
    }
}

Write-LauncherLog -Message "Launcher da khoi dong. RootPath: $RootPath"

# ============================================================
# TAO CUA SO CHINH
# ============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Material Control Auto Tools'
$form.Size = New-Object System.Drawing.Size(760, 660)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.BackColor = $ColorBg
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.TopMost = $false

# Tieu de
$topBar = New-Object System.Windows.Forms.Panel
$topBar.Location = New-Object System.Drawing.Point(0, 0)
$topBar.Size = New-Object System.Drawing.Size(760, 6)
$topBar.BackColor = $ColorAccent
$form.Controls.Add($topBar)

$appLabel = New-Object System.Windows.Forms.Label
$appLabel.Text = 'MATERIAL CONTROL'
$appLabel.Location = New-Object System.Drawing.Point(28, 24)
$appLabel.AutoSize = $true
$appLabel.ForeColor = $ColorAccent
$appLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    9,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($appLabel)

$header = New-Object System.Windows.Forms.Label
$header.Text = 'Material Control Auto Tools'
$header.Location = New-Object System.Drawing.Point(26, 42)
$header.AutoSize = $true
$header.ForeColor = $ColorText
$header.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    22,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($header)

$subHeader = New-Object System.Windows.Forms.Label
$subHeader.Text = 'Xin chao, {0}. Chon tool can chay ben duoi.' -f (Get-LauncherUserName)
$subHeader.Location = New-Object System.Drawing.Point(29, 84)
$subHeader.AutoSize = $true
$subHeader.ForeColor = $ColorSubText
$subHeader.Font = New-Object System.Drawing.Font('Segoe UI', 10.5)
$form.Controls.Add($subHeader)

$clockLabel = New-Object System.Windows.Forms.Label
$clockLabel.Text = Get-LauncherDateTimeText
$clockLabel.Location = New-Object System.Drawing.Point(500, 36)
$clockLabel.Size = New-Object System.Drawing.Size(218, 26)
$clockLabel.TextAlign = 'MiddleRight'
$clockLabel.ForeColor = $ColorText
$clockLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    10,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($clockLabel)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = $RootPath
$folderLabel.Location = New-Object System.Drawing.Point(408, 64)
$folderLabel.Size = New-Object System.Drawing.Size(310, 38)
$folderLabel.TextAlign = 'TopRight'
$folderLabel.ForeColor = $ColorSubText
$folderLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$form.Controls.Add($folderLabel)

$workdayLabel = New-Object System.Windows.Forms.Label
$workdayLabel.Location = New-Object System.Drawing.Point(29, 110)
$workdayLabel.Size = New-Object System.Drawing.Size(689, 18)
$workdayLabel.TextAlign = 'MiddleLeft'
$workdayLabel.ForeColor = $ColorSubText
$workdayLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.7)
$form.Controls.Add($workdayLabel)

$workdayTrack = New-Object System.Windows.Forms.Panel
$workdayTrack.Location = New-Object System.Drawing.Point(28, 134)
$workdayTrack.Size = New-Object System.Drawing.Size(690, 8)
$workdayTrack.BackColor = $ColorCard
$form.Controls.Add($workdayTrack)

$workdayFill = New-Object System.Windows.Forms.Panel
$workdayFill.Location = New-Object System.Drawing.Point(0, 0)
$workdayFill.Size = New-Object System.Drawing.Size(690, 8)
$workdayFill.BackColor = $ColorGreen
$workdayTrack.Controls.Add($workdayFill)

# Duong ke ngang
$line = New-Object System.Windows.Forms.Label
$line.Location = New-Object System.Drawing.Point(28, 152)
$line.Size = New-Object System.Drawing.Size(690, 2)
$line.BackColor = $ColorAccent
$form.Controls.Add($line)

$programCountLabel = New-Object System.Windows.Forms.Label
$programCountLabel.Text = 'Tools'
$programCountLabel.Location = New-Object System.Drawing.Point(30, 170)
$programCountLabel.AutoSize = $true
$programCountLabel.ForeColor = $ColorText
$programCountLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    11,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($programCountLabel)

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = 'Click nut tool de mo cua so rieng. Nhan F5 de tai lai danh sach.'
$hintLabel.Location = New-Object System.Drawing.Point(250, 173)
$hintLabel.Size = New-Object System.Drawing.Size(468, 22)
$hintLabel.TextAlign = 'MiddleRight'
$hintLabel.ForeColor = $ColorSubText
$hintLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.8)
$form.Controls.Add($hintLabel)

# Panel cuon chua cac nut
$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(28, 202)
$panel.Size = New-Object System.Drawing.Size(690, 333)
$panel.AutoScroll = $true
$panel.BackColor = $ColorPanel
$form.Controls.Add($panel)

function Update-WorkdayStatus {
    [CmdletBinding()]
    param()

    $status = Get-WorkdayStatus
    $workdayLabel.Text = $status.Text
    $workdayFill.Width = [Math]::Round(
        ($workdayTrack.ClientSize.Width * $status.Percent) / 100
    )

    if ($status.Percent -le 20) {
        $workdayFill.BackColor = $ColorBtnDel
    }
    elseif ($status.Percent -le 40) {
        $workdayFill.BackColor = $ColorWarning
    }
    else {
        $workdayFill.BackColor = $ColorGreen
    }
}

function Refresh-ButtonList {
    [CmdletBinding()]
    param()

    $panel.SuspendLayout()

    try {
        $panel.Controls.Clear()

        $programs = @(Load-Programs)
        $programCountLabel.Text = 'Tools ({0})' -f $programs.Count
        $y = 10

        if ($programs.Count -eq 0) {
            $emptyLabel = New-Object System.Windows.Forms.Label
            $emptyLabel.Text = (
                "Chua co chuong trinh nao.`r`n" +
                "Nhan '+ Them chuong trinh' de bat dau."
            )
            $emptyLabel.Location = New-Object System.Drawing.Point(22, 28)
            $emptyLabel.AutoSize = $true
            $emptyLabel.ForeColor = $ColorSubText
            $emptyLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10)

            $panel.Controls.Add($emptyLabel)
            return
        }

        foreach ($program in $programs) {
            # Nut chay chuong trinh
            $runButton = New-Object System.Windows.Forms.Button
            $runButton.Text = '  Run  |  ' + $program.Name
            $runButton.Tag = $program
            $runButton.Size = New-Object System.Drawing.Size(585, 46)
            $runButton.Location = New-Object System.Drawing.Point(12, $y)
            $runButton.TextAlign = 'MiddleLeft'
            $runButton.Font = New-Object System.Drawing.Font(
                'Segoe UI Semibold',
                10.5,
                [System.Drawing.FontStyle]::Bold
            )
            $runButton.FlatStyle = 'Flat'
            $runButton.FlatAppearance.BorderSize = 1
            $runButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(65, 65, 75)
            $runButton.BackColor = $ColorBtn
            $runButton.ForeColor = $ColorText
            $runButton.Cursor = 'Hand'
            $runButton.FlatAppearance.MouseOverBackColor = $ColorBtnHover

            $runButton.Add_Click({
                try {
                    Start-LauncherProgram -Program $this.Tag
                }
                catch {
                    Show-LauncherMessage `
                        -Title 'Loi' `
                        -Icon Error `
                        -Message $_.Exception.Message
                }
            })

            $panel.Controls.Add($runButton)

            # Nut xoa
            $deleteButton = New-Object System.Windows.Forms.Button
            $deleteButton.Text = 'X'
            $deleteButton.Tag = $program
            $deleteButton.Size = New-Object System.Drawing.Size(52, 46)
            $deleteButton.Location = New-Object System.Drawing.Point(606, $y)
            $deleteButton.Font = New-Object System.Drawing.Font(
                'Segoe UI',
                11,
                [System.Drawing.FontStyle]::Bold
            )
            $deleteButton.FlatStyle = 'Flat'
            $deleteButton.FlatAppearance.BorderSize = 0
            $deleteButton.BackColor = $ColorBtnDel
            $deleteButton.ForeColor = $ColorText
            $deleteButton.Cursor = 'Hand'

            $deleteButton.Add_Click({
                $selectedProgram = $this.Tag

                $confirmation = [System.Windows.Forms.MessageBox]::Show(
                    "Xoa '$($selectedProgram.Name)' khoi danh sach?",
                    'Xac nhan',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )

                if (
                    $confirmation -eq
                    [System.Windows.Forms.DialogResult]::Yes
                ) {
                    $remainingPrograms = @(
                        Load-Programs |
                            Where-Object {
                                -not (
                                    [string]::Equals(
                                        [string]$_.Name,
                                        [string]$selectedProgram.Name,
                                        [System.StringComparison]::OrdinalIgnoreCase
                                    ) -and
                                    [string]::Equals(
                                        [string]$_.Path,
                                        [string]$selectedProgram.Path,
                                        [System.StringComparison]::OrdinalIgnoreCase
                                    )
                                )
                            }
                    )

                    Save-Programs -Programs $remainingPrograms
                    Refresh-ButtonList
                }
            })

            $panel.Controls.Add($deleteButton)

            $y += 54
        }
    }
    finally {
        $panel.ResumeLayout()
    }
}

# ============================================================
# NUT THEM CHUONG TRINH
# ============================================================

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = '+  Them chuong trinh'
$btnAdd.Location = New-Object System.Drawing.Point(28, 552)
$btnAdd.Size = New-Object System.Drawing.Size(335, 48)
$btnAdd.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    10,
    [System.Drawing.FontStyle]::Bold
)
$btnAdd.FlatStyle = 'Flat'
$btnAdd.FlatAppearance.BorderSize = 0
$btnAdd.BackColor = $ColorGreen
$btnAdd.ForeColor = $ColorText
$btnAdd.Cursor = 'Hand'
$btnAdd.FlatAppearance.MouseOverBackColor = (
    [System.Drawing.Color]::FromArgb(60, 150, 105)
)

$btnAdd.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    try {
        $dialog.Filter = 'PowerShell Scripts (*.ps1)|*.ps1'
        $dialog.Title = 'Chon file .ps1'
        $dialog.InitialDirectory = $Paths.Scripts
        $dialog.CheckFileExists = $true
        $dialog.Multiselect = $false

        if (
            $dialog.ShowDialog($form) -ne
            [System.Windows.Forms.DialogResult]::OK
        ) {
            return
        }

        $selectedPath = [System.IO.Path]::GetFullPath($dialog.FileName)
        $defaultName = [System.IO.Path]::GetFileNameWithoutExtension(
            $selectedPath
        )

        $displayName = [Microsoft.VisualBasic.Interaction]::InputBox(
            'Nhap ten hien thi cho chuong trinh:',
            'Them chuong trinh',
            $defaultName
        )

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            return
        }

        $portablePath = ConvertTo-PortablePath -Path $selectedPath
        $programs = @(Load-Programs)

        $duplicate = @(
            $programs |
                Where-Object {
                    $existingResolvedPath = Resolve-ProgramPath `
                        -Path ([string]$_.Path)

                    [string]::Equals(
                        $existingResolvedPath,
                        $selectedPath,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        )

        if ($duplicate.Count -gt 0) {
            Show-LauncherMessage `
                -Title 'Chuong trinh da ton tai' `
                -Icon Information `
                -Message (
                    "Script nay da co trong danh sach:`r`n`r`n" +
                    "$selectedPath"
                )

            return
        }

        $programs += [PSCustomObject]@{
            Name = $displayName.Trim()
            Path = $portablePath
        }

        Save-Programs -Programs $programs
        Refresh-ButtonList

        Write-LauncherLog `
            -Message "Da them '$($displayName.Trim())': $portablePath"
    }
    catch {
        $errorMessage = $_.Exception.Message

        Write-LauncherLog `
            -Level ERROR `
            -Message "Loi them chuong trinh: $errorMessage"

        Show-LauncherMessage `
            -Title 'Loi them chuong trinh' `
            -Icon Error `
            -Message $errorMessage
    }
    finally {
        $dialog.Dispose()
    }
})

$form.Controls.Add($btnAdd)

# ============================================================
# NUT SUA FILE CAU HINH
# ============================================================

$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = 'Sua file cau hinh'
$btnEdit.Location = New-Object System.Drawing.Point(383, 552)
$btnEdit.Size = New-Object System.Drawing.Size(335, 48)
$btnEdit.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$btnEdit.FlatStyle = 'Flat'
$btnEdit.FlatAppearance.BorderSize = 1
$btnEdit.FlatAppearance.BorderColor = $ColorAccent
$btnEdit.BackColor = $ColorPanel
$btnEdit.ForeColor = $ColorText
$btnEdit.Cursor = 'Hand'

$btnEdit.Add_Click({
    try {
        Start-Process `
            -FilePath 'notepad.exe' `
            -ArgumentList "`"$ConfigPath`"" `
            -ErrorAction Stop
    }
    catch {
        Show-LauncherMessage `
            -Title 'Loi mo cau hinh' `
            -Icon Error `
            -Message $_.Exception.Message
    }
})

$form.Controls.Add($btnEdit)

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = 'Version {0} | Developed by Hugo Le Chi Quoc Hung' -f $AppVersion
$footerLabel.Location = New-Object System.Drawing.Point(28, 603)
$footerLabel.Size = New-Object System.Drawing.Size(690, 16)
$footerLabel.TextAlign = 'MiddleRight'
$footerLabel.ForeColor = $ColorSubText
$footerLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$footerLabel.TabStop = $false
$footerLabel.AccessibleName = 'Application version and developer'
$form.Controls.Add($footerLabel)

$clockTimer = New-Object System.Windows.Forms.Timer
$clockTimer.Interval = 1000
$clockTimer.Add_Tick({
    $clockLabel.Text = Get-LauncherDateTimeText
    Update-WorkdayStatus
})
$clockTimer.Start()

# Khi nguoi dung quay lai Launcher sau khi sua JSON bang Notepad,
# co the nhan F5 de tai lai danh sach.
$form.KeyPreview = $true

$form.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
        Refresh-ButtonList
    }
})

$form.Add_FormClosed({
    $clockTimer.Stop()
    $clockTimer.Dispose()
    Write-LauncherLog -Message 'Launcher da dong.'
})

Update-WorkdayStatus
Refresh-ButtonList

[void]$form.ShowDialog()
