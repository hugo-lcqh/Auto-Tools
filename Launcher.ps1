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
$AppVersion = '1.2.0'

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
# BANG MAU GIAO DIEN - INDUSTRIAL OPERATIONS
# ============================================================

$ColorBg             = [System.Drawing.Color]::FromArgb(248, 250, 252)
$ColorPanel          = [System.Drawing.Color]::White
$ColorHeader         = [System.Drawing.Color]::FromArgb(51, 65, 85)
$ColorHeaderText     = [System.Drawing.Color]::White
$ColorHeaderSubText  = [System.Drawing.Color]::FromArgb(226, 232, 240)
$ColorHeaderMuted    = [System.Drawing.Color]::FromArgb(203, 213, 225)
$ColorBtn            = [System.Drawing.Color]::FromArgb(51, 65, 85)
$ColorBtnHover       = [System.Drawing.Color]::FromArgb(30, 41, 59)
$ColorBtnDel         = [System.Drawing.Color]::FromArgb(185, 28, 28)
$ColorAccent         = [System.Drawing.Color]::FromArgb(245, 158, 11)
$ColorText           = [System.Drawing.Color]::FromArgb(15, 23, 42)
$ColorSubText        = [System.Drawing.Color]::FromArgb(71, 85, 105)
$ColorCard           = [System.Drawing.Color]::White
$ColorMuted          = [System.Drawing.Color]::FromArgb(241, 245, 249)
$ColorBorder         = [System.Drawing.Color]::FromArgb(203, 213, 225)
$ColorGreen          = [System.Drawing.Color]::FromArgb(4, 120, 87)
$ColorGreenHover     = [System.Drawing.Color]::FromArgb(6, 95, 70)
$ColorGreenSurface   = [System.Drawing.Color]::FromArgb(236, 253, 245)
$ColorWarning        = [System.Drawing.Color]::FromArgb(180, 83, 9)
$ColorWarningSurface = [System.Drawing.Color]::FromArgb(255, 251, 235)
$ColorDangerSurface  = [System.Drawing.Color]::FromArgb(254, 242, 242)
$ColorDangerPressed  = [System.Drawing.Color]::FromArgb(254, 226, 226)

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
        $text = 'Shift 07:30-16:00 | Not started'
    }
    elseif ($now -ge $end) {
        $percent = 0
        $text = 'Shift 07:30-16:00 | Complete'
    }
    else {
        $remaining = $end - $now
        $percent = [Math]::Round(($remaining.TotalSeconds / $totalSeconds) * 100)
        $text = 'Shift 07:30-16:00 | {0}h {1}m remaining ({2}%)' -f `
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
    $setupForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

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
    $textBox.AccessibleName = 'Supplier names and vendor codes'
    $setupForm.Controls.Add($textBox)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(16, 448)
    $statusLabel.Size = New-Object System.Drawing.Size(710, 22)
    $statusLabel.ForeColor = $ColorWarning
    $setupForm.Controls.Add($statusLabel)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Tao suppliers.csv'
    $btnSave.Location = New-Object System.Drawing.Point(466, 478)
    $btnSave.Size = New-Object System.Drawing.Size(126, 32)
    $btnSave.BackColor = $ColorGreen
    $btnSave.ForeColor = $ColorHeaderText
    $btnSave.FlatStyle = 'Flat'
    $btnSave.FlatAppearance.BorderSize = 0
    $btnSave.FlatAppearance.MouseOverBackColor = $ColorGreenHover
    $btnSave.AccessibleName = 'Create suppliers.csv'
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::None
    $setupForm.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Huy'
    $btnCancel.Location = New-Object System.Drawing.Point(606, 478)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 32)
    $btnCancel.BackColor = $ColorPanel
    $btnCancel.ForeColor = $ColorText
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.FlatAppearance.BorderColor = $ColorBtn
    $btnCancel.AccessibleName = 'Cancel supplier setup'
    $setupForm.Controls.Add($btnCancel)

    $setupForm.AcceptButton = $btnSave
    $setupForm.CancelButton = $btnCancel

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
$form.ClientSize = New-Object System.Drawing.Size(980, 680)
$form.MinimumSize = New-Object System.Drawing.Size(860, 620)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox = $true
$form.BackColor = $ColorBg
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.TopMost = $false
$form.KeyPreview = $true

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 3
$rootLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$rootLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        116
    )
))
$rootLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$rootLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        42
    )
))
$form.Controls.Add($rootLayout)

# Header
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Fill'
$headerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$headerPanel.BackColor = $ColorHeader
$rootLayout.Controls.Add($headerPanel, 0, 0)

$headerAccent = New-Object System.Windows.Forms.Panel
$headerAccent.Dock = 'Left'
$headerAccent.Width = 6
$headerAccent.BackColor = $ColorAccent
$headerPanel.Controls.Add($headerAccent)

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = 'Fill'
$headerLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$headerLayout.Padding = New-Object System.Windows.Forms.Padding(28, 16, 28, 12)
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
$headerLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        60
    )
))
$headerLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        40
    )
))
$headerPanel.Controls.Add($headerLayout)

$headerLeft = New-Object System.Windows.Forms.Panel
$headerLeft.Dock = 'Fill'
$headerLeft.Margin = New-Object System.Windows.Forms.Padding(0)
$headerLeft.BackColor = $ColorHeader
$headerLayout.Controls.Add($headerLeft, 0, 0)

$appLabel = New-Object System.Windows.Forms.Label
$appLabel.Text = 'MATERIAL CONTROL  /  OPERATIONS DESKTOP'
$appLabel.Location = New-Object System.Drawing.Point(0, 0)
$appLabel.AutoSize = $true
$appLabel.ForeColor = $ColorAccent
$appLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    8.5,
    [System.Drawing.FontStyle]::Bold
)
$headerLeft.Controls.Add($appLabel)

$header = New-Object System.Windows.Forms.Label
$header.Text = 'Automation Launcher'
$header.Location = New-Object System.Drawing.Point(0, 24)
$header.AutoSize = $true
$header.ForeColor = $ColorHeaderText
$header.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    21,
    [System.Drawing.FontStyle]::Bold
)
$headerLeft.Controls.Add($header)

$subHeader = New-Object System.Windows.Forms.Label
$subHeader.Text = 'Signed in as {0}  |  Select a validated workflow below.' -f (Get-LauncherUserName)
$subHeader.Location = New-Object System.Drawing.Point(2, 66)
$subHeader.AutoSize = $true
$subHeader.ForeColor = $ColorHeaderSubText
$subHeader.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$headerLeft.Controls.Add($subHeader)

$headerRight = New-Object System.Windows.Forms.TableLayoutPanel
$headerRight.Dock = 'Fill'
$headerRight.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$headerRight.Padding = New-Object System.Windows.Forms.Padding(0)
$headerRight.BackColor = $ColorHeader
$headerRight.ColumnCount = 1
$headerRight.RowCount = 2
$headerRight.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$headerRight.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        34
    )
))
$headerRight.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$headerLayout.Controls.Add($headerRight, 1, 0)

$clockLabel = New-Object System.Windows.Forms.Label
$clockLabel.Text = Get-LauncherDateTimeText
$clockLabel.Dock = 'Fill'
$clockLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$clockLabel.TextAlign = 'MiddleRight'
$clockLabel.ForeColor = $ColorHeaderText
$clockLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    9.5,
    [System.Drawing.FontStyle]::Bold
)
$clockLabel.AccessibleName = 'Current date and time'
$headerRight.Controls.Add($clockLabel, 0, 0)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = 'ROOT  {0}' -f $RootPath
$folderLabel.Dock = 'Fill'
$folderLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$folderLabel.TextAlign = 'TopRight'
$folderLabel.ForeColor = $ColorHeaderMuted
$folderLabel.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$folderLabel.AutoEllipsis = $true
$folderLabel.AccessibleName = 'Auto Tools root folder'
$headerRight.Controls.Add($folderLabel, 0, 1)

# Main content
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$contentPanel.Padding = New-Object System.Windows.Forms.Padding(26, 18, 26, 14)
$contentPanel.BackColor = $ColorBg
$rootLayout.Controls.Add($contentPanel, 0, 1)

$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.Dock = 'Fill'
$contentLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$contentLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$contentLayout.ColumnCount = 1
$contentLayout.RowCount = 4
$contentLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$contentLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        84
    )
))
$contentLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        42
    )
))
$contentLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$contentLayout.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        58
    )
))
$contentPanel.Controls.Add($contentLayout)

# Shift status card
$shiftPanel = New-Object System.Windows.Forms.Panel
$shiftPanel.Dock = 'Fill'
$shiftPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$shiftPanel.BackColor = $ColorPanel
$shiftPanel.BorderStyle = 'FixedSingle'
$contentLayout.Controls.Add($shiftPanel, 0, 0)

$shiftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$shiftLayout.Dock = 'Fill'
$shiftLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$shiftLayout.Padding = New-Object System.Windows.Forms.Padding(16, 10, 16, 10)
$shiftLayout.ColumnCount = 2
$shiftLayout.RowCount = 1
$shiftLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        190
    )
))
$shiftLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$shiftPanel.Controls.Add($shiftLayout)

$shiftIdentity = New-Object System.Windows.Forms.Panel
$shiftIdentity.Dock = 'Fill'
$shiftIdentity.Margin = New-Object System.Windows.Forms.Padding(0)
$shiftIdentity.BackColor = $ColorPanel
$shiftLayout.Controls.Add($shiftIdentity, 0, 0)

$shiftTitle = New-Object System.Windows.Forms.Label
$shiftTitle.Text = 'SHIFT CONTROL'
$shiftTitle.Location = New-Object System.Drawing.Point(0, 2)
$shiftTitle.AutoSize = $true
$shiftTitle.ForeColor = $ColorText
$shiftTitle.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    10,
    [System.Drawing.FontStyle]::Bold
)
$shiftIdentity.Controls.Add($shiftTitle)

$shiftHours = New-Object System.Windows.Forms.Label
$shiftHours.Text = 'Standard window  07:30 - 16:00'
$shiftHours.Location = New-Object System.Drawing.Point(0, 28)
$shiftHours.AutoSize = $true
$shiftHours.ForeColor = $ColorSubText
$shiftHours.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$shiftIdentity.Controls.Add($shiftHours)

$shiftDetails = New-Object System.Windows.Forms.TableLayoutPanel
$shiftDetails.Dock = 'Fill'
$shiftDetails.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$shiftDetails.Padding = New-Object System.Windows.Forms.Padding(0)
$shiftDetails.ColumnCount = 2
$shiftDetails.RowCount = 2
$shiftDetails.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$shiftDetails.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        112
    )
))
$shiftDetails.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        30
    )
))
$shiftDetails.RowStyles.Add((
    New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        12
    )
))
$shiftLayout.Controls.Add($shiftDetails, 1, 0)

$workdayLabel = New-Object System.Windows.Forms.Label
$workdayLabel.Dock = 'Fill'
$workdayLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$workdayLabel.TextAlign = 'MiddleLeft'
$workdayLabel.ForeColor = $ColorSubText
$workdayLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.8)
$workdayLabel.AccessibleName = 'Current shift status'
$shiftDetails.Controls.Add($workdayLabel, 0, 0)

$shiftPercentLabel = New-Object System.Windows.Forms.Label
$shiftPercentLabel.Dock = 'Fill'
$shiftPercentLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$shiftPercentLabel.TextAlign = 'MiddleRight'
$shiftPercentLabel.ForeColor = $ColorGreen
$shiftPercentLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    8.5,
    [System.Drawing.FontStyle]::Bold
)
$shiftPercentLabel.AccessibleName = 'Shift time remaining percentage'
$shiftDetails.Controls.Add($shiftPercentLabel, 1, 0)

$workdayTrack = New-Object System.Windows.Forms.Panel
$workdayTrack.Dock = 'Fill'
$workdayTrack.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$workdayTrack.BackColor = $ColorMuted
$shiftDetails.Controls.Add($workdayTrack, 0, 1)
$shiftDetails.SetColumnSpan($workdayTrack, 2)

$workdayFill = New-Object System.Windows.Forms.Panel
$workdayFill.Dock = 'Left'
$workdayFill.Width = 0
$workdayFill.BackColor = $ColorGreen
$workdayTrack.Controls.Add($workdayFill)

# Tool list heading
$listHeader = New-Object System.Windows.Forms.TableLayoutPanel
$listHeader.Dock = 'Fill'
$listHeader.Margin = New-Object System.Windows.Forms.Padding(0)
$listHeader.Padding = New-Object System.Windows.Forms.Padding(0)
$listHeader.ColumnCount = 2
$listHeader.RowCount = 1
$listHeader.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        40
    )
))
$listHeader.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        60
    )
))
$contentLayout.Controls.Add($listHeader, 0, 1)

$programCountLabel = New-Object System.Windows.Forms.Label
$programCountLabel.Text = 'AUTOMATION TOOLS'
$programCountLabel.Dock = 'Fill'
$programCountLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$programCountLabel.TextAlign = 'MiddleLeft'
$programCountLabel.ForeColor = $ColorText
$programCountLabel.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    10.5,
    [System.Drawing.FontStyle]::Bold
)
$listHeader.Controls.Add($programCountLabel, 0, 0)

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = 'Each workflow opens in its own PowerShell session  |  F5 refreshes this list'
$hintLabel.Dock = 'Fill'
$hintLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$hintLabel.TextAlign = 'MiddleRight'
$hintLabel.ForeColor = $ColorSubText
$hintLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$listHeader.Controls.Add($hintLabel, 1, 0)

# Responsive scrolling list
$programList = New-Object System.Windows.Forms.FlowLayoutPanel
$programList.Dock = 'Fill'
$programList.Margin = New-Object System.Windows.Forms.Padding(0)
$programList.Padding = New-Object System.Windows.Forms.Padding(10)
$programList.AutoScroll = $true
$programList.WrapContents = $false
$programList.FlowDirection = 'TopDown'
$programList.BackColor = $ColorMuted
$programList.BorderStyle = 'FixedSingle'
$programList.TabStop = $false
$contentLayout.Controls.Add($programList, 0, 2)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 100

# Bottom actions
$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.ColumnCount = 3
$actionsLayout.RowCount = 1
$actionsLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        190
    )
))
$actionsLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        190
    )
))
$actionsLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    )
))
$contentLayout.Controls.Add($actionsLayout, 0, 3)

function Update-WorkdayStatus {
    [CmdletBinding()]
    param()

    $status = Get-WorkdayStatus
    $workdayLabel.Text = $status.Text
    $shiftPercentLabel.Text = '{0}% REMAINING' -f $status.Percent
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

function Resize-ProgramRows {
    [CmdletBinding()]
    param()

    $rowWidth = [Math]::Max(
        560,
        $programList.ClientSize.Width - $programList.Padding.Horizontal - 24
    )

    foreach ($control in $programList.Controls) {
        if ($control.Name -eq 'ProgramRow') {
            $control.Width = $rowWidth
        }
    }
}

function Refresh-ButtonList {
    [CmdletBinding()]
    param()

    $programList.SuspendLayout()

    try {
        while ($programList.Controls.Count -gt 0) {
            $oldControl = $programList.Controls[0]
            $programList.Controls.RemoveAt(0)
            $oldControl.Dispose()
        }

        $programs = @(Load-Programs)
        $programCountLabel.Text = 'AUTOMATION TOOLS  /  {0:00} CONFIGURED' -f $programs.Count

        if ($programs.Count -eq 0) {
            $footerStatus.Text = 'LOCAL SYSTEM  /  NO TOOLS CONFIGURED'
            $footerStatus.ForeColor = $ColorWarning

            $emptyLabel = New-Object System.Windows.Forms.Label
            $emptyLabel.Text = (
                "NO WORKFLOWS CONFIGURED`r`n" +
                "Use ADD TOOL to register a PowerShell script."
            )
            $emptyLabel.Size = New-Object System.Drawing.Size(520, 58)
            $emptyLabel.Margin = New-Object System.Windows.Forms.Padding(14, 20, 0, 0)
            $emptyLabel.ForeColor = $ColorSubText
            $emptyLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10)
            $emptyLabel.AccessibleName = 'No automation tools configured'

            $programList.Controls.Add($emptyLabel)
            return
        }

        $programNumber = 0
        $availableProgramCount = 0
        foreach ($program in $programs) {
            $programNumber++
            $resolvedProgramPath = $null
            $isAvailable = $false

            try {
                $resolvedProgramPath = Resolve-ProgramPath -Path ([string]$program.Path)
                $isAvailable = Test-Path -LiteralPath $resolvedProgramPath -PathType Leaf
            }
            catch {
                $isAvailable = $false
            }

            if ($isAvailable) {
                $availableProgramCount++
                $statusText = 'READY'
                $statusColor = $ColorGreen
                $statusSurface = $ColorGreenSurface
            }
            else {
                $statusText = 'MISSING'
                $statusColor = $ColorWarning
                $statusSurface = $ColorWarningSurface
            }

            $programRow = New-Object System.Windows.Forms.Panel
            $programRow.Name = 'ProgramRow'
            $programRow.Size = New-Object System.Drawing.Size(820, 66)
            $programRow.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
            $programRow.BackColor = $ColorCard
            $programRow.BorderStyle = 'FixedSingle'

            $statusStrip = New-Object System.Windows.Forms.Panel
            $statusStrip.Dock = 'Left'
            $statusStrip.Width = 4
            $statusStrip.BackColor = $statusColor
            $programRow.Controls.Add($statusStrip)

            $numberLabel = New-Object System.Windows.Forms.Label
            $numberLabel.Text = '{0:00}' -f $programNumber
            $numberLabel.Location = New-Object System.Drawing.Point(18, 21)
            $numberLabel.Size = New-Object System.Drawing.Size(34, 20)
            $numberLabel.TextAlign = 'MiddleLeft'
            $numberLabel.ForeColor = $ColorSubText
            $numberLabel.Font = New-Object System.Drawing.Font(
                'Consolas',
                9.5,
                [System.Drawing.FontStyle]::Bold
            )
            $programRow.Controls.Add($numberLabel)

            $nameLabel = New-Object System.Windows.Forms.Label
            $nameLabel.Text = [string]$program.Name
            $nameLabel.Location = New-Object System.Drawing.Point(62, 9)
            $nameLabel.Size = New-Object System.Drawing.Size(460, 24)
            $nameLabel.Anchor = 'Top, Left, Right'
            $nameLabel.AutoEllipsis = $true
            $nameLabel.ForeColor = $ColorText
            $nameLabel.Font = New-Object System.Drawing.Font(
                'Segoe UI Semibold',
                10,
                [System.Drawing.FontStyle]::Bold
            )
            $programRow.Controls.Add($nameLabel)

            $pathLabel = New-Object System.Windows.Forms.Label
            $pathLabel.Text = [string]$program.Path
            $pathLabel.Location = New-Object System.Drawing.Point(62, 36)
            $pathLabel.Size = New-Object System.Drawing.Size(460, 19)
            $pathLabel.Anchor = 'Top, Left, Right'
            $pathLabel.AutoEllipsis = $true
            $pathLabel.ForeColor = $ColorSubText
            $pathLabel.Font = New-Object System.Drawing.Font('Consolas', 8.2)
            $toolTip.SetToolTip($pathLabel, [string]$program.Path)
            $programRow.Controls.Add($pathLabel)

            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Text = $statusText
            $statusLabel.Location = New-Object System.Drawing.Point(526, 20)
            $statusLabel.Size = New-Object System.Drawing.Size(72, 24)
            $statusLabel.Anchor = 'Top, Right'
            $statusLabel.TextAlign = 'MiddleCenter'
            $statusLabel.BackColor = $statusSurface
            $statusLabel.ForeColor = $statusColor
            $statusLabel.Font = New-Object System.Drawing.Font(
                'Segoe UI Semibold',
                8,
                [System.Drawing.FontStyle]::Bold
            )
            $statusLabel.AccessibleName = "Tool status: $statusText"
            $programRow.Controls.Add($statusLabel)

            $runButton = New-Object System.Windows.Forms.Button
            $runButton.Text = 'RUN'
            $runButton.Tag = $program
            $runButton.Size = New-Object System.Drawing.Size(88, 38)
            $runButton.Location = New-Object System.Drawing.Point(610, 13)
            $runButton.Anchor = 'Top, Right'
            $runButton.Font = New-Object System.Drawing.Font(
                'Segoe UI Semibold',
                9,
                [System.Drawing.FontStyle]::Bold
            )
            $runButton.FlatStyle = 'Flat'
            $runButton.FlatAppearance.BorderSize = 0
            $runButton.FlatAppearance.MouseOverBackColor = $ColorGreenHover
            $runButton.FlatAppearance.MouseDownBackColor = $ColorGreenHover
            $runButton.BackColor = $ColorGreen
            $runButton.ForeColor = $ColorHeaderText
            $runButton.Cursor = 'Hand'
            $runButton.TabIndex = ($programNumber * 2) - 2
            $runButton.AccessibleName = "Run $($program.Name)"
            $runButton.AccessibleDescription = "Open $($program.Name) in a separate PowerShell window."

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

            $programRow.Controls.Add($runButton)

            $deleteButton = New-Object System.Windows.Forms.Button
            $deleteButton.Text = 'REMOVE'
            $deleteButton.Tag = $program
            $deleteButton.Size = New-Object System.Drawing.Size(96, 38)
            $deleteButton.Location = New-Object System.Drawing.Point(706, 13)
            $deleteButton.Anchor = 'Top, Right'
            $deleteButton.Font = New-Object System.Drawing.Font(
                'Segoe UI Semibold',
                8.5,
                [System.Drawing.FontStyle]::Bold
            )
            $deleteButton.FlatStyle = 'Flat'
            $deleteButton.FlatAppearance.BorderSize = 1
            $deleteButton.FlatAppearance.BorderColor = $ColorBtnDel
            $deleteButton.FlatAppearance.MouseOverBackColor = $ColorDangerSurface
            $deleteButton.FlatAppearance.MouseDownBackColor = $ColorDangerPressed
            $deleteButton.BackColor = $ColorCard
            $deleteButton.ForeColor = $ColorBtnDel
            $deleteButton.Cursor = 'Hand'
            $deleteButton.TabIndex = ($programNumber * 2) - 1
            $deleteButton.AccessibleName = "Remove $($program.Name)"
            $deleteButton.AccessibleDescription = "Remove $($program.Name) from launcher-config.json after confirmation."

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

            $programRow.Controls.Add($deleteButton)
            $programList.Controls.Add($programRow)
        }

        if ($availableProgramCount -eq $programs.Count) {
            $footerStatus.Text = 'LOCAL SYSTEM  /  READY FOR OPERATION'
            $footerStatus.ForeColor = $ColorGreen
        }
        else {
            $missingProgramCount = $programs.Count - $availableProgramCount
            $footerStatus.Text = 'LOCAL SYSTEM  /  {0:00} TOOL(S) NEED ATTENTION' -f $missingProgramCount
            $footerStatus.ForeColor = $ColorWarning
        }

        Resize-ProgramRows
    }
    finally {
        $programList.ResumeLayout()
    }
}

$programList.Add_SizeChanged({ Resize-ProgramRows })

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = 'ADD TOOL'
$btnAdd.Dock = 'Fill'
$btnAdd.Margin = New-Object System.Windows.Forms.Padding(0, 10, 10, 4)
$btnAdd.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    9,
    [System.Drawing.FontStyle]::Bold
)
$btnAdd.FlatStyle = 'Flat'
$btnAdd.FlatAppearance.BorderSize = 0
$btnAdd.BackColor = $ColorGreen
$btnAdd.ForeColor = $ColorHeaderText
$btnAdd.Cursor = 'Hand'
$btnAdd.FlatAppearance.MouseOverBackColor = $ColorGreenHover
$btnAdd.FlatAppearance.MouseDownBackColor = $ColorGreenHover
$btnAdd.TabIndex = 100
$btnAdd.AccessibleName = 'Add PowerShell tool'
$btnAdd.AccessibleDescription = 'Select a PowerShell script and add it to launcher-config.json.'

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

$actionsLayout.Controls.Add($btnAdd, 0, 0)

$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = 'OPEN CONFIG'
$btnEdit.Dock = 'Fill'
$btnEdit.Margin = New-Object System.Windows.Forms.Padding(0, 10, 10, 4)
$btnEdit.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    9,
    [System.Drawing.FontStyle]::Bold
)
$btnEdit.FlatStyle = 'Flat'
$btnEdit.FlatAppearance.BorderSize = 1
$btnEdit.FlatAppearance.BorderColor = $ColorBtn
$btnEdit.FlatAppearance.MouseOverBackColor = $ColorMuted
$btnEdit.FlatAppearance.MouseDownBackColor = $ColorBorder
$btnEdit.BackColor = $ColorPanel
$btnEdit.ForeColor = $ColorText
$btnEdit.Cursor = 'Hand'
$btnEdit.TabIndex = 101
$btnEdit.AccessibleName = 'Open launcher configuration'
$btnEdit.AccessibleDescription = 'Open launcher-config.json in Notepad.'

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

$actionsLayout.Controls.Add($btnEdit, 1, 0)

$actionsHint = New-Object System.Windows.Forms.Label
$actionsHint.Text = 'Configuration changes are loaded with F5'
$actionsHint.Dock = 'Fill'
$actionsHint.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 4)
$actionsHint.TextAlign = 'MiddleRight'
$actionsHint.ForeColor = $ColorSubText
$actionsHint.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$actionsLayout.Controls.Add($actionsHint, 2, 0)

# Footer
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = 'Fill'
$footerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$footerPanel.BackColor = $ColorMuted
$rootLayout.Controls.Add($footerPanel, 0, 2)

$footerLine = New-Object System.Windows.Forms.Panel
$footerLine.Dock = 'Top'
$footerLine.Height = 1
$footerLine.BackColor = $ColorBorder
$footerPanel.Controls.Add($footerLine)

$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = 'Fill'
$footerLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$footerLayout.Padding = New-Object System.Windows.Forms.Padding(26, 4, 26, 3)
$footerLayout.ColumnCount = 2
$footerLayout.RowCount = 1
$footerLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        34
    )
))
$footerLayout.ColumnStyles.Add((
    New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        66
    )
))
$footerPanel.Controls.Add($footerLayout)

$footerStatus = New-Object System.Windows.Forms.Label
$footerStatus.Text = 'LOCAL SYSTEM  /  READY FOR OPERATION'
$footerStatus.Dock = 'Fill'
$footerStatus.Margin = New-Object System.Windows.Forms.Padding(0)
$footerStatus.TextAlign = 'MiddleLeft'
$footerStatus.ForeColor = $ColorGreen
$footerStatus.Font = New-Object System.Drawing.Font(
    'Segoe UI Semibold',
    7.8,
    [System.Drawing.FontStyle]::Bold
)
$footerLayout.Controls.Add($footerStatus, 0, 0)

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = 'Version {0} | Developed by Hugo Le Chi Quoc Hung | Phone: +84 39 5656 909' -f $AppVersion
$footerLabel.Dock = 'Fill'
$footerLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$footerLabel.TextAlign = 'MiddleRight'
$footerLabel.ForeColor = $ColorSubText
$footerLabel.Font = New-Object System.Drawing.Font('Segoe UI', 7.8)
$footerLabel.TabStop = $false
$footerLabel.AccessibleName = 'Application version and developer'
$footerLayout.Controls.Add($footerLabel, 1, 0)

$clockTimer = New-Object System.Windows.Forms.Timer
$clockTimer.Interval = 1000
$clockTimer.Add_Tick({
    $clockLabel.Text = Get-LauncherDateTimeText
    Update-WorkdayStatus
})
$clockTimer.Start()

$form.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
        Refresh-ButtonList
    }
})

$form.Add_FormClosed({
    $clockTimer.Stop()
    $clockTimer.Dispose()
    $toolTip.Dispose()
    Write-LauncherLog -Message 'Launcher da dong.'
})

Update-WorkdayStatus
Refresh-ButtonList

[void]$form.ShowDialog()
