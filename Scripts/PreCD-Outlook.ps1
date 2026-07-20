param(
    [ValidateSet("XC5","XN5","Both")][string]$Site,
    [string]$FilePath,
    [string]$BatchId,
    [string]$FilePathXc5,
    [string]$BatchIdXc5,
    [string]$FilePathXn5,
    [string]$BatchIdXn5,
    [int]$SignatureDelaySeconds = 3,
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"

if ($PSScriptRoot -and $PSScriptRoot.Trim()) {
    $ScriptDir = $PSScriptRoot
} else {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}

. (Join-Path $ScriptDir "AutoTools.Common.ps1")
$Paths = Initialize-AutoToolsPaths -StartPath $ScriptDir
$Config = Get-AutoToolsConfig -Paths $Paths

function HtmlEncode([string]$Value) {
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Normalize-Key([string]$Value) {
    return ([string]$Value).Trim().ToUpperInvariant()
}

function Format-PreCdDate([datetime]$Date) {
    $day = $Date.Day
    $suffix = "th"
    if (($day % 100) -notin 11,12,13) {
        switch ($day % 10) {
            1 { $suffix = "st" }
            2 { $suffix = "nd" }
            3 { $suffix = "rd" }
        }
    }
    return "{0} {1}{2}" -f $Date.ToString("MMM"), $day, $suffix
}

function Get-SpreadsheetCellMap($Row, [System.Xml.XmlNamespaceManager]$Ns) {
    $map = @{}
    $col = 1
    foreach ($cell in $Row.SelectNodes("ss:Cell", $Ns)) {
        $index = $cell.GetAttribute("Index", "urn:schemas-microsoft-com:office:spreadsheet")
        if ($index) { $col = [int]$index }
        $data = $cell.SelectSingleNode("ss:Data", $Ns)
        $map[$col] = if ($data) { [string]$data.InnerText } else { "" }
        $col++
    }
    return $map
}

function Get-PreCdRows([string]$XmlPath, [string]$MailSite, [string]$MailBatchId) {
    if (-not (Test-Path -LiteralPath $XmlPath -PathType Leaf)) {
        throw "File not found: $XmlPath"
    }

    [xml]$xml = Get-Content -LiteralPath $XmlPath -Raw -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")

    $rows = @($xml.SelectNodes("//ss:Worksheet[1]/ss:Table/ss:Row", $ns))
    $headerIndex = -1
    $headers = $null
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $map = Get-SpreadsheetCellMap $rows[$i] $ns
        if ((Normalize-Key $map[1]) -eq "VENDOR CODE" -and (Normalize-Key $map[6]).StartsWith("PART_NO")) {
            $headerIndex = $i
            $headers = $map
            break
        }
    }
    if ($headerIndex -lt 0) { throw "Cannot find Process-CDs header row in: $XmlPath" }

    $dateCols = @{}
    foreach ($key in $headers.Keys) {
        if ([int]$key -le 10) { continue }
        if ([string]::IsNullOrWhiteSpace($headers[$key])) { continue }
        $dateCols[[int]$key] = $headers[$key]
    }
    if ($dateCols.Count -eq 0) { throw "No request date columns found in: $XmlPath" }

    $out = @()
    for ($i = $headerIndex + 1; $i -lt $rows.Count; $i++) {
        $map = Get-SpreadsheetCellMap $rows[$i] $ns
        if ([string]::IsNullOrWhiteSpace($map[1]) -and [string]::IsNullOrWhiteSpace($map[6])) { continue }

        foreach ($dateCol in ($dateCols.Keys | Sort-Object)) {
            $qty = ([string]$map[$dateCol]).Trim()
            if ([string]::IsNullOrWhiteSpace($qty)) { continue }
            if ($qty -match '^[\d,.\s]+$' -and ([decimal]($qty -replace ',', '') -eq 0)) { continue }

            $dateText = $dateCols[$dateCol]
            $requestDate = $dateText
            $parsedDate = [datetime]::MinValue
            if ([datetime]::TryParse($dateText, [ref]$parsedDate)) {
                $requestDate = $parsedDate.ToString("yyyy-MM-dd")
            }

            $out += [pscustomobject]@{
                BatchID     = $MailBatchId
                Action      = "NEW"
                VendorCode  = ([string]$map[1]).Trim()
                VendorName  = ([string]$map[2]).Trim()
                ItemNo      = ([string]$map[6]).Trim()
                ItemRev     = ([string]$map[7]).Trim()
                RequestDate = $requestDate
                RequestTime = ""
                RequestQty  = $qty
                Site        = $MailSite
            }
        }
    }

    if ($out.Count -eq 0) { throw "No Pre-CD rows with quantity found in: $XmlPath" }
    return @($out)
}

function Get-SupplierMail([string]$SupplierPath, [string]$VendorCode, [string]$VendorName) {
    if (-not (Test-Path -LiteralPath $SupplierPath -PathType Leaf)) {
        throw "Cannot find supplier file: $SupplierPath"
    }

    $suppliers = Import-Csv -LiteralPath $SupplierPath -Encoding UTF8
    $vendorNameKey = Normalize-Key $VendorName
    $row = $suppliers | Where-Object {
        ([string]$_.VendorCode).Trim() -eq ([string]$VendorCode).Trim() -or
        (Normalize-Key $_.VendorName) -eq $vendorNameKey
    } | Select-Object -First 1

    if (-not $row) {
        return [pscustomobject]@{ To = ""; CC = ""; Warning = "No supplier row for $VendorName" }
    }

    return [pscustomobject]@{
        To = [string]$row.'Email To'
        CC = [string]$row.'Email CC'
        Warning = if ([string]::IsNullOrWhiteSpace([string]$row.'Email To')) { "Missing Email To in suppliers.csv for $VendorName" } else { "" }
    }
}

function New-PreCdTableHtml($Rows) {
    $trs = foreach ($r in $Rows) {
        "<tr><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.BatchID)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.Action)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.VendorCode)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.VendorName)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.ItemNo)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.ItemRev)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.RequestDate)</td><td style='border:1px solid #777;padding:2px 8px'>$(HtmlEncode $r.RequestTime)</td><td style='border:1px solid #777;padding:2px 8px;text-align:right'>$(HtmlEncode $r.RequestQty)</td></tr>"
    }

    return @"
<table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt">
<tr>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">BATCH ID</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">ACTION</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">VENDOR CODE</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">VENDOR NAME</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">ITEM NO</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">ITEM REV</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">REQUEST DATE</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">REQUEST TIME</th>
<th style="border:1px solid #777;padding:2px 8px;text-align:left">REQUEST QTY</th>
</tr>
$($trs -join "`r`n")
</table>
<p>&nbsp;</p>
"@
}

function New-PreCdHtml($Rows, [datetime]$MailDate) {
    $dateLabel = Format-PreCdDate $MailDate
    $siteText = (($Rows | Select-Object -ExpandProperty Site -Unique) -join "/")
    $sections = foreach ($group in ($Rows | Group-Object Site | Sort-Object Name)) {
        "<p>$(HtmlEncode $group.Name)</p>`r`n$(New-PreCdTableHtml -Rows $group.Group)"
    }

    return @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000">
<p><b>Dear supplier,</b></p>
<p>Pls refer below pre-CDs $(HtmlEncode $siteText) on $(HtmlEncode $dateLabel) . Please note with thanks.</p>
<ol>
<li>Please make sure your stock can cover below CDs quantity before opening it</li>
<li>If any items have PO shortage or any issue, pls raise immediately to not impact CD open/ delivery</li>
<li>Regarding the Customs Declaration, once the supplier opens the CDs, the supplier shall bear full responsibility for the completeness and accuracy of the goods in terms of both quantity and quality prior to the declaration submission. Supplier must ensure the delivery is carried out after completing customs declaration obligations at both ends and fully matching the ASN form. TTI shall not be held responsible for any deliveries made prior to the completion of customs formalities</li>
</ol>
$($sections -join "`r`n")
</div>
"@
}

function Show-PreCdInputDialog {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Pre-CD Outlook Mail"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ClientSize = New-Object System.Drawing.Size(760,230)

    $siteLabel = New-Object System.Windows.Forms.Label
    $siteLabel.Text = "Send site"
    $siteLabel.Location = New-Object System.Drawing.Point(14,18)
    $siteLabel.AutoSize = $true
    $form.Controls.Add($siteLabel)

    $siteBox = New-Object System.Windows.Forms.ComboBox
    $siteBox.DropDownStyle = "DropDownList"
    [void]$siteBox.Items.Add("XC5")
    [void]$siteBox.Items.Add("XN5")
    [void]$siteBox.Items.Add("Both")
    $siteBox.SelectedIndex = 0
    $siteBox.Location = New-Object System.Drawing.Point(105,14)
    $siteBox.Width = 90
    $form.Controls.Add($siteBox)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Choose Both when one supplier has XC5 and XN5 in the same email."
    $hint.Location = New-Object System.Drawing.Point(220,18)
    $hint.AutoSize = $true
    $form.Controls.Add($hint)

    $batchHeader = New-Object System.Windows.Forms.Label
    $batchHeader.Text = "Batch ID"
    $batchHeader.Location = New-Object System.Drawing.Point(105,42)
    $batchHeader.AutoSize = $true
    $form.Controls.Add($batchHeader)

    $fileHeader = New-Object System.Windows.Forms.Label
    $fileHeader.Text = "Process-CDs XML file"
    $fileHeader.Location = New-Object System.Drawing.Point(245,42)
    $fileHeader.AutoSize = $true
    $form.Controls.Add($fileHeader)

    function Add-InputRow([string]$rowSite, [int]$top) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $rowSite
        $label.Location = New-Object System.Drawing.Point(14,$top)
        $label.Size = New-Object System.Drawing.Size(70,24)
        $form.Controls.Add($label)

        $batch = New-Object System.Windows.Forms.TextBox
        $batch.Location = New-Object System.Drawing.Point(105,$top)
        $batch.Width = 125
        $form.Controls.Add($batch)

        $file = New-Object System.Windows.Forms.TextBox
        $file.Location = New-Object System.Drawing.Point(245,$top)
        $file.Width = 390
        $form.Controls.Add($file)

        $browse = New-Object System.Windows.Forms.Button
        $browse.Text = "Browse"
        $browse.Location = New-Object System.Drawing.Point(650,($top - 2))
        $browse.Width = 85
        $browse.Tag = [pscustomobject]@{
            Site = $rowSite
            TextBox = $file
        }
        $form.Controls.Add($browse)

        $browse.Add_Click({
            $target = $this.Tag
            $bu = if ($target.Site -eq "XC5") { "TC5" } else { "TN5" }
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = "Select $($target.Site) Process-CDs output XML"
            $dlg.Filter = "Excel XML (*.xml)|*.xml|All files (*.*)|*.*"
            $dlg.InitialDirectory = Join-Path $Config.ProcessCDsOutput $bu
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $target.TextBox.Text = $dlg.FileName
            }
        })

        return [pscustomobject]@{
            Site = $rowSite
            Label = $label
            Batch = $batch
            File = $file
            Browse = $browse
        }
    }

    $xc5Row = Add-InputRow -rowSite "XC5" -top 62
    $xn5Row = Add-InputRow -rowSite "XN5" -top 106

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Create Draft"
    $ok.Location = New-Object System.Drawing.Point(545,175)
    $ok.Width = 95
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Location = New-Object System.Drawing.Point(650,175)
    $cancel.Width = 85
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    function Update-InputRows {
        $selected = [string]$siteBox.SelectedItem
        foreach ($row in @($xc5Row, $xn5Row)) {
            $enabled = ($selected -eq "Both" -or $selected -eq $row.Site)
            $row.Label.Enabled = $enabled
            $row.Batch.Enabled = $enabled
            $row.File.Enabled = $enabled
            $row.Browse.Enabled = $enabled
        }
    }

    $siteBox.Add_SelectedIndexChanged({ Update-InputRows })
    Update-InputRows

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel
    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $selected = [string]$siteBox.SelectedItem
    $rows = @()
    foreach ($row in @($xc5Row, $xn5Row)) {
        if ($selected -eq "Both" -or $selected -eq $row.Site) {
            $rows += [pscustomobject]@{
                Site = $row.Site
                BatchId = $row.Batch.Text.Trim()
                FilePath = $row.File.Text.Trim()
            }
        }
    }
    return @($rows)
}

function Get-PreCdSelections {
    if (-not $Site) {
        $inputRows = Show-PreCdInputDialog
        if (-not $inputRows) { return @() }
        return @($inputRows)
    }

    if ($Site -eq "Both") {
        return @(
            [pscustomobject]@{ Site = "XC5"; BatchId = if ($BatchIdXc5) { $BatchIdXc5 } else { $BatchId }; FilePath = if ($FilePathXc5) { $FilePathXc5 } else { $FilePath } },
            [pscustomobject]@{ Site = "XN5"; BatchId = if ($BatchIdXn5) { $BatchIdXn5 } else { $BatchId }; FilePath = $FilePathXn5 }
        )
    }

    return @(
        [pscustomobject]@{ Site = $Site; BatchId = $BatchId; FilePath = $FilePath }
    )
}

function Assert-SameSupplier($Rows) {
    $vendorKeys = @($Rows | ForEach-Object { "{0}|{1}" -f $_.VendorCode, (Normalize-Key $_.VendorName) } | Select-Object -Unique)
    if ($vendorKeys.Count -gt 1) {
        throw "Selected files contain more than one supplier. Please create separate emails."
    }
}

if ($SelfTest) {
    $sample = Join-Path $Paths.Root "Output\Process-CDs\TC5\DONE_KHGEARS VIETNAM COMPANY LIMITED\KHGEARS VIETNAM COMPANY LIMITED.xml"
    $rows = Get-PreCdRows -XmlPath $sample -MailSite "XC5" -MailBatchId "1755952"
    $rows += Get-PreCdRows -XmlPath $sample -MailSite "XN5" -MailBatchId "1755953"
    Assert-SameSupplier -Rows $rows
    if ($rows.Count -lt 1) { throw "SelfTest failed: no rows parsed" }
    if ($rows[0].VendorCode -ne "11500") { throw "SelfTest failed: vendor code" }
    if ($rows[0].RequestDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw "SelfTest failed: request date" }
    $html = New-PreCdHtml -Rows $rows -MailDate (Get-Date)
    if ($html -notmatch 'XC5/XN5') { throw "SelfTest failed: combined site text" }
    Write-Host "SelfTest OK: parsed $($rows.Count) rows from sample XML and rendered combined XC5/XN5 mail."
    return
}

$selections = @(Get-PreCdSelections)
if ($selections.Count -eq 0) { return }

$rows = @()
foreach ($selection in $selections) {
    if ([string]::IsNullOrWhiteSpace([string]$selection.BatchId)) { throw "$($selection.Site) Batch ID is required." }
    if ([string]::IsNullOrWhiteSpace([string]$selection.FilePath)) { throw "$($selection.Site) XML file is required." }
    $rows += Get-PreCdRows -XmlPath $selection.FilePath -MailSite $selection.Site -MailBatchId $selection.BatchId
}
Assert-SameSupplier -Rows $rows

$vendorCode = $rows[0].VendorCode
$vendorName = $rows[0].VendorName
$mailDate = Get-Date
$dateLabel = Format-PreCdDate $mailDate
$subject = "{0}_Pre CDs {1}" -f $vendorName, $dateLabel
$supplier = Get-SupplierMail -SupplierPath $Config.SupplierMaster -VendorCode $vendorCode -VendorName $vendorName
$html = New-PreCdHtml -Rows $rows -MailDate $mailDate

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Display()
Start-Sleep -Seconds $SignatureDelaySeconds
$signature = [string]$mail.HTMLBody

$mail.To = $supplier.To
$mail.CC = $supplier.CC
$mail.Subject = $subject
$mail.HTMLBody = $html + $signature
$mail.Display()

if ($supplier.Warning) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show($supplier.Warning, "Pre-CD Outlook Mail", "OK", "Warning") | Out-Null
}

Write-Host "Draft created: $subject"
