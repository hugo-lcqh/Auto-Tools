<#
====================================================================
 AutoImport.ps1
 Tu dong hoa quy trinh cap nhat data ngay moi
 --------------------------------------------------------------------
 Cac buoc tu dong:
   1. Tim folder ngay moi nhat (nguon) trong thu muc Job Short
   2. Tao folder ngay hom nay, copy toan bo noi dung tu folder nguon
   3. Tim file Import *.zip moi nhat, giai nen va de
      vao folder Import cua folder moi
   4. Mo file working .xlsb, chay 5 macro theo thu tu, luu lai
====================================================================
#>

# ====================== CAU HINH (chinh o day neu can) ======================
. (Join-Path -Path $PSScriptRoot -ChildPath "AutoTools.Common.ps1")
$AutoToolsPaths  = Initialize-AutoToolsPaths -StartPath $PSScriptRoot
$AutoToolsConfig = Get-AutoToolsConfig -Paths $AutoToolsPaths

$RootFolder   = $AutoToolsConfig.JobShortRoot
# Folder rieng chua cac file .zip data
$ZipFolder    = $AutoToolsConfig.JobShortZipFolder
$ZipPattern   = "Import*.zip"          # mau ten file zip can tim (lay file moi nhat)
$WorkingFile  = "20240410 Jun JS.xlsb" # ten file working (giu nguyen o moi folder ngay)
$ImportSub    = "Import"               # ten folder con chua data import

# Danh sach macro chay theo thu tu (chi can ten macro, KHONG kem ten file)
$Macros = @(
    "UnfilterAllSheets",
    "ConvertWorkbooks",
    "Import_Text_File",
    "ClearDataInOverview",
    "Export_Daily"
)

# Quy uoc ten folder ngay: dd.MM  (vd ngay 4 thang 6 -> "04.06")
$TodayName = (Get-Date).ToString("dd.MM")
# ============================================================================

$ErrorActionPreference = "Stop"

function Write-Step($msg) { Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warn2($msg){ Write-Host "    [!]  $msg" -ForegroundColor Yellow }

function Get-ValidTodayZip {
    $zipFile = Get-LatestImportZip -DestinationFolder $ZipFolder -NamePattern $ZipPattern
    if ($zipFile -and $zipFile.LastWriteTime -ge (Get-Date).Date) {
        return $zipFile
    }

    return $null
}

function Get-LatestImportZip {
    param(
        [Parameter(Mandatory)]
        [string]$DestinationFolder,

        [Parameter(Mandatory)]
        [string]$NamePattern
    )

    if (-not (Test-Path -LiteralPath $DestinationFolder -PathType Container)) {
        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
    }

    return Get-ChildItem -LiteralPath $DestinationFolder -Filter $NamePattern -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

try {
    # ---------- KIEM TRA THU MUC GOC ----------
    if (-not (Test-Path $RootFolder)) {
        throw "Khong tim thay thu muc goc: $RootFolder"
    }

    # ---------- BUOC 1: TIM FOLDER NGAY NGUON (moi nhat) ----------
    Write-Step "Buoc 1: Tim folder ngay nguon"
    # Lay cac folder co ten dang dd.MM, bo qua folder 'Older' va folder hom nay neu da ton tai
    $dateFolders = Get-ChildItem -Path $RootFolder -Directory |
        Where-Object { $_.Name -match '^\d{2}\.\d{2}$' -and $_.Name -ne $TodayName }

    if (-not $dateFolders) {
        throw "Khong tim thay folder ngay nao (dang dd.MM) de lam nguon trong $RootFolder"
    }

    # Sap xep theo ngay sua doi moi nhat -> chon lam nguon
    $sourceFolder = $dateFolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Write-OK "Folder nguon: $($sourceFolder.Name)"

    # ---------- BUOC 2: TAO FOLDER NGAY HOM NAY + COPY ----------
    Write-Step "Buoc 2: Tao folder ngay hom nay ($TodayName) va copy noi dung"
    $newFolderPath = Join-Path $RootFolder $TodayName

    if (Test-Path $newFolderPath) {
        Write-Warn2 "Folder $TodayName da ton tai."
        $ans = Read-Host "    Ban co muon GHI DE (xoa va tao lai)? Go Y de tiep tuc, phim khac de DUNG"
        if ($ans -ne 'Y' -and $ans -ne 'y') {
            throw "Da dung theo yeu cau (khong ghi de folder $TodayName)."
        }
        Remove-Item -Path $newFolderPath -Recurse -Force
    }

    # Copy toan bo cau truc tu folder nguon sang folder moi
    Copy-Item -Path $sourceFolder.FullName -Destination $newFolderPath -Recurse -Force
    Write-OK "Da tao $TodayName va copy tu $($sourceFolder.Name)"

    # ---------- BUOC 3: TIM ZIP, GIAI NEN, DE VAO IMPORT ----------
    Write-Step "Buoc 3: Tim file zip moi nhat va giai nen vao Import"

    $zipFile = Get-ValidTodayZip
    if (-not $zipFile) {
        throw "Khong tim thay file '$ZipPattern' moi trong ngay hom nay tai $ZipFolder. Hay dat file zip moi vao folder nay truoc khi chay AutoImport."
    }
    Write-OK "File zip moi nhat: $($zipFile.Name)  (tai luc $($zipFile.LastWriteTime))"

    $importDir = Join-Path $newFolderPath $ImportSub
    if (-not (Test-Path $importDir)) {
        # Neu folder nguon khong co san Import thi tao moi
        New-Item -ItemType Directory -Path $importDir | Out-Null
    } else {
        # Xoa data cu trong Import truoc khi de data moi
        Get-ChildItem -Path $importDir -Recurse -Force | Remove-Item -Recurse -Force
        Write-OK "Da xoa data cu trong folder Import"
    }

    # Giai nen ra thu muc tam, roi xu ly truong hop zip co 1 folder bao ngoai
    $tempExtract = Join-Path $AutoToolsPaths.Temp ("AutoImport_" + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempExtract | Out-Null
    Expand-Archive -Path $zipFile.FullName -DestinationPath $tempExtract -Force

    # Neu zip giai nen ra chi co dung 1 folder cha -> lay noi dung ben trong folder do
    $extractedItems = Get-ChildItem -Path $tempExtract -Force
    if ($extractedItems.Count -eq 1 -and $extractedItems[0].PSIsContainer) {
        $sourceContent = $extractedItems[0].FullName
    } else {
        $sourceContent = $tempExtract
    }

    Copy-Item -Path (Join-Path $sourceContent '*') -Destination $importDir -Recurse -Force
    Remove-Item -Path $tempExtract -Recurse -Force
    Write-OK "Da giai nen va de data moi vao: $importDir"

    # ---------- BUOC 4: MO EXCEL, CHAY MACRO ----------
    Write-Step "Buoc 4: Mo file working va chay macro"
    $workingPath = Join-Path $newFolderPath $WorkingFile
    if (-not (Test-Path $workingPath)) {
        throw "Khong tim thay file working: $workingPath"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true          # de $false neu muon chay an
    $excel.DisplayAlerts = $false
    # Bat dau dong macro automation (tranh popup bao mat lam treo)
    $excel.AutomationSecurity = 1   # msoAutomationSecurityLow

    $wb = $excel.Workbooks.Open($workingPath)
    Write-OK "Da mo $WorkingFile"

    # ---- Khoi dong job nen tu dong bam OK cho cac popup MsgBox cua macro ----
    # Macro co the goi MsgBox (vd "Unfilter") va dung cho nguoi bam OK -> lam treo automation.
    # Job nen nay quet lien tuc, tim cua so popup cua Excel va gui phim Enter de bam OK.
    Write-Step "Khoi dong co che tu dong bam OK cho popup"
    $dismisser = Start-Job -ScriptBlock {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win {
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string c, string n);
    [DllImport("user32.dll")] public static extern IntPtr FindWindowEx(IntPtr p, IntPtr ch, string c, string n);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
}
"@
        # Lien tuc tim cua so popup co class '#32770' (dialog box chuan cua Windows/MsgBox)
        # va tieu de "Microsoft Excel" roi gui lenh dong (WM_CLOSE) / Enter.
        while ($true) {
            $hwnd = [Win]::FindWindow("#32770", "Microsoft Excel")
            if ($hwnd -ne [IntPtr]::Zero) {
                [Win]::SetForegroundWindow($hwnd) | Out-Null
                Start-Sleep -Milliseconds 150
                # Gui phim Enter de bam nut mac dinh (OK)
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            }
            Start-Sleep -Milliseconds 300
        }
    }
    Write-OK "Co che tu dong bam OK dang chay nen"

    try {
        foreach ($m in $Macros) {
            # Goi macro theo dinh dang 'TenFile'!TenMacro de chi dinh dung workbook
            $macroFullName = "'" + $WorkingFile + "'!" + $m
            Write-Host "    -> Dang chay macro: $m" -ForegroundColor White
            $excel.Run($macroFullName)
            Write-OK "Xong: $m"
        }
    }
    finally {
        # Dung job nen sau khi chay xong (du loi hay khong)
        if ($dismisser) {
            Stop-Job  $dismisser -ErrorAction SilentlyContinue
            Remove-Job $dismisser -Force -ErrorAction SilentlyContinue
            Write-OK "Da dung co che tu dong bam OK"
        }
    }

    # Luu va dong
    $wb.Save()
    $wb.Close($true)
    $excel.Quit()

    # Giai phong COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)    | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host " HOAN TAT! Folder moi: $TodayName" -ForegroundColor Green
    Write-Host " Data moi da duoc cap nhat va macro da chay xong." -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Green

    # ---------- GOI SCRIPT CHECK CD & SHORTAGE ----------
    # Chay sau khi da dong file working o tren, de tranh xung dot mo file.
    $checkScript = Join-Path $PSScriptRoot "CheckCD_Shortage.ps1"
    if (Test-Path $checkScript) {
        $ans = Read-Host "Ban co muon chay kiem tra CD & Shortage ngay bay gio? (Y/N)"
        if ($ans -eq 'Y' -or $ans -eq 'y') {
            Write-Host "`n>>> Dang chay CheckCD_Shortage..." -ForegroundColor Cyan
            # Goi truc tiep trong cung cua so; script Check se tu hoi 1 hay 2 tuan
            & $checkScript
        } else {
            Write-Host "Bo qua buoc kiem tra. Ban co the chay rieng CheckCD_Shortage.ps1 sau." -ForegroundColor Yellow
        }
    } else {
        Write-Host "[!] Khong tim thay CheckCD_Shortage.ps1 cung thu muc voi script nay." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host " LOI: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    # Don dep job nen neu con chay
    if ($dismisser) {
        Stop-Job  $dismisser -ErrorAction SilentlyContinue
        Remove-Job $dismisser -Force -ErrorAction SilentlyContinue
    }
    # Don dep Excel neu con mo do loi
    if ($excel) {
        try { $excel.Quit() } catch {}
    }
}
finally {
    Read-Host "Nhan Enter de dong cua so"
}
