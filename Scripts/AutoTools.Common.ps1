# Common path helpers for scripts inside the Auto Tools folder.

function Get-AutoToolsRoot {
    [CmdletBinding()]
    param(
        [string]$StartPath = $PSScriptRoot
    )

    if (-not [string]::IsNullOrWhiteSpace($env:AUTOTOOLS_ROOT)) {
        return [System.IO.Path]::GetFullPath($env:AUTOTOOLS_ROOT)
    }

    if ([string]::IsNullOrWhiteSpace($StartPath)) {
        $StartPath = (Get-Location).Path
    }

    $current = [System.IO.DirectoryInfo][System.IO.Path]::GetFullPath($StartPath)
    while ($null -ne $current) {
        $launcher = Join-Path -Path $current.FullName -ChildPath 'Launcher.ps1'
        $config = Join-Path -Path $current.FullName -ChildPath 'launcher-config.json'

        if (
            (Test-Path -LiteralPath $launcher -PathType Leaf) -and
            (Test-Path -LiteralPath $config -PathType Leaf)
        ) {
            return $current.FullName
        }

        $current = $current.Parent
    }

    throw 'Khong tim thay thu muc goc Auto Tools. Hay chay script ben trong folder Auto Tools hoac qua Launcher.ps1.'
}

function Initialize-AutoToolsPaths {
    [CmdletBinding()]
    param(
        [string]$StartPath = $PSScriptRoot
    )

    $root = Get-AutoToolsRoot -StartPath $StartPath

    $paths = [ordered]@{
        Root    = $root
        Scripts = Join-Path -Path $root -ChildPath 'Scripts'
        Config  = Join-Path -Path $root -ChildPath 'Config'
        Input   = Join-Path -Path $root -ChildPath 'Input'
        Output  = Join-Path -Path $root -ChildPath 'Output'
        Logs    = Join-Path -Path $root -ChildPath 'Logs'
        Temp    = Join-Path -Path $root -ChildPath 'Temp'
    }

    foreach ($directoryPath in $paths.Values) {
        if (-not (Test-Path -LiteralPath $directoryPath -PathType Container)) {
            $null = New-Item -Path $directoryPath -ItemType Directory -Force
        }
    }

    return [PSCustomObject]$paths
}

function Resolve-AutoToolsPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$RootPath
    )

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path.Trim().Trim('"'))
    if ([System.IO.Path]::IsPathRooted($expandedPath)) {
        return [System.IO.Path]::GetFullPath($expandedPath)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $RootPath -ChildPath $expandedPath))
}

function Split-AutoToolsConfigList {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split '[;|]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function ConvertTo-AutoToolsIntQuantity {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $clean = $Value -replace '[,\s]', ''
    $number = 0
    if ([int]::TryParse($clean, [ref]$number)) {
        return $number
    }

    return $null
}

function Get-AutoToolsConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Paths
    )

    $defaults = [ordered]@{
        JobShortRoot             = 'Input\Job Short'
        JobShortZipFolder        = 'Input\Data Job Short - Not Delete'
        ASCPRoot                 = 'Input\ASCP'
        SupplierCommitmentRoot   = 'Input\Supplier Commitment'
        SupplierCommitmentOutput = 'Output\Supplier Commitment'
        SupplierMaster           = 'Config\suppliers.csv'
        ProcessCDsTemplate       = 'Scripts\Process-CDs\Template_PreCD.xml'
        ProcessCDsInput          = 'Input\Process-CDs\cds_input.txt'
        ProcessCDsMoq            = 'Input\Process-CDs\MOQ_Yao I.txt'
        ProcessCDsOutput         = 'Output\Process-CDs'
        ProcessPOTemplate        = 'Scripts\Process-PO\PO_Template_5.xlsx'
        ProcessPOInput           = 'Input\Process-PO\input.txt'
        ProcessPOMoq             = 'Input\Process-PO\MOQ.txt'
        ProcessPOOutput          = 'Output\Process-PO'
    }

    $configPath = Join-Path -Path $Paths.Config -ChildPath 'autotools-paths.json'
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        try {
            $json = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($json)) {
                $custom = $json | ConvertFrom-Json
                foreach ($name in $custom.PSObject.Properties.Name) {
                    if ($defaults.Contains($name)) {
                        $defaults[$name] = [string]$custom.$name
                    }
                }
            }
        }
        catch {
            Write-Host "[Canh bao] Khong doc duoc Config\autotools-paths.json, dung cau hinh mac dinh. $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    $resolved = [ordered]@{}
    foreach ($key in $defaults.Keys) {
        $resolved[$key] = Resolve-AutoToolsPath -Path $defaults[$key] -RootPath $Paths.Root
    }

    return [PSCustomObject]$resolved
}
