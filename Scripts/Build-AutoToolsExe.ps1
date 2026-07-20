#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath = 'dist\AutoTools.exe',
    [string]$InstallDirName = 'AutoTools'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$buildRoot = Join-Path $root 'Temp\Build-AutoToolsExe'
$payloadRoot = Join-Path $buildRoot 'payload'
$payloadZip = Join-Path $buildRoot 'payload.zip'
$stubSourcePath = Join-Path $buildRoot 'AutoToolsStub.cs'
$outputFullPath = [System.IO.Path]::GetFullPath((Join-Path $root $OutputPath))
$outputDirectory = Split-Path -Path $outputFullPath -Parent
$zipFullPath = [System.IO.Path]::ChangeExtension($outputFullPath, '.zip')

if (Test-Path -LiteralPath $buildRoot) {
    Remove-Item -LiteralPath $buildRoot -Recurse -Force
}

$null = New-Item -Path $payloadRoot -ItemType Directory -Force
$null = New-Item -Path $outputDirectory -ItemType Directory -Force

function Copy-RelativeFile {
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath,

        [switch]$Optional
    )

    $source = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        if ($Optional) {
            return
        }

        throw "Khong tim thay file can dong goi: $RelativePath"
    }

    $destination = Join-Path $payloadRoot $RelativePath
    $destinationDirectory = Split-Path -Path $destination -Parent
    $null = New-Item -Path $destinationDirectory -ItemType Directory -Force
    Copy-Item -LiteralPath $source -Destination $destination -Force
}

function Copy-RelativeDirectoryFiles {
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath,

        [scriptblock]$Exclude = { $false }
    )

    $sourceRoot = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
        throw "Khong tim thay folder can dong goi: $RelativePath"
    }

    Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
        Where-Object { -not (& $Exclude $_) } |
        ForEach-Object {
            $relative = $_.FullName.Substring($root.Length).TrimStart('\', '/')
            Copy-RelativeFile $relative
        }
}

function New-PayloadDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    $null = New-Item -Path (Join-Path $payloadRoot $RelativePath) -ItemType Directory -Force
}

function Add-InputSeed {
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    $sourceRoot = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
        return
    }

    Copy-RelativeDirectoryFiles $RelativePath
}

@(
    'Launcher.ps1',
    'launcher-config.json',
    'AutoTools.txt',
    'HUONG_DAN_SETUP_AUTO_TOOLS.txt',
    'Config\autotools-paths.json',
    'Config\suppliers.csv'
) | ForEach-Object { Copy-RelativeFile $_ }

Copy-RelativeDirectoryFiles 'Scripts' {
    param($File)
    $File.FullName -match '\\Output\\' -or
    $File.FullName -match '\\Older CDs\\' -or
    $File.FullName -match '\\Older Output PO\\'
}

@(
    'Input\MR-Outlook',
    'Input\Process-CDs',
    'Input\Process-PO'
) | ForEach-Object { Add-InputSeed $_ }

foreach ($directory in @(
    'Input',
    'Input\ASCP',
    'Input\Data Job Short - Not Delete',
    'Input\Job Short',
    'Input\Supplier Commitment',
    'Output',
    'Output\Process-CDs',
    'Output\Process-PO',
    'Output\Supplier Commitment',
    'Logs',
    'Temp',
    'MR_Out'
)) {
    New-PayloadDirectory $directory
}

Compress-Archive -Path (Join-Path $payloadRoot '*') -DestinationPath $payloadZip -Force

$bootstrap = @'
$ErrorActionPreference = 'Stop'

$installRoot = Join-Path $env:LOCALAPPDATA '__INSTALL_DIR_NAME__'
$payloadZip = Join-Path $PSScriptRoot 'payload.zip'
$stageRoot = Join-Path $env:TEMP ('AutoToolsPayload_' + [guid]::NewGuid().ToString('N'))

New-Item -Path $installRoot -ItemType Directory -Force | Out-Null
New-Item -Path $stageRoot -ItemType Directory -Force | Out-Null

try {
    Expand-Archive -LiteralPath $payloadZip -DestinationPath $stageRoot -Force

    foreach ($name in @('Launcher.ps1', 'launcher-config.json', 'AutoTools.txt', 'HUONG_DAN_SETUP_AUTO_TOOLS.txt', 'Scripts', 'Config')) {
        $source = Join-Path $stageRoot $name
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination $installRoot -Recurse -Force
        }
    }

    foreach ($name in @(
        'Input',
        'Input\ASCP',
        'Input\Data Job Short - Not Delete',
        'Input\Job Short',
        'Input\Supplier Commitment',
        'Output',
        'Output\Process-CDs',
        'Output\Process-PO',
        'Output\Supplier Commitment',
        'Logs',
        'Temp',
        'MR_Out'
    )) {
        New-Item -Path (Join-Path $installRoot $name) -ItemType Directory -Force | Out-Null
    }

    $inputStageRoot = Join-Path $stageRoot 'Input'
    if (Test-Path -LiteralPath $inputStageRoot -PathType Container) {
        Get-ChildItem -LiteralPath $inputStageRoot -Recurse -File | ForEach-Object {
            $relative = $_.FullName.Substring($inputStageRoot.Length).TrimStart('\', '/')
            $destination = Join-Path (Join-Path $installRoot 'Input') $relative
            if (-not (Test-Path -LiteralPath $destination)) {
                New-Item -Path (Split-Path -Path $destination -Parent) -ItemType Directory -Force | Out-Null
                Copy-Item -LiteralPath $_.FullName -Destination $destination -Force
            }
        }
    }

    $env:AUTOTOOLS_ROOT = $installRoot
    Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (Join-Path $installRoot 'Launcher.ps1')
    )
}
finally {
    Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue
}
'@ -replace '__INSTALL_DIR_NAME__', ($InstallDirName -replace "'", "''")

$bootstrapBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bootstrap))

$stubSource = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal static class AutoToolsStub
{
    [STAThread]
    private static int Main()
    {
        string stageRoot = Path.Combine(Path.GetTempPath(), "AutoToolsExe_" + Guid.NewGuid().ToString("N"));

        try
        {
            Directory.CreateDirectory(stageRoot);

            string payloadZip = Path.Combine(stageRoot, "payload.zip");
            using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream("AutoTools.Payload.zip"))
            {
                if (input == null)
                {
                    throw new InvalidOperationException("Missing embedded payload.");
                }

                using (FileStream output = File.Create(payloadZip))
                {
                    input.CopyTo(output);
                }
            }

            string bootstrap = Encoding.UTF8.GetString(Convert.FromBase64String("$bootstrapBase64"));
            string bootstrapPath = Path.Combine(stageRoot, "bootstrap.ps1");
            File.WriteAllText(bootstrapPath, bootstrap, new UTF8Encoding(false));

            string powershell = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe"
            );

            ProcessStartInfo startInfo = new ProcessStartInfo(powershell);
            startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + bootstrapPath + "\"";
            startInfo.UseShellExecute = false;

            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();
                return process.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Auto Tools", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
        finally
        {
            try
            {
                if (Directory.Exists(stageRoot))
                {
                    Directory.Delete(stageRoot, true);
                }
            }
            catch
            {
            }
        }
    }
}
"@

Set-Content -LiteralPath $stubSourcePath -Value $stubSource -Encoding ASCII

$csc = @(
    (Join-Path $env:SystemRoot 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:SystemRoot 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1

if ([string]::IsNullOrWhiteSpace($csc)) {
    throw 'Khong tim thay csc.exe cua .NET Framework 4 tren may build nay.'
}

& $csc @(
    '/nologo',
    '/target:winexe',
    '/platform:anycpu',
    "/out:$outputFullPath",
    "/resource:$payloadZip,AutoTools.Payload.zip",
    '/reference:System.Windows.Forms.dll',
    $stubSourcePath
)

if ($LASTEXITCODE -ne 0) {
    throw "csc.exe failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
    throw "Khong tao duoc file exe: $outputFullPath"
}

Compress-Archive -LiteralPath $outputFullPath -DestinationPath $zipFullPath -Force

Write-Host "Done: $outputFullPath"
Write-Host "Zip : $zipFullPath"
