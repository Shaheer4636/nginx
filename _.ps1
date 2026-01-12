# Azure CLI + VC++ Runtime Downloader/Installer (best-effort without admin)
# Saves into %USERPROFILE%\Downloads\az-fix
# Tries non-admin install first; if blocked, tells you clearly.

$ErrorActionPreference = "Stop"

$base = Join-Path $env:USERPROFILE "Downloads\az-fix"
New-Item -ItemType Directory -Force -Path $base | Out-Null

function Download-File {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$OutFile
    )
    Write-Host "Downloading: $Url"
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
    Write-Host "Saved: $OutFile"
}

function Try-Install {
    param(
        [Parameter(Mandatory=$true)][string]$InstallerPath,
        [Parameter(Mandatory=$true)][string]$Name
    )

    Write-Host "`n--- Trying NON-ADMIN install: $Name ---"
    # For VC++ redistributables, /install /quiet /norestart is typical
    # Per-user install is not always supported; this may still require elevation.
    $args = "/install /quiet /norestart"

    try {
        $p = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru
        Write-Host "$Name exit code: $($p.ExitCode)"
        if ($p.ExitCode -eq 0) { return $true }
        else { return $false }
    } catch {
        Write-Host "$Name non-admin install failed: $($_.Exception.Message)"
        return $false
    }
}

function Try-Install-Admin {
    param(
        [Parameter(Mandatory=$true)][string]$InstallerPath,
        [Parameter(Mandatory=$true)][string]$Name
    )

    Write-Host "`n--- Trying ADMIN install (will prompt UAC): $Name ---"
    $args = "/install /quiet /norestart"
    try {
        $p = Start-Process -FilePath $InstallerPath -ArgumentList $args -Verb RunAs -Wait -PassThru
        Write-Host "$Name exit code: $($p.ExitCode)"
        return ($p.ExitCode -eq 0)
    } catch {
        Write-Host "$Name admin install failed or was blocked: $($_.Exception.Message)"
        return $false
    }
}

function Test-AzCliFromFolder {
    param([Parameter(Mandatory=$true)][string]$AzBinDir)

    $azCmd = Join-Path $AzBinDir "az.cmd"
    if (!(Test-Path $azCmd)) {
        Write-Host "az.cmd not found in $AzBinDir"
        return $false
    }

    Write-Host "`n--- Testing Azure CLI ---"
    try {
        $out = & $azCmd --version 2>&1
        Write-Host $out
        return $true
    } catch {
        Write-Host "Azure CLI test failed: $($_.Exception.Message)"
        return $false
    }
}

# Official Microsoft VC++ 2015-2022 Redistributable direct links
# (These are Microsoft's aka.ms redirectors)
$vcX64Url = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$vcX86Url = "https://aka.ms/vs/17/release/vc_redist.x86.exe"

$vcX64 = Join-Path $base "vc_redist.x64.exe"
$vcX86 = Join-Path $base "vc_redist.x86.exe"

Download-File -Url $vcX64Url -OutFile $vcX64
Download-File -Url $vcX86Url -OutFile $vcX86

$ok64 = Try-Install -InstallerPath $vcX64 -Name "VC++ 2015-2022 (x64)"
$ok86 = Try-Install -InstallerPath $vcX86 -Name "VC++ 2015-2022 (x86)"

if (-not $ok64 -or -not $ok86) {
    Write-Host "`nNon-admin install didn't fully succeed (expected on managed PCs)."
    Write-Host "If you have rights, we'll try elevation now."
    $ok64a = $true
    $ok86a = $true
    if (-not $ok64) { $ok64a = Try-Install-Admin -InstallerPath $vcX64 -Name "VC++ 2015-2022 (x64)" }
    if (-not $ok86) { $ok86a = Try-Install-Admin -InstallerPath $vcX86 -Name "VC++ 2015-2022 (x86)" }

    if (-not $ok64a -or -not $ok86a) {
        Write-Host "`n❌ VC++ runtime install is blocked / requires IT."
        Write-Host "Tell IT: install VC++ 2015-2022 Redistributable (x64 and x86). Error you hit: 0xc000007b."
    } else {
        Write-Host "`n✅ VC++ runtime installed with admin."
    }
} else {
    Write-Host "`n✅ VC++ runtime installed (non-admin)."
}

# Backup: Download Azure CLI ZIP (portable) and test
# Note: ZIP method still may fail if bundled python depends on missing runtimes, but often helps with MSI policy blocks.
$azZipUrl = "https://aka.ms/InstallAzureCliWindowsZip"
$azZip = Join-Path $base "azure-cli.zip"
Download-File -Url $azZipUrl -OutFile $azZip

$azExtract = Join-Path $base "azure-cli"
if (Test-Path $azExtract) { Remove-Item -Recurse -Force $azExtract }
Expand-Archive -Path $azZip -DestinationPath $azExtract -Force

# The ZIP typically contains a folder with 'bin'
$bin = Join-Path $azExtract "bin"
if (Test-Path $bin) {
    Test-AzCliFromFolder -AzBinDir $bin | Out-Null
} else {
    # Try to locate bin deeper just in case
    $found = Get-ChildItem -Path $azExtract -Recurse -Directory -Filter "bin" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { Test-AzCliFromFolder -AzBinDir $found.FullName | Out-Null }
    else { Write-Host "Could not find 'bin' folder in extracted Azure CLI ZIP." }
}

Write-Host "`nDone. Folder: $base"
