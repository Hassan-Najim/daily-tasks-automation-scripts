# =============================================================================
# Daily Tasks - Automation Suite launcher
#
# Run (zero install):   irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex
# Install command:      irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex; then choose Install
#                       or:  powershell -File launcher.ps1 -Install
# Uninstall command:    powershell -File launcher.ps1 -Uninstall
#
# Downloads the latest daily-tasks.exe from GitHub Releases (with SHA256
# verification), caches it in %LOCALAPPDATA%\daily-tasks, and runs it.
# Falls back to the cached copy when offline.
# =============================================================================
#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$Install,
    [switch]$Uninstall
)

$RepoOwner   = "Hassan-Najim"
$RepoName    = "daily-tasks-automation-scripts"
$ExeName     = "daily-tasks.exe"
$CommandName = "daily-tasks"
$InstallDir  = Join-Path $env:LOCALAPPDATA $CommandName
$ApiBase     = "https://api.github.com/repos/$RepoOwner/$RepoName"
$RawLauncher = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/main/launcher.ps1"

$ProgressPreference = "SilentlyContinue"
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

function Get-LatestRelease {
    try {
        return Invoke-RestMethod -Uri "$ApiBase/releases/latest" -TimeoutSec 15 -ErrorAction Stop
    } catch {
        return $null
    }
}

function Get-CachedVersion {
    $versionFile = Join-Path $InstallDir "version.txt"
    if (Test-Path $versionFile) {
        return (Get-Content $versionFile -Raw).Trim()
    }
    return $null
}

function Save-Release {
    param($Release)

    $exeAsset = $Release.assets | Where-Object { $_.name -eq $ExeName } | Select-Object -First 1
    if (-not $exeAsset) {
        throw "Asset '$ExeName' not found in release $($Release.tag_name)"
    }
    $sumAsset = $Release.assets | Where-Object { $_.name -eq "SHA256SUMS.txt" } | Select-Object -First 1

    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

    $tmpExe = Join-Path $env:TEMP "$CommandName-download.exe"
    Invoke-WebRequest -Uri $exeAsset.browser_download_url -OutFile $tmpExe -UseBasicParsing

    if ($sumAsset) {
        $tmpSum = Join-Path $env:TEMP "$CommandName-SHA256SUMS.txt"
        Invoke-WebRequest -Uri $sumAsset.browser_download_url -OutFile $tmpSum -UseBasicParsing
        $expected = $null
        foreach ($line in (Get-Content $tmpSum)) {
            if ($line -match $ExeName) {
                $expected = ($line -replace "^([0-9a-fA-F]+).*", '$1').ToLower()
                break
            }
        }
        if ($expected) {
            $actual = (Get-FileHash -Algorithm SHA256 -Path $tmpExe).Hash.ToLower()
            if ($actual -ne $expected) {
                Remove-Item $tmpExe -Force -ErrorAction SilentlyContinue
                throw "Checksum mismatch for $ExeName (expected $expected, got $actual)"
            }
        }
    }

    $destExe = Join-Path $InstallDir $ExeName
    if (Test-Path $destExe) {
        try {
            Remove-Item $destExe -Force -ErrorAction Stop
        } catch {
            Remove-Item $tmpExe -Force -ErrorAction SilentlyContinue
            throw "Cannot replace $destExe - close $CommandName and try again."
        }
    }
    Move-Item -Path $tmpExe -Destination $destExe
    Set-Content -Path (Join-Path $InstallDir "version.txt") -Value $Release.tag_name
}

function Install-Command {
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

    $localLauncher = Join-Path $InstallDir "launcher.ps1"
    if ($PSCommandPath -and (Test-Path $PSCommandPath)) {
        Copy-Item -Path $PSCommandPath -Destination $localLauncher -Force
    } else {
        Invoke-WebRequest -Uri $RawLauncher -OutFile $localLauncher -UseBasicParsing
    }

    $shimContent = "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"$localLauncher`" %*"
    Set-Content -Path (Join-Path $InstallDir "$CommandName.cmd") -Value $shimContent -Encoding ASCII

    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    if (-not $userPath) { $userPath = "" }
    $alreadyPresent = ($userPath -split ";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -eq $InstallDir }
    if (-not $alreadyPresent) {
        $newPath = if ($userPath.Trim()) { "$($userPath.TrimEnd(';'));$InstallDir" } else { $InstallDir }
        [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    }

    Write-Host ""
    Write-Host "Installed the '$CommandName' command." -ForegroundColor Green
    Write-Host "Open a NEW terminal and type: $CommandName"
}

function Uninstall-Command {
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    if ($userPath) {
        $kept = ($userPath -split ";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne $InstallDir }
        [Environment]::SetEnvironmentVariable("Path", ($kept -join ";"), "User")
    }
    if (Test-Path $InstallDir) {
        Remove-Item -Recurse -Force $InstallDir -ErrorAction SilentlyContinue
    }
    Write-Host "Removed the '$CommandName' command and cached files." -ForegroundColor Green
}

# =============================================================================
# Main flow
# =============================================================================

if ($Uninstall) {
    Uninstall-Command
    exit 0
}

$exePath = Join-Path $InstallDir $ExeName
$release = Get-LatestRelease
$cachedVersion = Get-CachedVersion

if ($release) {
    if ($cachedVersion -ne $release.tag_name) {
        Write-Host "Updating to $($release.tag_name)..." -ForegroundColor Cyan
        try {
            Save-Release $release
            $cachedVersion = $release.tag_name
        } catch {
            Write-Warning "Update failed: $($_.Exception.Message)"
            if (-not (Test-Path $exePath)) {
                Write-Error "No cached executable available. Aborting."
                exit 1
            }
            Write-Warning "Falling back to cached version $cachedVersion."
        }
    }
} else {
    if (Test-Path $exePath) {
        Write-Warning "Could not check for updates (offline or rate-limited). Using cached version $cachedVersion."
    }
}

if (-not (Test-Path $exePath)) {
    Write-Error "No release found and no cached executable. Check your internet connection or visit:"
    Write-Host "  https://github.com/$RepoOwner/$RepoName/releases"
    exit 1
}

if ($Install) {
    Install-Command
    exit 0
}

$shimPath = Join-Path $InstallDir "$CommandName.cmd"
if (-not (Test-Path $shimPath) -and $Host.Name -eq "ConsoleHost" -and -not [Console]::IsInputRedirected) {
    $answer = Read-Host "Install the '$CommandName' command so you can run it from any terminal? (y/N)"
    if ($answer -match "^[Yy]") {
        Install-Command
    }
}

Write-Host "Starting $CommandName (version $cachedVersion)..." -ForegroundColor Cyan
& $exePath
exit $LASTEXITCODE
