# =============================================================================
# Daily Tasks - Automation Suite launcher
#
# Run (zero install):   irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex
# Install command:      irm https://raw.githubusercontent.com/Hassan-Najim/daily-tasks-automation-scripts/main/launcher.ps1 | iex; then choose Install
#                       or:  powershell -File launcher.ps1 -Install
# Uninstall command:    powershell -File launcher.ps1 -Uninstall
# Force update:         daily-tasks -Update   (re-downloads even if up to date)
#
# Downloads the latest daily-tasks.exe from GitHub Releases (with SHA256
# verification), caches it in %LOCALAPPDATA%\daily-tasks, and runs it.
# The update check uses the releases/latest page redirect (not the GitHub
# API), so it is not affected by the API's 60-requests-per-hour limit.
# Falls back to the cached copy when offline (with a clear reason).
# The installed launcher copy refreshes itself from the repository.
# =============================================================================
#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Update
)

$RepoOwner   = "Hassan-Najim"
$RepoName    = "daily-tasks-automation-scripts"
$ExeName     = "daily-tasks.exe"
$CommandName = "daily-tasks"
$InstallDir  = Join-Path $env:LOCALAPPDATA $CommandName
$ApiBase     = "https://api.github.com/repos/$RepoOwner/$RepoName"
$RawLauncher = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/main/launcher.ps1"
$LatestPage  = "https://github.com/$RepoOwner/$RepoName/releases/latest"

$script:CheckError = ""

$ProgressPreference = "SilentlyContinue"
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

function Get-CheckFailureReason {
    param($ErrorRecord)

    $exception = $ErrorRecord.Exception
    if ($exception -is [System.Net.WebException]) {
        $response = $exception.Response
        if ($response) {
            $status = [int]$response.StatusCode
            if ($status -eq 403) { return "rate-limited by GitHub (HTTP 403)" }
            return "HTTP $status from github.com"
        }
        if ($exception.Status -eq [System.Net.WebExceptionStatus]::Timeout) {
            return "request timed out"
        }
        return "network unreachable (offline or DNS failure)"
    }
    if ($exception) { return $exception.Message }
    return "unknown error"
}

function Get-LatestTag {
    # Reads the redirect of the releases/latest page. This is the website,
    # not the API, so it is not rate limited like api.github.com.
    try {
        $response = Invoke-WebRequest -Uri $LatestPage -Method Head `
            -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        $finalUrl = $response.BaseResponse.ResponseUri.AbsoluteUri
        if ($finalUrl -match "/releases/tag/(.+?)/?$") {
            return $Matches[1]
        }
        $script:CheckError = "unexpected response from github.com"
    } catch {
        $script:CheckError = Get-CheckFailureReason $_
        $webResponse = $null
        try { $webResponse = $_.Exception.Response } catch { }
        if ($webResponse) {
            $location = $null
            try { $location = $webResponse.Headers["Location"] } catch { }
            if ($location -and $location -match "/releases/tag/(.+?)/?$") {
                return $Matches[1]
            }
        }
    }
    return $null
}

function Get-LatestRelease {
    # Fallback: the GitHub API (rate limited to 60 requests/hour per IP).
    try {
        return Invoke-RestMethod -Uri "$ApiBase/releases/latest" -TimeoutSec 15 -ErrorAction Stop
    } catch {
        if (-not $script:CheckError) {
            $script:CheckError = Get-CheckFailureReason $_
        }
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
    param([string]$Tag)

    $downloadBase = "https://github.com/$RepoOwner/$RepoName/releases/download/$Tag"

    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

    $tmpExe = Join-Path $env:TEMP "$CommandName-download.exe"
    Invoke-WebRequest -Uri "$downloadBase/$ExeName" -OutFile $tmpExe -UseBasicParsing

    $expected = $null
    $tmpSum = Join-Path $env:TEMP "$CommandName-SHA256SUMS.txt"
    Remove-Item $tmpSum -Force -ErrorAction SilentlyContinue
    try {
        Invoke-WebRequest -Uri "$downloadBase/SHA256SUMS.txt" -OutFile $tmpSum -UseBasicParsing
    } catch {
        Write-Warning "Could not download SHA256SUMS.txt; skipping checksum verification."
    }
    if (Test-Path $tmpSum) {
        foreach ($line in (Get-Content $tmpSum)) {
            if ($line -match $ExeName) {
                $expected = ($line -replace "^([0-9a-fA-F]+).*", '$1').ToLower()
                break
            }
        }
    }
    if ($expected) {
        $actual = (Get-FileHash -Algorithm SHA256 -Path $tmpExe).Hash.ToLower()
        if ($actual -ne $expected) {
            Remove-Item $tmpExe -Force -ErrorAction SilentlyContinue
            throw "Checksum mismatch for $ExeName (expected $expected, got $actual)"
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
    Set-Content -Path (Join-Path $InstallDir "version.txt") -Value $Tag
}

function Update-LocalLauncher {
    # Refresh the installed launcher copy from the repository so launcher
    # fixes reach already-installed users without reinstalling.
    $localLauncher = Join-Path $InstallDir "launcher.ps1"
    if (-not (Test-Path $localLauncher)) { return }
    try {
        $remote = Invoke-WebRequest -Uri $RawLauncher -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        if (-not $remote.Content) { return }
        $localText = Get-Content $localLauncher -Raw
        if (-not $localText) { $localText = "" }
        if ($remote.Content.Trim() -ne $localText.Trim()) {
            Set-Content -Path $localLauncher -Value $remote.Content -Encoding UTF8
            Write-Host "Launcher script updated." -ForegroundColor DarkGray
        }
    } catch { }
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

# Keep the installed launcher copy current.
Update-LocalLauncher

$exePath = Join-Path $InstallDir $ExeName
$cachedVersion = Get-CachedVersion

$latestTag = Get-LatestTag
if (-not $latestTag) {
    $release = Get-LatestRelease
    if ($release) {
        $latestTag = $release.tag_name
    }
}

if ($latestTag) {
    if (-not $cachedVersion) { $cachedVersion = "none" }
    Write-Host ("Latest: {0}   Installed: {1}" -f $latestTag, $cachedVersion)

    if ($Update -or $cachedVersion -ne $latestTag) {
        if ($Update) {
            Write-Host "Forcing re-download of $($latestTag)..." -ForegroundColor Cyan
        } else {
            Write-Host "Updating to $($latestTag)..." -ForegroundColor Cyan
        }
        try {
            Save-Release -Tag $latestTag
            $cachedVersion = $latestTag
            if ($Update) {
                Write-Host "Update complete: $cachedVersion" -ForegroundColor Green
                exit 0
            }
        } catch {
            Write-Warning "Update failed: $($_.Exception.Message)"
            if (-not (Test-Path $exePath)) {
                Write-Error "No cached executable available. Aborting."
                exit 1
            }
            Write-Warning "Falling back to cached version $cachedVersion."
            if ($Update) { exit 1 }
        }
    } else {
        Write-Host "Already up to date."
    }
} else {
    if (Test-Path $exePath) {
        Write-Warning ("Could not check for updates ({0}). Using cached version {1}." -f $script:CheckError, $cachedVersion)
    }
    if ($Update) {
        Write-Error "-Update requires an internet connection."
        exit 1
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
