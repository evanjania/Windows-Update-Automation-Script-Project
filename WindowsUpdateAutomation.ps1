# Windows Update Automation Script
# Author: Evan Jania
# Date: Dec 10 2025
# Description: Automates Windows updates with logging and error handling

param(
    [switch]$AutoReboot = $false,
    [string]$LogPath = "C:\Logs\WindowsUpdate"
)

$ScriptVersion = "1.0"
$LogFile = Join-Path $LogPath "Update_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Ensure log directory exists
if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'INFO'    { Write-Host $logEntry -ForegroundColor Cyan }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
    }
    
    Add-Content -Path $LogFile -Value $logEntry
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WindowsUpdateStatus {
    try {
        Write-Log "Checking for available Windows updates..." -Level INFO
        
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        
        Write-Log "Searching for updates (this may take a few minutes)..." -Level INFO
        $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
        
        $updateCount = $searchResult.Updates.Count
        
        if ($updateCount -eq 0) {
            Write-Log "No updates available. System is up to date." -Level SUCCESS
            return $null
        }
        
        Write-Log "Found $updateCount available update(s)" -Level INFO
        
        foreach ($update in $searchResult.Updates) {
            Write-Log "  - $($update.Title)" -Level INFO
        }
        return $searchResult.Updates
    }
    catch {
        Write-Log "Error checking for updates: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Start-UpdateDownload {
    param($Updates)
    
    try {
        Write-Log "Downloading updates..." -Level INFO
        
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $Updates
        
        $downloadResult = $downloader.Download()
        
        if ($downloadResult.ResultCode -eq 2) {
            Write-Log "All updates downloaded successfully" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Some updates failed to download. Result code: $($downloadResult.ResultCode)" -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "Error downloading updates: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Install-WindowsUpdates {
    param($Updates)
    
    try {
        Write-Log "Installing updates..." -Level INFO
        
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $installer = $updateSession.CreateUpdateInstaller()
        $installer.Updates = $Updates
        
        Write-Log "Installation in progress. This may take several minutes..." -Level INFO
        $installResult = $installer.Install()
        
        for ($i = 0; $i -lt $Updates.Count; $i++) {
            $update = $Updates.Item($i)
            $result = $installResult.GetUpdateResult($i)
            
            switch ($result.ResultCode) {
                2 { Write-Log "Installed: $($update.Title)" -Level SUCCESS }
                3 { Write-Log "Installed with errors: $($update.Title)" -Level WARNING }
                4 { Write-Log "Failed: $($update.Title)" -Level ERROR }
                5 { Write-Log "Aborted: $($update.Title)" -Level WARNING }
            }
        }
        
        if ($installResult.RebootRequired) {
            Write-Log "System reboot is required to complete installation" -Level WARNING
            return "REBOOT_REQUIRED"
        }
        
        Write-Log "Updates installed successfully" -Level SUCCESS
        return "SUCCESS"
    }
    catch {
        Write-Log "Error installing updates: $($_.Exception.Message)" -Level ERROR
        return "ERROR"
    }
}

# Main execution
Write-Log "========================================" -Level INFO
Write-Log "Windows Update Automation Script v$ScriptVersion" -Level INFO
Write-Log "Author: Evan Jania" -Level INFO
Write-Log "========================================" -Level INFO

if (!(Test-Administrator)) {
    Write-Log "This script must be run as Administrator!" -Level ERROR
    Write-Log "Please right-click PowerShell and select 'Run as Administrator'" -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Log "Running with Administrator privileges" -Level SUCCESS

$updates = Get-WindowsUpdateStatus

if ($null -eq $updates) {
    Write-Log "No updates to install" -Level INFO
    Write-Log "Log file saved to: $LogFile" -Level INFO
    Read-Host "Press Enter to exit"
    exit 0
}

$response = Read-Host "Do you want to proceed with downloading and installing these updates? (Y/N)"

if ($response -ne 'Y' -and $response -ne 'y') {
    Write-Log "Update process cancelled by user" -Level WARNING
    Read-Host "Press Enter to exit"
    exit 0
}

$downloadSuccess = Start-UpdateDownload -Updates $updates

if (!$downloadSuccess) {
    Write-Log "Download failed. Aborting installation." -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}

$installStatus = Install-WindowsUpdates -Updates $updates

Write-Log "========================================" -Level INFO
Write-Log "UPDATE COMPLETE" -Level INFO
Write-Log "Total Updates: $($updates.Count)" -Level INFO
Write-Log "Status: $installStatus" -Level INFO
Write-Log "Log File: $LogFile" -Level INFO
Write-Log "========================================" -Level INFO

if ($installStatus -eq "REBOOT_REQUIRED") {
    if ($AutoReboot) {
        Write-Log "Auto-reboot enabled. System will restart in 60 seconds..." -Level WARNING
        shutdown /r /t 60 /c "Windows updates installed. Restarting to complete."
    }
    else {
        $rebootChoice = Read-Host "Would you like to reboot now? (Y/N)"
        if ($rebootChoice -eq 'Y' -or $rebootChoice -eq 'y') {
            Write-Log "Initiating system restart..." -Level WARNING
            shutdown /r /t 10 /c "Restarting to complete Windows updates"
        }
        else {
            Write-Log "Please restart your computer soon to complete updates." -Level WARNING
        }
    }
}

Write-Log "Script execution completed" -Level SUCCESS
Read-Host "Press Enter to exit"
