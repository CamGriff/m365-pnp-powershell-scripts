# Script: Export-Multiple-Page.ps1
# Description: allows export of multiple pages listed in an asscociated json file
# Full documentation: https://www.camerongriffiths.com/scripts/exportmultiplepages
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions



# SharePoint Page Export Script - XML Output with CSV Logging
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = "Migrate_MultiplePages.json",
    [Parameter(Mandatory=$false)]
    [string]$ClientId = ""
)

# Get script directory for output files
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Generate log filename with timestamp
$LogFileName = "SharePoint_Export_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$LogPath = Join-Path $ScriptDir $LogFileName

# Initialize log array
$LogEntries = @()

# Function to add log entry
function Add-LogEntry {
    param(
        [string]$PageIdentity,
        [string]$Status,
        [string]$OutputPath = "",
        [string]$ErrorMessage = "",
        [string]$Duration = ""
    )

    $LogEntries += [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PageIdentity = $PageIdentity
        Status = $Status
        OutputPath = $OutputPath
        ErrorMessage = $ErrorMessage
        Duration = $Duration
        User = $env:USERNAME
        Computer = $env:COMPUTERNAME
    }
}

# Function to write log to CSV
function Write-LogToCsv {
    try {
        $LogEntries | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        Write-Host "Log exported to: $LogPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to write log file: $($_.Exception.Message)"
    }
}

# Check if config file exists
if (-not (Test-Path $ConfigFile)) {
    Write-Error "Configuration file '$ConfigFile' not found. Please ensure the JSON file exists."
    Add-LogEntry -PageIdentity "N/A" -Status "ERROR" -ErrorMessage "Configuration file not found: $ConfigFile"
    Write-LogToCsv
    exit 1
}

# Read and parse JSON configuration
try {
    $configContent = Get-Content $ConfigFile -Raw
    $config = $configContent | ConvertFrom-Json
    Write-Host "Loaded configuration for $($config.pages.Count) pages" -ForegroundColor Green
    Add-LogEntry -PageIdentity "CONFIG" -Status "SUCCESS" -ErrorMessage "Loaded $($config.pages.Count) pages from config"
}
catch {
    Write-Error "Failed to parse JSON configuration file: $($_.Exception.Message)"
    Add-LogEntry -PageIdentity "CONFIG" -Status "ERROR" -ErrorMessage "Failed to parse JSON: $($_.Exception.Message)"
    Write-LogToCsv
    exit 1
}

# Extract settings
$SourceSiteUrl = $config.settings.sourceSiteUrl
$BaseOutputPath = $config.settings.baseOutputPath

Write-Host "Source Site: $SourceSiteUrl" -ForegroundColor Cyan
Write-Host "Output Path: $BaseOutputPath" -ForegroundColor Cyan
Write-Host "Log File: $LogPath" -ForegroundColor Cyan

# Connect to SharePoint
try {
    Write-Host "Connecting to SharePoint site: $SourceSiteUrl" -ForegroundColor Yellow
    $connectionStart = Get-Date
    Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId $ClientId
    $connectionEnd = Get-Date
    $connectionDuration = "{0:F2}" -f ($connectionEnd - $connectionStart).TotalSeconds
    Write-Host "Connected successfully! ($connectionDuration seconds)" -ForegroundColor Green
    Add-LogEntry -PageIdentity "CONNECTION" -Status "SUCCESS" -ErrorMessage "Connected to $SourceSiteUrl" -Duration "$connectionDuration seconds"
}
catch {
    Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
    Add-LogEntry -PageIdentity "CONNECTION" -Status "ERROR" -ErrorMessage "Failed to connect: $($_.Exception.Message)"
    Write-LogToCsv
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $BaseOutputPath)) {
    try {
        New-Item -ItemType Directory -Path $BaseOutputPath -Force | Out-Null
        Write-Host "Created output directory: $BaseOutputPath" -ForegroundColor Gray
    }
    catch {
        Write-Error "Failed to create output directory: $($_.Exception.Message)"
        Add-LogEntry -PageIdentity "DIRECTORY" -Status "ERROR" -ErrorMessage "Failed to create directory: $BaseOutputPath"
        Write-LogToCsv
        exit 1
    }
}

# Process each page
$successCount = 0
$errorCount = 0

Write-Host "`nStarting page exports..." -ForegroundColor Yellow

foreach ($page in $config.pages) {
    $pageStart = Get-Date

    # Extract filename from identity for the XML output
    $pageName = $page.identity
    if ($pageName.Contains("/")) {
        $fileName = Split-Path $pageName -Leaf
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    } else {
        $fileName = $pageName -replace "\.aspx$", ""
    }

    $outputFileName = "$fileName.xml"
    $outputPath = Join-Path $BaseOutputPath $outputFileName

    try {
        Write-Host "Exporting: $($page.identity) -> $outputFileName" -ForegroundColor Cyan

        # Export the page using PnP
        Export-PnPPage -Identity $page.identity -Out $outputPath -Force

        $pageEnd = Get-Date
        $pageDuration = "{0:F2}" -f ($pageEnd - $pageStart).TotalSeconds

        Write-Host "✓ Success: $($page.identity) ($pageDuration seconds)" -ForegroundColor Green
        Add-LogEntry -PageIdentity $page.identity -Status "SUCCESS" -OutputPath $outputPath -Duration "$pageDuration seconds"
        $successCount++
    }
    catch {
        $pageEnd = Get-Date
        $pageDuration = "{0:F2}" -f ($pageEnd - $pageStart).TotalSeconds

        Write-Host "✗ Failed: $($page.identity) - $($_.Exception.Message)" -ForegroundColor Red
        Add-LogEntry -PageIdentity $page.identity -Status "ERROR" -OutputPath $outputPath -ErrorMessage $_.Exception.Message -Duration "$pageDuration seconds"
        $errorCount++
    }
}

# Summary
Write-Host "`n=== Export Summary ===" -ForegroundColor Yellow
Write-Host "Successfully exported: $successCount pages" -ForegroundColor Green
Write-Host "Failed exports: $errorCount pages" -ForegroundColor Red
Write-Host "Total processed: $($successCount + $errorCount) pages" -ForegroundColor White

# Add summary to log
Add-LogEntry -PageIdentity "SUMMARY" -Status "COMPLETE" -ErrorMessage "Success: $successCount, Failed: $errorCount, Total: $($successCount + $errorCount)"

# Write log to CSV
Write-LogToCsv

# Disconnect from SharePoint
try {
    Disconnect-PnPOnline
    Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
}
catch {
    Write-Warning "Didn't disconnect from SharePoint"
}
