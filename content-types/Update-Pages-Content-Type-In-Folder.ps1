# Script: Update-Pages-Content-Type-In-Folder.ps1
# Description: Creates SharePoint Online content types from JSON configuration file
# Full documentation: https://www.camerongriffiths.com/scripts/updatepagescontenttypeinafolder
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# PnP PowerShell script to update content type for all pages in a specific folder within Site Pages library
# Clean version using only the server relative URL method

# Variables
$ClientId = ""
$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
$ContentTypeName = "Supporting Documentation"
$LibraryName = "Site Pages"
$FolderName = "Supporting Documents"

# CSV output settings
$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = (Get-Location).Path
}
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvOutputPath = Join-Path $ScriptPath "ContentTypeUpdate_Results_$TimeStamp.csv"

# Initialize results array for CSV export
$Results = @()

# Connect to SharePoint site
Write-Host "Connecting to SharePoint site..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

try {
    # Get the Site Pages library
    Write-Host "Getting Site Pages library..." -ForegroundColor Yellow
    $Library = Get-PnPList -Identity $LibraryName

    if ($null -eq $Library) {
        Write-Host "Error: Could not find the '$LibraryName' library" -ForegroundColor Red
        return
    }
    Write-Host "Site Pages library found. Server relative URL: $($Library.RootFolder.ServerRelativeUrl)" -ForegroundColor Green

    # Get the folder using server relative URL
    Write-Host "Getting folder '$FolderName'..." -ForegroundColor Yellow
    try {
        $FolderServerRelativeUrl = $Library.RootFolder.ServerRelativeUrl + "/" + $FolderName
        $Folder = Get-PnPFolder -Url $FolderServerRelativeUrl -ErrorAction Stop
        Write-Host "✓ Found folder: $($Folder.ServerRelativeUrl)" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Could not find folder '$FolderName': $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Expected folder path: $FolderServerRelativeUrl" -ForegroundColor Yellow
        return
    }

    # Get the target content type
    Write-Host "Getting content type '$ContentTypeName'..." -ForegroundColor Yellow
    $ContentType = Get-PnPContentType -List $Library | Where-Object { $_.Name -eq $ContentTypeName }

    if ($null -eq $ContentType) {
        Write-Host "Error: Content type '$ContentTypeName' not found in the library" -ForegroundColor Red
        Write-Host "Available content types in this library:" -ForegroundColor Yellow
        Get-PnPContentType -List $Library | ForEach-Object { Write-Host "- $($_.Name)" }
        return
    }

    # Get all pages from the specific folder
    Write-Host "Getting all pages from the folder..." -ForegroundColor Yellow
    $Pages = Get-PnPListItem -List $Library -FolderServerRelativeUrl $Folder.ServerRelativeUrl -PageSize 1000

    if ($Pages.Count -eq 0) {
        Write-Host "No pages found in the '$FolderName' folder" -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($Pages.Count) pages to update in the '$FolderName' folder" -ForegroundColor Green

    # Counter for progress tracking
    $Counter = 0
    $SuccessCount = 0
    $ErrorCount = 0

    # Update content type for each page
    foreach ($Page in $Pages) {
        $Counter++
        $PageTitle = $Page.FieldValues.Title
        if ([string]::IsNullOrEmpty($PageTitle)) {
            $PageTitle = $Page.FieldValues.FileLeafRef
        }

        Write-Progress -Activity "Updating Content Type" -Status "Processing page $Counter of $($Pages.Count) in '$FolderName': $PageTitle" -PercentComplete (($Counter / $Pages.Count) * 100)

        try {
            # Get current content type before update
            $CurrentContentType = $Page.FieldValues.ContentType
            $CurrentContentTypeName = if ($CurrentContentType) { $CurrentContentType.Name } else { "Unknown" }

            # Update the content type
            Set-PnPListItem -List $Library -Identity $Page.Id -ContentType $ContentType.Name -ErrorAction Stop
            Write-Host "✓ Updated: $PageTitle" -ForegroundColor Green
            $SuccessCount++

            # Add successful result to CSV data
            $Results += [PSCustomObject]@{
                PageTitle = $PageTitle
                PageUrl = $Page.FieldValues.FileRef
                PreviousContentType = $CurrentContentTypeName
                NewContentType = $ContentTypeName
                Status = "Success"
                ErrorMessage = ""
                ProcessedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "✗ Failed to update: $PageTitle - Error: $ErrorMessage" -ForegroundColor Red
            $ErrorCount++

            # Add failed result to CSV data
            $Results += [PSCustomObject]@{
                PageTitle = $PageTitle
                PageUrl = $Page.FieldValues.FileRef
                PreviousContentType = if ($Page.FieldValues.ContentType) { $Page.FieldValues.ContentType.Name } else { "Unknown" }
                NewContentType = $ContentTypeName
                Status = "Failed"
                ErrorMessage = $ErrorMessage
                ProcessedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }

    # Export results to CSV
    Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
    try {
        $Results | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ CSV report saved to: $CsvOutputPath" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to export CSV: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Summary
    Write-Host "`n=== Update Summary ===" -ForegroundColor Cyan
    Write-Host "Total pages processed: $($Pages.Count)" -ForegroundColor White
    Write-Host "Successfully updated: $SuccessCount" -ForegroundColor Green
    Write-Host "Failed to update: $ErrorCount" -ForegroundColor Red
    Write-Host "Results exported to: $CsvOutputPath" -ForegroundColor Cyan

    if ($ErrorCount -eq 0) {
        Write-Host "`nAll pages in the '$FolderName' folder have been successfully updated to '$ContentTypeName' content type!" -ForegroundColor Green
    } else {
        Write-Host "`nUpdate completed with some errors. Please review the failed items above." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Disconnect from SharePoint
    Write-Host "`nDisconnecting from SharePoint..." -ForegroundColor Yellow
    Disconnect-PnPOnline
}
