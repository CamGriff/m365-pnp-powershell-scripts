# Error handling for the entire script
try {
    # Connect to your SharePoint site
    $ClientId = ""
    Connect-PnPOnline -Url "https://tenantName.sharepoint.com/sites/siteName" -Interactive -ClientId $ClientId
    Write-Host "Successfully connected to SharePoint." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to SharePoint: $($_.Exception.Message)" -ForegroundColor Red

    # Create error result for CSV
    $errorResult = [PSCustomObject]@{
        DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FolderPath = "N/A"
        FolderName = "N/A"
        Status = "Connection Error"
        Message = "Failed to connect to SharePoint: $($_.Exception.Message)"
        PreviousStatus = "N/A"
    }

    # Export error to CSV
    $csvFileName = "SharePoint_Folder_Approval_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvPath = Join-Path $PSScriptRoot $csvFileName
    @($errorResult) | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Error logged to: $csvPath" -ForegroundColor Yellow

    exit 1
}

# Initialize array to store results for CSV output
$results = @()

Write-Host "Searching for folders with PENDING status (2) in Site Pages library..." -ForegroundColor Cyan

# Get ALL items from Site Pages library (including nested folders)
# _ModerationStatus: 0 = Approved, 1 = Rejected, 2 = Pending, 3 = Draft
try {
    Write-Host "Getting ALL items from Site Pages library to find nested folders..." -ForegroundColor Cyan

    # Get ALL items with a high page size to ensure we get everything including subfolders
    $allItems = Get-PnPListItem -List "Site Pages" -PageSize 2000

    Write-Host "Retrieved $($allItems.Count) total items from Site Pages library." -ForegroundColor Green

    # Filter to only folders (including nested ones) - use FileSystemObjectType as primary method
    $allFolders = $allItems | Where-Object { $_.FileSystemObjectType -eq "Folder" }

    Write-Host "Found $($allFolders.Count) folders using FileSystemObjectType detection." -ForegroundColor Cyan

    # Show all folders we found with their paths
    Write-Host "`nAll folders found:" -ForegroundColor Yellow
    foreach ($folder in $allFolders) {
        $folderName = $folder["FileLeafRef"]
        $folderPath = $folder["FileRef"]
        $moderationStatus = $folder["_ModerationStatus"]

        Write-Host "  - '$folderName'" -ForegroundColor White
        Write-Host "    Path: '$folderPath' | Status: $moderationStatus" -ForegroundColor Gray
    }

    # Filter for folders with PENDING status (2) using PowerShell
    $foldersToApprove = $allFolders | Where-Object { $_["_ModerationStatus"] -eq 2 }

    Write-Host "`nFound $($foldersToApprove.Count) folders with PENDING status that need approval." -ForegroundColor Green

    # Show what we found that needs approval
    if ($foldersToApprove.Count -gt 0) {
        Write-Host "`nFolders with PENDING status (will be approved):" -ForegroundColor Red
        foreach ($folder in $foldersToApprove) {
            $folderName = $folder["FileLeafRef"]
            $folderPath = $folder["FileRef"]
            $moderationStatus = $folder["_ModerationStatus"]
            Write-Host "  - '$folderName' | Path: '$folderPath' | Status: $moderationStatus" -ForegroundColor White
        }
    }
}
catch {
    Write-Host "Failed to query Site Pages library: $($_.Exception.Message)" -ForegroundColor Red

    # Create error result for CSV
    $result = [PSCustomObject]@{
        DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FolderPath = "N/A"
        FolderName = "N/A"
        Status = "Query Error"
        Message = "Failed to query Site Pages library: $($_.Exception.Message)"
        PreviousStatus = "N/A"
    }
    $results += $result

    # Export error and exit
    $csvFileName = "SharePoint_Folder_Approval_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvPath = Join-Path $PSScriptRoot $csvFileName
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Error logged to: $csvPath" -ForegroundColor Yellow
    exit 1
}

if ($foldersToApprove.Count -eq 0) {
    Write-Host "No folders found that require approval." -ForegroundColor Green

    # Still create a CSV to document the run
    $result = [PSCustomObject]@{
        DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FolderPath = "N/A"
        FolderName = "N/A"
        Status = "No Action Required"
        Message = "All folders already approved"
        PreviousStatus = "N/A"
    }
    $results += $result
}
else {
    Write-Host "`nProcessing $($foldersToApprove.Count) folder(s) that need approval..." -ForegroundColor Yellow

    # Loop through each folder and approve it
    foreach ($folder in $foldersToApprove) {
        # Initialize variables to avoid scope issues
        $folderName = "Unknown"
        $folderPath = "Unknown"
        $currentStatus = "Unknown"

        # Create result object for this folder
        $result = [PSCustomObject]@{
            DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            FolderPath = ""
            FolderName = ""
            Status = ""
            Message = ""
            PreviousStatus = ""
        }

        try {
            # Safely extract folder properties with null checks
            $folderName = if ($folder["FileLeafRef"]) { $folder["FileLeafRef"] } else { "Unknown" }
            $folderPath = if ($folder["FileRef"]) { $folder["FileRef"] } else { "Unknown" }
            $currentStatus = if ($null -ne $folder["_ModerationStatus"]) { $folder["_ModerationStatus"] } else { "Unknown" }

            # Update result object with extracted values
            $result.FolderPath = $folderPath
            $result.FolderName = $folderName
            $result.PreviousStatus = $currentStatus

            Write-Host "Processing folder: '$folderName' (Current status: $currentStatus)"

            # Approve the folder
            $folder["_ModerationStatus"] = 0  # 0 = Approved
            $folder.Update()
            Invoke-PnPQuery
            Write-Host " Folder '$folderName' approved successfully" -ForegroundColor Green

            $result.Status = "Success"
            $result.Message = "Folder approved successfully (was status: $currentStatus)"
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host " Failed to approve folder '$folderName': $errorMessage" -ForegroundColor Red

            $result.FolderPath = $folderPath
            $result.FolderName = $folderName
            $result.PreviousStatus = $currentStatus
            $result.Status = "Error"
            $result.Message = "Failed to approve: $errorMessage"
        }

        # Add result to array
        $results += $result
    }
}

# Generate CSV output file
$csvFileName = "SharePoint_Folder_Approval_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$csvPath = Join-Path $PSScriptRoot $csvFileName

try {
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV results exported to: $csvPath" -ForegroundColor Cyan
}
catch {
    Write-Host "`nFailed to export CSV: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nFolder approval process completed."
