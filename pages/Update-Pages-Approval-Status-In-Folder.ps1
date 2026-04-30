# SharePoint Online - Update Page Approval Status Script
# This script updates the approval status of all pages in a specific document library

# Variables - Update these as needed
$ClientId = ""
$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
$LibraryName = "Site Pages"
$FolderName = ""  # Optional: Leave empty "" to process all pages
$ApprovalStatus = 0  # 0 for Approved

# Logging variables
$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptPath)) { $ScriptPath = Get-Location }
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $ScriptPath "SharePoint_Approval_Update_Log_$TimeStamp.txt"
$CsvFile = Join-Path $ScriptPath "SharePoint_Approval_Update_Results_$TimeStamp.csv"
$Results = @()

# Function to write to log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"

    # Write to console
    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }

    # Write to log file
    try {
        Add-Content -Path $LogFile -Value $LogEntry -ErrorAction SilentlyContinue
    }
    catch {
        # If logging fails, continue without stopping the script
    }
}

# Function to add result to CSV data
function Add-Result {
    param(
        $ItemId,
        $ItemName,
        $ItemUrl,
        $OriginalStatus,
        $NewStatus,
        $Status,
        $ErrorMessage = ""
    )

    $script:Results += [PSCustomObject]@{
        ItemId = $ItemId
        ItemName = $ItemName
        ItemUrl = $ItemUrl
        OriginalStatus = $OriginalStatus
        NewStatus = $NewStatus
        Status = $Status
        ErrorMessage = $ErrorMessage
        ProcessedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# Function to connect to SharePoint Online
function Connect-ToSharePoint {
    param($SiteUrl)

    try {
        Write-Log "Connecting to SharePoint Online site: $SiteUrl" "INFO"

        # Connect using interactive login (will prompt for credentials)
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

        Write-Log "Successfully connected to SharePoint Online" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error connecting to SharePoint: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to update page approval status
function Update-PageApprovalStatus {
    param(
        $LibraryName,
        $FolderName,
        $ApprovalStatus
    )

    try {
        if ([string]::IsNullOrEmpty($FolderName)) {
            Write-Log "Getting all items from library: $LibraryName" "INFO"
            $ListItems = Get-PnPListItem -List $LibraryName
        } else {
            Write-Log "Getting items from library: $LibraryName, folder: $FolderName" "INFO"

            # Get items from specific folder
            try {
                $ListItems = Get-PnPListItem -List $LibraryName -FolderServerRelativeUrl "/sites/devintranet/SitePages/$FolderName"
            }
            catch {
                Write-Log "Folder '$FolderName' not found with direct path. Trying alternative method..." "WARNING"
                # Alternative: Get all items and filter by folder path
                $AllItems = Get-PnPListItem -List $LibraryName
                $ListItems = $AllItems | Where-Object {
                    $_.FileSystemObjectType -eq "File" -and
                    ($_.FieldValues.FileDirRef -like "*/$FolderName" -or
                     $_.FieldValues.FileDirRef -like "*/$FolderName/*" -or
                     $_.FieldValues.FileRef -like "*/$FolderName/*")
                }
            }
        }

        if ($ListItems.Count -eq 0) {
            if ([string]::IsNullOrEmpty($FolderName)) {
                Write-Log "No items found in the library" "WARNING"
            } else {
                Write-Log "No items found in the folder '$FolderName'" "WARNING"
            }
            return
        }

        if ([string]::IsNullOrEmpty($FolderName)) {
            Write-Log "Found $($ListItems.Count) items in the library" "SUCCESS"
        } else {
            Write-Log "Found $($ListItems.Count) items in folder '$FolderName'" "SUCCESS"
        }

        $SuccessCount = 0
        $ErrorCount = 0

        # Loop through each item and update approval status
        foreach ($Item in $ListItems) {
            try {
                $ItemName = $Item["Title"]
                if ([string]::IsNullOrEmpty($ItemName)) {
                    $ItemName = $Item["FileLeafRef"]
                }

                $ItemUrl = $Item.FieldValues.FileRef
                $CurrentStatus = $Item.FieldValues["_ModerationStatus"]

                Write-Log "Processing: $ItemName (ID: $($Item.Id))" "INFO"
                Write-Log "  Current moderation status: $CurrentStatus" "INFO"

                # Update the approval status - try common approval status fields
                $UpdateValues = @{}

                if ($Item.FieldValues.ContainsKey("_ModerationStatus")) {
                    $UpdateValues["_ModerationStatus"] = $ApprovalStatus
                }
                if ($Item.FieldValues.ContainsKey("PublishingPageStatus")) {
                    $UpdateValues["PublishingPageStatus"] = $ApprovalStatus
                }
                if ($Item.FieldValues.ContainsKey("ApprovalStatus")) {
                    $UpdateValues["ApprovalStatus"] = $ApprovalStatus
                }

                # If no specific approval field found, try the moderation status
                if ($UpdateValues.Count -eq 0) {
                    $UpdateValues["_ModerationStatus"] = $ApprovalStatus
                }

                Set-PnPListItem -List $LibraryName -Identity $Item.Id -Values $UpdateValues

                Write-Log " Successfully updated approval status" "SUCCESS"
                Add-Result -ItemId $Item.Id -ItemName $ItemName -ItemUrl $ItemUrl -OriginalStatus $CurrentStatus -NewStatus "Approved" -Status "Success"
                $SuccessCount++
            }
            catch {
                $ErrorMsg = $_.Exception.Message
                Write-Log " Error updating item: $ErrorMsg" "ERROR"
                Add-Result -ItemId $Item.Id -ItemName $ItemName -ItemUrl $ItemUrl -OriginalStatus $CurrentStatus -NewStatus "Failed" -Status "Failed" -ErrorMessage $ErrorMsg
                $ErrorCount++
            }
        }

        # Summary
        Write-Log "`n=== UPDATE SUMMARY ===" "INFO"
        Write-Log "Total items processed: $($ListItems.Count)" "INFO"
        Write-Log "Successfully updated: $SuccessCount" "SUCCESS"
        Write-Log "Errors encountered: $ErrorCount" "ERROR"
    }
    catch {
        Write-Log "Error processing library: $($_.Exception.Message)" "ERROR"
    }
}

# Function to enable content approval on the library (if needed)
function Enable-ContentApproval {
    param($LibraryName)

    try {
        Write-Log "Checking if content approval is enabled for library: $LibraryName" "INFO"

        $List = Get-PnPList -Identity $LibraryName

        if (-not $List.EnableModeration) {
            Write-Log "Content approval is not enabled. Enabling now..." "WARNING"
            Set-PnPList -Identity $LibraryName -EnableContentApproval $true
            Write-Log "Content approval has been enabled" "SUCCESS"
        } else {
            Write-Log "Content approval is already enabled" "SUCCESS"
        }
    }
    catch {
        Write-Log "Error checking/enabling content approval: $($_.Exception.Message)" "ERROR"
    }
}

# Function to export results to CSV
function Export-ResultsToCsv {
    try {
        if ($Results.Count -gt 0) {
            $Results | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
            Write-Log "Results exported to CSV file: $CsvFile" "SUCCESS"
        } else {
            Write-Log "No results to export to CSV" "WARNING"
        }
    }
    catch {
        Write-Log "Error exporting results to CSV: $($_.Exception.Message)" "ERROR"
    }
}

# Main execution
try {
    # Initialize log file
    Write-Log "=== SharePoint Online Page Approval Status Update Script Started ===" "INFO"
    Write-Log "Script executed by: $env:USERNAME" "INFO"
    Write-Log "Execution time: $(Get-Date)" "INFO"
    Write-Log "Script location: $ScriptPath" "INFO"
    Write-Log "Log file: $LogFile" "INFO"
    Write-Log "CSV file: $CsvFile" "INFO"
    Write-Log "" "INFO"

    Write-Log "=== CONFIGURATION ===" "INFO"
    Write-Log "Site: $SiteUrl" "INFO"
    Write-Log "Library: $LibraryName" "INFO"
    if (-not [string]::IsNullOrEmpty($FolderName)) {
        Write-Log "Folder: $FolderName" "INFO"
    }
    Write-Log "New Approval Status: Approved (Value: $ApprovalStatus)" "INFO"
    Write-Log "" "INFO"

    # Confirm execution
    $Confirmation = Read-Host "Do you want to proceed with updating all pages? (Y/N)"
    if ($Confirmation -ne "Y" -and $Confirmation -ne "y") {
        Write-Log "Script execution cancelled by user" "WARNING"
        exit
    }

    # Connect to SharePoint
    if (-not (Connect-ToSharePoint -SiteUrl $SiteUrl)) {
        Write-Log "Failed to connect to SharePoint. Exiting." "ERROR"
        exit
    }

    # Enable content approval if needed
    Enable-ContentApproval -LibraryName $LibraryName

    # Update page approval status
    Update-PageApprovalStatus -LibraryName $LibraryName -FolderName $FolderName -ApprovalStatus $ApprovalStatus

    # Export results to CSV
    Export-ResultsToCsv

    Write-Log "" "INFO"
    Write-Log "=== SCRIPT EXECUTION COMPLETED ===" "SUCCESS"
    Write-Log "Log file saved to: $LogFile" "INFO"
    if ($Results.Count -gt 0) {
        Write-Log "Results exported to: $CsvFile" "INFO"
    }
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline
        Write-Log "Disconnected from SharePoint Online" "INFO"
    }
    catch {
        # Ignore disconnect errors
    }

    Write-Log "=== SCRIPT EXECUTION ENDED ===" "INFO"
}
