# SharePoint Online List Creation Script
# This script creates custom lists in SharePoint Online and logs operations

# Site URL
$siteUrl = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId = ""

# List of lists to create
$listsToCreate = @(
    "Intranet_Breadcrumbs"
)

# Log file path - creates in the same directory as script
$logFile = ".\SPO_List_Creation_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Create CSV file with headers
"Timestamp,ListName,Status,Message,ListUrl,ListId" | Out-File -FilePath $logFile

# Function to write to log file
function Write-ToLog {
    param(
        [string]$ListName,
        [string]$Status,
        [string]$Message,
        [string]$ListUrl = "",
        [string]$ListId = ""
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp,$ListName,$Status,$Message,$ListUrl,$ListId" | Out-File -FilePath $logFile -Append
}

try {
    # Ensure PnP PowerShell is loaded
    if (-not (Get-Module -Name "PnP.PowerShell")) {
        Write-Host "Loading PnP PowerShell module..." -ForegroundColor Yellow
        Import-Module PnP.PowerShell -ErrorAction Stop
    }

    # Connect to SharePoint Online with verbose error handling
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
    try {
        Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId -ErrorAction Stop

        # Validate connection by attempting to get web details
        $web = Get-PnPWeb -ErrorAction Stop
        Write-Host "Successfully connected to site: $($web.Title)" -ForegroundColor Green
        Write-ToLog -ListName "Connection" -Status "Success" -Message "Connected to SharePoint Online site: $($web.Title)"
    }
    catch {
        $errorMessage = "Failed to connect to SharePoint site. Error: $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Write-ToLog -ListName "Connection" -Status "Error" -Message $errorMessage
        throw $_
    }

    # Create each list
    foreach ($listName in $listsToCreate) {
        try {
            Write-Host "`nCreating list: $listName" -ForegroundColor Yellow

            # Check if list already exists with better error handling
            $existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

            if ($existingList) {
                Write-Host "List '$listName' already exists. Skipping..." -ForegroundColor Cyan
                Write-ToLog -ListName $listName -Status "Skipped" -Message "List already exists" -ListUrl $existingList.RootFolder.ServerRelativeUrl -ListId $existingList.Id
                continue
            }

            # Create new list with explicit error handling
            $newList = New-PnPList -Title $listName -Template GenericList -ErrorAction Stop

            # Verify list was created and get details
            $list = Get-PnPList -Identity $listName -ErrorAction Stop
            $listUrl = $list.RootFolder.ServerRelativeUrl
            $listId = $list.Id

            Write-Host "List '$listName' created successfully!" -ForegroundColor Green
            Write-Host "List URL: $listUrl" -ForegroundColor Green
            Write-Host "List ID: $listId" -ForegroundColor Green

            Write-ToLog -ListName $listName -Status "Success" -Message "List created successfully" -ListUrl $listUrl -ListId $listId

        } catch {
            $errorMessage = "Error creating list '$listName': $($_.Exception.Message)"
            Write-Host $errorMessage -ForegroundColor Red
            Write-ToLog -ListName $listName -Status "Error" -Message $errorMessage
        }
    }

} catch {
    $errorMessage = "Critical error: $($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    Write-ToLog -ListName "Critical" -Status "Error" -Message $errorMessage
} finally {
    # Disconnect from SharePoint Online
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "`nDisconnected from SharePoint Online" -ForegroundColor Yellow
    } catch {
        Write-Host "Error disconnecting from SharePoint Online" -ForegroundColor Red
    }

    Write-Host "`nOperation completed. Check log file at: $logFile" -ForegroundColor Yellow
}
