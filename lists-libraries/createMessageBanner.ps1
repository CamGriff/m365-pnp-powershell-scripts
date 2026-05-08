# Script: Create-MessageBanner-List.ps1
# Description: Creates the MessageBanner SharePoint list with all required columns. Ideally columns should be created as site columns in content types. 
# Once the list is created you will need to created managed properties
# Full documentation: https://www.camerongriffiths.com/scripts/createmessagebanner
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Parameters
$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId = ""
$ListName = "MessageBanner"
$ListDescription = "List-driven message banner for the intranet homepage"

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    Write-Host "Connected to SharePoint" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Create the list
try {
    $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "⚠ List '$ListName' already exists — skipping creation" -ForegroundColor Yellow
    }
    else {
        New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch
        Write-Host "List '$ListName' created" -ForegroundColor Green
    }
}
catch {
    Write-Host "Failed to create list: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Add columns
$columns = @(
    @{
        Name        = "TitleFR"
        Type        = "Text"
        Description = "French message title"
    },
    @{
        Name        = "Description"
        Type        = "Note"
        Description = "English message body"
    },
    @{
        Name        = "DescriptionFR"
        Type        = "Note"
        Description = "French message body"
    },
    @{
        Name        = "SeeMore"
        Type        = "URL"
        Description = "English link to related page"
    },
    @{
        Name        = "SeeMoreFR"
        Type        = "URL"
        Description = "French link to related page"
    }
)

foreach ($column in $columns) {
    try {
        $existing = Get-PnPField -List $ListName -Identity $column.Name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "Column '$($column.Name)' already exists — skipping" -ForegroundColor Yellow
        }
        else {
            Add-PnPField -List $ListName -DisplayName $column.Name -InternalName $column.Name -Type $column.Type -AddToDefaultView
            Write-Host "Column '$($column.Name)' created" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Failed to create column '$($column.Name)': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Add MessageLevel choice column
try {
    $existing = Get-PnPField -List $ListName -Identity "MessageLevel" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Column 'MessageLevel' already exists — skipping" -ForegroundColor Yellow
    }
    else {
        Add-PnPField -List $ListName -DisplayName "MessageLevel" -InternalName "MessageLevel" -Type Choice -AddToDefaultView -Choices @("Low", "Medium", "High")
        Write-Host "Column 'MessageLevel' created" -ForegroundColor Green
    }
}
catch {
    Write-Host "Failed to create column 'MessageLevel': $($_.Exception.Message)" -ForegroundColor Red
}

# Add MessageStatus choice column
try {
    $existing = Get-PnPField -List $ListName -Identity "MessageStatus" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Column 'MessageStatus' already exists — skipping" -ForegroundColor Yellow
    }
    else {
        Add-PnPField -List $ListName -DisplayName "MessageStatus" -InternalName "MessageStatus" -Type Choice -AddToDefaultView -Choices @("Active", "Inactive")
        Write-Host "Column 'MessageStatus' created" -ForegroundColor Green
    }
}
catch {
    Write-Host "Failed to create column 'MessageStatus': $($_.Exception.Message)" -ForegroundColor Red
}

# Add ExpirationDate column
try {
    $existing = Get-PnPField -List $ListName -Identity "ExpirationDate" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Column 'ExpirationDate' already exists — skipping" -ForegroundColor Yellow
    }
    else {
        Add-PnPField -List $ListName -DisplayName "ExpirationDate" -InternalName "ExpirationDate" -Type DateTime -AddToDefaultView
        Write-Host "Column 'ExpirationDate' created" -ForegroundColor Green
    }
}
catch {
    Write-Host "Failed to create column 'ExpirationDate': $($_.Exception.Message)" -ForegroundColor Red
}

# Set default value for MessageStatus to Inactive
try {
    $field = Get-PnPField -List $ListName -Identity "MessageStatus"
    $field.DefaultValue = "Inactive"
    $field.Update()
    Invoke-PnPQuery
    Write-Host "Default value for 'MessageStatus' set to Inactive" -ForegroundColor Green
}
catch {
    Write-Host "Could not set default value for MessageStatus: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "MessageBanner list setup complete" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Map MessageStatus to a RefinableString managed property (e.g. RefinableString100)" -ForegroundColor White
Write-Host "  2. Map MessageLevel to RefinableString101" -ForegroundColor White
Write-Host "  3. Map TitleFR to RefinableString105" -ForegroundColor White
Write-Host "  4. Map DescriptionFR to RefinableString106" -ForegroundColor White
Write-Host "  5. Trigger a reindex on the list" -ForegroundColor White
Write-Host "  6. Add the PnP Search Results web part to the homepage" -ForegroundColor White

Disconnect-PnPOnline
