# Configuration variables
$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId = ""  # Enter Client ID

# Connect to your site
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

# Array of site column internal names to check
$columnsToCheck = @(
    "Key",
    "Order",
    "Order1"
)

# Results array
$results = @()

Write-Host "Starting site column usage check..." -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "Checking $($columnsToCheck.Count) columns" -ForegroundColor Cyan
Write-Host "="*60

foreach ($columnInternalName in $columnsToCheck) {
    Write-Host "`nChecking column: $columnInternalName" -ForegroundColor Yellow

    # Get the site column
    $siteColumn = Get-PnPField -Identity $columnInternalName -ErrorAction SilentlyContinue

    if ($null -eq $siteColumn) {
        Write-Host "  x Site column not found!" -ForegroundColor Red
        $results += [PSCustomObject]@{
            ColumnInternalName = $columnInternalName
            ColumnDisplayName = "NOT FOUND"
            UsageType = "N/A"
            LocationName = "N/A"
            LocationUrl = "N/A"
            ContentTypeName = ""
        }
        continue
    }

    $usageFound = $false

    # Check all lists in the site
    $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false }

    foreach ($list in $lists) {
        # Get fields for this list
        $listFields = Get-PnPField -List $list.Title -ErrorAction SilentlyContinue

        # Check if the site column is used in this list
        $fieldInList = $listFields | Where-Object { $_.InternalName -eq $columnInternalName }

        if ($null -ne $fieldInList) {
            $usageFound = $true
            Write-Host "  Found in List: $($list.Title)" -ForegroundColor Green

            # Add to results
            $results += [PSCustomObject]@{
                ColumnInternalName = $columnInternalName
                ColumnDisplayName = $siteColumn.Title
                UsageType = "List/Library"
                LocationName = $list.Title
                LocationUrl = $list.RootFolder.ServerRelativeUrl
                ContentTypeName = ""
            }

            # Check if it's in any content types of this list
            $listContentTypes = Get-PnPContentType -List $list.Title
            foreach ($ct in $listContentTypes) {
                $ctFields = Get-PnPProperty -ClientObject $ct -Property Fields
                if ($ctFields.InternalName -contains $columnInternalName) {
                    Write-Host "    - Used in Content Type: $($ct.Name)" -ForegroundColor Cyan

                    $results += [PSCustomObject]@{
                        ColumnInternalName = $columnInternalName
                        ColumnDisplayName = $siteColumn.Title
                        UsageType = "List Content Type"
                        LocationName = $list.Title
                        LocationUrl = $list.RootFolder.ServerRelativeUrl
                        ContentTypeName = $ct.Name
                    }
                }
            }
        }
    }

    # Check site content types
    $siteContentTypes = Get-PnPContentType
    foreach ($ct in $siteContentTypes) {
        $ctFields = Get-PnPProperty -ClientObject $ct -Property Fields
        if ($ctFields.InternalName -contains $columnInternalName) {
            $usageFound = $true
            Write-Host "  Found in Site Content Type: $($ct.Name)" -ForegroundColor Green

            $results += [PSCustomObject]@{
                ColumnInternalName = $columnInternalName
                ColumnDisplayName = $siteColumn.Title
                UsageType = "Site Content Type"
                LocationName = "Site Level"
                LocationUrl = ""
                ContentTypeName = $ct.Name
            }
        }
    }

    if (-not $usageFound) {
        Write-Host "  Column not used anywhere in this site" -ForegroundColor Yellow

        $results += [PSCustomObject]@{
            ColumnInternalName = $columnInternalName
            ColumnDisplayName = $siteColumn.Title
            UsageType = "Not Used"
            LocationName = ""
            LocationUrl = ""
            ContentTypeName = ""
        }
    }
}

# Export to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = "C:\Temp\SiteColumnUsage_$timestamp.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`n="*60
Write-Host "Check complete!" -ForegroundColor Green
Write-Host "Results exported to: $csvPath" -ForegroundColor Green
Write-Host "Total columns checked: $($columnsToCheck.Count)" -ForegroundColor Cyan
Write-Host "Total usage records: $($results.Count)" -ForegroundColor Cyan
