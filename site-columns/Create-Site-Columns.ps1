# Script: Create-Site-Columns.ps1
# Description: Creates SharePoint Online site columns from JSON configuration file
# Full documentation: https://www.camerongriffiths.com/scripts/createsitecolumns
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Parameters
$SiteCollectionURL = "https://tenantName.sharepoint.com/sites/siteName"
$JSONFile = "createSiteColumns.json"
$ColumnGroup = "" # Column group name youw ant your columns categorised by
$ClientId = "" # Put your Tenant Client ID Here
$JSONFilePath = Resolve-Path $JSONFile
$LogFile = Join-Path (Split-Path $JSONFilePath -Parent) "SiteColumnCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp - $Message"
    
    # Write to console with color
    Write-Host $Message -ForegroundColor $Color
    
    # Write to log file
    Add-Content -Path $LogFile -Value $LogMessage
}

try {
    # Connect to PnP Online
    Write-Log "Connecting to SharePoint Online..." -Color Yellow
    Connect-PnPOnline -Url $SiteCollectionURL -Interactive -ClientId $ClientId

    # Read JSON file
    Write-Log "Reading JSON configuration file..." -Color Yellow
    $JSONContent = Get-Content -Path $JSONFile -Raw | ConvertFrom-Json

    # Process each site column
    Write-Log "Starting site column creation process..." -Color Yellow
    foreach ($column in $JSONContent.SiteColumns) {
        try {
            # Check if the column already exists
            $existingField = Get-PnPField -Identity $column.InternalName -ErrorAction SilentlyContinue

            if ($null -eq $existingField) {
                # Create column based on type
                switch ($column.Type) {
                    "TaxonomyField" {
                        Write-Log "Creating Managed Metadata field: $($column.DisplayName)" -Color Cyan
                        Add-PnPTaxonomyField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -TermSetPath $column.TermSetPath `
                            -Group $ColumnGroup `
                            -Required:$column.Required
                    }
                    
                    "Choice" {
                        Write-Log "Creating Choice field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type Choice `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                            -Choices $column.Choices

                       # Check if a default value is provided
                     if ($column.DefaultValue) {
                        Write-Log "Setting default value for $($column.DisplayName): $($column.DefaultValue)" -Color Cyan
                        Set-PnPField `
                            -Identity $column.InternalName `
                            -Values @{DefaultValue=$column.DefaultValue}
                        }     
                    }
                    
                    "DateTime" {
                        Write-Log "Creating DateTime field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type DateTime `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                    }   

                    
                    "Note" {
                        Write-Log "Creating Note field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type Note `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                            #-RichText:$column.RichText `
                            #-NumberOfLines $column.NumberOfLines
                    }

                    "Text" {
                        Write-Log "Creating Text field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type Text `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                    }
                    
                    "Number" {
                        Write-Log "Creating Number field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type Number `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                    }
                    
                    "Boolean" {
                        Write-Log "Creating Boolean field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type Boolean `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                    }
                    
                    "PersonField" {
                        Write-Log "Creating Person field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type User `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                          
                    }
                    
                    "URL" {
                        Write-Log "Creating URL field: $($column.DisplayName)" -Color Cyan
                        Add-PnPField `
                            -DisplayName $column.DisplayName `
                            -InternalName $column.InternalName `
                            -Type URL `
                            -Group $ColumnGroup `
                            -Required:$column.Required `
                    }
                    
                    default {
                        Write-Log "Unknown field type '$($column.Type)' for column '$($column.DisplayName)'" -Color Red
                        continue
                    }
                }
                Write-Log "Successfully created column: $($column.DisplayName)" -Color Green
            }
            else {
                Write-Log "Column '$($column.DisplayName)' already exists. Skipping..." -Color Yellow
            }
        }
        catch {
            Write-Log "Error creating column '$($column.DisplayName)': $_" -Color Red
        }
    }
}
catch {
    Write-Log "A critical error occurred: $_" -Color Red
}
finally {
    # Disconnect PnP Session
    Disconnect-PnPOnline
    Write-Log "Script execution completed. Log file saved to: $LogFile" -Color Yellow
}
