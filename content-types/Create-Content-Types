# SharePoint Online Content Types Creation Script
$SiteURL = "https://tenantName.sharepoint.com/sites/siteName"
$JSONFile = "createContentTypes.json"
$ClientId = "" # Your Client ID here

# Set log file path to same directory as JSON file
$JSONFilePath = Resolve-Path $JSONFile
$LogFile = Join-Path (Split-Path $JSONFilePath -Parent) "ContentTypeCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Initialize CSV with headers
"Timestamp,Action,ItemType,ItemName,Status,Message" | Out-File -FilePath $LogFile -Encoding UTF8

# Initialize a hashtable to track default content types by list
$script:DefaultContentTypes = @{}

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White",
        [string]$Action,
        [string]$ItemType,
        [string]$ItemName,
        [string]$Status
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host $Message -ForegroundColor $Color

    $csvLine = [PSCustomObject]@{
        Timestamp = $TimeStamp
        Action = $Action
        ItemType = $ItemType
        ItemName = $ItemName
        Status = $Status
        Message = $Message
    }

    $csvLine | Export-Csv -Path $LogFile -Append -NoTypeInformation
}

try {
    # Connect to SharePoint Online
    Write-Log -Message "Connecting to SharePoint site: $SiteURL" -Color Cyan -Action "Connect" -ItemType "SharePoint" -ItemName $SiteURL -Status "Started"
    Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $ClientId
    Write-Log -Message "Connected to SharePoint successfully" -Color Green -Action "Connect" -ItemType "SharePoint" -ItemName $SiteURL -Status "Success"

    # Read and parse JSON file
    Write-Log -Message "Reading JSON file: $JSONFile" -Color Cyan -Action "Read" -ItemType "JSON" -ItemName $JSONFile -Status "Started"
    try {
        $JSONContent = Get-Content -Path $JSONFile -Raw | ConvertFrom-Json
        Write-Log -Message "JSON file read successfully" -Color Green -Action "Read" -ItemType "JSON" -ItemName $JSONFile -Status "Success"
    }
    catch {
        throw "Failed to read or parse JSON file: $($_.Exception.Message)"
    }

    # Process each content type
    foreach ($ContentType in $JSONContent.ContentTypes) {
        Write-Log -Message "Processing content type: $($ContentType.Name)" -Color Cyan -Action "Process" -ItemType "ContentType" -ItemName $ContentType.Name -Status "Started"

        try {
            # Check if ListName is specified in the JSON
            if (-not $ContentType.ListName) {
                $errorMessage = "No ListName specified for content type '$($ContentType.Name)'. Please add a ListName property in the JSON."
                Write-Log -Message $errorMessage -Color Red -Action "Validate" -ItemType "ContentType" -ItemName $ContentType.Name -Status "Error"
                throw $errorMessage
            }

            # Get the parent content type object - use specified one or fallback to "Item"
            $parentContentTypeName = if ($ContentType.ParentContentType) { $ContentType.ParentContentType } else { "Item" }

            try {
                Write-Log -Message "Getting parent content type: $parentContentTypeName" -Color Cyan -Action "Get" -ItemType "ParentContentType" -ItemName $parentContentTypeName -Status "Started"
                $ParentCT = Get-PnPContentType -Identity $parentContentTypeName -ErrorAction Stop
                Write-Log -Message "Retrieved parent content type successfully" -Color Green -Action "Get" -ItemType "ParentContentType" -ItemName $parentContentTypeName -Status "Success"
            } catch {
                $errorMessage = "Parent content type '$parentContentTypeName' not found: $($_.Exception.Message)"
                Write-Log -Message $errorMessage -Color Red -Action "Get" -ItemType "ParentContentType" -ItemName $parentContentTypeName -Status "Error"
                throw $errorMessage
            }

            # Check if content type already exists
            $ExistingCT = Get-PnPContentType -Identity $ContentType.Name -ErrorAction SilentlyContinue

            if ($ExistingCT) {
                Write-Log -Message "Content type '$($ContentType.Name)' already exists" -Color Yellow -Action "Check" -ItemType "ContentType" -ItemName $ContentType.Name -Status "Exists"
                $CurrentCT = $ExistingCT
            }
            else {
                # Create new content type
                Write-Log -Message "Creating new content type: $($ContentType.Name)" -Color Cyan -Action "Create" -ItemType "ContentType" -ItemName $ContentType.Name -Status "InProgress"

                $CurrentCT = Add-PnPContentType `
                    -Name $ContentType.Name `
                    -Description $ContentType.Description `
                    -Group $ContentType.Group `
                    -ParentContentType $ParentCT

                Write-Log -Message "Content type '$($ContentType.Name)' created successfully" -Color Green -Action "Create" -ItemType "ContentType" -ItemName $ContentType.Name -Status "Success"
            }

            # Process columns for the content type
            foreach ($ColumnName in $ContentType.Columns) {
                # Determine if column is an object with Name/Required properties or just a string
                $IsRequired = $false
                $FieldName = ""

                if ($null -eq $ColumnName) {
                    Write-Log -Message "Null column found in content type '$($ContentType.Name)'" -Color Red -Action "Process" -ItemType "Column" -ItemName "NULL" -Status "Error"
                    continue
                }

                try {
                    if ($ColumnName.PSObject.Properties.Name -contains "Name") {
                        # Object format with properties
                        $FieldName = $ColumnName.Name
                        $IsRequired = [bool]($ColumnName.PSObject.Properties.Name -contains "Required" -and $ColumnName.Required -eq $true)
                    } else {
                        # String format
                        $FieldName = $ColumnName.ToString()
                    }
                } catch {
                    # Fallback to string conversion if PSObject access fails
                    $FieldName = $ColumnName.ToString()
                }

                # Ensure we have a valid field name
                if ([string]::IsNullOrEmpty($FieldName)) {
                    Write-Log -Message "Invalid or empty column name in content type '$($ContentType.Name)'" -Color Red -Action "Process" -ItemType "Column" -ItemName "Unknown" -Status "Error"
                    continue
                }

                Write-Log -Message "Processing column '$FieldName' for content type '$($ContentType.Name)'" -Color Cyan -Action "Process" -ItemType "Column" -ItemName $FieldName -Status "Started"

                try {
                    # Check if field exists at the site level
                    $Field = Get-PnPField -Identity $FieldName -ErrorAction SilentlyContinue

                    if ($Field) {
                        try {
                            # Add field to content type
                            Add-PnPFieldToContentType -Field $Field -ContentType $CurrentCT -ErrorAction Stop
                            Write-Log -Message "Added column '$FieldName' to content type '$($ContentType.Name)'" -Color Green -Action "Add" -ItemType "Column" -ItemName $FieldName -Status "Success"

                            # Set the field as required if specified
                            if ($IsRequired) {
                                try {
                                    # Get the context and field directly
                                    $ctx = Get-PnPContext

                                    # Set field as required in the content type
                                    $contentTypeField = $CurrentCT.Fields.GetByInternalNameOrTitle($FieldName)
                                    $ctx.Load($contentTypeField)
                                    $ctx.ExecuteQuery()

                                    # Set the required property and update
                                    $contentTypeField.Required = $true
                                    $contentTypeField.Update()
                                    $ctx.ExecuteQuery()

                                    Write-Log -Message "Set column '$FieldName' as required in content type '$($ContentType.Name)'" -Color Green -Action "Update" -ItemType "Column" -ItemName $FieldName -Status "Success"
                                } catch {
                                    Write-Log -Message "Error setting column '$FieldName' as required: $($_.Exception.Message)" -Color Red -Action "Update" -ItemType "Column" -ItemName $FieldName -Status "Error"
                                }
                            }
                        } catch {
                            Write-Log -Message "Error adding column '$FieldName' to content type: $($_.Exception.Message)" -Color Red -Action "Add" -ItemType "Column" -ItemName $FieldName -Status "Error"
                        }
                    } else {
                        Write-Log -Message "Column '$FieldName' not found in site" -Color Red -Action "Check" -ItemType "Column" -ItemName $FieldName -Status "NotFound"
                    }
                } catch {
                    Write-Log -Message "Error processing column '$FieldName': $($_.Exception.Message)" -Color Red -Action "Process" -ItemType "Column" -ItemName $FieldName -Status "Error"
                }
            }

            $ListName = $ContentType.ListName

            # Get target list/library
            Write-Log -Message "Getting list/library: $ListName" -Color Cyan -Action "Get" -ItemType "Library" -ItemName $ListName -Status "Started"
            try {
                $library = Get-PnPList -Identity $ListName -ErrorAction Stop
                Write-Log -Message "Retrieved list/library successfully" -Color Green -Action "Get" -ItemType "Library" -ItemName $ListName -Status "Success"

                # Enable content type management on the library if not already enabled
                if (-not $library.ContentTypesEnabled) {
                    Write-Log -Message "Enabling content types on library: $ListName" -Color Cyan -Action "Enable" -ItemType "Library" -ItemName $ListName -Status "Started"
                    Set-PnPList -Identity $ListName -EnableContentTypes $true
                    Write-Log -Message "Enabled content types on library successfully" -Color Green -Action "Enable" -ItemType "Library" -ItemName $ListName -Status "Success"
                }

                # Add content type to the list/library
                Write-Log -Message "Adding content type '$($ContentType.Name)' to list/library: $ListName" -Color Cyan -Action "Add" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Started"
                try {
                    Add-PnPContentTypeToList -List $ListName -ContentType $CurrentCT -ErrorAction Stop
                    Write-Log -Message "Added content type '$($ContentType.Name)' to list/library successfully" -Color Green -Action "Add" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Success"

                    # Check if content type should be set as default
                    if ($ContentType.Default -eq $true) {
                        # Check if we've already set a default content type for this list in this run
                        if ($script:DefaultContentTypes.ContainsKey($ListName)) {
                            $warningMessage = "Warning: Content type '$($ContentType.Name)' will be set as default for list '$ListName', but content type '$($script:DefaultContentTypes[$ListName])' was already set as default earlier. The last one processed will become the actual default."
                            Write-Log -Message $warningMessage -Color Yellow -Action "SetDefault" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Warning"
                        }

                        # Record this content type as the default for this list
                        $script:DefaultContentTypes[$ListName] = $ContentType.Name

                        Write-Log -Message "Setting content type '$($ContentType.Name)' as default for list/library: $ListName" -Color Cyan -Action "SetDefault" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Started"

                        try {
                            # Set the content type as default for the list
                            Set-PnPDefaultContentTypeToList -List $ListName -ContentType $CurrentCT.Name -ErrorAction Stop

                            Write-Log -Message "Set content type '$($ContentType.Name)' as default for list/library successfully" -Color Green -Action "SetDefault" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Success"
                        }
                        catch {
                            Write-Log -Message "Error setting content type '$($ContentType.Name)' as default: $($_.Exception.Message)" -Color Red -Action "SetDefault" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Error"
                        }
                    }
                }
                catch {
                    Write-Log -Message "Error adding content type '$($ContentType.Name)' to list/library: $($_.Exception.Message)" -Color Red -Action "Add" -ItemType "ContentTypeToLibrary" -ItemName $ContentType.Name -Status "Error"
                }
            }
            catch {
                Write-Log -Message "Error: List/library '$ListName' not found. Skipping content type addition." -Color Red -Action "Get" -ItemType "Library" -ItemName $ListName -Status "NotFound"
            }
        }
        catch {
            Write-Log -Message "Error processing content type '$($ContentType.Name)': $($_.Exception.Message)" -Color Red -Action "Process" -ItemType "ContentType" -ItemName $ContentType.Name -Status "Error"
        }
    }
}
catch {
    Write-Log -Message "Critical error: $($_.Exception.Message)" -Color Red -Action "Script" -ItemType "Global" -ItemName "Error" -Status "Error"
}
finally {
    # Disconnect from SharePoint
    if (Get-PnPConnection) {
        Write-Log -Message "Disconnecting from SharePoint..." -Color Cyan -Action "Disconnect" -ItemType "SharePoint" -ItemName $SiteURL -Status "Started"
        Disconnect-PnPOnline
        Write-Log -Message "Disconnected from SharePoint successfully" -Color Green -Action "Disconnect" -ItemType "SharePoint" -ItemName $SiteURL -Status "Success"
    }

    Write-Log -Message "Script execution completed. Log file saved to: $LogFile" -Color Yellow -Action "Script" -ItemType "Global" -ItemName "Complete" -Status "Success"
}
