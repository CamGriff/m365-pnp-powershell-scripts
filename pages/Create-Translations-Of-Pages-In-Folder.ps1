# Script: Create-Translation-Of-Pages-In-Folder.ps1
# Description: allows the creation of translated pages (not the content - just the page) for all pages in a folder
# Full documentation: https://www.camerongriffiths.com/scripts/createtranslationsofpagesinfolder
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions


# PnP PowerShell script to create French translation variants for all pages in a folder

            param(
                [Parameter(Mandatory=$false)]
                [string]$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName",
                
                [Parameter(Mandatory=$false)]
                [string]$ClientId = "",
                
                [Parameter(Mandatory=$false)]
                [string]$FolderName = "", # Define anem of target folder
                
                [Parameter(Mandatory=$false)]
                [int]$FrenchLCID = 1036 # Define language 
            )
            
            # Initialize CSV logging
            $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $csvLogFile = Join-Path $scriptPath "FrenchTranslation_Guidelines_$timestamp.csv"
            
            # Create CSV log array
            $logEntries = @()
            
            Write-Host "=== French Translation for Guidelines Folder ===" -ForegroundColor Cyan
            Write-Host "Target Folder: $FolderName" -ForegroundColor Yellow
            Write-Host "CSV log will be saved to: $csvLogFile" -ForegroundColor Cyan
            
            # Connect to SharePoint site
            try {
                Write-Host "`nConnecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
                Write-Host "Using Client ID: $ClientId" -ForegroundColor Yellow
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
                Write-Host "Connected successfully!" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
                exit 1
            }
            
            # Get the Site Pages library
            Write-Host "`nGetting Site Pages library..." -ForegroundColor Yellow
            $Library = Get-PnPList -Identity "Site Pages"
            
            if ($null -eq $Library) {
                Write-Host "Error: Could not find the Site Pages library" -ForegroundColor Red
                exit 1
            }
            
            Write-Host "Site Pages library found. Server relative URL: $($Library.RootFolder.ServerRelativeUrl)" -ForegroundColor Green
            
            # Get the target folder
            Write-Host "`nGetting folder '$FolderName'..." -ForegroundColor Yellow
            try {
                $FolderServerRelativeUrl = $Library.RootFolder.ServerRelativeUrl + "/" + $FolderName
                $Folder = Get-PnPFolder -Url $FolderServerRelativeUrl -ErrorAction Stop
                Write-Host "✓ Found folder: $($Folder.ServerRelativeUrl)" -ForegroundColor Green
            }
            catch {
                Write-Host "✗ Could not find folder '$FolderName': $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Expected folder path: $FolderServerRelativeUrl" -ForegroundColor Yellow
                exit 1
            }
            
            # Get all pages from the specific folder
            Write-Host "`nGetting all pages from the '$FolderName' folder..." -ForegroundColor Yellow
            $FolderPages = Get-PnPListItem -List $Library -FolderServerRelativeUrl $Folder.ServerRelativeUrl -PageSize 1000
            
            # Filter for .aspx pages only and exclude duplicates but don't exclude French translations upfront
            $pagesToTranslate = @()
            $processedPages = @{}
            
            foreach ($item in $FolderPages)
                if ($item.FieldValues.FileLeafRef -like "*.aspx") {
                    $fileName = $item.FieldValues.FileLeafRef
                    $fullPath = $item.FieldValues.FileRef
                    
                    # Only skip obvious French translation files (with _fr suffix)
                    # We'll check the translation status individually for each page later
                    if ($fileName -match "_fr\.aspx$") {
                        Write-Host "⚠ Skipping obvious French translation file: $fileName" -ForegroundColor Yellow
                        continue
                    }
                    
                    # Skip duplicates (same filename)
                    if ($processedPages.ContainsKey($fileName)) {
                        Write-Host "⚠ Skipping duplicate page: $fileName" -ForegroundColor Yellow
                        continue
                    }
                    
                    $pageObj = [PSCustomObject]@{
                        Name = $fileName
                        Title = if ($item.FieldValues.Title) { $item.FieldValues.Title } else { $fileName }
                        ServerRelativeUrl = $fullPath
                        FolderPath = $FolderName
                        PageId = $item.Id
                    }
                    
                    $pagesToTranslate += $pageObj
                    $processedPages[$fileName] = $true
                }
            }
            
            if ($pagesToTranslate.Count -eq 0) {
                Write-Host "No .aspx pages found in the '$FolderName' folder (or all pages already have French translations)" -ForegroundColor Yellow
                exit 0
            }
            
            Write-Host "Found $($pagesToTranslate.Count) pages to translate in folder '$FolderName'" -ForegroundColor Green
            
            # Display pages to be translated
            Write-Host "`nPages to translate to French (LCID: $FrenchLCID):" -ForegroundColor Cyan
            $pagesToTranslate | ForEach-Object { Write-Host "- $($_.Name)" -ForegroundColor White }
            
            # Confirm before proceeding
            $confirm = Read-Host "`nProceed with creating French translations for all pages in '$FolderName'? (y/N)"
            if ($confirm.ToLower() -ne 'y') {
                Write-Host "Operation cancelled" -ForegroundColor Yellow
                exit 0
            }
            
            # Create French translations for each page
            Write-Host "`nStarting translation process..." -ForegroundColor Yellow
            $successCount = 0
            $failCount = 0
            $skipCount = 0
            $counter = 0
            
            foreach ($page in $pagesToTranslate) {
                $counter++
                $startTime = Get-Date
                
                Write-Progress -Activity "Creating French Translations" -Status "Processing page $counter of $($pagesToTranslate.Count): $($page.Name)" -PercentComplete (($counter / $pagesToTranslate.Count) * 100)
                
                $logEntry = [PSCustomObject]@{
                    PageName = $page.Name
                    PageTitle = $page.Title
                    PageUrl = $page.ServerRelativeUrl
                    FolderPath = $page.FolderPath
                    FrenchLCID = $FrenchLCID
                    Status = ""
                    ErrorMessage = ""
                    ProcessedDateTime = $startTime.ToString("yyyy-MM-dd HH:mm:ss")
                    ProcessingTimeSeconds = 0
                }
                
                try {
                    Write-Host "`n[$counter/$($pagesToTranslate.Count)] Processing: $($page.Name)" -ForegroundColor Yellow
                    
                    # Extract the relative path without the leading site path for subfolder pages
                    # Convert from: /sites/devintranet/SitePages/Guidelines/PageName.aspx
                    # To: Guidelines/PageName (without extension and leading paths)
                    $pagePath = $page.ServerRelativeUrl
                    $sitePagesPart = "/SitePages/"
                    
                    if ($pagePath.Contains($sitePagesPart)) {
                        # Get everything after /SitePages/
                        $relativePath = $pagePath.Substring($pagePath.IndexOf($sitePagesPart) + $sitePagesPart.Length)
                        # Remove .aspx extension
                        $pageIdentity = $relativePath -replace "\.aspx$", ""
                    } else {
                        # Fallback to just the page name
                        $pageIdentity = $page.Name -replace "\.aspx$", ""
                    }
                    
                    Write-Host "Using page identity: '$pageIdentity'" -ForegroundColor Cyan
                    
                    # Check if French translation already exists by examining the _SPTranslatedLanguages field
                    $translationExists = $false
                    $skipReason = ""
                    
                    try {
                        # Get the list item for this page to check translation fields
                        $pageListItem = Get-PnPListItem -List $Library -Id $page.PageId
                        
                        if ($pageListItem) {
                            # Check the _SPTranslatedLanguages field
                            $translatedLanguages = $pageListItem.FieldValues["_SPTranslatedLanguages"]
                            
                            if ($translatedLanguages) {
                                Write-Host "Found translated languages data: $translatedLanguages" -ForegroundColor Cyan
                                
                                # Check if French (fr-fr or fr) is already in the translated languages
                                if ($translatedLanguages -match "fr-fr" -or $translatedLanguages -match '"fr"') {
                                    $translationExists = $true
                                    $skipReason = "French translation already exists (found in _SPTranslatedLanguages field)"
                                }
                            } else {
                                Write-Host "No translated languages found for this page" -ForegroundColor Cyan
                            }
                            
                            # Also check if this page is marked as a translation itself
                            $isTranslation = $pageListItem.FieldValues["_SPIsTranslation"]
                            $translationLanguage = $pageListItem.FieldValues["_SPTranslationLanguage"]
                            
                            if ($isTranslation -eq $true -and $translationLanguage -match "fr") {
                                $translationExists = $true
                                $skipReason = "This page is already a French translation of another page"
                            }
                        }
                    }
                    catch {
                        Write-Host "Could not check translation status, proceeding with creation attempt..." -ForegroundColor Yellow
                        # Continue with translation attempt if we can't check the fields
                    }
                    
                    if ($translationExists) {
                        Write-Host "⚠ SKIPPED: $skipReason" -ForegroundColor Yellow
                        
                        $logEntry.Status = "Skipped"
                        $logEntry.ErrorMessage = $skipReason
                        $logEntry.ProcessingTimeSeconds = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
                        $skipCount++
                    } else {
                        # No French translation exists, proceed with creation
                        Write-Host "Creating French translation..." -ForegroundColor Cyan
                        Set-PnPPage -Identity $pageIdentity -Translate -TranslationLanguageCodes $FrenchLCID
                        
                        $endTime = Get-Date
                        $processingTime = ($endTime - $startTime).TotalSeconds
                        
                        $logEntry.Status = "Success"
                        $logEntry.ProcessingTimeSeconds = [math]::Round($processingTime, 2)
                        
                        Write-Host "✓ French translation variant created for: $($page.Name)" -ForegroundColor Green
                        $successCount++
                    }
                }
                catch {
                    $endTime = Get-Date
                    $processingTime = ($endTime - $startTime).TotalSeconds
                    
                    $logEntry.Status = "Failed"
                    $logEntry.ErrorMessage = $_.Exception.Message
                    $logEntry.ProcessingTimeSeconds = [math]::Round($processingTime, 2)
                    
                    Write-Host "✗ Failed to create French translation for $($page.Name): $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
                
                # Add log entry to array
                $logEntries += $logEntry
            }
            
            # Clear progress bar
            Write-Progress -Activity "Creating French Translations" -Completed
            
            # Export to CSV
            Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
            try {
                $logEntries | Export-Csv -Path $csvLogFile -NoTypeInformation -Encoding UTF8
                Write-Host "✓ CSV log file created: $csvLogFile" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to create CSV log file: $($_.Exception.Message)"
            }
            
            # Summary
            Write-Host "`n" + "="*70 -ForegroundColor Cyan
            Write-Host "FRENCH TRANSLATION SUMMARY - GUIDELINES FOLDER" -ForegroundColor Cyan
            Write-Host "="*70 -ForegroundColor Cyan
            Write-Host "Folder processed: $FolderName" -ForegroundColor White
            Write-Host "Total pages found: $($pagesToTranslate.Count)" -ForegroundColor White
            Write-Host "Successfully translated: $successCount pages" -ForegroundColor Green
            Write-Host "Skipped (already translated): $skipCount pages" -ForegroundColor Yellow
            Write-Host "Failed translations: $failCount pages" -ForegroundColor Red
            Write-Host "Processing completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
            
            if ($successCount -gt 0) {
                Write-Host "`n✓ French page variants have been created with proper translation relationships" -ForegroundColor Green
                Write-Host "✓ Language switcher should now be available on translated pages" -ForegroundColor Green
                Write-Host "✓ You can now add French content to the translated pages" -ForegroundColor Green
            }
            
            # Show CSV log file location
            Write-Host "`n📋 Detailed log available at: $csvLogFile" -ForegroundColor Magenta
            
            # Show CSV content summary
            Write-Host "`nCSV Log Summary:" -ForegroundColor Cyan
            Write-Host "- Total Pages Processed: $($logEntries.Count)" -ForegroundColor White
            Write-Host "- Successful: $(($logEntries | Where-Object {$_.Status -eq 'Success'}).Count)" -ForegroundColor Green
            Write-Host "- Skipped: $(($logEntries | Where-Object {$_.Status -eq 'Skipped'}).Count)" -ForegroundColor Yellow
            Write-Host "- Failed: $(($logEntries | Where-Object {$_.Status -eq 'Failed'}).Count)" -ForegroundColor Red
            
            if ($logEntries | Where-Object {$_.Status -eq 'Failed'}) {
                Write-Host "`nFailed Pages:" -ForegroundColor Red
                $failedEntries = $logEntries | Where-Object {$_.Status -eq 'Failed'}
                $failedEntries | ForEach-Object {
                    Write-Host "  - $($_.PageName): $($_.ErrorMessage)" -ForegroundColor Yellow
                }
            }
            
            if ($logEntries | Where-Object {$_.Status -eq 'Skipped'}) {
                Write-Host "`nSkipped Pages:" -ForegroundColor Yellow
                $skippedEntries = $logEntries | Where-Object {$_.Status -eq 'Skipped'}
                $skippedEntries | ForEach-Object {
                    Write-Host "  - $($_.PageName): $($_.ErrorMessage)" -ForegroundColor Cyan
                }
            }
            
            Write-Host "`nScript completed!" -ForegroundColor Cyan
            Write-Host "You can now navigate to the folder in SharePoint to see the French variants." -ForegroundColor Green
            
            # Disconnect
            Disconnect-PnPOnline
