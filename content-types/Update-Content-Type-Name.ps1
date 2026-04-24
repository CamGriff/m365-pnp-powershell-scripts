# SharePoint Online Content Type Rename Script
            # This script renames a content type at site collection level and all associated list-level copies
            
            param(
                [Parameter(Mandatory=$false)]
                [string]$ClientId = "",
            
                [Parameter(Mandatory=$false)]
                [string]$SiteUrl = "https://tenantName.sharepoint.com/sites/siteName",
                
                [Parameter(Mandatory=$false)]
                [string]$CurrentContentTypeName = "Old Content type name",
                
                [Parameter(Mandatory=$false)]
                [string]$NewContentTypeName = "New Content type name",
                
                [Parameter(Mandatory=$false)]
                [string]$TargetLibrary = "Site Pages"
            )
            
            # Import PnP PowerShell module if not already loaded
            if (!(Get-Module -Name PnP.PowerShell -ListAvailable)) {
                Write-Host "PnP.PowerShell module not found. Please install it using: Install-Module -Name PnP.PowerShell" -ForegroundColor Red
                exit
            }
            
            try {
                # Connect to SharePoint Online
                Write-Host "Connecting to SharePoint Online site: $SiteUrl" -ForegroundColor Yellow
                Write-Host "Target Library: $TargetLibrary" -ForegroundColor Yellow
                Write-Host "Renaming Content Type: '$CurrentContentTypeName' → '$NewContentTypeName'" -ForegroundColor Yellow
                Write-Host ""
                
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
                
                # Step 1: Rename the site collection level content type
                Write-Host "Step 1: Renaming site collection content type '$CurrentContentTypeName' to '$NewContentTypeName'..." -ForegroundColor Green
                
                $siteContentType = Get-PnPContentType -Identity $CurrentContentTypeName -ErrorAction SilentlyContinue
                
                if ($siteContentType) {
                    Set-PnPContentType -Identity $CurrentContentTypeName -Name $NewContentTypeName
                    Write-Host "✓ Site collection content type renamed successfully" -ForegroundColor Green
                } else {
                    Write-Host "✗ Site collection content type '$CurrentContentTypeName' not found" -ForegroundColor Red
                    exit
                }
                
                # Step 2: Get all lists in the site collection (with focus on Site Pages library)
                Write-Host "Step 2: Finding all lists that use this content type (prioritizing '$TargetLibrary' library)..." -ForegroundColor Green
                
                $allLists = Get-PnPList
                $listsWithContentType = @()
                
                # First check the target library specifically
                $targetList = $allLists | Where-Object { $_.Title -eq $TargetLibrary }
                if ($targetList) {
                    Write-Host "  Checking target library: $TargetLibrary..." -ForegroundColor Cyan
                    try {
                        $listContentTypes = Get-PnPContentType -List $TargetLibrary -ErrorAction SilentlyContinue
                        
                        # Check if the content type exists in this list (by old name or new name)
                        $hasOldContentType = $listContentTypes | Where-Object { $_.Name -eq $CurrentContentTypeName }
                        $hasNewContentType = $listContentTypes | Where-Object { $_.Name -eq $NewContentTypeName }
                        
                        if ($hasOldContentType -or $hasNewContentType) {
                            $listsWithContentType += $targetList
                            Write-Host "    ✓ Found content type in target library: $TargetLibrary" -ForegroundColor Green
                        } else {
                            Write-Host "    ℹ Content type not found in target library: $TargetLibrary" -ForegroundColor Blue
                        }
                    }
                    catch {
                        Write-Host "    ✗ Error accessing target library: $TargetLibrary - $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "    ⚠️ Target library '$TargetLibrary' not found!" -ForegroundColor Yellow
                }
                
                # Then check all other lists
                Write-Host "  Checking other lists..." -ForegroundColor Cyan
                
                foreach ($list in $allLists) {
                    # Skip the target library as we already checked it
                    if ($list.Title -eq $TargetLibrary) {
                        continue
                    }
                    
                    try {
                        $listContentTypes = Get-PnPContentType -List $list.Title -ErrorAction SilentlyContinue
                        
                        # Check if the content type exists in this list (by old name or new name)
                        $hasOldContentType = $listContentTypes | Where-Object { $_.Name -eq $CurrentContentTypeName }
                        $hasNewContentType = $listContentTypes | Where-Object { $_.Name -eq $NewContentTypeName }
                        
                        if ($hasOldContentType -or $hasNewContentType) {
                            $listsWithContentType += $list
                            Write-Host "    Found content type in list: $($list.Title)" -ForegroundColor Cyan
                        }
                    }
                    catch {
                        # Skip lists that can't be accessed or don't support content types
                        continue
                    }
                }
                
                # Step 3: Rename content type in each list
                Write-Host "Step 3: Renaming content type in $($listsWithContentType.Count) lists..." -ForegroundColor Green
                
                $successCount = 0
                $errorCount = 0
                
                foreach ($list in $listsWithContentType) {
                    try {
                        Write-Host "  Processing list: $($list.Title)..." -ForegroundColor Yellow
                        
                        # Get the content type from the list
                        $listContentType = Get-PnPContentType -List $list.Title | Where-Object { 
                            $_.Name -eq $CurrentContentTypeName -or $_.Name -eq $NewContentTypeName 
                        }
                        
                        if ($listContentType -and $listContentType.Name -eq $CurrentContentTypeName) {
                            # Rename the content type at list level
                            Set-PnPContentType -Identity $listContentType.Id -List $list.Title -Name $NewContentTypeName
                            Write-Host "    ✓ Renamed in $($list.Title)" -ForegroundColor Green
                            $successCount++
                        }
                        elseif ($listContentType -and $listContentType.Name -eq $NewContentTypeName) {
                            Write-Host "    ℹ Already renamed in $($list.Title)" -ForegroundColor Blue
                            $successCount++
                        }
                    }
                    catch {
                        Write-Host "    ✗ Error processing $($list.Title): $($_.Exception.Message)" -ForegroundColor Red
                        $errorCount++
                    }
                }
                
                # Step 4: Summary
                Write-Host "`nRename Operation Summary:" -ForegroundColor Magenta
                Write-Host "=========================" -ForegroundColor Magenta
                Write-Host "Site Collection Content Type: Renamed successfully" -ForegroundColor Green
                Write-Host "Lists processed successfully: $successCount" -ForegroundColor Green
                Write-Host "Lists with errors: $errorCount" -ForegroundColor $(if($errorCount -gt 0){"Red"}else{"Green"})
                Write-Host "Total lists processed: $($successCount + $errorCount)" -ForegroundColor Cyan
                
                if ($errorCount -eq 0) {
                    Write-Host "`n🎉 Content type '$CurrentContentTypeName' has been successfully renamed to '$NewContentTypeName' across the entire site collection!" -ForegroundColor Green
                } else {
                    Write-Host "`n⚠️ Content type rename completed with some errors. Please review the error messages above." -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
            }
            finally {
                # Disconnect from SharePoint Online
                Disconnect-PnPOnline
                Write-Host "`nDisconnected from SharePoint Online." -ForegroundColor Gray
            }
            
       
