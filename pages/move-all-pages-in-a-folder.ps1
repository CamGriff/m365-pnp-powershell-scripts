# Script: Export-Multiple-Page.ps1
# Description: allows you to move all pages from one folder to another folder
# Full documentation: https://www.camerongriffiths.com/scripts/movesharepointpages
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Configuration
            $SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
            $SourceFolder = "SitePages/SourceFolder/"
            $TargetFolder = "SitePages/TargetFolder"
            $ClientId = ""
            # Connect to SharePoint
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
            
            # Get all pages from source folder
            $Pages = Get-PnPFolderItem -FolderSiteRelativeUrl $SourceFolder -ItemType File | Where-Object {$_.Name -like "*.aspx"}
            
            Write-Host "Found $($Pages.Count) pages to move" -ForegroundColor Yellow
            Write-Host ""
            
            # Loop through each page and move it
            foreach ($Page in $Pages) {
                Write-Host "Moving: $($Page.Name)" -ForegroundColor Cyan
                
                try {
                    Move-PnPFile -SourceUrl "$SourceFolder/$($Page.Name)" -TargetUrl "$TargetFolder/$($Page.Name)" -Force -OverwriteIfAlreadyExists
                    Write-Host "  ✓ Moved successfully" -ForegroundColor Green
                }
                catch {
                    Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Write-Host ""
            }
            
            Write-Host "Move complete!" -ForegroundColor Yellow
