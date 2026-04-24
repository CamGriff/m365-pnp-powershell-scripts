# Script: move-selected-pages-in-a-folder.ps1
# Description: allows you to move slected pages to another folder
# Full documentation: https://www.camerongriffiths.com/scripts/movesharepointpages
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions




# Configuration
            $ClientId = ""
            $SiteUrl = "https://tenantName.sharepoint.com/sites/siteName"
            $SourceFolder = "SitePages/SourceFolder"
            $TargetFolder = "SitePages/TargetFolder"
            $ErrorLogPath = ""
            
            # Array of files to move
            $FilesToMove = @(
                "fileName.aspx",
                "fileName1.aspx",
                "fileName2.aspx"
            )
            
            # Initialize error log array
            $ErrorLog = @()
            
            # Connect to SharePoint
            try {
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
                Write-Host "✓ Connected to SharePoint" -ForegroundColor Green
                Write-Host ""
            }
            catch {
                Write-Host "✗ Failed to connect to SharePoint: $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
            
            Write-Host "Moving $($FilesToMove.Count) pages..." -ForegroundColor Yellow
            Write-Host ""
            
            # Loop through each file and move it
            foreach ($FileName in $FilesToMove) {
                Write-Host "Moving: $FileName" -ForegroundColor Cyan
                
                try {
                    Move-PnPFile -SourceUrl "$SourceFolder/$FileName" -TargetUrl "$TargetFolder/$FileName" -Force -OverwriteIfAlreadyExists
                    Write-Host "  ✓ Moved successfully" -ForegroundColor Green
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-Host "  ✗ Error: $ErrorMessage" -ForegroundColor Red
                    
                    # Log error to array
                    $ErrorLog += [PSCustomObject]@{
                        FileName = $FileName
                        SourceFolder = $SourceFolder
                        TargetFolder = $TargetFolder
                        ErrorMessage = $ErrorMessage
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                Write-Host ""
            }
            
            # Export errors to CSV if any occurred
            if ($ErrorLog.Count -gt 0) {
                $ErrorLog | Export-Csv -Path $ErrorLogPath -NoTypeInformation -Encoding UTF8
                Write-Host "✗ $($ErrorLog.Count) errors occurred. Log saved to: $ErrorLogPath" -ForegroundColor Red
            }
            else {
                Write-Host "✓ All pages moved successfully - no errors!" -ForegroundColor Green
            }
            
            Write-Host ""
            Write-Host "Move complete!" -ForegroundColor Yellow
