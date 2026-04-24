# Script: Export-Megamenu-Navigation.ps1
# Description: Exports current values of the mega menu
# Full documentation: https://www.camerongriffiths.com/scripts/exportmegamenunavigation
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Configuration Variables
$SiteURL = "https://tenantName.sharepoint.com/sites/siteName"  # Set your SharePoint site URL here
$ClientId = ""  # Client ID here
$LogFile = "MegaMenu_Export_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$ExportFile = "MegaMenu_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Logging Function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Level = 'Info'
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = [PSCustomObject]@{
        Timestamp = $Timestamp
        Level = $Level
        Message = $Message
    }

    # Console output with color
    $Color = switch($Level) {
        'Info' { 'White' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
    }
    Write-Host "[$Timestamp] [$Level] $Message" -ForegroundColor $Color

    # Export to log file
    $LogEntry | Export-Csv -Path $LogFile -Append -NoTypeInformation
}

# Validate parameters
if ([string]::IsNullOrWhiteSpace($SiteURL)) {
    Write-Log "SiteURL is required. Please set the variable." -Level Error
    exit
}
if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Log "ClientId is required. Please set the variable." -Level Error
    exit
}

Write-Log "=== Starting MegaMenu Navigation Export ===" -Level Info
Write-Log "Site URL: $SiteURL" -Level Info

try {
    # Connect to SharePoint Online
    Write-Log "Attempting to connect to SharePoint Online..." -Level Info

    Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $ClientId -ErrorAction Stop

    Write-Log "Successfully connected to SharePoint Online" -Level Success

    # Initialize results array
    $NavigationResults = @()

    # Get MegaMenu navigation using REST API
    Write-Log "Retrieving MegaMenu navigation..." -Level Info

    try {
        # Use the navigation MenuState endpoint which contains MegaMenu data
        $restUrl = "/_api/navigation/MenuState"
        $menuState = Invoke-PnPSPRestMethod -Url $restUrl -Method Get

        Write-Log "Successfully retrieved MenuState data" -Level Success

        if ($menuState.Nodes -and $menuState.Nodes.Count -gt 0) {
            Write-Log "Found $($menuState.Nodes.Count) top-level navigation nodes" -Level Success

            foreach ($node in $menuState.Nodes) {
                # Add parent node
                $NavItem = [PSCustomObject]@{
                    Level = "Parent"
                    ParentTitle = ""
                    Title = $node.Title
                    Url = $node.SimpleUrl
                    Key = $node.Key
                    FriendlyUrlSegment = $node.FriendlyUrlSegment
                    IsExternal = $node.OpenInNewWindow
                    IsHidden = $node.IsHidden
                    IsVisible = -not $node.IsHidden
                }
                $NavigationResults += $NavItem
                Write-Log "Exported Parent: $($node.Title) - $($node.SimpleUrl)" -Level Info

                # Check for child nodes
                if ($node.Nodes -and $node.Nodes.Count -gt 0) {
                    Write-Log "Found $($node.Nodes.Count) children for: $($node.Title)" -Level Info

                    foreach ($child in $node.Nodes) {
                        $ChildItem = [PSCustomObject]@{
                            Level = "Child"
                            ParentTitle = $node.Title
                            Title = $child.Title
                            Url = $child.SimpleUrl
                            Key = $child.Key
                            FriendlyUrlSegment = $child.FriendlyUrlSegment
                            IsExternal = $child.OpenInNewWindow
                            IsHidden = $child.IsHidden
                            IsVisible = -not $child.IsHidden
                        }
                        $NavigationResults += $ChildItem
                        Write-Log "Exported Child: $($child.Title) (Parent: $($node.Title)) - $($child.SimpleUrl)" -Level Info

                        # Check for grandchildren (3rd level)
                        if ($child.Nodes -and $child.Nodes.Count -gt 0) {
                            Write-Log "Found $($child.Nodes.Count) grandchildren for: $($child.Title)" -Level Info

                            foreach ($grandchild in $child.Nodes) {
                                $GrandchildItem = [PSCustomObject]@{
                                    Level = "Grandchild"
                                    ParentTitle = $child.Title
                                    Title = $grandchild.Title
                                    Url = $grandchild.SimpleUrl
                                    Key = $grandchild.Key
                                    FriendlyUrlSegment = $grandchild.FriendlyUrlSegment
                                    IsExternal = $grandchild.OpenInNewWindow
                                    IsHidden = $grandchild.IsHidden
                                    IsVisible = -not $grandchild.IsHidden
                                }
                                $NavigationResults += $GrandchildItem
                                Write-Log "Exported Grandchild: $($grandchild.Title) (Parent: $($child.Title)) - $($grandchild.SimpleUrl)" -Level Info
                            }
                        }
                    }
                }
            }

            Write-Log "MegaMenu navigation export completed" -Level Success
        }
        else {
            Write-Log "No navigation nodes found in MenuState" -Level Warning
        }
    }
    catch {
        Write-Log "Error retrieving MegaMenu navigation: $($_.Exception.Message)" -Level Error
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Level Error
    }

    # Export to CSV
    if ($NavigationResults.Count -gt 0) {
        Write-Log "Exporting $($NavigationResults.Count) navigation items to CSV..." -Level Info
        $NavigationResults | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
        Write-Log "Export completed successfully: $ExportFile" -Level Success

        # Display summary
        Write-Log "`n=== Navigation Summary ===" -Level Info
        $parentCount = ($NavigationResults | Where-Object { $_.Level -eq "Parent" }).Count
        $childCount = ($NavigationResults | Where-Object { $_.Level -eq "Child" }).Count
        $grandchildCount = ($NavigationResults | Where-Object { $_.Level -eq "Grandchild" }).Count

        Write-Log "Parent Nodes: $parentCount" -Level Success
        Write-Log "Child Nodes: $childCount" -Level Success
        Write-Log "Grandchild Nodes: $grandchildCount" -Level Success
        Write-Log "Total Items: $($NavigationResults.Count)" -Level Success

        # Show first few items as preview
        Write-Log "`nPreview of exported items:" -Level Info
        $NavigationResults | Select-Object -First 10 | Format-Table Level, ParentTitle, Title, Url -AutoSize
    }
    else {
        Write-Log "No navigation items found to export" -Level Warning
    }

    # Summary
    Write-Log "n=== Export Summary ===" -Level Info
    Write-Log "Export File: $ExportFile" -Level Info
    Write-Log "Log File: $LogFile" -Level Info
}
catch {
    Write-Log "Critical error during navigation export: $($_.Exception.Message)" -Level Error
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Level Error
}
finally {
    # Disconnect from SharePoint
    try {
        Write-Log "Disconnecting from SharePoint Online..." -Level Info
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Log "Disconnected successfully" -Level Success
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level Warning
    }
}

Write-Log "n=== MegaMenu Navigation Export Process Completed ===" -Level Info
