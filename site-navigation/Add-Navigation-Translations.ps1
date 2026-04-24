# Script: Add-Navigation-Translations.ps1
# Description: Adds SharePoint language to mega menu navigation
# Full documentation: https://www.camerongriffiths.com/scripts/addnavigationtranslations
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Minimal SharePoint Navigation Translation Script  
# Handles multi-level navigation using Get-PnPNavigationNode method

# Site URL and Client ID
$SiteURL = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId = ""

# Translation mappings (English -> French)
$translations = @{
    # Level 1 nodes
    "Organisation" = "Organisation"
    "Life as a staff member" = "La vie d'un membre du personnel"
    # Level 2+ nodes (sub-navigation)
    "Diversity & employee resource groups" = "Diversité et groupes de ressources pour les employés"
    "Regulatory Framework" = "Cadre réglementaire"
    "Events" = "Événements"
}

# Setup error logging
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$errorLogPath = Join-Path $scriptDirectory "SharePoint_Navigation_Translation_Errors.csv"
$errorLog = @()

function Write-ErrorLog {
    param(
        [string]$SiteUrl,
        [string]$NodeTitle,
        [string]$ErrorMessage,
        [string]$ErrorType
    )
    
    $errorEntry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SiteUrl = $SiteUrl
        NodeTitle = $NodeTitle
        ErrorType = $ErrorType
        ErrorMessage = $ErrorMessage
    }
    
    $script:errorLog += $errorEntry
    Write-Host "⚠️  Error logged: $ErrorType - $ErrorMessage" -ForegroundColor Red
}

function Get-AllNavigationNodesRecursive {
    param(
        [Parameter(Mandatory=$true)]
        $NavigationNodeCollection,
        [string]$Level = "1",
        [string]$ParentTitle = "Root"
    )
    
    $allNodes = @()
    
    foreach ($node in $NavigationNodeCollection) {
        # Add current node
        $nodeInfo = @{
            Node = $node
            Level = $Level
            ParentTitle = $ParentTitle
            FullPath = if ($ParentTitle -eq "Root") { $node.Title } else { "$ParentTitle > $($node.Title)" }
        }
        $allNodes += $nodeInfo
        
        Write-Host "  $("  " * ([int]$Level - 1))Level $Level`: '$($node.Title)'" -ForegroundColor Gray
        
        # Get children using Get-PnPNavigationNode (the method that works)
        try {
            $childNodes = Get-PnPNavigationNode -Id $node.Id
            if ($childNodes.Children -and $childNodes.Children.Count -gt 0) {
                Write-Host "  $("  " * ([int]$Level - 1))  -> Found $($childNodes.Children.Count) children" -ForegroundColor DarkGreen
                $recursiveChildNodes = Get-AllNavigationNodesRecursive -NavigationNodeCollection $childNodes.Children -Level ([int]$Level + 1).ToString() -ParentTitle $node.Title
                $allNodes += $recursiveChildNodes
            }
        }
        catch {
            Write-Host "  $("  " * ([int]$Level - 1))  -> Could not load children: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-ErrorLog -SiteUrl $SiteURL -NodeTitle $node.Title -ErrorType "Child Loading" -ErrorMessage "Could not load child nodes: $($_.Exception.Message)"
        }
    }
    
    return $allNodes
}

try {
    # Connect to SharePoint
    Write-Host "Connecting to: $SiteURL" -ForegroundColor Green
    try {
        
        # Interactive authentication
        Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $ClientId
    }
    catch {
        Write-ErrorLog -SiteUrl $SiteURL -NodeTitle "N/A" -ErrorType "Connection" -ErrorMessage $_.Exception.Message
        throw
    }
    
    # Get CSOM context and load navigation
    try {
        $context = Get-PnPContext
        $web = $context.Web
        $context.Load($web.Navigation)
        $context.Load($web.Navigation.TopNavigationBar)
        $context.Load($web.Navigation.QuickLaunch)
        $context.ExecuteQuery()
        
        Write-Host "Found navigation collections:" -ForegroundColor Yellow
        Write-Host "  - TopNavigationBar: $($web.Navigation.TopNavigationBar.Count)" -ForegroundColor Gray
        Write-Host "  - QuickLaunch: $($web.Navigation.QuickLaunch.Count)" -ForegroundColor Gray
    }
    catch {
        Write-ErrorLog -SiteUrl $SiteURL -NodeTitle "N/A" -ErrorType "Navigation Loading" -ErrorMessage $_.Exception.Message
        throw
    }
    
    # Collect all navigation nodes recursively (including children)
    Write-Host "`nScanning navigation structure:" -ForegroundColor Yellow
    $allNodes = @()
    
    # Get all TopNavigationBar nodes and their children
    if ($web.Navigation.TopNavigationBar.Count -gt 0) {
        Write-Host "TopNavigationBar nodes:" -ForegroundColor Cyan
        $topNavNodes = Get-AllNavigationNodesRecursive -NavigationNodeCollection $web.Navigation.TopNavigationBar
        $allNodes += $topNavNodes
    }
    
    # Get all QuickLaunch nodes and their children  
    if ($web.Navigation.QuickLaunch.Count -gt 0) {
        Write-Host "QuickLaunch nodes:" -ForegroundColor Cyan
        $quickLaunchNodes = Get-AllNavigationNodesRecursive -NavigationNodeCollection $web.Navigation.QuickLaunch
        $allNodes += $quickLaunchNodes
    }
    
    Write-Host "`nTotal navigation nodes found (all levels): $($allNodes.Count)" -ForegroundColor Green

    # Apply translations
    Write-Host "`nApplying translations:" -ForegroundColor Green
    $successCount = 0
    foreach ($englishTitle in $translations.Keys) {
        $frenchTitle = $translations[$englishTitle]
        $targetNodeInfo = $allNodes | Where-Object { $_.Node.Title -eq $englishTitle }
        
        if ($targetNodeInfo) {
            try {
                $nodeLevel = $targetNodeInfo.Level
                $nodePath = $targetNodeInfo.FullPath
                Write-Host "Setting (Level $nodeLevel): '$nodePath' -> '$frenchTitle'" -ForegroundColor Cyan
                
                $targetNodeInfo.Node.TitleResource.SetValueForUICulture('fr-FR', $frenchTitle)
                $targetNodeInfo.Node.Update()
                $successCount++
            }
            catch {
                Write-ErrorLog -SiteUrl $SiteURL -NodeTitle $englishTitle -ErrorType "Translation Setting" -ErrorMessage $_.Exception.Message
            }
        } else {
            Write-Host "Node '$englishTitle' not found in navigation structure" -ForegroundColor Yellow
            Write-ErrorLog -SiteUrl $SiteURL -NodeTitle $englishTitle -ErrorType "Node Not Found" -ErrorMessage "Navigation node with title '$englishTitle' was not found in the site navigation (searched all levels)"
        }
    }
    
    # Execute all changes
    if ($successCount -gt 0) {
        try {
            Write-Host "Applying $successCount translations..." -ForegroundColor Green
            $context.ExecuteQuery()
            Write-Host "✓ Success! All translations applied." -ForegroundColor Green
        }
        catch {
            Write-ErrorLog -SiteUrl $SiteURL -NodeTitle "Multiple" -ErrorType "ExecuteQuery" -ErrorMessage $_.Exception.Message
            throw
        }
    } else {
        Write-Host "No translations applied - no matching nodes found." -ForegroundColor Red
        Write-ErrorLog -SiteUrl $SiteURL -NodeTitle "N/A" -ErrorType "No Matches" -ErrorMessage "No navigation nodes matched the translation keys provided"
    }
}
catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($errorLog.Count -eq 0) {
        Write-ErrorLog -SiteUrl $SiteURL -NodeTitle "N/A" -ErrorType "Critical Error" -ErrorMessage $_.Exception.Message
    }
}
finally {
    # Export error log to CSV if there are any errors
    if ($errorLog.Count -gt 0) {
        try {
            $errorLog | Export-Csv -Path $errorLogPath -NoTypeInformation -Encoding UTF8 -Force
            Write-Host "`n📄 Error log exported to: $errorLogPath" -ForegroundColor Yellow
            Write-Host "   $($errorLog.Count) error(s) logged" -ForegroundColor Yellow
        }
        catch {
            Write-Host "`n⚠️  Failed to export error log: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "`n✅ No errors occurred - no error log created" -ForegroundColor Green
    }
    
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline
    }
    catch {
        # Ignore disconnect errors
    }
    
    Write-Host "Done." -ForegroundColor Gray
}
