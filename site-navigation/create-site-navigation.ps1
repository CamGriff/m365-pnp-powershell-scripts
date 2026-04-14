# Script: create-site-navigation.ps1
# Description: Creates SharePoint Online navigation from JSON configuration file
# Full documentation: https://www.camerongriffiths.com/scripts/createnavigation
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# navigation-config.json and this script should be in the same directory
# Configuration
$siteUrl = "https://tenantName.sharepoint.com/sites/siteName" # You only need to specify the full site URL here
$ClientId = ""
$jsonFilePath = "createNavigation.json"
$logFilePath = ".\navigation-creation-log.csv"
$logEntries = @()

# Function to write to log file
function Write-ToLog {
    param(
        [string]$Action,
        [string]$NavigationNode,
        [string]$Status,
        [string]$Message
    )
    
    $logEntry = [PSCustomObject]@{
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Action = $Action
        NavigationNode = $NavigationNode
        Status = $Status
        Message = $Message
    }
    
    $logEntries += $logEntry
}

# Function to safely create navigation node
function Add-SafeNavigationNode {
    param(
        [string]$Title,
        [string]$Location = "QuickLaunch",
        [string]$Url,
        [int]$Parent
    )

    try {
        if ($Parent -and $Url) {
            $node = Add-PnPNavigationNode -Location $Location -Title $Title -Url $Url -Parent $Parent -ErrorAction Stop
        }
        elseif ($Parent) {
            $node = Add-PnPNavigationNode -Location $Location -Title $Title -Parent $Parent -ErrorAction Stop
        }
        else {
            $node = Add-PnPNavigationNode -Location $Location -Title $Title -ErrorAction Stop
        }
        Write-ToLog -Action "Create" -NavigationNode $Title -Status "Success" -Message "Created navigation node"
        return $node
    }
    catch {
        Write-Host "Warning: Could not create node '$Title' with URL '$Url'. Creating without URL validation." -ForegroundColor Yellow
        Write-ToLog -Action "Create" -NavigationNode $Title -Status "Warning" -Message "Created node without URL validation"
        
        # Fallback: Create node without URL validation
        if ($Parent) {
            return Add-PnPNavigationNode -Location $Location -Title $Title -Parent $Parent
        }
        else {
            return Add-PnPNavigationNode -Location $Location -Title $Title
        }
    }
}

# Function to build full URL from base URL and relative path
function Build-FullUrl {
    param(
        [string]$BaseUrl,
        [string]$RelativePath
    )
    
    # Ensure the base URL doesn't end with a slash and the relative path starts with one
    $BaseUrl = $BaseUrl.TrimEnd('/')
    if (-not [string]::IsNullOrEmpty($RelativePath)) {
        $RelativePath = $RelativePath.TrimStart('/')
        if (-not [string]::IsNullOrEmpty($RelativePath)) {
            return "$BaseUrl/$RelativePath"
        }
    }
    
    return $BaseUrl
}

try {
    # Connect to SharePoint Online
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId
    
    # Verify connection and permissions
    try {
        $site = Get-PnPWeb
        Write-Host "Successfully connected to site: $($site.Title)" -ForegroundColor Green
        
        # Test navigation access
        $testNav = Get-PnPNavigationNode -Location QuickLaunch
        Write-Host "Successfully accessed navigation" -ForegroundColor Green
    } catch {
        Write-Host "Error accessing site or navigation. Please verify permissions." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit
    }
    
    Write-ToLog -Action "Connection" -NavigationNode "N/A" -Status "Success" -Message "Connected to SharePoint Online"

    # Read JSON configuration file
    $navConfig = Get-Content $jsonFilePath -Raw | ConvertFrom-Json
    Write-ToLog -Action "ConfigRead" -NavigationNode "N/A" -Status "Success" -Message "Navigation configuration loaded"

    # Get the base URL from the config
    # Use the site URL from the script configuration as sitebase
    $siteBaseUrl = $siteUrl

    # Clear existing navigation
    Write-Host "Clearing existing navigation..." -ForegroundColor Yellow
    $existingNodes = Get-PnPNavigationNode -Location QuickLaunch
    foreach ($node in $existingNodes) {
        Remove-PnPNavigationNode -Identity $node.Id -Force
        Write-ToLog -Action "Remove" -NavigationNode $node.Title -Status "Success" -Message "Removed existing navigation node"
    }

    # Create Level 1 navigation (Labels)
    foreach ($level1 in $navConfig.level1) {
        Write-Host "Creating Level 1 node: $($level1.title)" -ForegroundColor Green
        
        # Build full URL for level 1 if it has a URL
        $level1Url = $null
        if ($level1.url) {
            $level1Url = Build-FullUrl -BaseUrl $siteBaseUrl -RelativePath $level1.url
        }
        
        $level1Node = Add-SafeNavigationNode -Location QuickLaunch -Title $level1.title -Url $level1Url
        
        # Create Level 2 navigation
        if ($level1Node -and $level1.children) {
            foreach ($level2 in $level1.children) {
                Write-Host "Creating Level 2 node: $($level2.title)" -ForegroundColor Cyan
                
                # Build full URL for level 2
                $level2Url = $null
                if ($level2.url) {
                    $level2Url = Build-FullUrl -BaseUrl $siteBaseUrl -RelativePath $level2.url
                }
                
                $level2Node = Add-SafeNavigationNode -Location QuickLaunch -Title $level2.title -Url $level2Url -Parent $level1Node.Id

                # Create Level 3 navigation
                if ($level2Node -and $level2.children) {
                    foreach ($level3 in $level2.children) {
                        Write-Host "Creating Level 3 node: $($level3.title)" -ForegroundColor White
                        
                        # Build full URL for level 3
                        $level3Url = $null
                        if ($level3.url) {
                            $level3Url = Build-FullUrl -BaseUrl $siteBaseUrl -RelativePath $level3.url
                        }
                        
                        $level3Node = Add-SafeNavigationNode -Location QuickLaunch -Title $level3.title -Url $level3Url -Parent $level2Node.Id
                    }
                }
            }
        }
    }

    # Verify navigation was created
    $finalNodes = Get-PnPNavigationNode -Location QuickLaunch
    Write-Host "Created $($finalNodes.Count) navigation nodes" -ForegroundColor Green

    # Export log to CSV
    $logEntries | Export-Csv -Path $logFilePath -NoTypeInformation
    Write-Host "Navigation creation completed. Check $logFilePath for details." -ForegroundColor Green

} catch {
    Write-ToLog -Action "Error" -NavigationNode "N/A" -Status "Failed" -Message $_.Exception.Message
    $logEntries | Export-Csv -Path $logFilePath -NoTypeInformation
    Write-Host "An error occurred. Check $logFilePath for details." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
} finally {
    Disconnect-PnPOnline
    Write-ToLog -Action "Connection" -NavigationNode "N/A" -Status "Success" -Message "Disconnected from SharePoint Online"
}
  
