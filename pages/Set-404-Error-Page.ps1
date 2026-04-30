# Set your Client ID (optional - leave empty for interactive authentication)
$ClientId = ""

# Connect to your SharePoint Online site
Connect-PnPOnline -Url "https://tenantName.sharepoint.com/sites/siteName" -Interactive -ClientId $ClientId

# Set the property bag value to configure the custom 404 error page
Set-PnPPropertyBagValue -Key "vti_filenotfoundpage" -Value "/sites/siteName/SitePages/pageName.aspx"
          
