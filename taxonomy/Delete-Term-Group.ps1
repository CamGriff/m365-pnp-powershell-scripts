#Config Variables
$AdminCenterURL = "https://tenantName-admin.sharepoint.com"
$ClientId = "" #Your tenant Client ID

#Connect to PnP Online
Connect-PnPOnline -Url $AdminCenterURL -Interactive -ClientId $ClientId

#Delete the Term Group
Remove-PnPTaxonomyItem "Intranet_Taxonomy" -Force
