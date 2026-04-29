#Parameters
$TenantAdminURL = "https://tenantName-admin.sharepoint.com"
$CSVFilePath = "C:\Path\To\SiteCollectionAdmin.csv"
$ClientId = ¨¨

Try {
    #Connect to Admin Center
    Connect-PnPOnline -Url $TenantAdminURL -Interactive -ClientId $ClientId

    #Get data from the CSV file
    $CSVData = Import-Csv $CSVFilePath

    #Iterate through each row in the CSV
    ForEach($Row in $CSVData)
    {
        Try{
            #Add Site collection Admin
            Set-PnPTenantSite -Url $Row.SiteURL -Owners $Row.SiteCollectionAdmin
            Write-host "Added Site collection Administrator to $($Row.SiteURL)" -f Green
        }
        Catch {
            write-host -f Yellow "`tError Adding Site Collection Admin to $($Row.SiteURL) :" $_.Exception.Message
        }
    }
}
Catch {
    write-host -f Red "`tError:" $_.Exception.Message
}
