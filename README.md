# M365 PnP PowerShell Scripts
 
A collection of validated PnP PowerShell scripts for Microsoft 365 and SharePoint Online, built and tested on enterprise deployments.
 
Full documentation, usage guides and examples for each script at
👉 [camerongriffiths.com](https://www.camerongriffiths.com/en/resources)
 
---
 
## Scripts
 
### Site Columns
- [Create-Site-Columns.ps1](./site-columns/Create-Site-Columns.ps1) — Bulk create SharePoint site columns from JSON configuration | [Full docs](https://www.camerongriffiths.com/scripts/createsitecolumns)
### Content Types
- [Create-Content-Types.ps1](./content-types/Create-Content-Types.ps1) — Create SharePoint content types programmatically from JSON configuration | [Full docs](https://www.camerongriffiths.com/scripts/createcontenttypes)
- [Update-Content-Type-Name.ps1](./content-types/Update-Content-Type-Name.ps1) — Rename an existing SharePoint content type across a site collection | [Full docs](https://www.camerongriffiths.com/scripts/updatecontenttypename)
- [Update-Pages-Content-Type-In-Folder.ps1](./content-types/Update-Pages-Content-Type-In-Folder.ps1) — Bulk update the content type of pages within a specified folder | [Full docs](https://www.camerongriffiths.com/scripts/updatepagescontenttypeinafolder)
### Navigation
- [Create-Navigation.ps1](./site-navigation/Create-Navigation.ps1) — Create SharePoint site navigation nodes programmatically | [Full docs](https://www.camerongriffiths.com/scripts/createnavigation)
- [Add-Navigation-Translations.ps1](./site-navigation/Add-Navigation-Translations.ps1) — Add multilingual translations to SharePoint navigation nodes | [Full docs](https://www.camerongriffiths.com/scripts/addnavigationtranslations)
- [Export-Megamenu-Navigation.ps1](./site-navigation/Export-Megamenu-Navigation.ps1) — Export SharePoint megamenu navigation structure to JSON | [Full docs](https://www.camerongriffiths.com/scripts/exportMegamenuNavigation)
- [Site-Navigation.ps1](./site-navigation/Site-Navigation.ps1) — Manage and update SharePoint Online site navigation programmatically | [Full docs](https://www.camerongriffiths.com/scripts/sitenavigation)
### Pages
- [Export-Multiple-Pages.ps1](./pages/Export-Multiple-Pages.ps1) — Export multiple SharePoint modern pages using PnP provisioning templates | [Full docs](https://www.camerongriffiths.com/scripts/exportmultiplepages)
- [Import-Multiple-Pages.ps1](./pages/Import-Multiple-Pages.ps1) — Import multiple SharePoint modern pages from PnP provisioning templates | [Full docs](https://www.camerongriffiths.com/scripts/importmultiplepages)
- [Move-SharePoint-Pages.ps1](./pages/Move-SharePoint-Pages.ps1) — Move SharePoint pages between folders or libraries with metadata preservation | [Full docs](https://www.camerongriffiths.com/scripts/movesharepointpages)
- [Create-Translations-Of-Pages-In-Folder.ps1](./pages/Create-Translations-Of-Pages-In-Folder.ps1) — Bulk create multilingual translation pages for all pages within a folder | [Full docs](https://www.camerongriffiths.com/scripts/createtranslationsofpagesinfolder)
- [Update-Pages-Approval-Status-In-Folder.ps1](./pages/Update-Pages-Approval-Status-In-Folder.ps1) — Bulk update the approval status of pages within a specified folder | [Full docs](https://www.camerongriffiths.com/scripts/updatepagesapprovalstatusinfolder)
- [Set-404-Error-Page.ps1](./pages/Set-404-Error-Page.ps1) — Configure a custom 404 error page for a SharePoint site | [Full docs](https://www.camerongriffiths.com/scripts/set404errorpage)
### Lists & Libraries
- [Create-List.ps1](./lists-libraries/Create-List.ps1) — Create a SharePoint list with custom columns from JSON configuration | [Full docs](https://www.camerongriffiths.com/scripts/createlist)
- [Create-Folder-Structure.ps1](./lists-libraries/Create-Folder-Structure.ps1) — Bulk create a folder structure in a SharePoint library from JSON | [Full docs](https://www.camerongriffiths.com/scripts/createfolderstructure)
- [Update-Folder-Approval-Status.ps1](./lists-libraries/Update-Folder-Approval-Status.ps1) — Update the approval status of a folder in a SharePoint library | [Full docs](https://www.camerongriffiths.com/scripts/updatefolderapprovalstatus)
- [Update-Folder-Approval-Statuses.ps1](./lists-libraries/Update-Folder-Approval-Statuses.ps1) — Bulk update approval statuses across multiple folders in a library | [Full docs](https://www.camerongriffiths.com/scripts/updatefolderapprovalstatuses)
- [Update-URLs-In-An-Excel-Column.ps1](./lists-libraries/Update-URLs-In-An-Excel-Column.ps1) — Bulk update URL column values in SharePoint lists using data from Excel | [Full docs](https://www.camerongriffiths.com/scripts/updateurlsinanexcelcolumn)
- [Find-Column-Usage.ps1](./lists-libraries/Find-Column-Usage.ps1) — Identify where a specific site column is used across lists and libraries | [Full docs](https://www.camerongriffiths.com/scripts/findColumnUsage)
### Groups & Permissions
- [Add-Group-Members.ps1](./groups-permissions/Add-Group-Members.ps1) — Add members to a SharePoint group from a CSV file | [Full docs](https://www.camerongriffiths.com/scripts/addgroupmembers)
- [Create-Group-And-Add-Group-Members.ps1](./groups-permissions/Create-Group-And-Add-Group-Members.ps1) — Create a SharePoint group and populate it with members in one operation | [Full docs](https://www.camerongriffiths.com/scripts/creategroup&addgroupmembers)
- [Create-Group-And-Add-Users.ps1](./groups-permissions/Create-Group-And-Add-Users.ps1) — Create a SharePoint group and add individual users directly | [Full docs](https://www.camerongriffiths.com/scripts/creategroup&addusers)
### Taxonomy & Term Store
- [Create-Term-Store.ps1](./taxonomy/Create-Term-Store.ps1) — Create term groups, term sets and terms in the SharePoint term store from JSON | [Full docs](https://www.camerongriffiths.com/scripts/createtermstore)
- [Delete-Term-Group.ps1](./taxonomy/Delete-Term-Group.ps1) — Delete a term group and all its term sets from the SharePoint term store | [Full docs](https://www.camerongriffiths.com/scripts/deletetermgroup)
### Site Templates
- [Site-Template-Export.ps1](./site-templates/Site-Template-Export.ps1) — Export a SharePoint site as a PnP provisioning template | [Full docs](https://www.camerongriffiths.com/scripts/sitetemplateexport)
- [Site-Template-Apply.ps1](./site-templates/Site-Template-Apply.ps1) — Apply a PnP provisioning template to a SharePoint site | [Full docs](https://www.camerongriffiths.com/scripts/sitetemplateapply)
### Search & Crawl
- [Get-Site-Crawl-Logs.ps1](./search/Get-Site-Crawl-Logs.ps1) — Retrieve and export SharePoint search crawl logs for a site collection | [Full docs](https://www.camerongriffiths.com/scripts/getsitecrawllogs)
- [Get-Site-Search-Settings.ps1](./search/Get-Site-Search-Settings.ps1) — Export search configuration settings from a SharePoint site | [Full docs](https://www.camerongriffiths.com/scripts/getsitesearchsettings)
### Site Analysis & Utilities
- [Compare-Two-Site-Settings.ps1](./utilities/Compare-Two-Site-Settings.ps1) — Compare configuration settings between two SharePoint site collections | [Full docs](https://www.camerongriffiths.com/scripts/comparetwositesettings)
- [Find-All-Links-On-Site-Pages.ps1](./utilities/Find-All-Links-On-Site-Pages.ps1) — Extract and export all hyperlinks found across SharePoint site pages | [Full docs](https://www.camerongriffiths.com/scripts/findalllinkssitepages)
- [Find-Spage-Path-Lengths.ps1](./utilities/Find-Spage-Path-Lengths.ps1) — Identify SharePoint pages with URL paths approaching the 260 character limit | [Full docs](https://www.camerongriffiths.com/scripts/findspagepathlengths)
- [Folder-Path-Length.ps1](./utilities/Folder-Path-Length.ps1) — Check folder path lengths in a SharePoint library for migration compatibility | [Full docs](https://www.camerongriffiths.com/scripts/folderPathLength)
- [Site-Settings-Export.ps1](./utilities/Site-Settings-Export.ps1) — Export key SharePoint site settings and configuration to CSV | [Full docs](https://www.camerongriffiths.com/scripts/siteSettingsExport)
---
 
## Requirements
 
- PnP.PowerShell module (`Install-Module PnP.PowerShell`)
- PowerShell 7+
- Site Collection Administrator or higher
- Microsoft 365 tenant with SharePoint Online
- PnP Azure AD app registration with appropriate permissions
---
 
## About
 
Scripts by [Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain. Specialising in SharePoint Online, PnP PowerShell, SPFx and Power Platform.
 
Each script on this repo has a corresponding documentation page at [camerongriffiths.com](https://www.camerongriffiths.com/en/resources) with full usage guides, prerequisites, parameter explanations and JSON configuration examples.
