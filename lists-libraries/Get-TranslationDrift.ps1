# ============================================================
# Get-TranslationDrift.ps1
# Detects translation drift between English source pages and
# their French (fr-FR) translations in the Site Pages library.
# Populates the TranslationDrift list for dashboard reporting.
# ============================================================

$SiteUrl      = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId     = ""
$ListName     = "TranslationDrift"
$TargetLocale = "fr-fr"
$StaleDays    = 7  # Days drift before a page is considered Stale rather than In Sync

# ── Connect ──────────────────────────────────────────────────
try {
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    Write-Host "Connected successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# ── Get site owner for translator field ───────────────────────
try {
    $owners = Get-PnPGroupMember -Group "devintranet Owners" -ErrorAction Stop
    $firstOwner      = $owners | Where-Object { $_.LoginName -notlike "*sharepoint*" } | Select-Object -First 1
    $translatorName  = $firstOwner.Title
    $translatorEmail = $firstOwner.Email
    Write-Host "Site owner (translator): $translatorName <$translatorEmail>" -ForegroundColor Cyan
}
catch {
    Write-Host "Could not retrieve site owner — translator fields will be blank: $($_.Exception.Message)" -ForegroundColor Yellow
    $translatorName  = ""
    $translatorEmail = ""
}

# ── Get all Site Pages ────────────────────────────────────────
try {
    Write-Host "Retrieving Site Pages..." -ForegroundColor Cyan
    $allPages = Get-PnPListItem -List "Site Pages" -PageSize 500 -Fields "Title","FileRef","FileLeafRef","Modified","_SPTranslationSourceItemId","_SPIsTranslation","_SPTranslationLanguage","_SPTranslatedLanguages","UniqueId","_ModerationStatus" -ErrorAction Stop
    Write-Host "Retrieved $($allPages.Count) pages" -ForegroundColor Green
}
catch {
    Write-Host "Failed to retrieve Site Pages: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

# ── Separate English source pages and French translations ─────
$sourcePages = $allPages | Where-Object {
    ($_.FieldValues["_SPIsTranslation"] -ne $true) -and
    $_.FieldValues["FileRef"] -notmatch "/fr/" -and
    $_.FieldValues["FileRef"] -match "\.aspx$" -and
    $_.FieldValues["Title"] -ne $null -and
    $_.FieldValues["Title"] -ne "" -and
    $_.FieldValues["_ModerationStatus"] -eq 0  # 0 = Approved
}

$frenchPages = $allPages | Where-Object {
    $_.FieldValues["_SPIsTranslation"] -eq $true -and
    $_.FieldValues["_SPTranslationLanguage"] -eq $TargetLocale
}

Write-Host "English source pages: $($sourcePages.Count)" -ForegroundColor Cyan
Write-Host "French translated pages: $($frenchPages.Count)" -ForegroundColor Cyan

# ── Build French pages lookup by source GUID ─────────────────
$frenchLookup = @{}
foreach ($frPage in $frenchPages) {
    $sourceId = $frPage.FieldValues["_SPTranslationSourceItemId"]
    if ($sourceId) {
        $frenchLookup[$sourceId.ToString().ToLower()] = $frPage
    }
}

# ── Get existing list items for upsert comparison ────────────
Write-Host "Loading existing drift list items..." -ForegroundColor Cyan
$existingItems = Get-PnPListItem -List $ListName -PageSize 500 -Fields "PageGuid","TranslationLanguage" -ErrorAction SilentlyContinue
$existingLookup = @{}
foreach ($ei in $existingItems) {
    $key = "$($ei.FieldValues['PageGuid'].ToString().ToLower())|$($ei.FieldValues['TranslationLanguage'].ToString().ToLower())"
    $existingLookup[$key] = $ei
}
Write-Host "Existing items loaded: $($existingItems.Count)" -ForegroundColor Cyan

# ── Process each English source page ─────────────────────────
$now = Get-Date
$processed = 0
$errors = 0

foreach ($sourcePage in $sourcePages) {
    $pageGuid    = $sourcePage.FieldValues["UniqueId"].ToString().ToLower()
    $pageTitle   = $sourcePage.FieldValues["Title"]
    $pageUrl     = $sourcePage.FieldValues["FileRef"]
    $pageModified = $sourcePage.FieldValues["Modified"]

    Write-Host "Processing: $pageTitle" -ForegroundColor Cyan

    # Determine drift status
    $frPage          = $frenchLookup[$pageGuid.ToLower()]
    $translatedLangs = $sourcePage.FieldValues["_SPTranslatedLanguages"]

    if (-not $frPage) {
        # No French translation exists
        $driftStatus          = "Missing"
        $translationPageUrl   = ""
        $translationModified  = $null
        $daysDrift            = $null
    }
    else {
        $translationPageUrl  = $frPage.FieldValues["FileRef"]
        $translationModified = $frPage.FieldValues["Modified"]
        # Positive drift = English newer than French (stale). Negative = French newer (unexpected).
        $daysDrift           = [math]::Round(($pageModified - $translationModified).TotalDays, 0)
        if ($daysDrift -lt 0) { $daysDrift = 0 }  # French newer than English — treat as In Sync

        if ($daysDrift -le 0) {
            $driftStatus = "In Sync"
        }
        elseif ($daysDrift -le $StaleDays) {
            $driftStatus = "In Sync"
        }
        else {
            $driftStatus = "Stale"
        }
    }

    # ── Upsert into TranslationDrift list ────────────────────
    try {
        # Check if a row already exists for this page + language
        $lookupKey = "$($pageGuid.ToLower())|$($TargetLocale.ToLower())"
        $existingItem = $existingLookup[$lookupKey]

        $itemValues = @{
            "Title"               = $pageTitle
            "DefaultPageTitle"    = $pageTitle
            "DefaultPageUrl"      = $pageUrl
            "DefaultPageModified" = $pageModified.ToString("yyyy-MM-ddTHH:mm:ssZ")
            "TranslationLanguage" = $TargetLocale
            "TranslationPageUrl"  = $translationPageUrl
            "DriftStatus"         = $driftStatus
            "TranslatorName"      = $translatorName
            "TranslatorEmail"     = $translatorEmail
            "SiteUrl"             = $SiteUrl
            "LastChecked"         = $now.ToString("yyyy-MM-ddTHH:mm:ssZ")
            "PageGuid"            = $pageGuid
        }

        # Only set date fields if they have values
        if ($translationModified) {
            $itemValues["TranslationModified"] = $translationModified.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        if ($null -ne $daysDrift) {
            $itemValues["DaysDrift"] = $daysDrift
        }

        if ($existingItem) {
            # Update existing row — preserve NudgeSent/NudgeDate
            try {
                Set-PnPListItem -List $ListName -Identity $existingItem.Id -Values $itemValues -ErrorAction Stop | Out-Null
                Write-Host "  Updated: $driftStatus ($daysDrift days)" -ForegroundColor Green
            }
            catch {
                # Item may have been deleted — fall back to create
                Write-Host "  Item not found, creating instead..." -ForegroundColor Yellow
                $itemValues["NudgeSent"] = $false
                Add-PnPListItem -List $ListName -Values $itemValues -ErrorAction Stop | Out-Null
                Write-Host "  Created: $driftStatus" -ForegroundColor Green
            }
        }
        else {
            # Create new row
            $itemValues["NudgeSent"] = $false
            Add-PnPListItem -List $ListName -Values $itemValues -ErrorAction Stop | Out-Null
            Write-Host "  Created: $driftStatus" -ForegroundColor Green
        }

        $processed++
    }
    catch {
        Write-Host "  Failed to write list item for '$pageTitle': $($_.Exception.Message)" -ForegroundColor Red
        $errors++
    }

    Start-Sleep -Milliseconds 200
}

# ── Summary ───────────────────────────────────────────────────
Write-Host "`n── Run complete ────────────────────────────────" -ForegroundColor Cyan
Write-Host "Pages processed : $processed" -ForegroundColor Green
Write-Host "Errors          : $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
Write-Host "View the list   : $SiteUrl/Lists/$ListName" -ForegroundColor Cyan

# ── NOTE: Scheduling ─────────────────────────────────────────
# To run this script on a schedule, consider Azure Automation Runbooks.
# Replace -Interactive with certificate-based authentication:
# Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant "yourtenant.onmicrosoft.com" -CertificatePath "cert.pfx" -CertificatePassword $securePassword
# Then configure a monthly schedule in Azure Automation to trigger the Runbook.

try {
    Disconnect-PnPOnline
    Write-Host "Disconnected." -ForegroundColor Cyan
}
catch {
    Write-Host "Disconnect error: $($_.Exception.Message)" -ForegroundColor Yellow
}
