# ============================================================
# Create-TranslationDriftList.ps1
# Creates the Translation Drift tracking list and all columns
# ============================================================

$SiteUrl  = "https://tenantName.sharepoint.com/sites/siteName"
$ClientId = ""
$ListName = "TranslationDrift"

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

# ── Create List ──────────────────────────────────────────────
try {
    Write-Host "Checking for existing list..." -ForegroundColor Cyan
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

    if ($list) {
        Write-Host "List '$ListName' already exists — skipping creation" -ForegroundColor Yellow
    }
    else {
        Write-Host "Creating list '$ListName'..." -ForegroundColor Cyan
        $list = New-PnPList -Title $ListName -Template GenericList -ErrorAction Stop
        if (-not $list) {
            throw "List creation returned null."
        }
        Write-Host "List created" -ForegroundColor Green
        Start-Sleep -Seconds 5
    }
}
catch {
    Write-Host "Failed to create list '$ListName': $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

# ── Helper function ───────────────────────────────────────────
function Add-ColumnIfNotExists {
    param(
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Type,
        [string[]]$Choices = @()
    )

    try {
        $existing = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  Column '$DisplayName' already exists — skipping" -ForegroundColor Yellow
            return
        }

        switch ($Type) {
            "Text" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type Text -AddToDefaultView -ErrorAction Stop | Out-Null
            }
            "DateTime" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type DateTime -AddToDefaultView -ErrorAction Stop | Out-Null
            }
            "Number" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type Number -AddToDefaultView -ErrorAction Stop | Out-Null
            }
            "Boolean" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type Boolean -AddToDefaultView -ErrorAction Stop | Out-Null
            }
            "Choice" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type Choice -Choices $Choices -AddToDefaultView -ErrorAction Stop | Out-Null
            }
            "URL" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type URL -AddToDefaultView -ErrorAction Stop | Out-Null
            }
        }
        Write-Host "  Created column '$DisplayName'" -ForegroundColor Green
        Start-Sleep -Milliseconds 300
    }
    catch {
        Write-Host "  Failed to create column '$DisplayName': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ── Core Page Identification ──────────────────────────────────
Write-Host "`nCreating core page identification columns..." -ForegroundColor Cyan
Add-ColumnIfNotExists -DisplayName "Default Page Title"    -InternalName "DefaultPageTitle"    -Type "Text"
Add-ColumnIfNotExists -DisplayName "Default Page Url"      -InternalName "DefaultPageUrl"      -Type "URL"
Add-ColumnIfNotExists -DisplayName "Default Page Modified" -InternalName "DefaultPageModified" -Type "DateTime"
Add-ColumnIfNotExists -DisplayName "Translation Language"  -InternalName "TranslationLanguage" -Type "Text"
Add-ColumnIfNotExists -DisplayName "Translation Page Url"  -InternalName "TranslationPageUrl"  -Type "URL"
Add-ColumnIfNotExists -DisplayName "Translation Modified"  -InternalName "TranslationModified" -Type "DateTime"

# ── Drift Calculation ─────────────────────────────────────────
Write-Host "`nCreating drift calculation columns..." -ForegroundColor Cyan
Add-ColumnIfNotExists -DisplayName "Days Drift"   -InternalName "DaysDrift"   -Type "Number"
Add-ColumnIfNotExists -DisplayName "Drift Status" -InternalName "DriftStatus" -Type "Choice" -Choices @("In Sync", "Stale", "Missing")

# ── Translator / Ownership ────────────────────────────────────
Write-Host "`nCreating translator columns..." -ForegroundColor Cyan
Add-ColumnIfNotExists -DisplayName "Translator Name"  -InternalName "TranslatorName"  -Type "Text"
Add-ColumnIfNotExists -DisplayName "Translator Email" -InternalName "TranslatorEmail" -Type "Text"

# ── Governance Metadata ───────────────────────────────────────
Write-Host "`nCreating governance metadata columns..." -ForegroundColor Cyan
Add-ColumnIfNotExists -DisplayName "Site Url"     -InternalName "SiteUrl"     -Type "Text"
Add-ColumnIfNotExists -DisplayName "Last Checked" -InternalName "LastChecked" -Type "DateTime"
Add-ColumnIfNotExists -DisplayName "Page Guid"    -InternalName "PageGuid"    -Type "Text"

# ── Nudge Tracking ────────────────────────────────────────────
Write-Host "`nCreating nudge tracking columns..." -ForegroundColor Cyan
Add-ColumnIfNotExists -DisplayName "Nudge Sent" -InternalName "NudgeSent" -Type "Boolean"
Add-ColumnIfNotExists -DisplayName "Nudge Date" -InternalName "NudgeDate" -Type "DateTime"

# ── Done ──────────────────────────────────────────────────────
Write-Host "`nList '$ListName' is ready." -ForegroundColor Green
Write-Host "URL: $SiteUrl/Lists/$ListName" -ForegroundColor Cyan

try {
    Disconnect-PnPOnline
    Write-Host "Disconnected." -ForegroundColor Cyan
}
catch {
    Write-Host "Disconnect error: $($_.Exception.Message)" -ForegroundColor Yellow
}
