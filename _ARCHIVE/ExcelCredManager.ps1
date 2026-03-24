<#
.SYNOPSIS
    Excel credential scanner and cleaner for SharePoint / Power BI / SQL Server auth stacks.
.PARAMETER Action
    Scan  — enumerate all relevant token stores, write cred_report.csv
    Clear — delete entries marked Stale or Expired in cred_report.csv, then re-run Scan
.PARAMETER ReportPath
    Where to write the CSV. Default: %TEMP%\cred_report.csv
#>
param(
    [ValidateSet('Scan','Clear')]
    [string]$Action = 'Scan',
    [string]$ReportPath = "$env:TEMP\cred_report.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'   # non-existent paths are expected

# ── Constants ────────────────────────────────────────────────────────────────
$NOW          = Get-Date
$STALE_DAYS   = 1        # LastWriteTime age threshold when no expiry parseable
$SOON_HOURS   = 1        # "Expiring Soon" window

# ── Token store paths ─────────────────────────────────────────────────────────
$STORES = @{
    OneAuth     = "$env:LOCALAPPDATA\Microsoft\OneAuth\accounts"
    TokenBroker = "$env:LOCALAPPDATA\Microsoft\TokenBroker\Accounts"
    PowerQuery  = "$env:LOCALAPPDATA\Microsoft\Power Query"
    MsalCache   = "$env:LOCALAPPDATA\Microsoft\MicrosoftEdge\User\Default\IndexedDB"  # fallback, rarely used by Excel
}

# ── Helpers ──────────────────────────────────────────────────────────────────
function Get-ExpiryStatus {
    param([nullable[datetime]]$Expiry, [datetime]$LastWrite)
    if ($Expiry) {
        if ($Expiry -lt $NOW)                    { return 'Expired' }
        if ($Expiry -lt $NOW.AddHours($SOON_HOURS)) { return 'Expiring Soon' }
        return 'Valid'
    }
    # No parseable expiry — fall back to LastWriteTime age
    $ageDays = ($NOW - $LastWrite).TotalDays
    if ($ageDays -gt $STALE_DAYS) { return "Stale ($([math]::Round($ageDays,1))d old)" }
    return 'Recent (no expiry field)'
}

function Parse-OneAuthExpiry {
    param([string]$JsonPath)
    try {
        $j = Get-Content $JsonPath -Raw | ConvertFrom-Json
        # OneAuth stores expiry in several possible fields depending on Office version
        foreach ($field in @('extended_expires_on','expires_on','access_token_expiry','expiry_time','exp')) {
            $val = $j.$field
            if (-not $val) { $val = $j.access_token.$field }
            if ($val) {
                # Unix epoch integer
                if ($val -match '^\d+$') {
                    return [DateTimeOffset]::FromUnixTimeSeconds([long]$val).LocalDateTime
                }
                # ISO string
                try { return [datetime]$val } catch {}
            }
        }
    } catch {}
    return $null
}

function Parse-TokenBrokerExpiry {
    param([string]$JsonPath)
    try {
        $j = Get-Content $JsonPath -Raw | ConvertFrom-Json
        foreach ($field in @('expiry','expirationTime','access_token_expiry','exp','expires_on')) {
            $val = $j.$field
            if ($val) {
                if ($val -match '^\d+$') {
                    return [DateTimeOffset]::FromUnixTimeSeconds([long]$val).LocalDateTime
                }
                try { return [datetime]$val } catch {}
            }
        }
    } catch {}
    return $null
}

function New-Record {
    param([string]$Store, [string]$Target, [string]$User,
          [datetime]$LastWrite, [nullable[datetime]]$Expiry, [string]$TokenType)
    $status = Get-ExpiryStatus -Expiry $Expiry -LastWrite $LastWrite
    [PSCustomObject]@{
        Store      = $Store
        Target     = $Target
        User       = $User
        LastWrite  = $LastWrite.ToString('yyyy-MM-dd HH:mm:ss')
        Expiry     = if ($Expiry) { $Expiry.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        Status     = $status
        TokenType  = $TokenType
        FilePath   = ''   # set by caller for file-based stores
    }
}

# ── SCAN ─────────────────────────────────────────────────────────────────────
function Invoke-Scan {
    $records = [System.Collections.Generic.List[object]]::new()

    # 1. OneAuth accounts (SharePoint + Power BI AAD tokens)
    if (Test-Path $STORES.OneAuth) {
        Get-ChildItem $STORES.OneAuth -Filter '*.json' -File | ForEach-Object {
            $expiry = Parse-OneAuthExpiry $_.FullName
            try   { $j = Get-Content $_.FullName -Raw | ConvertFrom-Json }
            catch { $j = $null }
            $user = if ($j) { $j.username ?? $j.preferred_username ?? $j.upn ?? '' } else { '' }
            $r = New-Record -Store 'OneAuth' -Target $_.BaseName -User $user `
                            -LastWrite $_.LastWriteTime -Expiry $expiry -TokenType 'AAD OAuth'
            $r.FilePath = $_.FullName
            $records.Add($r)
        }
    }

    # 2. TokenBroker accounts (WAM — used by SQL AAD + PBI in newer Office builds)
    if (Test-Path $STORES.TokenBroker) {
        Get-ChildItem $STORES.TokenBroker -Filter '*.json' -File -Recurse | ForEach-Object {
            $expiry = Parse-TokenBrokerExpiry $_.FullName
            $r = New-Record -Store 'TokenBroker' -Target $_.BaseName -User '' `
                            -LastWrite $_.LastWriteTime -Expiry $expiry -TokenType 'WAM/AAD'
            $r.FilePath = $_.FullName
            $records.Add($r)
        }
    }

    # 3. Power Query cache (connection metadata, not auth tokens — but stale cache causes refresh failures)
    if (Test-Path $STORES.PowerQuery) {
        Get-ChildItem $STORES.PowerQuery -Recurse -File |
            Where-Object { $_.Extension -in @('.json','.cache','') } | ForEach-Object {
            $r = New-Record -Store 'PowerQuery' -Target $_.Name -User '' `
                            -LastWrite $_.LastWriteTime -Expiry $null -TokenType 'PQ Cache'
            $r.FilePath = $_.FullName
            $records.Add($r)
        }
    }

    # 4. Credential Manager — filter Office/SharePoint/SQL entries
    $credKeywords = @('MicrosoftOffice','Microsoft_OC','Office16','Office15',
                      'PowerBI','SharePoint','OneDrive','microsoftonline',
                      'login.windows.net','AADToken','MicrosoftSqlServer')
    $raw = cmdkey /list 2>$null
    $curTarget = ''; $curType = ''; $curUser = ''
    foreach ($line in $raw) {
        $t = $line.Trim()
        if ($t -like 'Target:*') {
            $curTarget = $t -replace '^Target:\s*',''
            $curType = ''; $curUser = ''
        } elseif ($t -like 'Type:*')  { $curType = $t -replace '^Type:\s*','' }
        elseif ($t -like 'User:*') {
            $curUser = $t -replace '^User:\s*',''
            $match = $credKeywords | Where-Object { $curTarget -like "*$_*" }
            if ($match) {
                $r = [PSCustomObject]@{
                    Store     = 'CredentialManager'
                    Target    = $curTarget
                    User      = $curUser
                    LastWrite = ''
                    Expiry    = ''
                    Status    = 'Found'        # no timestamp available from cmdkey
                    TokenType = $curType
                    FilePath  = ''
                }
                $records.Add($r)
            }
        }
    }

    $records | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Scan complete. $($records.Count) entries written to $ReportPath"
}

# ── CLEAR ─────────────────────────────────────────────────────────────────────
function Invoke-Clear {
    if (-not (Test-Path $ReportPath)) {
        Write-Error "No report found at $ReportPath — run Scan first."
        return
    }

    $records = Import-Csv $ReportPath
    $toClear = $records | Where-Object { $_.Status -match 'Expired|Stale|Found' }

    if (-not $toClear) {
        Write-Host "Nothing to clear."
        return
    }

    $deleted = 0; $failed = 0

    foreach ($r in $toClear) {
        if ($r.Store -eq 'CredentialManager') {
            $result = cmdkey /delete:"$($r.Target)" 2>&1
            if ($LASTEXITCODE -eq 0) { $deleted++ } else { $failed++ }
        } else {
            # File-based stores
            if ($r.FilePath -and (Test-Path $r.FilePath)) {
                try {
                    Remove-Item $r.FilePath -Force
                    $deleted++
                } catch { $failed++ }
            } elseif ($r.Store -eq 'PowerQuery' -and (Test-Path $STORES.PowerQuery)) {
                # Nuke entire PQ cache dir — individual file tracking unreliable across renames
                try {
                    Remove-Item $STORES.PowerQuery -Recurse -Force
                    $deleted++
                    break   # done, don't iterate further PQ entries
                } catch { $failed++ }
            }
        }
    }

    Write-Host "Cleared: $deleted  Failed: $failed"
    # Re-scan so the workbook sees fresh state
    Invoke-Scan
}

# ── Entry point ───────────────────────────────────────────────────────────────
switch ($Action) {
    'Scan'  { Invoke-Scan  }
    'Clear' { Invoke-Clear }
}
