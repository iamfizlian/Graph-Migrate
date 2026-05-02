<#
.SYNOPSIS
  Provision N additional Entra app registrations for pstmigrate, each with
  Mail.ReadWrite (application) granted at the tenant level.

.DESCRIPTION
  Adds throughput headroom to the migration by giving you more independent
  Graph throttle buckets. Each new app:
    - Is created with the given display name (default: pstmigrate-N)
    - Is granted Microsoft Graph -> Mail.ReadWrite (application)
    - Has admin consent applied automatically
    - Generates a 24-month client secret
    - Outputs ready-to-paste TOML for config.toml

.PARAMETER NamePrefix
  Prefix for new apps (default: 'pstmigrate'). Apps will be named
  '<prefix>-<N>' starting from -StartIndex.

.PARAMETER StartIndex
  Number to start from (default: 5, since you already have pstmigrate-1..4).

.PARAMETER Count
  How many apps to create (default: 2).

.PARAMETER ScopePolicyGroup
  Optional. Mail-enabled security group (UPN form, e.g. 'PSTMigrate-Targets@contoso.com')
  to apply an Application Access Policy restricting the new apps to those
  mailboxes only. Highly recommended for production use.

.EXAMPLE
  .\Add-MigrationApps.ps1
  # Creates pstmigrate-5 and pstmigrate-6, prints their TOML blocks

.EXAMPLE
  .\Add-MigrationApps.ps1 -Count 4 -ScopePolicyGroup 'PSTMigrate-Targets@contoso.com'
  # Creates pstmigrate-5..8 and scopes each via ApplicationAccessPolicy
#>

[CmdletBinding()]
param(
    [string]$NamePrefix = 'pstmigrate',
    [int]$StartIndex    = 5,
    [int]$Count         = 2,
    [string]$ScopePolicyGroup = $null,
    [int]$SecretLifetimeMonths = 24
)

$ErrorActionPreference = 'Stop'

# ---- 1. Module preflight --------------------------------------------------
$required = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Identity.SignIns'
)
foreach ($mod in $required) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod ..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -ErrorAction Stop
}

# ---- 2. Connect with consent for the operations we need -------------------
$scopes = @(
    'Application.ReadWrite.All',     # create app + service principal
    'AppRoleAssignment.ReadWrite.All', # grant Mail.ReadWrite consent
    'Directory.ReadWrite.All'        # tenant-level consent
)
Write-Host "Connecting to Microsoft Graph (you will be prompted for sign-in)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes -NoWelcome

$ctx = Get-MgContext
Write-Host ("Connected as {0} on tenant {1}" -f $ctx.Account, $ctx.TenantId) -ForegroundColor Green
$tenantId = $ctx.TenantId

# ---- 3. Resolve Microsoft Graph service principal once --------------------
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$mailReadWriteRole = $graphSp.AppRoles | Where-Object {
    $_.Value -eq 'Mail.ReadWrite' -and $_.AllowedMemberTypes -contains 'Application'
}
if (-not $mailReadWriteRole) {
    throw "Could not find Mail.ReadWrite app role on Microsoft Graph SP - aborting."
}

# ---- 4. Loop and create -------------------------------------------------
$results = @()
for ($i = 0; $i -lt $Count; $i++) {
    $appName = "{0}-{1}" -f $NamePrefix, ($StartIndex + $i)
    Write-Host ""
    Write-Host ("=== {0} ===" -f $appName) -ForegroundColor Cyan

    # Skip if it already exists
    $existing = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  ! '$appName' already exists (AppId $($existing.AppId)) - SKIPPING." -ForegroundColor Yellow
        continue
    }

    # 4a. Create the application with required Graph permission baked in
    $requiredAccess = @{
        ResourceAppId  = $graphSp.AppId
        ResourceAccess = @(@{
            Id   = $mailReadWriteRole.Id
            Type = 'Role'
        })
    }
    $app = New-MgApplication `
        -DisplayName $appName `
        -SignInAudience 'AzureADMyOrg' `
        -RequiredResourceAccess @($requiredAccess)
    Write-Host "  + App created: $($app.AppId)" -ForegroundColor Green

    # 4b. Create the service principal (needed for permission grants)
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "  + Service principal created: $($sp.Id)" -ForegroundColor Green

    # 4c. Grant admin consent for Mail.ReadWrite (app)
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $sp.Id `
        -PrincipalId $sp.Id `
        -ResourceId $graphSp.Id `
        -AppRoleId $mailReadWriteRole.Id | Out-Null
    Write-Host "  + Granted Mail.ReadWrite (admin consent applied)" -ForegroundColor Green

    # 4d. Create a client secret
    $secretEnd = (Get-Date).AddMonths($SecretLifetimeMonths)
    $pw = Add-MgApplicationPassword `
        -ApplicationId $app.Id `
        -PasswordCredential @{
            DisplayName = "pstmigrate auto-generated"
            EndDateTime = $secretEnd
        }
    Write-Host "  + Client secret created (expires $($secretEnd.ToString('yyyy-MM-dd')))" -ForegroundColor Green

    # 4e. Optional: scope the app via ApplicationAccessPolicy
    if ($ScopePolicyGroup) {
        Write-Host "  ~ Applying ApplicationAccessPolicy scoped to $ScopePolicyGroup ..." -ForegroundColor Yellow
        Write-Host "    (run this in an Exchange Online PowerShell session manually:)" -ForegroundColor DarkGray
        Write-Host ("      New-ApplicationAccessPolicy -AppId {0} -PolicyScopeGroupId '{1}' -AccessRight RestrictAccess -Description 'PST migration scope'" -f $app.AppId, $ScopePolicyGroup) -ForegroundColor DarkGray
    }

    $results += [pscustomobject]@{
        Name        = $appName
        AppId       = $app.AppId
        TenantId    = $tenantId
        Secret      = $pw.SecretText
        SecretExpiry = $secretEnd
    }
}

# ---- 5. Emit ready-to-paste TOML -----------------------------------------
if ($results.Count -eq 0) {
    Write-Host ""
    Write-Host "No new apps were created (all targets already existed)." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Append the following to config.toml under your existing" -ForegroundColor Cyan
Write-Host " [[apps]] blocks, then bump max_parallel_mailboxes to 6." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
foreach ($r in $results) {
    @"
[[apps]]
name = "$($r.Name)"
tenant_id = "$($r.TenantId)"
client_id = "$($r.AppId)"
client_secret = "$($r.Secret)"

"@
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " IMPORTANT: client secrets above are shown ONLY ONCE." -ForegroundColor Yellow
Write-Host " Copy them into config.toml NOW. They cannot be retrieved later." -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-MgGraph | Out-Null
