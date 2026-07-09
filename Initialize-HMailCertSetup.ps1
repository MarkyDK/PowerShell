#Requires -RunAsAdministrator
#Requires -Version 5.1
# =============================================================================
# Author  : Thomas Nielsen Hoff-Hansen, assisted by GitHub Copilot
# Created : 2026-07-09
#
# DISCLAIMER: This script is provided as-is without warranty of any kind.
# The author accepts no responsibility for any issues, data loss, or unintended
# changes that may result from running this script. Always test in a non-
# production environment before executing against live systems.
# Version 1.0
# =============================================================================

<#
.SYNOPSIS
    One-time setup for automated Let's Encrypt SSL certificate management for
    HMailServer using Posh-ACME and Azure DNS validation.

.DESCRIPTION
    Performs the complete one-time setup required for fully automated, unattended
    certificate issuance and renewal:

      Step 1 - Install required PowerShell modules (Posh-ACME, Az.Accounts, Az.Resources)
      Step 2 - Configure the POSHACME_HOME system environment variable
      Step 3 - Create the certificate deployment and log directories with locked-down permissions
      Step 4 - Create the Azure App Registration, Service Principal, and a least-privilege
               custom role ("DNS TXT Contributor") scoped to the DNS zone resource group
      Step 5 - Validate the Azure DNS plugin against the Let's Encrypt Staging server
      Step 6 - Request the production certificate from Let's Encrypt
      Step 7 - Deploy the PEM certificate files to the target directory
      Step 8 - Register the daily renewal task in Windows Task Scheduler

    After this script completes, follow the printed instructions to configure
    HMailServer to use the deployed certificate files (one-time, via the Admin GUI).

    The renewal script (Invoke-HMailCertRenewal.ps1) is then run daily by Task Scheduler.
    It checks whether a renewal is due and, if so, copies the new files and restarts
    the HMailServer Windows service automatically.

.PARAMETER MailHostname
    The fully qualified hostname to issue the certificate for (e.g. mail.contoso.com).

.PARAMETER ContactEmail
    Email address for the Let's Encrypt account. Let's Encrypt sends expiry
    notifications to this address.

.PARAMETER AzSubscriptionId
    Azure Subscription ID (GUID) that contains the DNS zone.

.PARAMETER AzTenantId
    Azure Active Directory Tenant ID (GUID).

.PARAMETER AzResourceGroup
    Name of the Azure Resource Group that contains the DNS zone.

.PARAMETER AzDnsZone
    The Azure DNS zone name (e.g. contoso.com).

.PARAMETER CertDir
    Local directory where the deployed PEM certificate files are stored.
    HMailServer is configured to read from this directory.
    Defaults to C:\Certs\HMailServer.

.PARAMETER PoshAcmeHome
    Directory used as the Posh-ACME data and config root (POSHACME_HOME).
    All ACME account, order, and certificate data is stored here.
    Defaults to C:\PoshACME.

.PARAMETER LogDir
    Directory where renewal log files are written by the renewal script.
    Defaults to C:\Logs\HMailCert.

.PARAMETER RenewalScriptPath
    Full path to the renewal script (Invoke-HMailCertRenewal.ps1).
    Used when registering the Task Scheduler task.
    Defaults to Invoke-HMailCertRenewal.ps1 in the same folder as this script.

.PARAMETER HMailServiceName
    The Windows service name for HMailServer.
    Defaults to 'hMailServer'. Run 'Get-Service *hMail*' to confirm the name.

.PARAMETER SkipAzureSetup
    Skip creating the Azure custom role and Service Principal.
    Use this flag if the SP and role already exist. You will be prompted to enter
    the existing AppId and client secret interactively.

.PARAMETER SkipStaging
    Skip the Let's Encrypt Staging validation test and request the production
    certificate directly. Not recommended for the initial setup.

.EXAMPLE
    .\Initialize-HMailCertSetup.ps1 `
        -MailHostname    mail.contoso.com `
        -ContactEmail    admin@contoso.com `
        -AzSubscriptionId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
        -AzTenantId       'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy' `
        -AzResourceGroup  'dns-rg' `
        -AzDnsZone        'contoso.com'

.EXAMPLE
    # Re-run using a pre-existing Service Principal (skip Azure SP/role creation)
    .\Initialize-HMailCertSetup.ps1 `
        -MailHostname    mail.contoso.com `
        -ContactEmail    admin@contoso.com `
        -AzSubscriptionId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
        -AzTenantId       'yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy' `
        -AzResourceGroup  'dns-rg' `
        -AzDnsZone        'contoso.com' `
        -SkipAzureSetup
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [string]$MailHostname,

    [Parameter(Mandatory)]
    [string]$ContactEmail,

    [Parameter(Mandatory)]
    [string]$AzSubscriptionId,

    [Parameter(Mandatory)]
    [string]$AzTenantId,

    [Parameter(Mandatory)]
    [string]$AzResourceGroup,

    [Parameter(Mandatory)]
    [string]$AzDnsZone,

    [string]$CertDir           = 'C:\Certs\HMailServer',
    [string]$PoshAcmeHome      = 'C:\PoshACME',
    [string]$LogDir            = 'C:\Logs\HMailCert',
    [string]$RenewalScriptPath = '',
    [string]$HMailServiceName  = 'hMailServer',

    [switch]$SkipAzureSetup,
    [switch]$SkipStaging
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helper: Write a section header to the console
# ---------------------------------------------------------------------------
function Write-Step {
    param([string]$Message)
    Write-Host ("`n" + ('=' * 72)) -ForegroundColor DarkGray
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Helper: Write a success indicator
# ---------------------------------------------------------------------------
function Write-OK {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Helper: Write an info line
# ---------------------------------------------------------------------------
function Write-Info {
    param([string]$Message)
    Write-Host "  [..] $Message" -ForegroundColor Gray
}

# Resolve the default renewal script path (same folder as this script)
if (-not $RenewalScriptPath) {
    $RenewalScriptPath = Join-Path $PSScriptRoot 'Invoke-HMailCertRenewal.ps1'
}

Write-Host "`nHMailServer Let's Encrypt SSL Setup" -ForegroundColor White
Write-Host ('=' * 72) -ForegroundColor DarkGray
Write-Host "  Mail hostname    : $MailHostname"
Write-Host "  Contact email    : $ContactEmail"
Write-Host "  Azure DNS zone   : $AzDnsZone (in resource group '$AzResourceGroup')"
Write-Host "  Certificate dir  : $CertDir"
Write-Host "  Posh-ACME home   : $PoshAcmeHome"
Write-Host "  Log dir          : $LogDir"
Write-Host "  HMail service    : $HMailServiceName"
Write-Host "  Skip Azure setup : $SkipAzureSetup"
Write-Host "  Skip staging     : $SkipStaging"
Write-Host ('=' * 72) -ForegroundColor DarkGray

# =============================================================================
# STEP 1 — Install required PowerShell modules
# =============================================================================
Write-Step 'Step 1 of 8 — Installing required PowerShell modules'

$requiredModules = @('Posh-ACME', 'Az.Accounts', 'Az.Resources')
foreach ($modName in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $modName | Select-Object -First 1
    if ($installed) {
        Write-OK "$modName already installed (version $($installed.Version))"
    } else {
        Write-Info "Installing $modName from PowerShell Gallery..."
        Install-Module -Name $modName -Scope AllUsers -Force -AllowClobber -Repository PSGallery
        Write-OK "$modName installed"
    }
}

Import-Module Posh-ACME -ErrorAction Stop

# =============================================================================
# STEP 2 — Configure POSHACME_HOME environment variable
# =============================================================================
Write-Step 'Step 2 of 8 — Configuring POSHACME_HOME'

if (-not (Test-Path -LiteralPath $PoshAcmeHome)) {
    New-Item -ItemType Directory -Path $PoshAcmeHome -Force | Out-Null
    Write-OK "Created directory: $PoshAcmeHome"
}

[System.Environment]::SetEnvironmentVariable('POSHACME_HOME', $PoshAcmeHome, 'Machine')
$env:POSHACME_HOME = $PoshAcmeHome
Write-OK "POSHACME_HOME set to '$PoshAcmeHome' (machine-wide persistent)"
Write-Info 'Note: Open a new PowerShell session for the machine variable to take effect outside this script.'

# =============================================================================
# STEP 3 — Create local directories and restrict permissions on cert dir
# =============================================================================
Write-Step 'Step 3 of 8 — Creating directories'

foreach ($dir in @($CertDir, $LogDir)) {
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-OK "Created: $dir"
    } else {
        Write-OK "Already exists: $dir"
    }
}

# Remove inherited permissions from the cert dir and restrict to Administrators
# and SYSTEM only. This protects the private key file from unprivileged reads.
$acl = Get-Acl -Path $CertDir
$acl.SetAccessRuleProtection($true, $false)  # disable inheritance, remove inherited ACEs

$adminRule  = [System.Security.AccessControl.FileSystemAccessRule]::new(
    'BUILTIN\Administrators',
    'FullControl',
    'ContainerInherit,ObjectInherit',
    'None',
    'Allow')
$systemRule = [System.Security.AccessControl.FileSystemAccessRule]::new(
    'NT AUTHORITY\SYSTEM',
    'FullControl',
    'ContainerInherit,ObjectInherit',
    'None',
    'Allow')

$acl.AddAccessRule($adminRule)
$acl.AddAccessRule($systemRule)
Set-Acl -Path $CertDir -AclObject $acl
Write-OK "Permissions on '$CertDir' restricted to Administrators and SYSTEM"

# =============================================================================
# STEP 4 — Azure Service Principal setup
# =============================================================================
Write-Step 'Step 4 of 8 — Azure Service Principal setup'

if ($SkipAzureSetup) {
    # Use an existing Service Principal — prompt for credentials interactively
    Write-Host '  -SkipAzureSetup specified. Using an existing Service Principal.' -ForegroundColor Yellow
    Write-Host '  Enter the Service Principal credentials when prompted:' -ForegroundColor Yellow
    Write-Host '    Username = Application (Client) ID' -ForegroundColor DarkGray
    Write-Host '    Password = Client Secret'           -ForegroundColor DarkGray
    $appCred = Get-Credential -Message 'Service Principal: Username = AppId, Password = Client Secret'
    Write-OK "Service Principal credentials captured for AppId: $($appCred.UserName)"
} else {
    # Create the Azure App Registration, custom role, and role assignment
    Import-Module Az.Accounts, Az.Resources -ErrorAction Stop

    Write-Info 'Connecting to Azure...'
    Connect-AzAccount -Subscription $AzSubscriptionId -Tenant $AzTenantId | Out-Null
    Write-OK "Connected to Azure subscription '$AzSubscriptionId'"

    # --- Custom role: DNS TXT Contributor ---
    # This role allows only TXT record management on Azure DNS zones, following
    # the principle of least privilege. Based on the Posh-ACME Azure plugin guide:
    # https://poshac.me/docs/v4/Plugins/Azure/
    $roleName    = 'DNS TXT Contributor'
    $existingRole = Get-AzRoleDefinition -Name $roleName -ErrorAction SilentlyContinue

    if (-not $existingRole) {
        Write-Info "Creating custom role '$roleName'..."
        $roleDef = Get-AzRoleDefinition -Name 'DNS Zone Contributor'
        $roleDef.Id          = $null
        $roleDef.Name        = $roleName
        $roleDef.Description = 'Manage DNS TXT records only. Used by Posh-ACME for ACME DNS-01 challenge validation.'

        $roleDef.Actions.RemoveRange(0, $roleDef.Actions.Count)
        $roleDef.Actions.Add('Microsoft.Network/dnsZones/TXT/*')
        $roleDef.Actions.Add('Microsoft.Network/dnsZones/read')
        $roleDef.Actions.Add('Microsoft.Authorization/*/read')
        $roleDef.Actions.Add('Microsoft.Insights/alertRules/*')
        $roleDef.Actions.Add('Microsoft.ResourceHealth/availabilityStatuses/read')
        $roleDef.Actions.Add('Microsoft.Resources/deployments/read')
        $roleDef.Actions.Add('Microsoft.Resources/subscriptions/resourceGroups/read')

        $roleDef.AssignableScopes.Clear()
        $roleDef.AssignableScopes.Add("/subscriptions/$AzSubscriptionId")

        $existingRole = New-AzRoleDefinition $roleDef
        Write-OK "Custom role '$roleName' created"
    } else {
        Write-OK "Custom role '$roleName' already exists"
    }

    # --- Service Principal ---
    $spDisplayName = "PoshACME-HMailServer-$($AzDnsZone -replace '\.', '-')"
    $existingSP    = Get-AzADServicePrincipal -DisplayName $spDisplayName -ErrorAction SilentlyContinue

    if (-not $existingSP) {
        Write-Info "Creating Service Principal '$spDisplayName'..."
        $notBefore = Get-Date
        $notAfter  = $notBefore.AddYears(5)

        $sp = New-AzADServicePrincipal -DisplayName $spDisplayName `
                                        -StartDate $notBefore `
                                        -EndDate   $notAfter

        $spSecret = $sp.PasswordCredentials.SecretText

        Write-OK "Service Principal created: $spDisplayName"
        Write-Host ''
        Write-Host ('  ' + ('*' * 68)) -ForegroundColor Yellow
        Write-Host '  IMPORTANT: Save these credentials — the secret cannot be retrieved again.' -ForegroundColor Yellow
        Write-Host ("  AppId  : {0}" -f $sp.AppId)    -ForegroundColor White
        Write-Host ("  Secret : {0}" -f $spSecret)    -ForegroundColor White
        Write-Host ('  ' + ('*' * 68)) -ForegroundColor Yellow
        Write-Host ''
    } else {
        Write-OK "Service Principal '$spDisplayName' already exists (AppId: $($existingSP.AppId))"
        Write-Host '  A new secret will NOT be generated for an existing SP.' -ForegroundColor Yellow
        $sp        = $existingSP
        $spSecret  = $null
    }

    # --- Role assignment ---
    $existingAssignment = Get-AzRoleAssignment `
        -ApplicationId     $sp.AppId `
        -RoleDefinitionName $roleName `
        -ResourceGroupName  $AzResourceGroup `
        -ErrorAction SilentlyContinue

    if (-not $existingAssignment) {
        # Brief delay to allow the Service Principal to propagate in Azure AD
        Write-Info 'Waiting 20 seconds for Service Principal propagation...'
        Start-Sleep -Seconds 20

        New-AzRoleAssignment `
            -ApplicationId     $sp.AppId `
            -RoleDefinitionName $roleName `
            -ResourceGroupName  $AzResourceGroup | Out-Null

        Write-OK "Role '$roleName' assigned to '$spDisplayName' on resource group '$AzResourceGroup'"
    } else {
        Write-OK "Role assignment already exists for '$spDisplayName'"
    }

    # --- Build the PSCredential for Posh-ACME plugin args ---
    if ($spSecret) {
        $spSecretSecure = ConvertTo-SecureString $spSecret -AsPlainText -Force
    } else {
        $spSecretSecure = Read-Host `
            "Enter the existing client secret for Service Principal '$($sp.AppId)'" `
            -AsSecureString
    }
    $appCred = [pscredential]::new($sp.AppId, $spSecretSecure)
}

# Build the Posh-ACME Azure plugin parameter hashtable.
# These are persisted by Posh-ACME (SecureString values are DPAPI-encrypted on Windows)
# and re-used automatically on every subsequent Submit-Renewal call.
$pArgs = @{
    AZSubscriptionId = $AzSubscriptionId
    AZTenantId       = $AzTenantId
    AZAppCred        = $appCred
}

# =============================================================================
# STEP 5 — Staging validation
# =============================================================================
if (-not $SkipStaging) {
    Write-Step "Step 5 of 8 — Let's Encrypt Staging validation"
    Write-Info 'Requesting a staging certificate to validate the Azure DNS plugin...'
    Write-Info 'This creates a temporary TXT record in Azure DNS and verifies ownership.'

    Set-PAServer LE_STAGE

    $stagingCert = New-PACertificate $MailHostname `
        -AcceptTOS   `
        -Contact     $ContactEmail `
        -Plugin      Azure `
        -PluginArgs  $pArgs `
        -Verbose

    if ($stagingCert) {
        Write-OK "Staging certificate issued successfully"
        Write-OK "  Subject    : $($stagingCert.Subject)"
        Write-OK "  Issuer     : $($stagingCert.Issuer)"
        Write-OK "  Expires    : $($stagingCert.NotAfter.ToString('yyyy-MM-dd'))"
        Write-Host ''
        Write-Host '  Staging validation passed. The Azure DNS plugin is configured correctly.' -ForegroundColor Green
    } else {
        throw "Staging certificate request did not return a certificate object. Check verbose output above."
    }
} else {
    Write-Step "Step 5 of 8 — Let's Encrypt Staging validation"
    Write-Host "  -SkipStaging specified. Skipping staging validation." -ForegroundColor Yellow
}

# =============================================================================
# STEP 6 — Production certificate request
# =============================================================================
Write-Step "Step 6 of 8 — Production certificate request (Let's Encrypt)"
Write-Info "Requesting production certificate for '$MailHostname'..."
Write-Info "This will create and delete a temporary _acme-challenge TXT record in Azure DNS."

Set-PAServer LE_PROD

$cert = New-PACertificate $MailHostname `
    -AcceptTOS   `
    -Contact     $ContactEmail `
    -Plugin      Azure `
    -PluginArgs  $pArgs `
    -Verbose

if (-not $cert) {
    throw "Production certificate request did not return a certificate object. Check verbose output above."
}

Write-OK "Production certificate issued"
Write-OK "  Subject    : $($cert.Subject)"
Write-OK "  Issuer     : $($cert.Issuer)"
Write-OK "  Expires    : $($cert.NotAfter.ToString('yyyy-MM-dd'))"
Write-OK "  CertFile   : $($cert.CertFile)"
Write-OK "  FullChain  : $($cert.FullChainFile)"
Write-OK "  KeyFile    : $($cert.KeyFile)"

# =============================================================================
# STEP 7 — Deploy certificate files
# =============================================================================
Write-Step 'Step 7 of 8 — Deploying certificate files'

$destCert = Join-Path $CertDir 'certificate.cer'
$destKey  = Join-Path $CertDir 'privatekey.key'

# Use the full chain (leaf cert + intermediate CA) so HMailServer presents the
# complete chain to connecting clients. The private key is deployed separately.
Copy-Item -LiteralPath $cert.FullChainFile -Destination $destCert -Force
Copy-Item -LiteralPath $cert.KeyFile       -Destination $destKey  -Force

Write-OK "Certificate (full chain) deployed : $destCert"
Write-OK "Private key deployed              : $destKey"

# =============================================================================
# STEP 8 — Register Windows Task Scheduler task
# =============================================================================
Write-Step 'Step 8 of 8 — Registering Windows Task Scheduler task'

if (-not (Test-Path -LiteralPath $RenewalScriptPath)) {
    Write-Host "  WARNING: Renewal script not found at '$RenewalScriptPath'." -ForegroundColor Yellow
    Write-Host "  The task will still be registered, but update the script path before the first run." -ForegroundColor Yellow
}

$taskName = 'HMailServer-CertRenewal'

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Info "Task '$taskName' already exists — removing and recreating."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Build the scheduled task action — passes all relevant paths as parameters so the
# renewal script does not need to be edited after installation.
$psArgs = '-NonInteractive -ExecutionPolicy Bypass -File "{0}" -MailHostname "{1}" -CertDir "{2}" -PoshAcmeHome "{3}" -LogDir "{4}" -HMailServiceName "{5}"' -f `
    $RenewalScriptPath, $MailHostname, $CertDir, $PoshAcmeHome, $LogDir, $HMailServiceName

$action   = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $psArgs

# Random minute past 3 AM — avoids load spikes on Let's Encrypt servers at round hours
$triggerMinute = Get-Random -Minimum 5 -Maximum 55
$trigger  = New-ScheduledTaskTrigger -Daily -At ('03:{0}' -f $triggerMinute.ToString('D2'))

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# S4U logon type: the task runs as the current user even when they are not
# interactively logged on, without requiring the password to be stored.
# The Posh-ACME plugin args are DPAPI-encrypted for the current user identity,
# so this task MUST run under the same user account that ran this setup script.
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal   = New-ScheduledTaskPrincipal `
    -UserId    $currentUser `
    -LogonType S4U `
    -RunLevel  Highest

Register-ScheduledTask `
    -TaskName   $taskName `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -Principal  $principal `
    -Description ("Automatically renews the Let's Encrypt SSL certificate for HMailServer ({0}) via Posh-ACME and Azure DNS." -f $MailHostname) | Out-Null

Write-OK "Task '$taskName' registered"
Write-OK "  Runs as  : $currentUser (S4U — no stored password required)"
Write-OK ("  Schedule : daily at 03:{0}" -f $triggerMinute.ToString('D2'))
Write-OK "  Script   : $RenewalScriptPath"

# =============================================================================
# SUMMARY — Manual HMailServer configuration instructions
# =============================================================================
$separator = '=' * 72
$nl        = [System.Environment]::NewLine

$instructions = @"

$separator
  ACTION REQUIRED: Configure HMailServer SSL Certificate (one-time step)
$separator

  The certificate files have been deployed. You must now configure
  HMailServer to use them. This is done once via the Admin GUI.
  All future renewals are handled automatically by Task Scheduler.

  -----------------------------------------------------------------------
  A) Configure the SSL Certificate in HMailServer Administrator
  -----------------------------------------------------------------------

  1. Open the HMailServer Administrator application.

  2. Navigate to:
         Settings > Advanced > SSL certificates

  3. Click [Add] and complete the fields:
         Name             : LetsEncrypt
         Certificate file : $destCert
         Private key file : $destKey

  4. Click [Save].

  -----------------------------------------------------------------------
  B) Assign the Certificate to the SMTP Port
  -----------------------------------------------------------------------

  5. Navigate to:
         Settings > Protocols > SMTP > Ports

  6. Select your TLS-enabled port:
         Port 465  (SMTPS / Implicit TLS)
         Port 587  (Submission / STARTTLS)

  7. In the port settings, set the SSL Certificate to 'LetsEncrypt'.

  8. Click [Save].

  -----------------------------------------------------------------------
  C) Restart HMailServer and Verify
  -----------------------------------------------------------------------

  9. Restart the HMailServer Windows service:
         Restart-Service $HMailServiceName

 10. Verify the certificate is presented correctly:
         openssl s_client -connect ${MailHostname}:587 -starttls smtp

  -----------------------------------------------------------------------
  Paths and configuration summary
  -----------------------------------------------------------------------

  Certificate file (full chain) : $destCert
  Private key file              : $destKey
  Certificate expires           : $($cert.NotAfter.ToString('yyyy-MM-dd'))

  Posh-ACME data directory      : $PoshAcmeHome
  Renewal script                : $RenewalScriptPath
  Log directory                 : $LogDir
  Scheduled task                : $taskName (daily 03:$($triggerMinute.ToString('D2')))

  To force an immediate renewal (e.g. for testing):
      & '$RenewalScriptPath' -MailHostname '$MailHostname' ``
          -CertDir '$CertDir' ``
          -PoshAcmeHome '$PoshAcmeHome' ``
          -LogDir '$LogDir' ``
          -Force

$separator

"@

Write-Host $instructions -ForegroundColor White
Write-Host 'Setup complete.' -ForegroundColor Green
