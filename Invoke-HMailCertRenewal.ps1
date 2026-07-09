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
    Renews the Let's Encrypt SSL certificate for HMailServer and deploys the
    updated certificate files.

.DESCRIPTION
    This script is intended to be run daily by Windows Task Scheduler as part
    of the automated certificate management solution set up by Initialize-HMailCertSetup.ps1.

    On each run it:
      1. Configures the Posh-ACME data directory (POSHACME_HOME)
      2. Targets the Let's Encrypt production server and the correct certificate order
      3. Calls Submit-Renewal — Posh-ACME renews only if the certificate is within
         30 days of expiry (or immediately if -Force is used)
      4. If a renewal occurred:
           - Copies the updated full-chain certificate to the deployment directory
           - Copies the updated private key to the deployment directory
           - Restarts the HMailServer Windows service so it loads the new certificate
      5. Writes a timestamped entry to a monthly log file regardless of outcome

    The script exits with code 0 on success or no-op, and code 1 on any error.
    Task Scheduler records this exit code in the task history.

    Azure DNS credentials are stored encrypted by Posh-ACME in the order directory
    (DPAPI-encrypted on Windows). No credentials are needed in this script.

.PARAMETER MailHostname
    The fully qualified hostname the certificate was issued for.
    Must match the order name used during the initial setup (e.g. mail.contoso.com).

.PARAMETER CertDir
    Local directory containing the deployed certificate PEM files.
    Must match the path configured in HMailServer.
    Defaults to C:\Certs\HMailServer.

.PARAMETER PoshAcmeHome
    Directory used as the Posh-ACME data and config root.
    Must match the POSHACME_HOME value set during initial setup.
    Defaults to C:\PoshACME.

.PARAMETER LogDir
    Directory where log files are written.
    A monthly log file (yyyy-MM_renewal.log) is appended to on each run.
    Defaults to C:\Logs\HMailCert.

.PARAMETER HMailServiceName
    The Windows service name for HMailServer.
    Defaults to 'hMailServer'. Run 'Get-Service *hMail*' to confirm the name.

.PARAMETER Force
    Force a renewal even if the certificate is not yet within the renewal window.
    Use this to test the full renewal and deploy flow without waiting for expiry.

.EXAMPLE
    # Typical Task Scheduler invocation (all parameters passed by the setup script):
    .\Invoke-HMailCertRenewal.ps1 `
        -MailHostname    mail.contoso.com `
        -CertDir         'C:\Certs\HMailServer' `
        -PoshAcmeHome    'C:\PoshACME' `
        -LogDir          'C:\Logs\HMailCert' `
        -HMailServiceName hMailServer

.EXAMPLE
    # Force a renewal to test the full deploy flow (e.g. after initial setup):
    .\Invoke-HMailCertRenewal.ps1 -MailHostname mail.contoso.com -Force
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$MailHostname,

    [string]$CertDir           = 'C:\Certs\HMailServer',
    [string]$PoshAcmeHome      = 'C:\PoshACME',
    [string]$LogDir            = 'C:\Logs\HMailCert',
    [string]$HMailServiceName  = 'hMailServer',

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
# Logging
# =============================================================================

# Ensure the log directory exists before any logging attempt
if (-not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Monthly log file: one file per calendar month to keep log sizes manageable
$logFile = Join-Path $LogDir ((Get-Date -Format 'yyyy-MM') + '_renewal.log')

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $entry = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level, $Message
    Add-Content -LiteralPath $logFile -Value $entry -Encoding UTF8

    switch ($Level) {
        'ERROR' { Write-Host $entry -ForegroundColor Red    }
        'WARN'  { Write-Host $entry -ForegroundColor Yellow }
        default { Write-Host $entry }
    }
}

# =============================================================================
# Main
# =============================================================================

Write-Log ('--- Renewal check started | Hostname: {0} | Force: {1} ---' -f $MailHostname, $Force.IsPresent)

try {
    # -------------------------------------------------------------------------
    # Configure Posh-ACME data directory
    # The machine environment variable is set by Initialize-HMailCertSetup.ps1,
    # but we also set it in-process to guarantee correct behaviour when invoked
    # by Task Scheduler (S4U sessions may not inherit the machine variable).
    # -------------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $PoshAcmeHome)) {
        throw "Posh-ACME home directory not found: '$PoshAcmeHome'. Run Initialize-HMailCertSetup.ps1 first."
    }
    $env:POSHACME_HOME = $PoshAcmeHome

    # -------------------------------------------------------------------------
    # Import Posh-ACME
    # -------------------------------------------------------------------------
    if (-not (Get-Module -ListAvailable -Name 'Posh-ACME')) {
        throw "Posh-ACME module is not installed. Run Initialize-HMailCertSetup.ps1 first."
    }
    Import-Module Posh-ACME -ErrorAction Stop

    # -------------------------------------------------------------------------
    # Target the Let's Encrypt production server and the correct order.
    # Set-PAOrder also sets the active server to whichever server the order
    # belongs to, so this is safe even if PAServer was last set to LE_STAGE.
    # -------------------------------------------------------------------------
    Set-PAServer LE_PROD -ErrorAction Stop
    Set-PAOrder $MailHostname -ErrorAction Stop

    # Confirm the order exists and is in a usable state
    $order = Get-PAOrder
    if (-not $order) {
        throw "No Posh-ACME order found for '$MailHostname' on LE_PROD. Run Initialize-HMailCertSetup.ps1 first."
    }
    Write-Log ("Current certificate expires: {0:yyyy-MM-dd}" -f (Get-PACertificate).NotAfter)

    # -------------------------------------------------------------------------
    # Attempt renewal
    # Posh-ACME will only renew if the certificate is within the renewal window
    # (default: 30 days before expiry). Pass -Force to override this check.
    # Submit-Renewal returns the PACertificate object ONLY when a renewal
    # actually occurred; it returns nothing if no renewal was needed.
    # -------------------------------------------------------------------------
    Write-Log ('Calling Submit-Renewal (Force: {0})' -f $Force.IsPresent)
    $cert = Submit-Renewal -Force:$Force -ErrorAction Stop

    if ($cert) {
        # ---------------------------------------------------------------------
        # Renewal occurred — deploy the updated certificate files
        # ---------------------------------------------------------------------
        Write-Log ('Certificate renewed. New expiry: {0:yyyy-MM-dd}' -f $cert.NotAfter)

        # Validate the source files exist before deploying
        if (-not (Test-Path -LiteralPath $cert.FullChainFile)) {
            throw "Expected FullChainFile not found at '$($cert.FullChainFile)'"
        }
        if (-not (Test-Path -LiteralPath $cert.KeyFile)) {
            throw "Expected KeyFile not found at '$($cert.KeyFile)'"
        }

        # Verify the deployment directory is still accessible
        if (-not (Test-Path -LiteralPath $CertDir)) {
            throw "Certificate deployment directory not found: '$CertDir'"
        }

        $destCert = Join-Path $CertDir 'certificate.cer'
        $destKey  = Join-Path $CertDir 'privatekey.key'

        # Copy full chain (leaf cert + intermediate CA) as the certificate file.
        # HMailServer needs the full chain to present a complete trust path to clients.
        Copy-Item -LiteralPath $cert.FullChainFile -Destination $destCert -Force
        Copy-Item -LiteralPath $cert.KeyFile       -Destination $destKey  -Force

        Write-Log "Certificate files deployed: $CertDir"

        # ---------------------------------------------------------------------
        # Restart HMailServer to load the new certificate from disk.
        # HMailServer reads SSL certificate files at startup; a restart is
        # required to pick up the renewed certificate.
        # ---------------------------------------------------------------------
        $svc = Get-Service -Name $HMailServiceName -ErrorAction Stop

        if ($svc.Status -eq 'Running') {
            Write-Log "Restarting service '$HMailServiceName'..."
            Restart-Service -Name $HMailServiceName -Force -ErrorAction Stop
            Write-Log "Service '$HMailServiceName' restarted successfully"
        } elseif ($svc.Status -eq 'Stopped') {
            Write-Log "Service '$HMailServiceName' is stopped. Starting it..." 'WARN'
            Start-Service -Name $HMailServiceName -ErrorAction Stop
            Write-Log "Service '$HMailServiceName' started"
        } else {
            Write-Log "Service '$HMailServiceName' is in state '$($svc.Status)'. Manual intervention may be required." 'WARN'
        }

        Write-Log '--- Renewal completed successfully ---'

    } else {
        # ---------------------------------------------------------------------
        # No renewal — certificate is still within its validity window
        # ---------------------------------------------------------------------
        $currentCert = Get-PACertificate
        $daysLeft    = [math]::Round(($currentCert.NotAfter - (Get-Date)).TotalDays)
        Write-Log ("Certificate is not yet due for renewal ($daysLeft day(s) remaining). No action taken.")
        Write-Log '--- Renewal check completed (no-op) ---'
    }

    exit 0

} catch {
    Write-Log "UNHANDLED ERROR: $($_.Exception.Message)" 'ERROR'
    Write-Log "Script stack trace: $($_.ScriptStackTrace)"  'ERROR'

    # Write a blank separator line to the log for readability
    Add-Content -LiteralPath $logFile -Value '' -Encoding UTF8
    Write-Log '--- Renewal check failed ---' 'ERROR'

    exit 1
}
