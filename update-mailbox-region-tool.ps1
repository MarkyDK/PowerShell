# =============================================================================
# Author  : Thomas Nielsen Hoff-Hansen, assisted by GitHub Copilot
# Created : 2026-07-06 on a rainy day in my summer vacation 2026 here in Denmark.
#
# DISCLAIMER: This script is provided as-is without warranty of any kind.
# The author accepts no responsibility for any issues, data loss, or unintended
# changes that may result from running this script. Always test in a non-
# production environment before executing against live mailboxes.
# Version 1.0
# =============================================================================

<#
.SYNOPSIS
    Reports on and optionally updates Exchange Online mailbox regional settings.

.DESCRIPTION
    In report mode (default), connects to Exchange Online, reads current regional settings
    for each target mailbox, and displays a per-mailbox comparison of current vs. planned
    settings. A CSV log is always written.

    Pass -Execute to apply the changes. The script auto-disconnects after executing changes unless
    -NoDisconnect is specified. In report mode the session is left open.

    Log files are written to -LogPath (defaults to the current directory) with the filename:
        yyyyMMdd_HHmmss_Report.csv   or   yyyyMMdd_HHmmss_Execute.csv

.PARAMETER All
    Target all mailboxes, excluding Discovery Mailboxes. This is the default targeting mode.

.PARAMETER Identity
    Target a single mailbox by primary SMTP address, alias, or display name.

.PARAMETER CsvPath
    Path to a CSV file. Must contain an 'Identity' column.
    The value can be a primary SMTP address, alias, or display name.

.PARAMETER Mailboxes
    A pre-built collection of mailbox objects or identity strings (primary SMTP address, alias, or display name).

.PARAMETER MailboxType
    Filter by mailbox type. Accepts one or more of: UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox.
    Defaults to all four types.

.PARAMETER Country
    Country whose regional settings to apply. Defaults to 'Denmark'.
    Tab completion is supported for 29 countries (27 EU member states + Norway + United Kingdom).

.PARAMETER Language
    Override the country's default language code (e.g. 'de-DE').

.PARAMETER TimeZone
    Override the country's default Windows timezone ID (e.g. 'W. Europe Standard Time').

.PARAMETER DateFormat
    Override the country's default date format string (e.g. 'dd.MM.yyyy').

.PARAMETER TimeFormat
    Override the country's default time format string (e.g. 'HH:mm').

.PARAMETER LocalizeDefaultFolderName
    Rename default mailbox folders (Inbox, Sent Items, etc.) to match the target language.
    Enabled by default. Pass $false to disable.

.PARAMETER Execute
    Apply the regional settings. Without this switch the script runs in report mode only.
    Mailboxes whose settings already match the target are skipped unless -ForceUpdate is also set.

.PARAMETER ForceUpdate
    When combined with -Execute, apply settings to all mailboxes even if settings appear to already be correct.

.PARAMETER LogPath
    Folder where the CSV log file is saved. Defaults to the current directory.
    The folder must already exist; the script will not create it.

.PARAMETER NoDisconnect
    Skip the automatic disconnect from Exchange Online after -Execute completes.

.EXAMPLE
    .\update-mailbox-region-tool.ps1
    Report mode: all mailboxes with Danish settings shown as planned values.

.EXAMPLE
    .\update-mailbox-region-tool.ps1 -Country Germany -All -Execute
    Apply German regional settings to all mailboxes.

.EXAMPLE
    .\update-mailbox-region-tool.ps1 -Identity user@contoso.com -Country Sweden
    Report mode for a single mailbox, previewing Swedish settings.

.EXAMPLE
    .\update-mailbox-region-tool.ps1 -All -Execute -LogPath C:\Logs -NoDisconnect
    Apply Danish settings, save log to C:\Logs, and stay connected when done.
#>

[CmdletBinding(DefaultParameterSetName = 'All')]
param (
    # --- Targeting (mutually exclusive) ---
    [Parameter(ParameterSetName = 'All')]
    [switch]$All,

    [Parameter(ParameterSetName = 'Identity', Mandatory)]
    [string]$Identity,

    [Parameter(ParameterSetName = 'CsvPath', Mandatory)]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Mailboxes', Mandatory)]
    [object[]]$Mailboxes,

    [ValidateSet('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')]
    [string[]]$MailboxType = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'),

    # --- Region ---
    # Note: when adding a country here, also add it to $CountryConfig in the Begin block.
    [ValidateSet(
        'Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic',
        'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary',
        'Ireland', 'Italy', 'Latvia', 'Lithuania', 'Luxembourg', 'Malta',
        'Netherlands', 'Norway', 'Poland', 'Portugal', 'Romania', 'Slovakia',
        'Slovenia', 'Spain', 'Sweden', 'United Kingdom'
    )]
    [string]$Country = 'Denmark',
    [string]$Language,
    [string]$TimeZone,
    [string]$DateFormat,
    [string]$TimeFormat,
    [bool]$LocalizeDefaultFolderName = $true,

    # --- Behavior ---
    [switch]$Execute,
    [switch]$ForceUpdate,
    [string]$LogPath,
    [switch]$NoDisconnect
)

Begin {
    # =============================================================================
    # COUNTRY CONFIG
    # To add a new country: add entry to this hashtable.
    # TimeZone values must be Windows timezone IDs. Run Get-TimeZone -ListAvailable to verify.
    # =============================================================================
    $CountryConfig = [ordered]@{
        'Denmark'        = @{ Capital = 'Copenhagen';      Language = 'da-DK'; DateFormat = 'dd-MM-yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'Romance Standard Time'          }
        'Austria'        = @{ Capital = 'Vienna';          Language = 'de-AT'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Belgium'        = @{ Capital = 'Brussels';        Language = 'fr-BE'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'Romance Standard Time'          }  # nl-BE also valid
        'Bulgaria'       = @{ Capital = 'Sofia';           Language = 'bg-BG'; DateFormat = "dd.MM.yyyy '$([char]0x0433).'"; TimeFormat = 'HH:mm'; TimeZone = 'FLE Standard Time' }
        'Croatia'        = @{ Capital = 'Zagreb';          Language = 'hr-HR'; DateFormat = "dd.MM.yyyy."; TimeFormat = 'HH:mm'; TimeZone = 'Central European Standard Time' }
        'Cyprus'         = @{ Capital = 'Nicosia';         Language = 'el-CY'; DateFormat = 'd/M/yyyy';   TimeFormat = 'HH:mm'; TimeZone = 'E. Europe Standard Time'        }
        'Czech Republic' = @{ Capital = 'Prague';          Language = 'cs-CZ'; DateFormat = "dd.MM.yyyy";  TimeFormat = 'H:mm'; TimeZone = 'Central Europe Standard Time'   }
        'Estonia'        = @{ Capital = 'Tallinn';         Language = 'et-EE'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'FLE Standard Time'              }
        'Finland'        = @{ Capital = 'Helsinki';        Language = 'fi-FI'; DateFormat = 'd.M.yyyy';   TimeFormat = 'H.mm';  TimeZone = 'FLE Standard Time'              }
        'France'         = @{ Capital = 'Paris';           Language = 'fr-FR'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'Romance Standard Time'          }
        'Germany'        = @{ Capital = 'Berlin';          Language = 'de-DE'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Greece'         = @{ Capital = 'Athens';          Language = 'el-GR'; DateFormat = 'd/M/yyyy';   TimeFormat = 'HH:mm'; TimeZone = 'GTB Standard Time'              }
        'Hungary'        = @{ Capital = 'Budapest';        Language = 'hu-HU'; DateFormat = 'yyyy. MM. dd.'; TimeFormat = 'H:mm'; TimeZone = 'Central Europe Standard Time'   }
        'Ireland'        = @{ Capital = 'Dublin';          Language = 'en-IE'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'GMT Standard Time'              }
        'Italy'          = @{ Capital = 'Rome';            Language = 'it-IT'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Latvia'         = @{ Capital = 'Riga';            Language = 'lv-LV'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'FLE Standard Time'              }
        'Lithuania'      = @{ Capital = 'Vilnius';         Language = 'lt-LT'; DateFormat = 'yyyy-MM-dd'; TimeFormat = 'HH:mm'; TimeZone = 'FLE Standard Time'              }
        'Luxembourg'     = @{ Capital = 'Luxembourg City'; Language = 'fr-LU'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'Romance Standard Time'          }
        'Malta'          = @{ Capital = 'Valletta';        Language = 'mt-MT'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Netherlands'    = @{ Capital = 'Amsterdam';       Language = 'nl-NL'; DateFormat = 'd-M-yyyy';   TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Poland'         = @{ Capital = 'Warsaw';          Language = 'pl-PL'; DateFormat = 'd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'Central European Standard Time' }
        'Portugal'       = @{ Capital = 'Lisbon';          Language = 'pt-PT'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'GMT Standard Time'              }
        'Romania'        = @{ Capital = 'Bucharest';       Language = 'ro-RO'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'GTB Standard Time'              }
        'Slovakia'       = @{ Capital = 'Bratislava';      Language = 'sk-SK'; DateFormat = 'd. M. yyyy'; TimeFormat = 'H:mm'; TimeZone = 'Central Europe Standard Time'   }
        'Slovenia'       = @{ Capital = 'Ljubljana';       Language = 'sl-SI'; DateFormat = 'd. MM. yyyy';  TimeFormat = 'HH:mm'; TimeZone = 'Central Europe Standard Time'   }
        'Spain'          = @{ Capital = 'Madrid';          Language = 'es-ES'; DateFormat = 'dd/MM/yyyy';   TimeFormat = 'H:mm'; TimeZone = 'Romance Standard Time'          }
        'Sweden'         = @{ Capital = 'Stockholm';       Language = 'sv-SE'; DateFormat = 'yyyy-MM-dd'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'Norway'         = @{ Capital = 'Oslo';            Language = 'nb-NO'; DateFormat = 'dd.MM.yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'W. Europe Standard Time'         }
        'United Kingdom' = @{ Capital = 'London';          Language = 'en-GB'; DateFormat = 'dd/MM/yyyy'; TimeFormat = 'HH:mm'; TimeZone = 'GMT Standard Time'              }
    }

    # Helper: truncate and right-pad a string to a fixed column width for console table output
    function Format-Column ([string]$Value, [int]$Width) {
        if ($Value.Length -le $Width) { return $Value.PadRight($Width) }
        return ($Value.Substring(0, $Width - 3) + '...').PadRight($Width)
    }

    # Resolve effective region settings: country defaults, overridden by any explicitly passed params
    $cfg                 = $CountryConfig[$Country]
    $effectiveLanguage   = if ($PSBoundParameters.ContainsKey('Language'))   { $Language }   else { $cfg.Language   }
    $effectiveTimeZone   = if ($PSBoundParameters.ContainsKey('TimeZone'))   { $TimeZone }   else { $cfg.TimeZone   }
    $effectiveDateFormat = if ($PSBoundParameters.ContainsKey('DateFormat')) { $DateFormat } else { $cfg.DateFormat }
    $effectiveTimeFormat = if ($PSBoundParameters.ContainsKey('TimeFormat')) { $TimeFormat } else { $cfg.TimeFormat }

    # Resolve log folder - capture $PWD now before any potential location changes
    $logFolder = if ($PSBoundParameters.ContainsKey('LogPath')) {
        if (-not (Test-Path -LiteralPath $LogPath -PathType Container)) {
            throw "LogPath folder does not exist: '$LogPath'"
        }
        $LogPath
    } else {
        $PWD.Path
    }

    # Build auto-generated log filename
    $modeLabel = if ($Execute) { 'Execute' } else { 'Report' }
    $logFile   = Join-Path $logFolder ("$(Get-Date -Format 'yyyyMMdd_HHmmss')_${modeLabel}.csv")

    # Initialise log collection (shared across Begin/Process/End)
    $logEntries = [System.Collections.Generic.List[PSObject]]::new()

    # Connect to Exchange Online if no active session exists.
    # Get-ConnectionInformation (EXO module v3+) is a fast local check with no round-trip.
    # Falls back to Get-OrganizationConfig for older module versions.
    $isConnected = $false
    try {
        $connInfo    = Get-ConnectionInformation -ErrorAction Stop
        $isConnected = $null -ne $connInfo -and @($connInfo).Count -gt 0
    } catch {
        try { $null = Get-OrganizationConfig -ErrorAction Stop; $isConnected = $true } catch {}
    }
    if (-not $isConnected) {
        Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false
    }

    # Mode banner
    if ($Execute) {
        Write-Host "`n[ EXECUTE MODE - changes will be applied ]" -ForegroundColor Yellow
    } else {
        Write-Host "`n[ REPORT MODE - read only. Use -Execute to apply. ]" -ForegroundColor Cyan
    }
    Write-Host ("Country  : {0} ({1})"                                                              -f $Country, $cfg.Capital)
    Write-Host ("Settings : Language={0}  TimeZone={1}  DateFormat={2}  TimeFormat={3}"             -f $effectiveLanguage, $effectiveTimeZone, $effectiveDateFormat, $effectiveTimeFormat)
    Write-Host ("Log file : {0}`n"                                                                   -f $logFile)
}

Process {
    # Build target mailbox list based on which parameter set is active
    $rawMailboxes = switch ($PSCmdlet.ParameterSetName) {
        'All' {
            Write-Host 'Retrieving mailboxes...' -ForegroundColor Cyan
            Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $MailboxType
        }
        'Identity' {
            Write-Host "Retrieving mailbox: $Identity" -ForegroundColor Cyan
            try { Get-Mailbox -Identity $Identity -ErrorAction Stop }
            catch { throw "Could not find mailbox '$Identity': $_" }
        }
        'CsvPath' {
            if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
                throw "CsvPath file not found: '$CsvPath'"
            }
            Write-Host "Reading mailboxes from CSV: $CsvPath" -ForegroundColor Cyan
            Import-Csv -LiteralPath $CsvPath | ForEach-Object {
                if (-not $_.PSObject.Properties['Identity'] -or -not $_.Identity) {
                    Write-Warning 'Skipping CSV row: missing Identity value.'
                    return
                }
                try { Get-Mailbox -Identity $_.Identity -ErrorAction Stop }
                catch { Write-Warning "Could not resolve mailbox '$($_.Identity)': $_" }
            }
        }
        'Mailboxes' {
            foreach ($item in $Mailboxes) {
                if ($item -is [string]) {
                    try { Get-Mailbox -Identity $item -ErrorAction Stop }
                    catch { Write-Warning "Could not resolve mailbox '$item': $_" }
                } else {
                    $item
                }
            }
        }
    }

    # For -All the type filter is already applied server-side by -RecipientTypeDetails;
    # for other parameter sets filter locally to enforce the type restriction.
    $targetList = if ($PSCmdlet.ParameterSetName -eq 'All') {
        @($rawMailboxes | Where-Object { $_ })
    } else {
        @($rawMailboxes | Where-Object { $_ -and $_.RecipientTypeDetails -in $MailboxType })
    }
    if ($targetList.Count -eq 0) {
        Write-Warning 'No target mailboxes found.'
        return
    }
    Write-Host ("{0} mailbox(es) found (type: {1}).`n" -f $targetList.Count, ($MailboxType -join ', '))

    # Column widths for the console table
    $w = @{ Name = 25; Alias = 10; Smtp = 30; CurLang = 9; CurTZ = 28; CurDate = 11; CurTime = 9; NewLang = 9; NewTZ = 28; NewDate = 11; NewTime = 9 }

    $header = (
        (Format-Column 'DisplayName'   $w.Name)    + ' ' +
        (Format-Column 'Alias'         $w.Alias)   + ' ' +
        (Format-Column 'PrimarySmtp'   $w.Smtp)    + ' ' +
        (Format-Column 'Old.Language'  $w.CurLang) + ' ' +
        (Format-Column 'Old.TimeZone'  $w.CurTZ)   + ' ' +
        (Format-Column 'Old.DateFmt'   $w.CurDate) + ' ' +
        (Format-Column 'Old.TimeFmt'   $w.CurTime) + ' ' +
        (Format-Column 'New.Language'  $w.NewLang) + ' ' +
        (Format-Column 'New.TimeZone'  $w.NewTZ)   + ' ' +
        (Format-Column 'New.DateFmt'   $w.NewDate) + ' ' +
        (Format-Column 'New.TimeFmt'   $w.NewTime) + ' ' +
        'Status'
    )
    Write-Host $header -ForegroundColor White
    Write-Host ('-' * $header.Length) -ForegroundColor DarkGray

    $total = $targetList.Count
    $index = 0
    $progressActivity = if ($Execute) { 'Updating mailbox regional settings' } else { 'Reading mailbox regional settings' }

    foreach ($mailbox in $targetList) {
        $index++
        $mbxId    = $mailbox.PrimarySmtpAddress
        $alias    = $mailbox.Alias
        $dispName = $mailbox.DisplayName
        $mbxType  = [string]$mailbox.RecipientTypeDetails

        Write-Progress -Activity $progressActivity `
                       -Status   ("[$index / $total] $dispName") `
                       -CurrentOperation $mbxId `
                       -PercentComplete  ([math]::Round($index / [math]::Max($total, 1) * 100))

        # Read current regional configuration
        try {
            $regional = Get-MailboxRegionalConfiguration -Identity $mbxId -ErrorAction Stop
        } catch {
            Write-Warning "Could not read regional config for '$mbxId': $_"
            continue
        }

        # Language is returned as a CultureInfo object; casting to string gives the BCP-47 tag (e.g. da-DK)
        $curLang = [string]$regional.Language
        $curTZ   = [string]$regional.TimeZone
        $curDate = [string]$regional.DateFormat
        $curTime = [string]$regional.TimeFormat

        # Determine whether this mailbox requires a settings change.
        # Exchange Online does not reliably return Language and TimeFormat for all mailbox types
        # (e.g. SharedMailbox). Empty fields are skipped in the comparison to avoid perpetual
        # false positives. However if ALL fields are empty the mailbox is unconfigured and always
        # needs an update.
        $allEmpty    = $curLang -eq '' -and $curTZ -eq '' -and $curDate -eq '' -and $curTime -eq ''
        $langMatch   = $curLang  -eq '' -or $curLang  -eq $effectiveLanguage
        $tzMatch     = $curTZ    -eq '' -or $curTZ    -eq $effectiveTimeZone
        $dateMatch   = $curDate  -eq '' -or $curDate  -eq $effectiveDateFormat
        $timeMatch   = $curTime  -eq '' -or $curTime  -eq $effectiveTimeFormat
        $needsUpdate = $allEmpty -or -not ($langMatch -and $tzMatch -and $dateMatch -and $timeMatch)

        $status      = ''
        $message     = ''
        $updateStatus = ''

        if ($Execute) {
            if ($needsUpdate -or $ForceUpdate) {
                try {
                    $setParams = @{
                        Identity    = $mbxId
                        Language    = $effectiveLanguage
                        TimeZone    = $effectiveTimeZone
                        DateFormat  = $effectiveDateFormat
                        TimeFormat  = $effectiveTimeFormat
                        ErrorAction = 'Stop'
                    }
                    if ($LocalizeDefaultFolderName) { $setParams['LocalizeDefaultFolderName'] = $true }
                    Set-MailboxRegionalConfiguration @setParams
                    $status       = 'Success'
                    $updateStatus = 'Updated'
                } catch {
                    $status       = 'Failed'
                    $message      = $_.Exception.Message
                    $updateStatus = 'UpdateFailed'
                }
            } else {
                $status       = 'Skipped'
                $updateStatus = 'AlreadyUpToDate'
            }
        } else {
            $status       = if ($needsUpdate) { 'WouldApply' } else { 'UpToDate' }
            $updateStatus = if ($needsUpdate) { 'Yes' } else { 'No' }
        }

        $rowColor = switch ($status) {
            'Success'    { 'Green'    }
            'Failed'     { 'Red'      }
            'WouldApply' { 'Cyan'     }
            'UpToDate'   { 'Green'    }
            'Skipped'    { 'DarkGray' }
        }

        $row = (
            (Format-Column $dispName            $w.Name)    + ' ' +
            (Format-Column $alias               $w.Alias)   + ' ' +
            (Format-Column $mbxId               $w.Smtp)    + ' ' +
            (Format-Column $curLang             $w.CurLang) + ' ' +
            (Format-Column $curTZ               $w.CurTZ)   + ' ' +
            (Format-Column $curDate             $w.CurDate) + ' ' +
            (Format-Column $curTime             $w.CurTime) + ' ' +
            (Format-Column $effectiveLanguage   $w.NewLang) + ' ' +
            (Format-Column $effectiveTimeZone   $w.NewTZ)   + ' ' +
            (Format-Column $effectiveDateFormat $w.NewDate) + ' ' +
            (Format-Column $effectiveTimeFormat $w.NewTime) + ' ' +
            $status
        )
        Write-Host $row -ForegroundColor $rowColor

        $logEntries.Add([PSCustomObject]@{
            Timestamp          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            PrimarySmtpAddress = $mbxId
            Alias              = $alias
            DisplayName        = $dispName
            MailboxType        = $mbxType
            OldLanguage        = $curLang
            OldTimeZone        = $curTZ
            OldDateFormat      = $curDate
            OldTimeFormat      = $curTime
            NewLanguage        = $effectiveLanguage
            NewTimeZone        = $effectiveTimeZone
            NewDateFormat      = $effectiveDateFormat
            NewTimeFormat      = $effectiveTimeFormat
            NeedsUpdate        = $updateStatus
            Status             = $status
            Message            = $message
        })
    }

    Write-Progress -Activity $progressActivity -Completed

    Write-Host ('-' * $header.Length) -ForegroundColor DarkGray

    if ($Execute) {
        $successCount = @($logEntries | Where-Object { $_.Status -eq 'Success' }).Count
        $skippedCount = @($logEntries | Where-Object { $_.Status -eq 'Skipped' }).Count
        $failCount    = @($logEntries | Where-Object { $_.Status -eq 'Failed'  }).Count
        Write-Host ("Completed: {0} updated, {1} skipped (already correct), {2} failed.`n" -f $successCount, $skippedCount, $failCount)
    } else {
        $needsUpdateCount = @($logEntries | Where-Object { $_.NeedsUpdate -eq 'Yes' }).Count
        Write-Host ("Report complete: {0} of {1} mailbox(es) need updating.`n" -f $needsUpdateCount, $logEntries.Count)
    }
}

End {
    # Write CSV log (always)
    if ($logEntries.Count -gt 0) {
        # ConvertTo-Csv + Set-Content writes UTF-8 WITH BOM on both PS5.1 and PS7,
        # ensuring correct display of non-ASCII characters (e.g. accented names) when opened in Excel.
        $logEntries | ConvertTo-Csv -NoTypeInformation | Set-Content -LiteralPath $logFile -Encoding UTF8
        Write-Host ("Log saved to: {0}" -f $logFile) -ForegroundColor Cyan
    }

    # Disconnect only after Execute (report mode leaves the session open)
    if ($Execute -and -not $NoDisconnect) {
        Write-Host 'Disconnecting from Exchange Online...' -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
