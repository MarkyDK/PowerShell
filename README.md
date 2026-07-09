# PowerShellScripts
PowerShell scripts that I build are put here.

# Update-Mailbox-Region-Tool.ps1
Update-Mailbox-Region-Tool.ps1 is a PowerShell administration tool for Exchange Online that simplifies the management of mailbox regional settings across one or many mailboxes. It can be used to audit existing regional configuration, preview planned changes, and apply standardized regional settings in a controlled and repeatable manner.

By default, the script operates in report mode, allowing administrators to verify the current mailbox configuration before making any changes. When the -Execute parameter is specified, the script updates mailbox regional settings, including language, time zone, date format, time format, and optionally localizes the default mailbox folder names.

Key Features
Reports current and planned mailbox regional settings before making changes.
Supports report-only and execution modes.
Targets:
All supported mailbox types
Individual mailboxes
Multiple mailboxes from a CSV file
Pre-built mailbox collections
Supports User, Shared, Room, and Equipment mailboxes.
Applies predefined regional settings for all EU member states, Norway, and the United Kingdom.
Allows overriding default country settings with custom:
Language
Time zone
Date format
Time format
Optionally renames default mailbox folders (Inbox, Sent Items, Calendar, etc.) to the selected language.
Skips mailboxes that already match the desired configuration unless -ForceUpdate is specified.
Generates detailed timestamped CSV logs for both reporting and execution.
Automatically disconnects from Exchange Online after execution (optional).
Typical Use Cases
Standardize mailbox regional settings after Microsoft 365 tenant migrations.
Configure mailbox localization for users in multiple countries.
Audit mailbox regional settings without making changes.
Correct inconsistent language or time zone settings across an organization.
Bulk configure mailbox settings from a CSV file.
Safety

The script is designed with a preview-first approach. Administrators can review all proposed changes in report mode before committing them with -Execute. Every run produces a detailed CSV log for auditing and troubleshooting purposes.

This tool is intended for Microsoft 365 and Exchange Online administrators who need a reliable, repeatable, and scalable method of managing mailbox regional settings.
