
#Shared Mailbox Audit Scripts
These scripts identify inactive shared mailboxes in Exchange Online using a fixed 90 day Unified Audit Log lookback period.
They are designed for read-only auditing and mailbox lifecycle review. No changes are made to mailboxes or permissions.


#Included Scripts
List all Shared.ps1
Exports all shared mailboxes from Exchange Online.
Generates:

ALLUSERS.csv
This script is expected to be run first.


#Shared Mailbox Audit.ps1
Checks shared mailbox activity in Unified Audit Logs.
Behavior:

Reads mailbox UPNs from ALLUSERS.csv
Uses a fixed 90 day lookback period
Exports only mailboxes with no detected activity
Generates:

90DayInactiveAudit.csv


Requirements

PowerShell 5.1 or PowerShell 7+
ExchangeOnlineManagement module
Interactive sign-in with access to:Exchange Online mailbox data
Unified Audit Logs


Security Design

Read-only cmdlets only
Modern authentication
No stored credentials or secrets
Strict error handling enabled
Time-bounded and user-scoped audit queries
Explicit disconnect from Exchange Online and Compliance sessions
CSV outputs may contain mailbox identifiers and should be handled according to organizational data policies.


Common Issues

Execution policy may block local scripts
Missing permissions may prevent audit results for some mailboxes
Audit log retention limits restrict visibility of older activity
Large tenants may experience long runtimes


Recommended Usage

Run List all Shared.ps1 to generate ALLUSERS.csv
Run Shared Mailbox Audit.ps1 to generate 90DayInactiveAudit.csv
Review results for cleanup, access review, or lifecycle decisions


Disclaimer
Scripts are provided as-is for administrative and audit purposes. Review usage and outputs in accordance with organizational security and compliance policies.

