# Shared Mailbox Audit Scripts

These scripts identify inactive shared mailboxes in Exchange Online using a fixed 90 day Unified Audit Log lookback period.
They are designed for read-only auditing and mailbox lifecycle review. No changes are made to mailboxes or permissions, and the scripts do not store any sensitive information or secrets. 


## Included Scripts

### List all Shared.ps1

Exports all shared mailboxes from Exchange Online.
Writes information to ALLUSERS.csv at the root of the script.
This script is expected to be run first, as the next script uses this csv file. 

### Shared Mailbox Audit.ps1

Checks shared mailbox activity in Unified Audit Logs.

  **Behavior:**

1. Reads mailbox UPNs from ALLUSERS.csv
2. Uses a fixed 90 day lookback period
3. Exports only mailboxes with no detected activity
4. Writes information to 90DayInactiveAudit.csv


## Requirements

1. PowerShell 5.1 or PowerShell 7+
2. ExchangeOnlineManagement module
3. Interactive sign-in with access to Exchange Online mailbox data and Unified Audit Logs


## Security Design

- Read-only cmdlets only
- Modern authentication
- No stored credentials or secrets
- Strict error handling enabled
- Time-bounded and user-scoped audit queries
- Explicit disconnect from Exchange Online and Compliance sessions
- CSV outputs may contain mailbox identifiers and should be handled according to organizational data policies.


## Common Issues

- Execution policy may block local scripts
- Missing permissions may prevent audit results for some mailboxes
- Audit log retention limits restrict visibility of older activity
- Large tenants may experience long runtimes. Use -Verbose when running this if you would like to verify the script is working correctly. 


## Recommended Usage

1. Run List all Shared.ps1 to generate ALLUSERS.csv
2. Run Shared Mailbox Audit.ps1 to generate 90DayInactiveAudit.csv
3. Review results for cleanup, access review, or lifecycle decisions


## **Disclaimer**
Scripts are provided as-is for administrative and audit purposes. Review usage and outputs in accordance with organizational security and compliance policies specific to your environment. 

