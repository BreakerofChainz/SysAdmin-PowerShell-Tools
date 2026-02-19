# Add Multiple Mailbox Permissions for One User

This script grants a single user Full Access, Send As, and Send On Behalf permissions on multiple Exchange Online mailboxes. It is intended for administrative use and makes changes in Exchange Online.

## Included Scripts

### AddMultipleMailboxesOneUser.ps1

Reads a list of mailbox identities, prompts which permissions to assign, and applies only missing permissions.

**Behavior:**

1. Resolves or creates the mailbox list file (`Emails.txt` or custom `-EmailsFilePath`).
2. Connects to Exchange Online using modern authentication.
3. Prompts for Full Access, Send As, and Send On Behalf selections.
4. Resolves the target user with `Get-Recipient`.
5. Processes each mailbox:
   - Resolves mailbox identity
   - Grants selected permissions only if not already present
6. Outputs a result object per mailbox with success and error details.
7. Disconnects from Exchange Online.

## Requirements

1. PowerShell 5.1 or PowerShell 7+
2. ExchangeOnlineManagement module
3. Interactive sign-in with rights to modify mailbox permissions in Exchange Online

## Security Design

- Modern authentication only
- No stored credentials or secrets
- Strict error handling enabled
- Supports `-WhatIf` via `SupportsShouldProcess`
- Explicit disconnect from Exchange Online
- Output may contain mailbox and user identifiers; handle per organizational policy

## Output

- Console output containing:
  - Resolved mailbox and user identities
  - Granted vs already-present permissions
  - Success flag
  - Error message if applicable
- May create an empty `Emails.txt` if the file does not exist

## Common Issues / Limitations

- Execution policy may block script execution
- Missing Exchange Online roles can prevent permission changes
- Invalid or empty mailbox list halts processing
- Incorrect identities cause resolution errors

## Recommended Usage

1. Populate `Emails.txt` with one mailbox identity per line.
2. Run the script and select desired permissions when prompted.
3. Review output or export results for auditing.

## Disclaimer

Scripts are provided as-is for administrative use. Validate functionality and follow organizational security and compliance requirements.
