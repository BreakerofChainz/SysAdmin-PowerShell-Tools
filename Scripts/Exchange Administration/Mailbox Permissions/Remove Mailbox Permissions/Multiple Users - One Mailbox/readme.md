
# Exchange Online Mailbox Permission Removal

This script removes mailbox permissions in Exchange Online for multiple users. It supports removing Full Access, Send As, and Send On Behalf. The script makes changes to mailbox permissions and is safe to re-run due to built‑in checks.

## Included Scripts

### Remove-MailboxPermissions-MultiUser.ps1
Removes selected permission types from a specified mailbox for all users listed in a text file.

**Behavior:**
1. Resolves the target mailbox.
2. Loads users from a text file (default: Users.txt).
3. Prompts to remove Full Access, Send As, and Send On Behalf.
4. For each user:
   - Resolves user identity.
   - Removes selected permissions if present.
   - Outputs a result object summarizing actions.
5. Disconnects from Exchange Online after processing.

## Requirements
1. PowerShell 5.1 or PowerShell 7+
2. ExchangeOnlineManagement module
3. Interactive sign‑in with rights to modify mailbox permissions

## Security Design
- Modern authentication only
- No stored credentials or secrets
- Strict error handling
- Per‑user isolation of errors
- Explicit disconnect from Exchange Online
- Output may contain user and mailbox identifiers; handle per organizational policy

## Output
- One object per processed user showing:
  - Resolved identity
  - Whether each permission type was removed or not present
  - Success state and any error message
- Creates Users.txt if missing

## Common Issues / Limitations
- Missing permissions can block removals
- Execution policy may prevent running scripts
- Users must be resolvable via Get‑Recipient
- Script requires interactive Y/N input

## Recommended Usage
1. Populate Users.txt with one user per line.
2. Run the script with -SharedMailbox.
3. Select which permissions to remove when prompted.
4. Optionally export results using PowerShell pipeline commands.

## Disclaimer
Scripts are provided as‑is for administrative use. Validate and use according to your organization's security and compliance requirements.

