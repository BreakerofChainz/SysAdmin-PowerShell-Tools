# Multi‑User Mailbox Permission Script

This script assigns mailbox permissions in Exchange Online for multiple users using a text file input. It supports Full Access, Send As, and Send On Behalf permissions and makes direct changes to mailbox configuration.

## Included Scripts

### AddMultipleUsersOneMailbox.ps1
Grants selected permission types to multiple users on a single mailbox.

**Behavior:**
1. Connects to Exchange Online
2. Resolves the target mailbox
3. Loads user identifiers from Users.txt or a specified file
4. Prompts which permissions to apply
5. Adds only missing permissions
6. Outputs a status object per user
7. Disconnects from Exchange Online

## Requirements

1. PowerShell 5.1 or PowerShell 7+
2. ExchangeOnlineManagement module
3. Interactive authentication with rights to manage mailbox permissions

## Security Design

- Modern authentication only
- No stored credentials or secrets
- Strict error handling and safe disconnect logic
- Outputs contain mailbox and user identifiers only
- No log files generated unless operator exports them

## Output

- Structured PowerShell objects containing:
  - Mailbox identity
  - User identity (input and resolved)
  - Permission results
  - Success or error state

## Common Issues / Limitations

- Users.txt must contain valid identifiers
- Missing roles can prevent permission assignment
- Large user lists may increase runtime
- Script requires interactive prompts and is not fully automation‑friendly

## Recommended Usage

1. Populate Users.txt with one user per line
2. Run the script with `-SharedMailbox mailbox@domain.com`
3. Select desired permissions at prompts
4. Pipe results to CSV if reporting is needed

## Disclaimer

Scripts are provided as‑is for administrative use. Always validate in non‑production environments and ensure compliance with organizational security policies.
