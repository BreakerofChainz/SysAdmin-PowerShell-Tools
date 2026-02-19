# Mailbox Permission Removal Script

This script removes Full Access, Send As, and Send On Behalf permissions for a specified user on a specified mailbox in Exchange Online.  
It makes changes to mailbox permissions and uses modern authentication. The script automatically disconnects from Exchange Online upon completion.

## Included Scripts

### RemoveOneMailbox.ps1

Removes selected permissions from a single mailbox.

- Prompts for removal of Full Access, Send As, and Send On Behalf  
- Resolves both the mailbox and the target user  
- Removes only permissions that already exist  
- Produces a summary object showing which permissions were removed or not present  
- Automatically disconnects the Exchange Online session when finished

## Behavior

1. Connects to Exchange Online using modern authentication  
2. Resolves the mailbox and target user  
3. Removes permissions only when they are explicitly assigned  
4. Matches trustees using identifiers such as ExternalDirectoryObjectId, GUID, SMTP, or UPN  
5. Updates GrantSendOnBehalfTo for Send On Behalf removal  
6. Outputs a structured summary object  
7. Disconnects from Exchange Online automatically in the final step

## Requirements

1. PowerShell 5.1 or PowerShell 7+  
2. ExchangeOnlineManagement module  
3. Interactive sign‑in with rights to modify mailbox permissions  

## Security Design

- Modern authentication with no stored credentials or secrets  
- Strict error handling enabled  
- Uses ShouldProcess for safe, controlled permission removal  
- Automatic session cleanup to prevent lingering Exchange Online connections  
- Outputs only mailbox and user identifiers required for administrative review  

## Output

- A single PSCustomObject containing:
  - Mailbox and User SMTP values  
  - FullAccessRemoved / FullAccessNotPresent  
  - SendAsRemoved / SendAsNotPresent  
  - SendOnBehalfRemoved / SendOnBehalfNotPresent  

## Common Issues / Limitations

- Insufficient permissions may prevent removal of certain rights  
- Local execution policy may block script execution  
- Large environments may experience slower mailbox/recipient resolution  
- Inconsistent module versions may affect cmdlet behavior  

## Recommended Usage

1. Run the script with `-SharedMailbox` and `-User`  
2. Select which permissions to remove when prompted  
3. Use `-WhatIf` during testing to validate expected behavior  

## Disclaimer

Scripts are provided as‑is for administrative use. Review and use them in accordance with your organization’s security, access, and compliance policies.
