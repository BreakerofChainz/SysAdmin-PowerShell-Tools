# Exchange Online Mailbox Permission Script

This script grants Full Access, Send As, and Send On Behalf permissions to a specified user on an Exchange Online mailbox. It uses modern authentication, interactive permission selection, and automatic session cleanup. The script modifies mailbox permissions.

## Included Scripts

### AddOneMailbox.ps1

Grants selected mailbox permissions through interactive prompts.

**Behavior:**

1. Prompts for Full Access, Send As, and Send On Behalf additions  
2. Connects to Exchange Online  
3. Resolves the target mailbox and user  
4. Adds only missing permissions  
5. Outputs a summary object showing what was granted or already present  
6. Attempts to disconnect from Exchange Online in all cases

## Requirements

1. PowerShell 5.1 or PowerShell 7+  
2. ExchangeOnlineManagement module  
3. Interactive sign-in with rights to modify mailbox permissions

## Security Design

- Modern authentication  
- No stored credentials or secrets  
- Idempotent permission checks  
- Strict error handling  
- Automatic disconnect from Exchange Online  
- Output contains mailbox identifiers; handle according to organizational policies

## Output

- A single `PSCustomObject` reporting:  
  - Full Access (granted/already present)  
  - Send As (granted/already present)  
  - Send On Behalf (granted/already present)  
- No files are written to disk

## Common Issues / Limitations

- Missing Exchange Online roles may block permission changes  
- Execution policy may prevent script execution  
- Interactive prompts make it unsuitable for unattended automation

## Recommended Usage

1. Run the script with `-SharedMailbox` and `-User` parameters  
2. Answer prompts for which permissions to add  
3. Review the output summary for confirmation

## Disclaimer

Scripts are provided as-is. Review usage and outputs in accordance with organizational security and compliance requirements.
