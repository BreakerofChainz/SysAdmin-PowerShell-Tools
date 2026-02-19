# Multi-Mailbox Permission Removal Script

This script removes Full Access, Send As, and Send On Behalf permissions for a single user across multiple Exchange Online mailboxes.  
It makes administrative changes to permissions and uses modern authentication. The script automatically disconnects from Exchange Online when finished.

## Included Scripts

### RemoveMultipleMailboxesOneUser.ps1

Reads mailbox identities from a text file and removes selected permissions for one user.

- Prompts for which permission types to remove  
- Resolves the user once before processing  
- Matches trustees using Identity, the provided User value, SMTP address, and Alias  
- Skips invalid or missing ACEs without failing execution  
- Outputs one result object per mailbox  

## Behavior

1. Resolves script directory and validates or creates Emails.txt  
2. Connects to Exchange Online via modern authentication  
3. Prompts for selected permission removals  
4. Resolves the user from the provided input  
5. Reads mailbox identities (ignores blank/commented lines)  
6. For each mailbox: resolves identity, removes selected permissions if present, and returns a result object  
7. Automatically disconnects from Exchange Online  

## Requirements

1. PowerShell 5.1 or PowerShell 7+  
2. ExchangeOnlineManagement module  
3. Interactive sign-in with rights to modify mailbox permissions  

## Security Design

- Modern authentication  
- No stored credentials or secrets  
- Strict error handling and safe trustee matching  
- Automatic session cleanup  
- Output contains only mailbox identifiers and removal status  

## Output

One PSCustomObject per mailbox, including:

- Mailbox input and resolved address  
- User input and resolved address  
- Flags showing whether each permission was removed or not present  
- Success state and any error message  

## Common Issues / Limitations

- Incorrect mailbox identifiers may fail resolution  
- Insufficient permissions may block some removals  
- Local script execution policy may prevent launch  
- Large mailbox lists may increase runtime  

## Recommended Usage

1. Populate Emails.txt with one mailbox identity per line  
2. Run the script with `-User`  
3. Choose which permissions to remove when prompted  
4. Review the result output for any failed mailboxes  

## **Disclaimer**

This script is provided as-is for administrative use. Review usage and outputs according to your organization’s security and compliance policies.
