# Shared Mailbox Permission Manager

An administrative PowerShell 7 GUI application for adding or removing shared mailbox permissions in Exchange Online.  
The script performs write operations to Exchange Online and provides a controlled, logged, and security‑aware workflow for managing Full Access, Send As, and Send On Behalf permissions at scale.


## Included Scripts

### SharedMailboxPermissions.ps1

A single WPF‑based tool that consolidates multiple mailbox permission operations into one interface.

**Purpose:**
- Connects interactively to Exchange Online
- Adds or removes Full Access, Send As, and Send On Behalf permissions
- Processes multiple mailboxes and users in bulk
- Provides real‑time progress, activity logs, and CSV export
- Automatically disconnects and resets on idle, minimize, or workstation lock

**Behavior:**

1. Validates PowerShell 7 and relaunches in STA mode if required  
2. Loads a WPF GUI with fields for:
   - Operation mode (Add or Remove)
   - Permission types
   - Manual or file‑based mailbox and user lists  
3. Establishes an Exchange Online session using modern authentication  
4. Executes permission changes in a background STA runspace  
5. Performs idempotency checks:
   - Skips existing permissions when adding  
   - Skips absent permissions when removing  
6. Displays console‑style logs and result records in a DataGrid  
7. Exports results to a configurable CSV file  
8. Automatically disconnects and clears state when:
   - Idle timeout is reached  
   - The window is minimized  
   - The workstation is locked  
9. Tears down all runspaces and sessions on exit  


## Requirements

- PowerShell 7 (mandatory; script enforces this)
- ExchangeOnlineManagement module
- Interactive modern authentication to Exchange Online
- Appropriate administrative permissions to modify mailbox delegation
- Desktop environment capable of running WPF


## Security Design

- Modern authentication only; no basic auth, stored credentials, or secrets  
- Error messages sanitized to remove:
  - UPNs
  - GUIDs and object IDs
  - Hostnames
  - IP addresses  
- Idle timeout (default 10 minutes) triggers automatic disconnect and data purge  
- Auto‑disconnect on minimize or workstation lock to prevent unintended exposure  
- Exchange Online actions run in an isolated STA runspace  
- Explicit disconnect logic and session cleanup  
- All output stored only in memory until exported by the user  


## Output

The script generates:

- **PermissionResults.csv** (default name, configurable)  
  Contains:
  - Mailbox  
  - User  
  - Operation  
  - FullAccessResult  
  - SendAsResult  
  - SendOnBehalfResult  
  - Status  

- Activity log visible in the GUI  
- Real‑time progress and action messages  


## Common Issues / Limitations

- Script cannot run in Windows PowerShell 5.1  
- Requires admin rights to change mailbox permissions  
- Large mailbox/user combinations may increase processing time  
- Exchange Online propagation delays can affect immediate visibility of changes  
- WPF interface requires a full desktop environment  


## Recommended Usage

1. Launch the script in PowerShell 7  
2. Click **Connect** and complete interactive sign‑in  
3. Enter or load mailbox and user lists  
4. Choose Add or Remove mode and select permission types  
5. Click **Execute** to begin processing  
6. Export results once processing completes  


## Disclaimer

Scripts are provided as‑is for administrative automation.  
Review functionality, outputs, and security implications according to your organization’s policies and compliance requirements.
