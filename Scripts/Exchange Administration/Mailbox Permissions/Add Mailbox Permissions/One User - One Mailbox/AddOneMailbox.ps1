<#
.SYNOPSIS
    Connects securely to Exchange Online and grants Full Access, Send As,
    and optionally Send On Behalf permissions to a user on a mailbox.

.DESCRIPTION
    This script:
      1. Connects to Exchange Online using modern authentication.
      2. Resolves the target mailbox and user/recipient.
      3. Interactively prompts whether to ADD:
            - Full Access
            - Send As
            - Send On Behalf
      4. Grants, for the selected access types:
            - Full Access (with AutoMapping) if missing.
            - Send As if missing.
            - SendOnBehalf rights if -GrantSendOnBehalf is specified and missing.

    The script is idempotent: it checks for existing permissions and only adds
    what is missing, so it is safe to run multiple times.

.PARAMETER SharedMailbox
    SMTP address, alias, or UPN of the mailbox to which permissions will be granted.
    Example: group@domain.com

.PARAMETER User
    User who should receive the permissions (alias, UPN, SAM, or SMTP).
    Example: UPN or user@domain.com

.PARAMETER GrantSendOnBehalf
    When specified, the script is allowed to grant SendOnBehalfTo on the mailbox
    for the user. The actual granting also depends on the interactive prompt.

.EXAMPLE
    .\Add-MailboxPermissions.ps1 -SharedMailbox "group@domain.com" -User "UPN"

.EXAMPLE
    .\Add-MailboxPermissions.ps1 -SharedMailbox "group@domain.com" -User "UPN" -GrantSendOnBehalf -Verbose

.EXAMPLE
    .\Add-MailboxPermissions.ps1 -SharedMailbox "group@domain.com" -User "UPN" -WhatIf -Verbose

.NOTES
    - Uses Connect-ExchangeOnline with modern authentication.
    - Assumes the ExchangeOnlineManagement module is already installed.
    - Recommended for Windows PowerShell 5.1 or PowerShell 7+.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SharedMailbox,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$User,

    [Parameter(Mandatory = $false)]
    [switch]$GrantSendOnBehalf,

    [Parameter(Mandatory = $false)]
    [switch]$DisconnectWhenDone
)

begin {
    #====================================================================
    # Section 0: Initial Setup & Helper Functions
    #====================================================================

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Write-Verbose "Starting Add-MailboxPermissions script for mailbox '$SharedMailbox' and user '$User'."

    function New-PermissionResult {
        param(
            [string]$Mailbox,
            [string]$User,
            [bool]$FullAccessGranted,
            [bool]$FullAccessAlreadyPresent,
            [bool]$SendAsGranted,
            [bool]$SendAsAlreadyPresent,
            [bool]$SendOnBehalfGranted,
            [bool]$SendOnBehalfAlreadyPresent
        )

        [PSCustomObject]@{
            Mailbox                    = $Mailbox
            User                       = $User
            FullAccessGranted          = $FullAccessGranted
            FullAccessAlreadyPresent   = $FullAccessAlreadyPresent
            SendAsGranted              = $SendAsGranted
            SendAsAlreadyPresent       = $SendAsAlreadyPresent
            SendOnBehalfGranted        = $SendOnBehalfGranted
            SendOnBehalfAlreadyPresent = $SendOnBehalfAlreadyPresent
        }
    }

    function Read-YesNo {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Message
        )

        while ($true) {
            $response = Read-Host "$Message (Y/N)"
            if ([string]::IsNullOrWhiteSpace($response)) {
                continue
            }

            $normalized = $response.Trim().ToUpperInvariant()

            if ($normalized -eq 'Y' -or $normalized -eq 'YES') {
                return $true
            }
            elseif ($normalized -eq 'N' -or $normalized -eq 'NO') {
                return $false
            }
            else {
                Write-Host "Please enter 'Y' or 'N'." -ForegroundColor Yellow
            }
        }
    }

    #====================================================================
    # Section 1: Interactive Permission Selection
    #====================================================================

    Write-Host ""
    Write-Host "Permission addition selection for mailbox '$SharedMailbox' and user '$User':" -ForegroundColor Cyan

    $script:AddFullAccess   = Read-YesNo -Message "Do you want to ADD Full Access"
    $script:AddSendAs       = Read-YesNo -Message "Do you want to ADD Send As"
    $script:AddSendOnBehalf = Read-YesNo -Message "Do you want to ADD Send On Behalf"

    Write-Host ""
    Write-Host "Summary of selected additions:" -ForegroundColor Cyan
    Write-Host ("  Full Access     : {0}" -f ($(if ($script:AddFullAccess)   { 'Yes' } else { 'No' })))
    Write-Host ("  Send As         : {0}" -f ($(if ($script:AddSendAs)       { 'Yes' } else { 'No' })))
    Write-Host ("  Send On Behalf  : {0}" -f ($(if ($script:AddSendOnBehalf) { 'Yes' } else { 'No' })))
    Write-Host ""
}

process {
    #====================================================================
    # Section 2: Connect to Exchange Online & Resolve Objects
    #====================================================================

    $exoConnected = $false

    try {
        Write-Verbose "Connecting to Exchange Online using modern authentication..."
        # Secure, modern auth connection to Exchange Online 
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $exoConnected = $true
        Write-Verbose "Connected to Exchange Online."

        Write-Verbose "Resolving mailbox '$SharedMailbox'..."
        $mailbox = Get-Mailbox -Identity $SharedMailbox -ErrorAction Stop

        if (-not $mailbox) {
            throw "Mailbox '$SharedMailbox' could not be found in Exchange Online."
        }

        Write-Verbose "Mailbox resolved as '$($mailbox.Identity)' (Primary SMTP: $($mailbox.PrimarySmtpAddress))."

        Write-Verbose "Resolving user '$User'..."
        $userRecipient = Get-Recipient -Identity $User -ErrorAction Stop

        if (-not $userRecipient) {
            throw "User '$User' could not be resolved as a recipient in Exchange Online."
        }

        Write-Verbose "User resolved as '$($userRecipient.Identity)' (Primary SMTP: $($userRecipient.PrimarySmtpAddress))."

        $mailboxId = $mailbox.Identity
        $userId    = $userRecipient.Identity

        #================================================================
        # Section 3: Initialize Result Flags
        #====================================================================

        $fullAccessGranted = $false
        $fullAccessAlready = $false
        $sendAsGranted     = $false
        $sendAsAlready     = $false
        $sobGranted        = $false
        $sobAlready        = $false

        #================================================================
        # Section 4: Full Access (conditional on prompt)
        #====================================================================

        if ($script:AddFullAccess) {
            Write-Verbose "Checking existing Full Access permissions on '$mailboxId' for '$userId'..."
            $existingFullAccess = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue

            if ($existingFullAccess -and
                $existingFullAccess.AccessRights -contains 'FullAccess' -and
                -not $existingFullAccess.IsInherited) {

                Write-Verbose "Full Access already present for '$userId' on '$mailboxId'."
                $fullAccessAlready = $true
            }
            else {
                if ($PSCmdlet.ShouldProcess($mailboxId, "Grant Full Access to '$userId'")) {
                    Write-Verbose "Granting Full Access for '$userId' on '$mailboxId'..."
                    Add-MailboxPermission -Identity $mailboxId -User $userId -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                    $fullAccessGranted = $true
                    Write-Verbose "Full Access granted."
                }
            }
        }
        else {
            Write-Verbose "Full Access addition not selected; skipping."
        }

        #================================================================
        # Section 5: Send As (conditional on prompt)
        #====================================================================

        if ($script:AddSendAs) {
            Write-Verbose "Checking existing Send As permissions on '$mailboxId' for '$userId'..."

            $existingSendAs = Get-RecipientPermission -Identity $mailboxId -ErrorAction SilentlyContinue |
                              Where-Object { $_.Trustee -eq $userId -and $_.AccessRights -contains 'SendAs' }

            if ($existingSendAs) {
                Write-Verbose "Send As already present for '$userId' on '$mailboxId'."
                $sendAsAlready = $true
            }
            else {
                if ($PSCmdlet.ShouldProcess($mailboxId, "Grant Send As to '$userId'")) {
                    Write-Verbose "Granting Send As for '$userId' on '$mailboxId'..."
                    Add-RecipientPermission -Identity $mailboxId -Trustee $userId -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    $sendAsGranted = $true
                    Write-Verbose "Send As granted."
                }
            }
        }
        else {
            Write-Verbose "Send As addition not selected; skipping."
        }

        #================================================================
        # Section 6: Send On Behalf (logic updated to rely on prompt and work reliably)
        #====================================================================

        if ($script:AddSendOnBehalf) {
            if ($GrantSendOnBehalf.IsPresent) {
                Write-Verbose "GrantSendOnBehalf switch specified and Send On Behalf addition selected. Proceeding to grant SendOnBehalf."
            }
            else {
                Write-Verbose "Send On Behalf addition selected in prompt (Y), proceeding even though -GrantSendOnBehalf switch was not specified."
            }

            # Re-query mailbox to ensure we have the latest GrantSendOnBehalfTo list
            $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

            $currentSobList = @()
            if ($mailboxForSob.GrantSendOnBehalfTo) {
                # Normalize entries to strings (typically distinguished names)
                $currentSobList = @($mailboxForSob.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() })
            }

            $userDn = $userRecipient.DistinguishedName

            if ($currentSobList -contains $userDn) {
                Write-Verbose "SendOnBehalf already present for '$userId' on '$mailboxId'."
                $sobAlready = $true
            }
            else {
                if ($PSCmdlet.ShouldProcess($mailboxId, "Grant SendOnBehalf to '$userId'")) {
                    Write-Verbose "Granting SendOnBehalf for '$userId' on '$mailboxId'..."

                    $newSobList = @()
                    if ($currentSobList.Count -gt 0) {
                        $newSobList += $currentSobList
                    }
                    $newSobList += $userDn

                    Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                    $sobGranted = $true
                    Write-Verbose "SendOnBehalf granted."
                }
            }
        }
        else {
            if ($GrantSendOnBehalf.IsPresent) {
                Write-Verbose "GrantSendOnBehalf specified but Send On Behalf addition not selected in prompt; skipping SendOnBehalf."
            }
            else {
                Write-Verbose "Send On Behalf addition not selected; skipping."
            }
        }

        #================================================================
        # Section 7: Output Summary Object
        #====================================================================

        $result = New-PermissionResult `
            -Mailbox $mailbox.PrimarySmtpAddress `
            -User $userRecipient.PrimarySmtpAddress `
            -FullAccessGranted $fullAccessGranted `
            -FullAccessAlreadyPresent $fullAccessAlready `
            -SendAsGranted $sendAsGranted `
            -SendAsAlreadyPresent $sendAsAlready `
            -SendOnBehalfGranted $sobGranted `
            -SendOnBehalfAlreadyPresent $sobAlready

        Write-Output $result
    }
    catch {
        Write-Error "An error occurred while granting permissions to '$User' on '$SharedMailbox': $($_.Exception.Message)"
        throw
    }
    finally {
        if ($script:ExoConnected) {
            Write-Verbose "Disconnecting from Exchange Online..."
            try {
                Disconnect-ExchangeOnline -Confirm:$false
            }
            catch {
                Write-Warning "An error occurred while disconnecting from Exchange Online: $($_.Exception.Message)"
            }
            Write-Verbose "Disconnected from Exchange Online."
        }
    }
}

end {
    Write-Verbose "Add-MailboxPermissions script execution completed."
}
