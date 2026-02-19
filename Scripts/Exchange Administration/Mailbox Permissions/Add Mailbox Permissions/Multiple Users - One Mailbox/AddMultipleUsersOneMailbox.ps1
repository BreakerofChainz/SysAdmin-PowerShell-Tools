<#
.SYNOPSIS
    Grants Full Access, Send As, and/or Send On Behalf permissions on a mailbox to multiple users from a text file.

.DESCRIPTION
    This script:
      - Connects securely to Exchange Online using Connect-ExchangeOnline (modern authentication).
      - Resolves a single target mailbox (typically a shared mailbox).
      - Reads a list of users from a text file (default: Users.txt in the same directory as this script).
      - Interactively prompts whether to grant:
            * Full Access
            * Send As
            * Send On Behalf
      - For each selected permission type and each user, grants:
            * Full Access (with AutoMapping) if missing,
            * Send As if missing,
            * SendOnBehalfTo if missing.

    The script is idempotent:
      - It checks for existing permissions and does not create duplicates.
      - It is safe to re-run for the same mailbox and users.

.PARAMETER SharedMailbox
    SMTP address, alias, or UPN of the mailbox to which permissions will be granted.
    Example: group@domain.com

.PARAMETER UsersFilePath
    Optional path to a text file that contains one user identifier per line.
    If not specified, the script uses 'Users.txt' in the same directory as this script.
    Blank lines and lines starting with '#' are ignored.

.EXAMPLE
    .\Add-MailboxPermissions-MultiUser.ps1 -SharedMailbox "group@domain.com"

.EXAMPLE
    .\Add-MailboxPermissions-MultiUser.ps1 -SharedMailbox "group@domain.com" -Verbose

.EXAMPLE
    .\Add-MailboxPermissions-MultiUser.ps1 -SharedMailbox "group@domain.com" `
        -UsersFilePath "C:\Temp\Users.txt" -WhatIf -Verbose

.NOTES
    - Assumes ExchangeOnlineManagement is installed.
    - Designed for Exchange Online (EXO).
    - Recommended for Windows PowerShell 5.1 or PowerShell 7+.
#>

#====================================================================
# Section 0: Parameters & Initial Configuration
#====================================================================

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SharedMailbox,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$UsersFilePath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Write-Verbose "Starting Add-MailboxPermissions-MultiUser for mailbox '$SharedMailbox'."

    #================================================================
    # Section 1: Resolve Script Root & Users File Path, Connect EXO
    #================================================================

    try {
        if ($MyInvocation.MyCommand.Path) {
            $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
        }
        else {
            $scriptRoot = (Get-Location).ProviderPath
        }
        Write-Verbose "Script root resolved to: $scriptRoot"
    }
    catch {
        Write-Error "Failed to resolve script root. Error: $($_.Exception.Message)"
        throw
    }

    try {
        if (-not $PSBoundParameters.ContainsKey('UsersFilePath')) {
            $UsersFilePath = Join-Path -Path $scriptRoot -ChildPath 'Users.txt'
            Write-Verbose "UsersFilePath not specified. Using default: $UsersFilePath"
        }
        else {
            $UsersFilePath = (Resolve-Path -Path $UsersFilePath -ErrorAction Stop).ProviderPath
            Write-Verbose "Using custom UsersFilePath: $UsersFilePath"
        }
    }
    catch {
        Write-Error "Failed to resolve UsersFilePath. Error: $($_.Exception.Message)"
        throw
    }

    if (-not (Test-Path -Path $UsersFilePath)) {
        try {
            Write-Verbose "Users file not found. Creating a new file at: $UsersFilePath"
            New-Item -Path $UsersFilePath -ItemType File -Force | Out-Null
            Write-Warning "A new Users.txt file has been created at: $UsersFilePath"
            Write-Warning "Add one user identifier per line (e.g. 'user1' or 'user1@domain.com'), then re-run the script."
        }
        catch {
            Write-Error "Failed to create users file at '$UsersFilePath'. Error: $($_.Exception.Message)"
        }

        throw "Users file was missing and has been created at '$UsersFilePath'. Populate it with users and re-run the script."
    }

    Write-Verbose "Users file path in use: '$UsersFilePath'."

    $script:ExoConnected = $false
    try {
        Write-Verbose "Connecting to Exchange Online using modern authentication..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $script:ExoConnected = $true
        Write-Verbose "Connected to Exchange Online."
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        throw
    }

    #================================================================
    # Section 2: Helper Functions
    #================================================================

    function New-PermissionResult {
        param(
            [string]$Mailbox,
            [string]$UserInput,
            [string]$UserResolved,
            [bool]$FullAccessGranted,
            [bool]$FullAccessAlreadyPresent,
            [bool]$SendAsGranted,
            [bool]$SendAsAlreadyPresent,
            [bool]$SendOnBehalfGranted,
            [bool]$SendOnBehalfAlreadyPresent,
            [bool]$Success,
            [string]$ErrorMessage
        )

        [PSCustomObject]@{
            Mailbox                    = $Mailbox
            UserInput                  = $UserInput
            UserResolved               = $UserResolved
            FullAccessGranted          = $FullAccessGranted
            FullAccessAlreadyPresent   = $FullAccessAlreadyPresent
            SendAsGranted              = $SendAsGranted
            SendAsAlreadyPresent       = $SendAsAlreadyPresent
            SendOnBehalfGranted        = $SendOnBehalfGranted
            SendOnBehalfAlreadyPresent = $SendOnBehalfAlreadyPresent
            Success                    = $Success
            ErrorMessage               = $ErrorMessage
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

    #================================================================
    # Section 3: Interactive Permission Selection
    #================================================================

    Write-Host ""
    Write-Host "Permission selection for mailbox '$SharedMailbox':" -ForegroundColor Cyan

    $script:GrantFullAccess   = Read-YesNo -Message "Do you want to grant Full Access"
    $script:GrantSendAs       = Read-YesNo -Message "Do you want to grant Send As"
    $script:GrantSendOnBehalf = Read-YesNo -Message "Do you want to grant Send On Behalf"

    Write-Host ""
    Write-Host "Summary of selected permissions:" -ForegroundColor Cyan
    Write-Host ("  Full Access     : {0}" -f ($(if ($script:GrantFullAccess)   { 'Yes' } else { 'No' })))
    Write-Host ("  Send As         : {0}" -f ($(if ($script:GrantSendAs)       { 'Yes' } else { 'No' })))
    Write-Host ("  Send On Behalf  : {0}" -f ($(if ($script:GrantSendOnBehalf) { 'Yes' } else { 'No' })))
    Write-Host ""
}

process {
    #================================================================
    # Section 4: Read Users File & Resolve Shared Mailbox
    #================================================================

    $results = @()

    try {
        $userLines = Get-Content -Path $UsersFilePath -ErrorAction Stop

        # Ensure array semantics even when only one line exists
        $userIdentifiers = @(
            foreach ($line in $userLines) {
                $trimmed = $line.Trim()
                if (-not [string]::IsNullOrWhiteSpace($trimmed) -and -not $trimmed.StartsWith('#')) {
                    $trimmed
                }
            }
        )

        if (-not $userIdentifiers -or $userIdentifiers.Count -eq 0) {
            throw "No valid user entries found in file '$UsersFilePath'."
        }

        Write-Verbose "Loaded $($userIdentifiers.Count) user entries from file."

        Write-Verbose "Resolving mailbox '$SharedMailbox'..."
        $mailbox = Get-Mailbox -Identity $SharedMailbox -ErrorAction Stop

        if (-not $mailbox) {
            throw "Mailbox '$SharedMailbox' could not be found in Exchange Online."
        }

        $mailboxId          = $mailbox.Identity
        $mailboxPrimarySmtp = $mailbox.PrimarySmtpAddress.ToString()

        Write-Verbose "Mailbox resolved as '$mailboxId' (Primary SMTP: $mailboxPrimarySmtp)."

        #================================================================
        # Section 5: Process Each User & Apply Permissions
        #================================================================

        foreach ($userInput in $userIdentifiers) {
            Write-Verbose "Processing user '$userInput'..."

            $fullAccessGranted = $false
            $fullAccessAlready = $false
            $sendAsGranted     = $false
            $sendAsAlready     = $false
            $sobGranted        = $false
            $sobAlready        = $false
            $success           = $false
            $errorMessage      = $null
            $resolvedUserSmtp  = $null

            try {
                $userRecipient = Get-Recipient -Identity $userInput -ErrorAction Stop
                if (-not $userRecipient) {
                    throw "User '$userInput' could not be resolved as a recipient in Exchange Online."
                }

                $userId           = $userRecipient.Identity
                $resolvedUserSmtp = $userRecipient.PrimarySmtpAddress.ToString()

                Write-Verbose "User '$userInput' resolved as '$userId' (Primary SMTP: $resolvedUserSmtp)."

                # --- Full Access ---
                if ($script:GrantFullAccess) {
                    Write-Verbose "Checking Full Access permissions for '$userId' on '$mailboxId'..."
                    $existingFullAccess = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue

                    if ($existingFullAccess -and $existingFullAccess.AccessRights -contains 'FullAccess' -and -not $existingFullAccess.IsInherited) {
                        Write-Verbose "Full Access already present."
                        $fullAccessAlready = $true
                    }
                    else {
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Grant Full Access to '$userId'")) {
                            Write-Verbose "Granting Full Access..."
                            Add-MailboxPermission -Identity $mailboxId -User $userId -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                            $fullAccessGranted = $true
                            Write-Verbose "Full Access granted."
                        }
                    }
                }
                else {
                    Write-Verbose "Full Access not selected; skipping for this user."
                }

                # --- Send As ---
                if ($script:GrantSendAs) {
                    Write-Verbose "Checking Send As permissions for '$userId' on '$mailboxId'..."
                    $existingSendAs = Get-RecipientPermission -Identity $mailboxId -ErrorAction SilentlyContinue |
                                      Where-Object { $_.Trustee -eq $userId -and $_.AccessRights -contains 'SendAs' }

                    if ($existingSendAs) {
                        Write-Verbose "Send As already present."
                        $sendAsAlready = $true
                    }
                    else {
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Grant Send As to '$userId'")) {
                            Write-Verbose "Granting Send As..."
                            Add-RecipientPermission -Identity $mailboxId -Trustee $userId -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            $sendAsGranted = $true
                            Write-Verbose "Send As granted."
                        }
                    }
                }
                else {
                    Write-Verbose "Send As not selected; skipping for this user."
                }

                # --- Send On Behalf ---
                if ($script:GrantSendOnBehalf) {
                    Write-Verbose "Send On Behalf selected; checking current SendOnBehalf configuration..."

                    # Re-read mailbox to ensure we have the latest GrantSendOnBehalfTo list
                    $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

                    $currentSobIdentities = @()
                    if ($mailboxForSob.GrantSendOnBehalfTo) {
                        foreach ($entry in $mailboxForSob.GrantSendOnBehalfTo) {
                            try {
                                $sobRecipient = Get-Recipient -Identity $entry -ErrorAction Stop
                                if ($sobRecipient -and $sobRecipient.PrimarySmtpAddress) {
                                    $currentSobIdentities += $sobRecipient.PrimarySmtpAddress.ToString().ToLowerInvariant()
                                }
                            }
                            catch {
                                Write-Warning "Failed to resolve existing SendOnBehalf entry '$entry': $($_.Exception.Message)"
                            }
                        }
                    }

                    $userSobIdentity = $resolvedUserSmtp.ToLowerInvariant()

                    if ($currentSobIdentities -contains $userSobIdentity) {
                        Write-Verbose "SendOnBehalf already present for '$userSobIdentity'. Skipping assignment."
                        $sobAlready = $true
                    }
                    else {
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Grant SendOnBehalf to '$userSobIdentity'")) {
                            Write-Verbose "Granting SendOnBehalf..."

                            $newSobList = @()
                            if ($currentSobIdentities.Count -gt 0) {
                                $newSobList += $currentSobIdentities
                            }
                            $newSobList += $userSobIdentity

                            Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                            $sobGranted = $true
                            Write-Verbose "SendOnBehalf granted."
                        }
                    }
                }
                else {
                    Write-Verbose "Send On Behalf not selected; skipping for this user."
                }

                $success = $true
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Error "Failed to process user '$userInput' for mailbox '$SharedMailbox': $errorMessage"
            }

            $results += New-PermissionResult `
                -Mailbox $mailboxPrimarySmtp `
                -UserInput $userInput `
                -UserResolved $resolvedUserSmtp `
                -FullAccessGranted $fullAccessGranted `
                -FullAccessAlreadyPresent $fullAccessAlready `
                -SendAsGranted $sendAsGranted `
                -SendAsAlreadyPresent $sendAsAlready `
                -SendOnBehalfGranted $sobGranted `
                -SendOnBehalfAlreadyPresent $sobAlready `
                -Success $success `
                -ErrorMessage $errorMessage
        }

        $results
    }
    catch {
        Write-Error "An error occurred in Add-MailboxPermissions-MultiUser: $($_.Exception.Message)"
        throw
    }
    finally {
        #================================================================
        # Section 6: Disconnect from Exchange Online
        #================================================================

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
    Write-Verbose "Add-MailboxPermissions-MultiUser script execution completed."
}
