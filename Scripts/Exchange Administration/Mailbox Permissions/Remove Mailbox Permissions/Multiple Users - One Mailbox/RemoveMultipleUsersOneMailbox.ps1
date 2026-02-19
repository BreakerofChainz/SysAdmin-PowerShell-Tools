<#
.SYNOPSIS
    Removes Full Access, Send As, and/or Send On Behalf permissions on a mailbox
    from multiple users listed in a text file.

.DESCRIPTION
    This script:
      - Connects securely to Exchange Online using Connect-ExchangeOnline.
      - Resolves a single target mailbox (typically a shared mailbox).
      - Reads a list of users from a text file (default: Users.txt in the same directory as this script).
      - Interactively prompts whether to remove:
            * Full Access
            * Send As
            * Send On Behalf
      - For each selected permission type and each user, removes:
            * Full Access if explicitly assigned (non‑inherited),
            * Send As if assigned,
            * SendOnBehalfTo if assigned.

    The script is idempotent:
      - It checks for existing permissions and does nothing if they are already absent.
      - It is safe to re-run for the same mailbox and users.

.PARAMETER SharedMailbox
    SMTP address, alias, or UPN of the mailbox from which permissions will be removed.
    Example: group@domain.com

.PARAMETER UsersFilePath
    Optional path to a text file that contains one user identifier per line.
    If not specified, the script uses 'Users.txt' in the same directory as this script.
    Blank lines and lines starting with '#' are ignored.

.EXAMPLE
    .\Remove-MailboxPermissions-MultiUser.ps1 -SharedMailbox "group@domain.com"

.EXAMPLE
    .\Remove-MailboxPermissions-MultiUser.ps1 -SharedMailbox "group@domain.com" -Verbose

.EXAMPLE
    .\Remove-MailboxPermissions-MultiUser.ps1 `
        -SharedMailbox "group@domain.com" `
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

    Write-Verbose "Starting Remove-MailboxPermissions-MultiUser for mailbox '$SharedMailbox'."

    #================================================================
    # Section 1: Resolve Script Root & Users File Path, Connect EXO
    #====================================================================

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
    #====================================================================

    function New-PermissionResult {
        param(
            [string]$Mailbox,
            [string]$UserInput,
            [string]$UserResolved,
            [bool]$FullAccessRemoved,
            [bool]$FullAccessNotPresent,
            [bool]$SendAsRemoved,
            [bool]$SendAsNotPresent,
            [bool]$SendOnBehalfRemoved,
            [bool]$SendOnBehalfNotPresent,
            [bool]$Success,
            [string]$ErrorMessage
        )

        [PSCustomObject]@{
            Mailbox                = $Mailbox
            UserInput              = $UserInput
            UserResolved           = $UserResolved
            FullAccessRemoved      = $FullAccessRemoved
            FullAccessNotPresent   = $FullAccessNotPresent
            SendAsRemoved          = $SendAsRemoved
            SendAsNotPresent       = $SendAsNotPresent
            SendOnBehalfRemoved    = $SendOnBehalfRemoved
            SendOnBehalfNotPresent = $SendOnBehalfNotPresent
            Success                = $Success
            ErrorMessage           = $ErrorMessage
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
    #====================================================================

    Write-Host ""
    Write-Host "Permission removal selection for mailbox '$SharedMailbox':" -ForegroundColor Cyan

    $script:RemoveFullAccess   = Read-YesNo -Message "Do you want to REMOVE Full Access"
    $script:RemoveSendAs       = Read-YesNo -Message "Do you want to REMOVE Send As"
    $script:RemoveSendOnBehalf = Read-YesNo -Message "Do you want to REMOVE Send On Behalf"

    Write-Host ""
    Write-Host "Summary of selected removals:" -ForegroundColor Cyan
    Write-Host ("  Full Access     : {0}" -f ($(if ($script:RemoveFullAccess)   { 'Yes' } else { 'No' })))
    Write-Host ("  Send As         : {0}" -f ($(if ($script:RemoveSendAs)       { 'Yes' } else { 'No' })))
    Write-Host ("  Send On Behalf  : {0}" -f ($(if ($script:RemoveSendOnBehalf) { 'Yes' } else { 'No' })))
    Write-Host ""
}

process {
    #================================================================
    # Section 4: Read Users File & Resolve Shared Mailbox
    #====================================================================

    $results = @()

    try {
        $userLines = Get-Content -Path $UsersFilePath -ErrorAction Stop

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
        # Section 5: Process Each User & Remove Permissions
        #====================================================================

        foreach ($userInput in $userIdentifiers) {
            Write-Verbose "Processing user '$userInput' for permission removal..."

            $fullAccessRemoved     = $false
            $fullAccessNotPresent  = $false
            $sendAsRemoved         = $false
            $sendAsNotPresent      = $false
            $sobRemoved            = $false
            $sobNotPresent         = $false
            $success               = $false
            $errorMessage          = $null
            $resolvedUserSmtp      = $null

            try {
                $userRecipient = Get-Recipient -Identity $userInput -ErrorAction Stop
                if (-not $userRecipient) {
                    throw "User '$userInput' could not be resolved as a recipient in Exchange Online."
                }

                $userId = $userRecipient.Identity

                if ($userRecipient.PSObject.Properties['PrimarySmtpAddress'] -and
                    $userRecipient.PrimarySmtpAddress) {

                    $resolvedUserSmtp = $userRecipient.PrimarySmtpAddress.ToString()
                }

                Write-Verbose "User '$userInput' resolved as '$userId' (Primary SMTP: $resolvedUserSmtp)."

                # --- Remove Full Access ---
                if ($script:RemoveFullAccess) {
                    Write-Verbose "Checking Full Access permissions for '$userId' on '$mailboxId'..."

                    $existingFullAccess = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue

                    if ($existingFullAccess -and
                        $existingFullAccess.AccessRights -contains 'FullAccess' -and
                        -not $existingFullAccess.IsInherited) {

                        Write-Verbose "Explicit Full Access found; removing..."
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Full Access for '$userId'")) {
                            Remove-MailboxPermission -Identity $mailboxId -User $userId -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                            $fullAccessRemoved = $true
                            Write-Verbose "Full Access removed."
                        }
                    }
                    else {
                        Write-Verbose "No explicit Full Access found for '$userId'; skipping removal."
                        $fullAccessNotPresent = $true
                    }
                }
                else {
                    Write-Verbose "Full Access removal not selected; skipping for this user."
                }

                # --- Remove Send As (robust candidate-based removal) ---
                if ($script:RemoveSendAs) {
                    Write-Verbose "Attempting Send As removal for '$userId' on '$mailboxId'..."

                    $candidateTrustees = @()

                    if ($userId)           { $candidateTrustees += $userId }
                    if ($resolvedUserSmtp) { $candidateTrustees += $resolvedUserSmtp }

                    $props = $userRecipient.PSObject.Properties
                    foreach ($propName in 'UserPrincipalName','Alias','Name') {
                        if ($props[$propName] -and $props[$propName].Value) {
                            $candidateTrustees += $props[$propName].Value
                        }
                    }

                    # Deduplicate and filter empties
                    $candidateTrustees = $candidateTrustees |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                        Select-Object -Unique

                    if (-not $candidateTrustees -or $candidateTrustees.Count -eq 0) {
                        Write-Verbose "No candidate trustee values generated for '$userId'; skipping Send As removal."
                        $sendAsNotPresent = $true
                    }
                    else {
                        $anyRemoved = $false

                        foreach ($candidate in $candidateTrustees) {
                            Write-Verbose "Trying to remove Send As using trustee '$candidate'..."

                            if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Send As for '$candidate'")) {
                                try {
                                    Remove-RecipientPermission -Identity $mailboxId `
                                                               -Trustee $candidate `
                                                               -AccessRights SendAs `
                                                               -Confirm:$false `
                                                               -ErrorAction Stop

                                    Write-Verbose "Send As removed for trustee '$candidate'."
                                    $anyRemoved    = $true
                                    $sendAsRemoved = $true
                                }
                                catch {
                                    # Most likely "object not found" or no matching permission; safe to ignore for this candidate
                                    Write-Verbose "No Send As permission found (or removable) for trustee '$candidate': $($_.Exception.Message)"
                                }
                            }
                        }

                        if (-not $anyRemoved) {
                            Write-Verbose "No Send As permissions were removed for '$userId'; treating as not present."
                            $sendAsNotPresent = $true
                        }
                    }
                }
                else {
                    Write-Verbose "Send As removal not selected; skipping for this user."
                }

                # --- Remove Send On Behalf ---
                if ($script:RemoveSendOnBehalf) {
                    Write-Verbose "Send On Behalf removal selected; checking current configuration..."

                    $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

                    $currentSobIdentities = @()
                    if ($mailboxForSob.GrantSendOnBehalfTo) {
                        foreach ($entry in $mailboxForSob.GrantSendOnBehalfTo) {
                            try {
                                $sobRecipient = Get-Recipient -Identity $entry -ErrorAction Stop
                                if ($sobRecipient -and
                                    $sobRecipient.PSObject.Properties['PrimarySmtpAddress'] -and
                                    $sobRecipient.PrimarySmtpAddress) {

                                    $currentSobIdentities += $sobRecipient.PrimarySmtpAddress.ToString().ToLowerInvariant()
                                }
                            }
                            catch {
                                Write-Warning "Failed to resolve existing SendOnBehalf entry '$entry': $($_.Exception.Message)"
                            }
                        }
                    }

                    if (-not $resolvedUserSmtp) {
                        Write-Verbose "User '$userId' does not have a PrimarySmtpAddress; skipping SendOnBehalf removal for this user."
                        $sobNotPresent = $true
                    }
                    else {
                        $userSobIdentity = $resolvedUserSmtp.ToLowerInvariant()

                        if ($currentSobIdentities -contains $userSobIdentity) {
                            Write-Verbose "SendOnBehalf is currently assigned for '$userSobIdentity'; removing..."
                            if ($PSCmdlet.ShouldProcess($mailboxId, "Remove SendOnBehalf for '$userSobIdentity'")) {
                                $newSobList = $currentSobIdentities | Where-Object { $_ -ne $userSobIdentity }
                                Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                                $sobRemoved = $true
                                Write-Verbose "SendOnBehalf removed for '$userSobIdentity'."
                            }
                        }
                        else {
                            Write-Verbose "SendOnBehalf is not assigned for '$userSobIdentity'; skipping removal."
                            $sobNotPresent = $true
                        }
                    }
                }
                else {
                    Write-Verbose "Send On Behalf removal not selected; skipping for this user."
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
                -FullAccessRemoved $fullAccessRemoved `
                -FullAccessNotPresent $fullAccessNotPresent `
                -SendAsRemoved $sendAsRemoved `
                -SendAsNotPresent $sendAsNotPresent `
                -SendOnBehalfRemoved $sobRemoved `
                -SendOnBehalfNotPresent $sobNotPresent `
                -Success $success `
                -ErrorMessage $errorMessage
        }

        $results
    }
    catch {
        Write-Error "An error occurred in Remove-MailboxPermissions-MultiUser: $($_.Exception.Message)"
        throw
    }
    finally {
        #================================================================
        # Section 6: Disconnect from Exchange Online
        #====================================================================

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
    Write-Verbose "Remove-MailboxPermissions-MultiUser script execution completed."
}

