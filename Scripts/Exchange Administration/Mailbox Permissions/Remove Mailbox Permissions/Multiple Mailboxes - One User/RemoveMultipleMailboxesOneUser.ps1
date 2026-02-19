<#
.SYNOPSIS
    Removes Full Access, Send As, and/or Send On Behalf permissions for one user from multiple mailboxes listed in a text file.

.DESCRIPTION
    This script:
      - Connects securely to Exchange Online using Connect-ExchangeOnline (modern authentication).
      - Resolves a single target user/recipient.
      - Reads a list of mailbox identities from a text file (default: Emails.txt in the script directory).
      - Interactively prompts whether to remove the following permissions:
            * Full Access
            * Send As
            * Send On Behalf
      - For each selected permission type and each mailbox, removes if present:
            * Full Access (non-inherited),
            * Send As,
            * SendOnBehalfTo entry.

    Trustee / identity matching:
      - For Send As and Send On Behalf, the script only uses:
            * recipient Identity
            * the UPN/value passed into -User
            * Primary SMTP address
            * Alias
      - It attempts removals / matches using these values only.
      - If a trustee / ACE is not found or is duplicate/mismatched, it quietly skips and continues
        (no fatal errors because of missing or duplicate trustees).

    The script is idempotent:
      - Safe to re-run for the same user and mailboxes.
      - If a permission is already absent, it is treated as "not present" and processing continues.

.PARAMETER User
    User whose permissions should be removed (alias, UPN, SAM, or SMTP).
    Example: UPN or user@domain.com

.PARAMETER EmailsFilePath
    Optional path to a text file that contains one mailbox identity per line
    (SMTP address, alias, or UPN).
    If not specified, the script uses 'Emails.txt' in the same directory as this script.
    Blank lines and lines starting with '#' are ignored.

.EXAMPLE
    .\Remove-MailboxPermissions-MultiMailbox.ps1 -User "UPN"

.EXAMPLE
    .\Remove-MailboxPermissions-MultiMailbox.ps1 -User "UPN" -Verbose

.EXAMPLE
    .\Remove-MailboxPermissions-MultiMailbox.ps1 -User "UPN" -EmailsFilePath "C:\Temp\SharedMailboxes.txt" -WhatIf

.NOTES
    - Assumes ExchangeOnlineManagement is installed.
    - Designed for Exchange Online (EXO).
    - Recommended for Windows PowerShell 5.1 or PowerShell 7+.
    - Example shared/group mailbox format: group@domain.com
#>

#====================================================================
# Section 0: Parameters & Initial Configuration
#====================================================================

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$User,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailsFilePath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Write-Verbose "Starting Remove-MailboxPermissions-MultiMailbox for user '$User'."

    #================================================================
    # Section 1: Resolve Script Root & Emails File Path, Connect EXO
    #================================================================

    # Resolve script root folder
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

    # Resolve EmailsFilePath (default to Emails.txt in script root if not provided)
    try {
        if (-not $PSBoundParameters.ContainsKey('EmailsFilePath')) {
            $EmailsFilePath = Join-Path -Path $scriptRoot -ChildPath 'Emails.txt'
            Write-Verbose "EmailsFilePath not specified. Using default: $EmailsFilePath"
        }
        else {
            $EmailsFilePath = (Resolve-Path -Path $EmailsFilePath -ErrorAction Stop).ProviderPath
            Write-Verbose "Using custom EmailsFilePath: $EmailsFilePath"
        }
    }
    catch {
        Write-Error "Failed to resolve EmailsFilePath. Error: $($_.Exception.Message)"
        throw
    }

    # Ensure Emails file exists; create and instruct user if missing
    if (-not (Test-Path -Path $EmailsFilePath)) {
        try {
            Write-Verbose "Emails file not found. Creating a new file at: $EmailsFilePath"
            New-Item -Path $EmailsFilePath -ItemType File -Force | Out-Null
            Write-Warning "A new Emails file has been created at: $EmailsFilePath"
            Write-Warning "Add one mailbox identity per line (e.g. 'sharedmbx@domain.com' or alias) and re-run the script."
        }
        catch {
            Write-Error "Failed to create emails file at '$EmailsFilePath'. Error: $($_.Exception.Message)"
        }

        throw "Emails file was missing and has been created at '$EmailsFilePath'. Populate it with mailbox identities and re-run the script."
    }

    Write-Verbose "Emails file path in use: '$EmailsFilePath'."

    # Connect to Exchange Online using modern authentication
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

    # Creates a standardized result object for each mailbox processed
    function New-PermissionResult {
        param(
            [string]$MailboxInput,
            [string]$MailboxResolved,
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
            MailboxInput           = $MailboxInput
            MailboxResolved        = $MailboxResolved
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

    # Prompts user for a Yes/No answer and returns $true or $false
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
    Write-Host "Permission removal selection for user '$User':" -ForegroundColor Cyan

    $script:RemoveFullAccess   = Read-YesNo -Message "Do you want to remove Full Access"
    $script:RemoveSendAs       = Read-YesNo -Message "Do you want to remove Send As"
    $script:RemoveSendOnBehalf = Read-YesNo -Message "Do you want to remove Send On Behalf"

    Write-Host ""
    Write-Host "Summary of selected permission removals:" -ForegroundColor Cyan
    Write-Host ("  Full Access     : {0}" -f ($(if ($script:RemoveFullAccess)   { 'Yes' } else { 'No' })))
    Write-Host ("  Send As         : {0}" -f ($(if ($script:RemoveSendAs)       { 'Yes' } else { 'No' })))
    Write-Host ("  Send On Behalf  : {0}" -f ($(if ($script:RemoveSendOnBehalf) { 'Yes' } else { 'No' })))
    Write-Host ""
}

process {
    #================================================================
    # Section 4: Read Mailbox List & Resolve Target User
    #================================================================

    $results = @()

    try {
        # Read and filter mailbox identifiers from file
        $emailLines = Get-Content -Path $EmailsFilePath -ErrorAction Stop

        $mailboxIdentifiers = foreach ($line in $emailLines) {
            $trimmed = $line.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed) -and -not $trimmed.StartsWith('#')) {
                $trimmed
            }
        }

        if (-not $mailboxIdentifiers -or @($mailboxIdentifiers).Count -eq 0) {
            throw "No valid mailbox entries found in file '$EmailsFilePath'."
        }

        Write-Verbose "Loaded $(@($mailboxIdentifiers).Count) mailbox entries from file."

        # Resolve target user once
        Write-Verbose "Resolving user '$User'..."
        $userRecipient = Get-Recipient -Identity $User -ErrorAction Stop

        if (-not $userRecipient) {
            throw "User '$User' could not be resolved as a recipient in Exchange Online."
        }

        $userId           = $userRecipient.Identity
        $userResolvedSmtp = $userRecipient.PrimarySmtpAddress.ToString()
        $userAlias        = $userRecipient.Alias

        Write-Verbose "User resolved as '$userId' (Primary SMTP: $userResolvedSmtp, Alias: $userAlias)."

        # Build candidate trustee values used for Send As and Send On Behalf
        # Only: Identity, UPN (original input), SMTP, Alias
        $trusteeCandidates = @(
            $userId,
            $User,
            $userResolvedSmtp,
            $userAlias
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        Write-Verbose "Trustee/identity candidates for '$User': $($trusteeCandidates -join ', ')"

        #================================================================
        # Section 5: Process Each Mailbox & Remove Permissions
        #================================================================

        foreach ($mailboxInput in $mailboxIdentifiers) {
            Write-Verbose "Processing mailbox '$mailboxInput'..."

            $fullAccessRemoved      = $false
            $fullAccessNotPresent   = $false
            $sendAsRemoved          = $false
            $sendAsNotPresent       = $false
            $sobRemoved             = $false
            $sobNotPresent          = $false
            $success                = $false
            $errorMessage           = $null
            $resolvedMailboxSmtp    = $null

            try {
                # Resolve mailbox
                $mailbox = Get-Mailbox -Identity $mailboxInput -ErrorAction Stop

                if (-not $mailbox) {
                    throw "Mailbox '$mailboxInput' could not be found in Exchange Online."
                }

                $mailboxId           = $mailbox.Identity
                $resolvedMailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()

                Write-Verbose "Mailbox '$mailboxInput' resolved as '$mailboxId' (Primary SMTP: $resolvedMailboxSmtp)."

                #--------------------------------------------------------
                # Remove Full Access if selected
                #--------------------------------------------------------
                if ($script:RemoveFullAccess) {
                    Write-Verbose "Checking Full Access permissions for '$userId' on '$mailboxId'..."

                    $existingFullAccess = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue |
                                          Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }

                    if ($existingFullAccess) {
                        Write-Verbose "Full Access permission found. Removing..."
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Full Access from '$userId'")) {
                            try {
                                Remove-MailboxPermission -Identity $mailboxId -User $userId -AccessRights FullAccess -Confirm:$false -ErrorAction Stop | Out-Null
                                $fullAccessRemoved = $true
                                Write-Verbose "Full Access removed."
                            }
                            catch {
                                $faError = $_.Exception.Message
                                # If permission is already gone or ACE does not exist, treat as non-fatal
                                if ($faError -like '*The specified permission entry does not exist*' -or
                                    $faError -like '*AccessControlEntry*does not exist*') {
                                    Write-Verbose "Full Access ACE for '$userId' on '$mailboxId' not found during removal. Skipping."
                                    $fullAccessNotPresent = $true
                                }
                                else {
                                    throw
                                }
                            }
                        }
                    }
                    else {
                        Write-Verbose "Full Access permission not present. Nothing to remove."
                        $fullAccessNotPresent = $true
                    }
                }
                else {
                    Write-Verbose "Full Access removal not selected; skipping for this mailbox."
                }

                #--------------------------------------------------------
                # Remove Send As if selected (Identity/UPN/SMTP/Alias)
                #--------------------------------------------------------
                if ($script:RemoveSendAs) {
                    Write-Verbose "Attempting Send As removal for '$User' on '$mailboxId' using trustees: $($trusteeCandidates -join ', ')"

                    $anySendAsRemoved = $false

                    foreach ($trustee in $trusteeCandidates) {
                        if ([string]::IsNullOrWhiteSpace($trustee)) {
                            continue
                        }

                        Write-Verbose "Checking/removing Send As for trustee '$trustee' on '$mailboxId'..."

                        if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Send As from '$trustee'")) {
                            try {
                                Remove-RecipientPermission -Identity $mailboxId -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                                $anySendAsRemoved = $true
                                Write-Verbose "Send As removed for trustee '$trustee'."
                            }
                            catch {
                                $saError = $_.Exception.Message

                                # Expected benign errors (no ACE, invalid trustee, duplicates) - skip, do not fail mailbox
                                if ($saError -like '*AccessControlEntry*does not exist*' -or
                                    $saError -like '*The Access Control Entry being changed does not exist*' -or
                                    $saError -like '*is not a valid trustee*' -or
                                    $saError -like '*Cannot find*' -or
                                    $saError -like '*couldn''t be found*') {

                                    Write-Verbose "Send As ACE for trustee '$trustee' on '$mailboxId' not found or not applicable. Skipping."
                                }
                                else {
                                    throw
                                }
                            }
                        }
                    }

                    if ($anySendAsRemoved) {
                        $sendAsRemoved = $true
                    }
                    else {
                        Write-Verbose "No Send As ACEs removed for any trustee candidate on '$mailboxId'. Likely not present."
                        $sendAsNotPresent = $true
                    }
                }
                else {
                    Write-Verbose "Send As removal not selected; skipping for this mailbox."
                }

                #--------------------------------------------------------
                # Remove Send On Behalf if selected (Identity/UPN/SMTP/Alias)
                #--------------------------------------------------------
                if ($script:RemoveSendOnBehalf) {
                    Write-Verbose "Send On Behalf removal selected; checking current GrantSendOnBehalfTo for '$mailboxId'..."

                    # Re-query mailbox to ensure latest GrantSendOnBehalfTo list
                    $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

                    $currentSobList = @()
                    if ($mailboxForSob.GrantSendOnBehalfTo) {
                        # Normalize to string array
                        $currentSobList = @($mailboxForSob.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() })
                    }

                    if (-not $currentSobList -or @($currentSobList).Count -eq 0) {
                        Write-Verbose "No Send On Behalf entries configured on '$mailboxId'. Nothing to remove."
                        $sobNotPresent = $true
                    }
                    else {
                        Write-Verbose "Current GrantSendOnBehalfTo entries: $(@($currentSobList) -join ', ')"

                        # Determine which entries match any of our identity candidates
                        $sobMatches = @($currentSobList | Where-Object { $_ -in $trusteeCandidates })

                        if ($sobMatches -and @($sobMatches).Count -gt 0) {
                            Write-Verbose "Matched Send On Behalf entries for removal: $(@($sobMatches) -join ', ')"

                            if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Send On Behalf from '$User'")) {
                                # Build new list excluding matching entries
                                $newSobList = @($currentSobList | Where-Object { $_ -notin $trusteeCandidates })

                                try {
                                    if (@($newSobList).Count -gt 0) {
                                        Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                                    }
                                    else {
                                        # If list is empty, clear the property
                                        Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $null -ErrorAction Stop
                                    }
                                    $sobRemoved = $true
                                    Write-Verbose "Send On Behalf removed for '$User' on '$mailboxId'."
                                }
                                catch {
                                    $sobError = $_.Exception.Message

                                    # Treat duplication / already-clean scenarios as non-fatal
                                    if ($sobError -like '*GrantSendOnBehalfTo*duplicated recipient identity*' -or
                                        $sobError -like '*The specified directory object is not present*') {
                                        Write-Verbose "Non-fatal Send On Behalf update issue on '$mailboxId'. Likely already cleaned. Skipping."
                                        $sobNotPresent = $true
                                    }
                                    else {
                                        throw
                                    }
                                }
                            }
                        }
                        else {
                            Write-Verbose "No Send On Behalf entries matched identity candidates for '$User' on '$mailboxId'."
                            $sobNotPresent = $true
                        }
                    }
                }
                else {
                    Write-Verbose "Send On Behalf removal not selected; skipping for this mailbox."
                }

                $success = $true
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Error "Failed to process mailbox '$mailboxInput' for user '$User': $errorMessage"
            }

            $results += New-PermissionResult `
                -MailboxInput $mailboxInput `
                -MailboxResolved $resolvedMailboxSmtp `
                -UserInput $User `
                -UserResolved $userResolvedSmtp `
                -FullAccessRemoved $fullAccessRemoved `
                -FullAccessNotPresent $fullAccessNotPresent `
                -SendAsRemoved $sendAsRemoved `
                -SendAsNotPresent $sendAsNotPresent `
                -SendOnBehalfRemoved $sobRemoved `
                -SendOnBehalfNotPresent $sobNotPresent `
                -Success $success `
                -ErrorMessage $errorMessage
        }

        # Output all results for this run
        $results
    }
    catch {
        Write-Error "An error occurred in Remove-MailboxPermissions-MultiMailbox: $($_.Exception.Message)"
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
    Write-Verbose "Remove-MailboxPermissions-MultiMailbox script execution completed."
}
