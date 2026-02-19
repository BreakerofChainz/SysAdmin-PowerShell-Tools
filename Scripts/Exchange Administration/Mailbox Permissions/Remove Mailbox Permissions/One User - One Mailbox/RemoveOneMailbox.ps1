<#
.SYNOPSIS
    Connects securely to Exchange Online and removes Full Access, Send As,
    and optionally Send On Behalf permissions for a user on a mailbox.

.DESCRIPTION
    This script:
      1. Connects to Exchange Online using modern authentication
         via Connect-ExchangeOnline (no -DisableWAM).
      2. Resolves the target mailbox and user/recipient.
      3. Interactively prompts whether to REMOVE:
            - Full Access
            - Send As
            - Send On Behalf
      4. Removes, for the selected access types:
            - Full Access if explicitly assigned (non-inherited).
            - Send As if assigned.
            - SendOnBehalf rights if assigned.

    The script is idempotent: it checks for existing permissions and only removes
    what is present, so it is safe to run multiple times.

.PARAMETER SharedMailbox
    SMTP address, alias, or UPN of the mailbox from which permissions will be removed.
    Example: group@domain.com

.PARAMETER User
    User who should have the permissions removed (alias, UPN, SAM, or SMTP).
    Example: UPN or user@domain.com


.EXAMPLE
    .\Remove-MailboxPermissions.ps1 -SharedMailbox "group@domain.com" -User "UPN"

.EXAMPLE
    .\Remove-MailboxPermissions.ps1 -SharedMailbox "group@domain.com" -User "UPN" -WhatIf -Verbose

.NOTES
    - Uses Connect-ExchangeOnline with modern authentication (no -DisableWAM).
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
    [switch]$DisconnectWhenDone
)

begin {
    #====================================================================
    # Section 0: Initial Setup & Helper Functions
    #====================================================================

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Write-Verbose "Starting Remove-MailboxPermissions script for mailbox '$SharedMailbox' and user '$User'."

    function New-PermissionResult {
        param(
            [string]$Mailbox,
            [string]$User,
            [bool]$FullAccessRemoved,
            [bool]$FullAccessNotPresent,
            [bool]$SendAsRemoved,
            [bool]$SendAsNotPresent,
            [bool]$SendOnBehalfRemoved,
            [bool]$SendOnBehalfNotPresent
        )

        [PSCustomObject]@{
            Mailbox                = $Mailbox
            User                   = $User
            FullAccessRemoved      = $FullAccessRemoved
            FullAccessNotPresent   = $FullAccessNotPresent
            SendAsRemoved          = $SendAsRemoved
            SendAsNotPresent       = $SendAsNotPresent
            SendOnBehalfRemoved    = $SendOnBehalfRemoved
            SendOnBehalfNotPresent = $SendOnBehalfNotPresent
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

    function Get-SafePropertyValue {
        <#
        .SYNOPSIS
            Safely gets a property value from an object even under Set-StrictMode -Version Latest.

        .DESCRIPTION
            Returns $null if the property does not exist on the object instead of throwing.

        .PARAMETER InputObject
            The object to inspect.

        .PARAMETER PropertyName
            The property name to retrieve.
        #>
        param(
            [Parameter(Mandatory = $true)]
            [object]$InputObject,

            [Parameter(Mandatory = $true)]
            [string]$PropertyName
        )

        if (-not $InputObject) {
            return $null
        }

        $props = $InputObject.PSObject.Properties
        if ($props[$PropertyName]) {
            return $props[$PropertyName].Value
        }
        else {
            return $null
        }
    }

    #====================================================================
    # Section 1: Interactive Permission Selection
    #====================================================================

    Write-Host ""
    Write-Host "Permission removal selection for mailbox '$SharedMailbox' and user '$User':" -ForegroundColor Cyan

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
    #====================================================================
    # Section 2: Connect to Exchange Online & Resolve Objects
    #====================================================================

    $exoConnected = $false

    try {
        Write-Verbose "Connecting to Exchange Online using modern authentication..."
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

        $mailboxId        = $mailbox.Identity
        $userId           = $userRecipient.Identity
        $resolvedUserSmtp = $null

        if ($userRecipient.PSObject.Properties['PrimarySmtpAddress'] -and
            $userRecipient.PrimarySmtpAddress) {

            $resolvedUserSmtp = $userRecipient.PrimarySmtpAddress.ToString()
        }

        #================================================================
        # Section 3: Initialize Result Flags
        #====================================================================

        $fullAccessRemoved     = $false
        $fullAccessNotPresent  = $false
        $sendAsRemoved         = $false
        $sendAsNotPresent      = $false
        $sobRemoved            = $false
        $sobNotPresent         = $false

        #================================================================
        # Section 4: Full Access (conditional on prompt)
        #====================================================================

        if ($script:RemoveFullAccess) {
            Write-Verbose "Checking existing Full Access permissions on '$mailboxId' for '$userId'..."
            $existingFullAccess = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue

            if ($existingFullAccess -and
                $existingFullAccess.AccessRights -contains 'FullAccess' -and
                -not $existingFullAccess.IsInherited) {

                Write-Verbose "Explicit Full Access found for '$userId'; removing..."
                if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Full Access from '$userId'")) {
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
            Write-Verbose "Full Access removal not selected; skipping."
        }

        #================================================================
        # Section 5: Send As (robust trustee resolution with safe property access)
        #====================================================================

        if ($script:RemoveSendAs) {
            Write-Verbose "Checking existing Send As permissions on '$mailboxId' for '$userId'..."

            # Get all Send As permissions on the mailbox once
            try {
                $allSendAsPerms = Get-RecipientPermission -Identity $mailboxId -ErrorAction SilentlyContinue |
                                  Where-Object { $_.AccessRights -contains 'SendAs' }
            }
            catch {
                Write-Verbose "Failed to get recipient permissions for '$mailboxId': $($_.Exception.Message)"
                $allSendAsPerms = @()
            }

            if (-not $allSendAsPerms -or $allSendAsPerms.Count -eq 0) {
                Write-Verbose "No Send As permissions exist on '$mailboxId' at all."
                $sendAsNotPresent = $true
            }
            else {
                Write-Verbose ("Found {0} Send As permission entries on mailbox '{1}'." -f $allSendAsPerms.Count, $mailboxId)

                $matchingSendAsPerms = @()

                foreach ($perm in $allSendAsPerms) {
                    if (-not $perm) { continue }

                    $trusteeString = $perm.Trustee.ToString()
                    Write-Verbose ("Evaluating Send As trustee '{0}'..." -f $trusteeString)

                    $trusteeRecipient = $null
                    try {
                        $trusteeRecipient = Get-Recipient -Identity $perm.Trustee -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Verbose ("Get-Recipient failed for trustee '{0}': {1}" -f $trusteeString, $_.Exception.Message)
                    }

                    $isMatch = $false

                    if ($trusteeRecipient) {
                        # Safely get identifiers & key properties for comparison
                        $userExtId      = Get-SafePropertyValue -InputObject $userRecipient     -PropertyName 'ExternalDirectoryObjectId'
                        $trusteeExtId   = Get-SafePropertyValue -InputObject $trusteeRecipient -PropertyName 'ExternalDirectoryObjectId'
                        $userGuid       = Get-SafePropertyValue -InputObject $userRecipient     -PropertyName 'Guid'
                        $trusteeGuid    = Get-SafePropertyValue -InputObject $trusteeRecipient -PropertyName 'Guid'
                        $userSmtp       = Get-SafePropertyValue -InputObject $userRecipient     -PropertyName 'PrimarySmtpAddress'
                        $trusteeSmtp    = Get-SafePropertyValue -InputObject $trusteeRecipient -PropertyName 'PrimarySmtpAddress'
                        $userUpn        = Get-SafePropertyValue -InputObject $userRecipient     -PropertyName 'UserPrincipalName'
                        $trusteeUpn     = Get-SafePropertyValue -InputObject $trusteeRecipient -PropertyName 'UserPrincipalName'
                        $userAlias      = Get-SafePropertyValue -InputObject $userRecipient     -PropertyName 'Alias'
                        $trusteeAlias   = Get-SafePropertyValue -InputObject $trusteeRecipient -PropertyName 'Alias'

                        # Compare by stable identifiers first (ExternalDirectoryObjectId / Guid)
                        if ($userExtId -and $trusteeExtId -and ($userExtId -eq $trusteeExtId)) {
                            $isMatch = $true
                            Write-Verbose "Matched by ExternalDirectoryObjectId."
                        }
                        elseif ($userGuid -and $trusteeGuid -and ($userGuid -eq $trusteeGuid)) {
                            $isMatch = $true
                            Write-Verbose "Matched by Guid."
                        }
                        # Fall back to SMTP / UPN / Alias comparisons (case-insensitive)
                        elseif ($userSmtp -and $trusteeSmtp -and
                                ($userSmtp.ToString().ToLowerInvariant() -eq $trusteeSmtp.ToString().ToLowerInvariant())) {
                            $isMatch = $true
                            Write-Verbose "Matched by PrimarySmtpAddress."
                        }
                        elseif ($userUpn -and $trusteeUpn -and
                                ($userUpn.ToLowerInvariant() -eq $trusteeUpn.ToLowerInvariant())) {
                            $isMatch = $true
                            Write-Verbose "Matched by UserPrincipalName."
                        }
                        elseif ($userAlias -and $trusteeAlias -and
                                ($userAlias.ToLowerInvariant() -eq $trusteeAlias.ToLowerInvariant())) {
                            $isMatch = $true
                            Write-Verbose "Matched by Alias."
                        }
                    }
                    else {
                        # If we can't resolve trustee as a recipient, last resort: string comparisons
                        $candidateStrings = @()

                        if ($userId) {
                            $candidateStrings += $userId
                        }

                        if ($resolvedUserSmtp) {
                            $candidateStrings += $resolvedUserSmtp
                        }

                        $userUpnFallback      = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'UserPrincipalName'
                        $userAliasFallback    = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'Alias'
                        $userNameFallback     = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'Name'
                        $userDnFallback       = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'DistinguishedName'

                        if ($userUpnFallback)   { $candidateStrings += $userUpnFallback }
                        if ($userAliasFallback) { $candidateStrings += $userAliasFallback }
                        if ($userNameFallback)  { $candidateStrings += $userNameFallback }
                        if ($userDnFallback)    { $candidateStrings += $userDnFallback }

                        $candidateStrings = $candidateStrings |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                            Select-Object -Unique

                        $candidateLower = $candidateStrings | ForEach-Object { $_.ToString().ToLowerInvariant() }
                        $trusteeLower   = $trusteeString.ToLowerInvariant()

                        if ($candidateLower -contains $trusteeLower) {
                            $isMatch = $true
                            Write-Verbose "Matched trustee by fallback string comparison."
                        }
                    }

                    if ($isMatch) {
                        Write-Verbose ("Trustee '{0}' identified as matching user '{1}'." -f $trusteeString, $userId)
                        $matchingSendAsPerms += $perm
                    }
                    else {
                        Write-Verbose ("Trustee '{0}' does not match user '{1}'." -f $trusteeString, $userId)
                    }
                }

                if ($matchingSendAsPerms.Count -gt 0) {
                    $trusteesToRemove = $matchingSendAsPerms |
                        Select-Object -ExpandProperty Trustee -Unique

                    foreach ($trustee in $trusteesToRemove) {
                        Write-Verbose "Send As permission found for trustee '$trustee'; removing..."
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Remove Send As from '$trustee'")) {
                            Remove-RecipientPermission -Identity $mailboxId -Trustee $trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                            $sendAsRemoved = $true
                            Write-Verbose "Send As removed for '$trustee'."
                        }
                    }
                }
                else {
                    Write-Verbose "No Send As permission entries matched the specified user; skipping removal."
                    $sendAsNotPresent = $true
                }
            }
        }
        else {
            Write-Verbose "Send As removal not selected; skipping."
        }

        #====================================================================
        # Section 6: Send On Behalf (robust trustee resolution)
        #====================================================================

        if ($script:RemoveSendOnBehalf) {
            Write-Verbose "Send On Behalf removal selected; checking current configuration..."

            # Re-query mailbox to get latest GrantSendOnBehalfTo list
            $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

            $currentSobList = @()
            if ($mailboxForSob.GrantSendOnBehalfTo) {
                $currentSobList = @($mailboxForSob.GrantSendOnBehalfTo)
            }

            if ($currentSobList.Count -eq 0) {
                Write-Verbose "No Send-On-Behalf delegates assigned to this mailbox."
                $sobNotPresent = $true
            }
            else {
                Write-Verbose ("Found {0} Send-On-Behalf entries." -f $currentSobList.Count)

                $matchingSobEntries = @()

                foreach ($sobEntry in $currentSobList) {
                    $sobString = $sobEntry.ToString()
                    Write-Verbose ("Evaluating SOB entry: '{0}'..." -f $sobString)

                    # Resolve SOB trustee to a recipient object
                    $sobRecipient = $null
                    try {
                        $sobRecipient = Get-Recipient -Identity $sobEntry -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Verbose ("Failed to resolve SOB entry '{0}': {1}" -f $sobString, $_.Exception.Message)
                    }

                    $isMatch = $false

                    if ($sobRecipient) {
                        # Safely retrieve key props
                        $userExtId    = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'ExternalDirectoryObjectId'
                        $sobExtId     = Get-SafePropertyValue -InputObject $sobRecipient  -PropertyName 'ExternalDirectoryObjectId'

                        $userGuid     = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'Guid'
                        $sobGuid      = Get-SafePropertyValue -InputObject $sobRecipient  -PropertyName 'Guid'

                        $userSmtp     = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'PrimarySmtpAddress'
                        $sobSmtp      = Get-SafePropertyValue -InputObject $sobRecipient  -PropertyName 'PrimarySmtpAddress'

                        $userUpn      = Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'UserPrincipalName'
                        $sobUpn       = Get-SafePropertyValue -InputObject $sobRecipient  -PropertyName 'UserPrincipalName'

                        # Compare unique identifiers FIRST
                        if ($userExtId -and $sobExtId -and $userExtId -eq $sobExtId) {
                            $isMatch = $true
                            Write-Verbose "Matched SOB entry by ExternalDirectoryObjectId."
                        }
                        elseif ($userGuid -and $sobGuid -and $userGuid -eq $sobGuid) {
                            $isMatch = $true
                            Write-Verbose "Matched SOB entry by Guid."
                        }
                        elseif ($userSmtp -and $sobSmtp -and
                                ($userSmtp.ToString().ToLowerInvariant() -eq $sobSmtp.ToString().ToLowerInvariant())) {
                            $isMatch = $true
                            Write-Verbose "Matched SOB entry by PrimarySmtpAddress."
                        }
                        elseif ($userUpn -and $sobUpn -and
                                ($userUpn.ToLowerInvariant() -eq $sobUpn.ToLowerInvariant())) {
                            $isMatch = $true
                            Write-Verbose "Matched SOB entry by UserPrincipalName."
                        }
                    }
                    else {
                        # Fallback string-based matching if Get-Recipient fails
                        $candidateStrings = @(
                            $userId,
                            $resolvedUserSmtp,
                            (Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'UserPrincipalName'),
                            (Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'Alias'),
                            (Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'Name'),
                            (Get-SafePropertyValue -InputObject $userRecipient -PropertyName 'DistinguishedName')
                        ) | Where-Object { $_ }

                        $candidateLower = $candidateStrings | ForEach-Object { $_.ToLowerInvariant() }
                        $sobLower = $sobString.ToLowerInvariant()

                        if ($candidateLower -contains $sobLower) {
                            $isMatch = $true
                            Write-Verbose "Matched SOB entry by fallback string comparison."
                        }
                    }

                    if ($isMatch) {
                        Write-Verbose ("SOB entry '{0}' matches user '{1}'." -f $sobString, $User)
                        $matchingSobEntries += $sobEntry
                    }
                }

                if ($matchingSobEntries.Count -gt 0) {
                    Write-Verbose ("Removing {0} matching SOB entries..." -f $matchingSobEntries.Count)

                    $newSobList = $currentSobList | Where-Object { $matchingSobEntries -notcontains $_ }

                    if ($PSCmdlet.ShouldProcess($mailboxId, "Remove SendOnBehalf from '$User'")) {
                        Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                        $sobRemoved = $true
                        Write-Verbose "Send-On-Behalf removed successfully."
                    }
                }
                else {
                    Write-Verbose "User does not appear in Send-On-Behalf list."
                    $sobNotPresent = $true
                }
            }
        }
        else {
            Write-Verbose "Send On Behalf removal not selected; skipping."
        }

        #================================================================
        # Section 7: Output Summary Object
        #====================================================================

        $result = New-PermissionResult `
            -Mailbox $mailbox.PrimarySmtpAddress `
            -User $userRecipient.PrimarySmtpAddress `
            -FullAccessRemoved $fullAccessRemoved `
            -FullAccessNotPresent $fullAccessNotPresent `
            -SendAsRemoved $sendAsRemoved `
            -SendAsNotPresent $sendAsNotPresent `
            -SendOnBehalfRemoved $sobRemoved `
            -SendOnBehalfNotPresent $sobNotPresent

        Write-Output $result
    }
    catch {
        Write-Error "An error occurred while removing permissions from '$User' on '$SharedMailbox': $($_.Exception.Message)"
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
    Write-Verbose "Remove-MailboxPermissions script execution completed."
}
