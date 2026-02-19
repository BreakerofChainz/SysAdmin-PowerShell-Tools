<#
.SYNOPSIS
    Grants Full Access, Send As, and/or Send On Behalf permissions for one user on multiple mailboxes from a text file.

.DESCRIPTION
    This script:
      - Connects securely to Exchange Online using Connect-ExchangeOnline (modern authentication).
      - Resolves a single target user/recipient.
      - Reads a list of mailbox identities from a text file (default: Emails.txt in the script directory).
      - Interactively prompts whether to grant the following permissions:
            * Full Access
            * Send As
            * Send On Behalf
      - For each selected permission type and each mailbox, grants:
            * Full Access (with AutoMapping) if missing,
            * Send As if missing,
            * SendOnBehalfTo if missing.

    The script is idempotent:
      - It checks for existing permissions and does not create duplicates.
      - It is safe to re-run for the same user and mailboxes.

.PARAMETER User
    User who should receive permissions (alias, UPN, SAM, or SMTP).
    Example: UPN or user@domain.com

.PARAMETER EmailsFilePath
    Optional path to a text file that contains one mailbox identity per line
    (SMTP address, alias, or UPN).
    If not specified, the script uses 'Emails.txt' in the same directory as this script.
    Blank lines and lines starting with '#' are ignored.

.EXAMPLE
    .\Add-MailboxPermissions-MultiMailbox.ps1 -User "UPN"

.EXAMPLE
    .\Add-MailboxPermissions-MultiMailbox.ps1 -User "UPN" -Verbose

.EXAMPLE
    .\Add-MailboxPermissions-MultiMailbox.ps1 -User "UPN" -EmailsFilePath "C:\Temp\SharedMailboxes.txt" -WhatIf

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
    [string]$User,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailsFilePath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Write-Verbose "Starting Add-MailboxPermissions-MultiMailbox for user '$User'."

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
            MailboxInput               = $MailboxInput
            MailboxResolved            = $MailboxResolved
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
    Write-Host "Permission selection for user '$User':" -ForegroundColor Cyan

    $script:GrantFullAccess      = Read-YesNo -Message "Do you want to grant Full Access"
    $script:GrantSendAs          = Read-YesNo -Message "Do you want to grant Send As"
    $script:GrantSendOnBehalf    = Read-YesNo -Message "Do you want to grant Send On Behalf"

    Write-Host ""
    Write-Host "Summary of selected permissions:" -ForegroundColor Cyan
    Write-Host ("  Full Access     : {0}" -f ($(if ($script:GrantFullAccess) { 'Yes' } else { 'No' })))
    Write-Host ("  Send As         : {0}" -f ($(if ($script:GrantSendAs) { 'Yes' } else { 'No' })))
    Write-Host ("  Send On Behalf  : {0}" -f ($(if ($script:GrantSendOnBehalf) { 'Yes' } else { 'No' })))
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

        if (-not $mailboxIdentifiers -or $mailboxIdentifiers.Count -eq 0) {
            throw "No valid mailbox entries found in file '$EmailsFilePath'."
        }

        Write-Verbose "Loaded $($mailboxIdentifiers.Count) mailbox entries from file."

        # Resolve target user once
        Write-Verbose "Resolving user '$User'..."
        $userRecipient = Get-Recipient -Identity $User -ErrorAction Stop

        if (-not $userRecipient) {
            throw "User '$User' could not be resolved as a recipient in Exchange Online."
        }

        $userId           = $userRecipient.Identity
        $userResolvedSmtp = $userRecipient.PrimarySmtpAddress.ToString()

        Write-Verbose "User resolved as '$userId' (Primary SMTP: $userResolvedSmtp)."

        #================================================================
        # Section 5: Process Each Mailbox & Apply Permissions
        #================================================================

        foreach ($mailboxInput in $mailboxIdentifiers) {
            Write-Verbose "Processing mailbox '$mailboxInput'..."

            $fullAccessGranted = $false
            $fullAccessAlready = $false
            $sendAsGranted     = $false
            $sendAsAlready     = $false
            $sobGranted        = $false
            $sobAlready        = $false
            $success           = $false
            $errorMessage      = $null
            $resolvedMailboxSmtp = $null

            try {
                # Resolve mailbox
                $mailbox = Get-Mailbox -Identity $mailboxInput -ErrorAction Stop

                if (-not $mailbox) {
                    throw "Mailbox '$mailboxInput' could not be found in Exchange Online."
                }

                $mailboxId           = $mailbox.Identity
                $resolvedMailboxSmtp = $mailbox.PrimarySmtpAddress.ToString()

                Write-Verbose "Mailbox '$mailboxInput' resolved as '$mailboxId' (Primary SMTP: $resolvedMailboxSmtp)."

                # Check and grant Full Access if selected
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
                    Write-Verbose "Full Access not selected; skipping for this mailbox."
                }

                # Check and grant Send As if selected
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
                    Write-Verbose "Send As not selected; skipping for this mailbox."
                }

                # Check and optionally grant Send On Behalf if selected
                if ($script:GrantSendOnBehalf) {
                    Write-Verbose "Send On Behalf selected; checking current SendOnBehalf configuration..."

                    # Re-query mailbox to ensure latest GrantSendOnBehalfTo list
                    $mailboxForSob = Get-Mailbox -Identity $mailboxId -ErrorAction Stop

                    $currentSobList = @()
                    if ($mailboxForSob.GrantSendOnBehalfTo) {
                        # Normalize to string identities (typically DNs)
                        $currentSobList = @($mailboxForSob.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() })
                    }

                    $userDn = $userRecipient.DistinguishedName

                    if ($currentSobList -contains $userDn) {
                        Write-Verbose "SendOnBehalf already present."
                        $sobAlready = $true
                    }
                    else {
                        if ($PSCmdlet.ShouldProcess($mailboxId, "Grant SendOnBehalf to '$userId'")) {
                            Write-Verbose "Granting SendOnBehalf..."

                            # Build new list and ensure uniqueness to avoid duplicate identity errors
                            $newSobList = @()
                            if ($currentSobList.Count -gt 0) {
                                $newSobList += $currentSobList
                            }
                            $newSobList += $userDn
                            $newSobList = $newSobList | Select-Object -Unique

                            try {
                                Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                                $sobGranted = $true
                                Write-Verbose "SendOnBehalf granted."
                            }
                            catch {
                                $sobError = $_.Exception.Message
                                # If this is a duplicate-identity error, treat it as "already present" and skip
                                if ($sobError -like '*GrantSendOnBehalfTo*duplicated recipient identity*') {
                                    Write-Verbose "Duplicate SendOnBehalf identity detected for '$userId' on '$mailboxId'. Treating as already present and skipping."
                                    $sobAlready = $true
                                    # Do NOT rethrow; allow processing to continue
                                }
                                else {
                                    # Unknown error - rethrow to be handled by outer catch
                                    throw
                                }
                            }
                        }
                    }
                }
                else {
                    Write-Verbose "Send On Behalf not selected; skipping for this mailbox."
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
                -FullAccessGranted $fullAccessGranted `
                -FullAccessAlreadyPresent $fullAccessAlready `
                -SendAsGranted $sendAsGranted `
                -SendAsAlreadyPresent $sendAsAlready `
                -SendOnBehalfGranted $sobGranted `
                -SendOnBehalfAlreadyPresent $sobAlready `
                -Success $success `
                -ErrorMessage $errorMessage
        }

        # Output all results for this run
        $results
    }
    catch {
        Write-Error "An error occurred in Add-MailboxPermissions-MultiMailbox: $($_.Exception.Message)"
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
    Write-Verbose "Add-MailboxPermissions-MultiMailbox script execution completed."
}
