<#
SYNOPSIS
Checks mailbox activity in Unified Audit Logs for the last 90 days and exports inactive mailboxes to CSV.

DESCRIPTION
This script:
1. Automatically uses ALLUSERS.csv located in the same folder as this script
   as the default input file. No prompts for time frame; lookback is fixed
   to 90 days by default.
2. Reads the 'UserPrincipalName' column from the CSV.
3. Connects to Exchange Online and the Compliance/Unified Audit endpoint using
   interactive sign-in.
4. Resolves DisplayName for each mailbox from Exchange Online.
5. For each UPN, searches Unified Audit Logs for the last N days (default: 90)
   for any of these operations:

   MailboxLogin
   FolderBind
   MessageBind
   SendAs
   SendOnBehalf
   Create
   Update
   SoftDelete
   HardDelete
   Move
   MoveToDeletedItems
   CalendarSharing
   MeetingResponse
   MeetingCancel
   MeetingForward
   MeetingDeleted

6. Exports ONLY mailboxes with NO activity in that period to '90DayInactiveAudit.csv'
   in the script directory (or a custom path if provided), including DisplayName.

PARAMETER InputCsvPath
Optional. Path to the input CSV file that contains a 'UserPrincipalName' column.
Defaults to ALLUSERS.csv located in the same folder as this script.

PARAMETER OutputCsvPath
Optional. Path to the output CSV file.
Defaults to '90DayInactiveAudit.csv' in the same folder as this script.

PARAMETER LookbackDays
Optional. Number of days to look back from the time the script is run.
Defaults to 90. The script will NOT prompt; it just uses this value.

EXAMPLE
.\SharedMailboxAudit.ps1

Uses ALLUSERS.csv from the script folder, checks the last 90 days, and writes
90DayInactiveAudit.csv in the same folder.

EXAMPLE
.\SharedMailboxAudit.ps1 -OutputCsvPath "C:\Temp\Inactive.csv"

Uses ALLUSERS.csv from the script folder as input and writes inactive mailboxes
(with DisplayName) to C:\Temp\Inactive.csv.

EXAMPLE
.\SharedMailboxAudit.ps1 -InputCsvPath "C:\Lists\SharedMailboxes.csv" -Verbose

Overrides the default input file and shows detailed progress, still using 90 days lookback.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath 'ALLUSERS.csv'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath = (Join-Path -Path $PSScriptRoot -ChildPath '90DayInactiveAudit.csv'),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$LookbackDays = 90
)

#====================================================================
# Section 1: Initial Setup
#====================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Starting mailbox audit activity check script."
Write-Verbose "Input CSV: $InputCsvPath"
Write-Verbose "Output CSV: $OutputCsvPath"
Write-Verbose "Lookback days: $LookbackDays"

#====================================================================
# Section 2: Helper Functions
#====================================================================

function Test-ModuleInstalled {
    # Ensures a required module is installed and imported.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        throw "Required module '$ModuleName' is not installed. Install it using: Install-Module $ModuleName -Scope CurrentUser"
    }

    if (-not (Get-Module -Name $ModuleName)) {
        Write-Verbose "Importing module '$ModuleName'."
        Import-Module $ModuleName -ErrorAction Stop
    }
}

function Connect-ExchangeAndCompliance {
    # Connects to Exchange Online and Compliance/Unified Audit endpoints.
    [CmdletBinding()]
    param()

    Write-Verbose "Connecting to Exchange Online (interactive sign-in)..."

    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Verbose "Connected to Exchange Online."
    }
    catch {
        throw "Failed to connect to Exchange Online. Error: $_"
    }

    Write-Verbose "Connecting to Compliance/Unified Audit endpoint (interactive sign-in)..."

    try {
        Connect-IPPSSession -ErrorAction Stop | Out-Null
        Write-Verbose "Connected to Compliance/Unified Audit endpoint."
    }
    catch {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        } catch { }

        throw "Failed to connect to Compliance/Unified Audit endpoint. Error: $_"
    }
}

function Disconnect-ExchangeAndCompliance {
    # Disconnects from Exchange Online and removes compliance sessions.
    [CmdletBinding()]
    param()

    Write-Verbose "Disconnecting from Exchange Online and Compliance session."

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Error while disconnecting Exchange Online: $_"
    }

    try {
        Get-PSSession |
            Where-Object { $_.ComputerName -like '*.ps.compliance.protection.outlook.com' } |
            Remove-PSSession -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Error while removing IPP/Compliance PSSessions: $_"
    }
}

function Get-MailboxDisplayNameMap {
    # Builds a lookup table of UserPrincipalName -> DisplayName from Exchange Online.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$UserPrincipalNames
    )

    $map = @{}
    $total = $UserPrincipalNames.Count
    $index = 0

    foreach ($upn in $UserPrincipalNames) {
        $index++

        if ([string]::IsNullOrWhiteSpace($upn)) {
            Write-Verbose "Skipping empty UPN at index $index during DisplayName resolution."
            continue
        }

        $percent = [int](($index / $total) * 100)
        Write-Progress -Activity "Resolving mailbox DisplayNames" `
                       -Status "Processing $upn ($index of $total)" `
                       -PercentComplete $percent

        Write-Verbose "[$index/$total] Resolving DisplayName for '$upn'."

        try {
            # Using Get-EXOMailbox for modern EXO endpoint
            $mbx = Get-EXOMailbox -Identity $upn -ErrorAction Stop

            if ($mbx) {
                $map[$upn.ToLower()] = $mbx.DisplayName
                Write-Verbose "Resolved DisplayName '$($mbx.DisplayName)' for '$upn'."
            }
            else {
                Write-Verbose "No mailbox object returned for '$upn'."
                if (-not $map.ContainsKey($upn.ToLower())) {
                    $map[$upn.ToLower()] = $null
                }
            }
        }
        catch {
            Write-Warning "Failed to resolve mailbox for '$upn'. Error: $($_.Exception.Message)"
            if (-not $map.ContainsKey($upn.ToLower())) {
                $map[$upn.ToLower()] = $null
            }
        }
    }

    Write-Progress -Activity "Resolving mailbox DisplayNames" -Completed -Status "Completed"

    return $map
}

function Get-MailboxAuditStatus {
    # Determines mailbox audit activity for each UPN in the specified lookback window.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$UserPrincipalNames,

        [Parameter(Mandatory = $true)]
        [int]$LookbackDays,

        [Parameter(Mandatory = $true)]
        [hashtable]$DisplayNameMap
    )

    # Operations considered valid activity
    $operations = @(
        'MailboxLogin'
        'FolderBind'
        'MessageBind'
        'SendAs'
        'SendOnBehalf'
        'Create'
        'Update'
        'SoftDelete'
        'HardDelete'
        'Move'
        'MoveToDeletedItems'
        'CalendarSharing'
        'MeetingResponse'
        'MeetingCancel'
        'MeetingForward'
        'MeetingDeleted'
    )

    # Unified Audit expects UTC; compute UTC range
    $endUtc   = (Get-Date).ToUniversalTime()
    $startUtc = $endUtc.AddDays(-1 * $LookbackDays)

    Write-Verbose "Audit window (UTC): Start = $startUtc, End = $endUtc"
    Write-Verbose "Operations being checked: $($operations -join ', ')"

    $results = New-Object System.Collections.Generic.List[object]
    $total   = $UserPrincipalNames.Count
    $index   = 0

    foreach ($upn in $UserPrincipalNames) {
        $index++

        if ([string]::IsNullOrWhiteSpace($upn)) {
            Write-Verbose "Skipping empty UPN at index $index during audit check."
            continue
        }

        $percent = [int](($index / $total) * 100)
        Write-Progress -Activity "Checking mailbox audit activity" `
                       -Status "Processing $upn ($index of $total)" `
                       -PercentComplete $percent

        Write-Verbose "[$index/$total] Checking audit activity for '$upn'."

        # Lookup DisplayName (if available)
        $displayName = $null
        $lookupKey = $upn.ToLower()
        if ($DisplayNameMap.ContainsKey($lookupKey)) {
            $displayName = $DisplayNameMap[$lookupKey]
        }

        try {
            # Request only 1 record per user since we just need to know if activity exists.
            $auditRecord = Search-UnifiedAuditLog `
                -StartDate  $startUtc `
                -EndDate    $endUtc `
                -UserIds    $upn `
                -Operations $operations `
                -ResultSize 1 `
                -ErrorAction Stop

            if ($auditRecord) {
                $lastActivity = ($auditRecord | Sort-Object CreationDate -Descending | Select-Object -First 1).CreationDate

                $results.Add(
                    [PSCustomObject]@{
                        UserPrincipalName               = $upn
                        DisplayName                     = $displayName
                        HasActivityInLookbackWindow     = $true
                        LastActivity                    = $lastActivity
                        AuditQuerySucceeded             = $true
                        AuditQueryError                 = $null
                    }
                )
            }
            else {
                $results.Add(
                    [PSCustomObject]@{
                        UserPrincipalName               = $upn
                        DisplayName                     = $displayName
                        HasActivityInLookbackWindow     = $false
                        LastActivity                    = $null
                        AuditQuerySucceeded             = $true
                        AuditQueryError                 = $null
                    }
                )
            }
        }
        catch {
            Write-Warning "Audit query failed for '$upn'. Error: $($_.Exception.Message)"

            $results.Add(
                [PSCustomObject]@{
                    UserPrincipalName               = $upn
                    DisplayName                     = $displayName
                    HasActivityInLookbackWindow     = $null
                    LastActivity                    = $null
                    AuditQuerySucceeded             = $false
                    AuditQueryError                 = $_.Exception.Message
                }
            )
        }
    }

    Write-Progress -Activity "Checking mailbox audit activity" -Completed -Status "Completed"

    return $results
}

#====================================================================
# Section 3: Main Script Logic
#====================================================================

try {
    # Validate input CSV path (default is ALLUSERS.csv in the script folder)
    if (-not (Test-Path -Path $InputCsvPath -PathType Leaf)) {
        throw "Input CSV file '$InputCsvPath' does not exist or is not a file. Ensure ALLUSERS.csv is in the script folder: $PSScriptRoot"
    }

    # Ensure required module is available
    Test-ModuleInstalled -ModuleName 'ExchangeOnlineManagement'

    # Load CSV
    Write-Verbose "Importing CSV file '$InputCsvPath'."
    $csvData = Import-Csv -Path $InputCsvPath

    if (-not $csvData) {
        throw "No rows were found in '$InputCsvPath'."
    }

    if (-not ($csvData | Get-Member -Name 'UserPrincipalName' -MemberType NoteProperty)) {
        throw "The input CSV '$InputCsvPath' does not contain a 'UserPrincipalName' column."
    }

    # Extract UPNs and de-duplicate
    $upns = $csvData |
        Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName.Trim() -ne '' } |
        Select-Object -ExpandProperty UserPrincipalName -Unique

    if (-not $upns -or $upns.Count -eq 0) {
        throw "No valid UserPrincipalName values were found in '$InputCsvPath'."
    }

    Write-Verbose "Total unique UPNs loaded from CSV: $($upns.Count)"

    # Connect to Exchange Online and Compliance/Audit
    Connect-ExchangeAndCompliance

    # Build DisplayName map from Exchange Online
    Write-Verbose "Building mailbox DisplayName lookup table."
    $displayNameMap = Get-MailboxDisplayNameMap -UserPrincipalNames $upns

    # Query Unified Audit Logs for mailbox activity
    $mailboxAuditStatus = Get-MailboxAuditStatus `
        -UserPrincipalNames $upns `
        -LookbackDays $LookbackDays `
        -DisplayNameMap $displayNameMap

    # Determine inactive mailboxes (only those where query succeeded and no activity)
    $inactiveMailboxes = $mailboxAuditStatus |
        Where-Object {
            $_.AuditQuerySucceeded -eq $true -and
            $_.HasActivityInLookbackWindow -eq $false
        }

    # Export inactive mailboxes to CSV
    Write-Verbose "Preparing to export inactive mailboxes to '$OutputCsvPath'."

    if ($inactiveMailboxes.Count -gt 0) {
        $inactiveMailboxes |
            Select-Object UserPrincipalName, DisplayName, LastActivity, HasActivityInLookbackWindow |
            Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

        Write-Host "Export completed. Inactive mailboxes: $($inactiveMailboxes.Count) of $($upns.Count)." -ForegroundColor Green
        Write-Host "Output file: $OutputCsvPath"
    }
    else {
        Write-Warning "No inactive mailboxes were found in the last $LookbackDays days (or only failed queries). Exporting a header-only CSV."

        $headerOnly = [PSCustomObject]@{
            UserPrincipalName           = $null
            DisplayName                 = $null
            LastActivity                = $null
            HasActivityInLookbackWindow = $null
        }

        $headerOnly | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Empty (header-only) CSV created at: $OutputCsvPath" -ForegroundColor Yellow
    }

    # Return structured results to the pipeline for interactive usage
    Write-Verbose "Returning full mailbox audit status objects to the pipeline."
    $mailboxAuditStatus
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
}
finally {
    Disconnect-ExchangeAndCompliance
    Write-Verbose "Script completed."

}
