<#
.SYNOPSIS
Exports all Shared Mailboxes in Exchange Online to a CSV file.

.DESCRIPTION
This script connects to Exchange Online, retrieves ALL mailboxes with
RecipientTypeDetails "SharedMailbox" (no extra filtering), and exports
them to a CSV file.

The script is structured with clear section breaks, similar to other audit
scripts, to make it easy to read and maintain.

.PARAMETER OutputPath
Full path to the CSV file to create. If not supplied, a file named
"ALLUSERS.csv" will be created in the current directory.

.PARAMETER StartTranscript
If supplied, the script will start a transcript to log all console output.

.PARAMETER TranscriptPath
Optional custom path for the transcript file. If not provided but
-StartTranscript is used, a timestamped log will be created in the current
directory.

.EXAMPLE
.\ListallShared.ps1

Connects to Exchange Online and creates a timestamped CSV file in the
current directory containing all shared mailboxes.

.EXAMPLE
.\ListallShared.ps1 -OutputPath "C:\Temp\ListallShared.csv"

Connects to Exchange Online and exports all shared mailboxes to the
specified CSV path.

.EXAMPLE
.\ListallShared.ps1 -StartTranscript -Verbose

Runs the script with detailed verbose output and logs everything
to a transcript file.
#>

[CmdletBinding(SupportsShouldProcess = $false)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]
    $StartTranscript,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TranscriptPath
)

# =====================================================================
# Section 0: Script Safety & Defaults
# =====================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Starting Shared Mailbox export script."

# Build default output path if not provided
if (-not $OutputPath) {
    $fileName = "ALLUSERS.csv"
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath $fileName
}

# Ensure the target directory exists
try {
    $directory = Split-Path -Path $OutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -Path $directory)) {
        Write-Verbose "Target directory does not exist. Creating: $directory"
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}
catch {
    Write-Error "Failed to validate or create output directory for path '$OutputPath'. Error: $($_.Exception.Message)"
    return
}

# =====================================================================
# Section 1: Optional Transcript / Logging Setup
# =====================================================================

if ($StartTranscript) {
    try {
        if (-not $TranscriptPath) {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $TranscriptPath = Join-Path -Path (Get-Location) -ChildPath ("ALLUSERS.log" -f $timestamp)
        }

        Write-Verbose "Starting transcript at path: $TranscriptPath"
        Start-Transcript -Path $TranscriptPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Failed to start transcript. Error: $($_.Exception.Message)"
    }
}

# =====================================================================
# Section 2: Connect to Exchange Online
# =====================================================================

try {
    Write-Verbose "Connecting to Exchange Online..."
    # Requires Exchange Online Management module (EXO V3).
    Connect-ExchangeOnline -ShowProgress:$false -ErrorAction Stop
    Write-Verbose "Successfully connected to Exchange Online."
}
catch {
    Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
    if ($StartTranscript) {
        try {
            Stop-Transcript | Out-Null
        }
        catch {
            Write-Warning "Failed to stop transcript after connection error. Error: $($_.Exception.Message)"
        }
    }
    return
}

# =====================================================================
# Section 3: Retrieve All Shared Mailboxes
# =====================================================================

Write-Verbose "Retrieving all shared mailboxes from Exchange Online..."

$mailboxes = $null

try {
    # Get all shared mailboxes (no additional filtering)
    $mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop

    if (-not $mailboxes) {
        Write-Warning "No shared mailboxes were found in the environment."
    }
    else {
        Write-Verbose ("Retrieved {0} shared mailbox(es)." -f $mailboxes.Count)
    }
}
catch {
    Write-Error "Failed to retrieve shared mailboxes. Error: $($_.Exception.Message)"

    # Cleanup due to failure
    try {
        Write-Verbose "Disconnecting from Exchange Online due to error..."
        Disconnect-ExchangeOnline -Confirm:$false
    }
    catch {
        Write-Warning "Failed to cleanly disconnect from Exchange Online. Error: $($_.Exception.Message)"
    }

    if ($StartTranscript) {
        try {
            Stop-Transcript | Out-Null
        }
        catch {
            Write-Warning "Failed to stop transcript cleanly after retrieval error. Error: $($_.Exception.Message)"
        }
    }

    return
}

# =====================================================================
# Section 4: Shape Data for CSV Export
# =====================================================================

Write-Verbose "Shaping shared mailbox data for CSV export."

# Customize this selection to match your existing 'Shared Mailbox Audit.ps1'
# if needed. These are common, useful fields for an audit CSV.
$exportData = $mailboxes | Select-Object `
    DisplayName,
    PrimarySmtpAddress,
    Alias,
    Identity,
    OrganizationalUnit,
    RecipientTypeDetails,
    WhenCreated,
    EmailAddresses,
    ProhibitSendQuota,
    ProhibitSendReceiveQuota

# =====================================================================
# Section 5: Export to CSV
# =====================================================================

try {
    Write-Verbose "Exporting shared mailbox data to CSV: $OutputPath"

    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export complete." -ForegroundColor Green
    Write-Host "Shared mailbox CSV created at: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to export shared mailbox data to CSV. Error: $($_.Exception.Message)"
}
finally {

# =====================================================================    
#Section 6: Cleanup (Disconnect / Stop Transcript)
# =====================================================================
   

    try {
        Write-Verbose "Disconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false
    }
    catch {
        Write-Warning "Failed to cleanly disconnect from Exchange Online. Error: $($_.Exception.Message)"
    }

    if ($StartTranscript) {
        try {
            Write-Verbose "Stopping transcript."
            Stop-Transcript | Out-Null
        }
        catch {
            Write-Warning "Failed to stop transcript cleanly. Error: $($_.Exception.Message)"
        }
    }

    Write-Verbose "Shared mailbox export script finished."
}


