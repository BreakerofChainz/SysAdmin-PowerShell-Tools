<#
.SYNOPSIS
    Shared Mailbox Permission Manager — Interactive WPF GUI for adding and
    removing Full Access, Send As, and Send On Behalf permissions on Exchange
    Online shared mailboxes.

.DESCRIPTION
    Allows exchange administrators to add and remove users from shared mailboxes in bulk, or
    one at a time. 

    All Exchange Online operations run in a dedicated STA background runspace
    so the UI remains responsive throughout.

    Workflow:
      1. Click Connect            — Authenticates to Exchange Online via
                                    interactive modern auth (OAuth).
      2. Select Operation Mode    — Choose Add Permissions or Remove Permissions.
      3. Enter Mailbox Emails     — Type addresses directly, or load from a file
                                    (default: Emails.txt at the script root).
      4. Enter User UPNs          — Type UPNs directly, or load from a file
                                    (default: Users.txt at the script root).
      5. Select Permission Types  — Full Access, Send As, Send On Behalf.
      6. Click Execute             — Resolves identities, applies changes with
                                    idempotency checks, and logs all results.
      7. Click Export Results      — Writes a CSV report of all operations.

    Security Features:
      - PowerShell 7 required (hard gate).
      - Modern authentication only (Connect-ExchangeOnline).
      - 10-minute idle timeout with auto-disconnect and UI reset.
      - Workstation lock and window minimize detection.
      - Error messages sanitized to strip tenant-identifying details.
      - No secrets, credentials, or user data stored anywhere.

.PARAMETER DefaultEmailsFilePath
    Pre-populates the mailbox emails file path field.
    Default: Emails.txt in the script folder.

.PARAMETER DefaultUsersFilePath
    Pre-populates the user UPNs file path field.
    Default: Users.txt in the script folder.

.PARAMETER DefaultOutputPath
    Pre-populates the CSV export path field.
    Default: PermissionResults.csv in the script folder.

.PARAMETER IdleTimeoutMinutes
    Minutes of inactivity before auto-disconnect. Must be 1-120. Default: 10.

.EXAMPLE
    .\SharedMailboxPermissionManager.ps1

    Launches with default settings (10-minute idle timeout).

.EXAMPLE
    .\SharedMailboxPermissionManager.ps1 -IdleTimeoutMinutes 15 -Verbose

    Launches with a 15-minute idle timeout and verbose console logging.

.EXAMPLE
    .\SharedMailboxPermissionManager.ps1 -DefaultOutputPath "C:\Logs\Results.csv"

    Launches with a custom default export path.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DefaultEmailsFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'Emails.txt'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DefaultUsersFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'Users.txt'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DefaultOutputPath = (Join-Path -Path $PSScriptRoot -ChildPath 'PermissionResults.csv'),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 120)]
    [int]$IdleTimeoutMinutes = 10
)

#=================================================================
# Section -1:  PowerShell 7 Requirement & STA Self-Relaunch
#=================================================================
# Hard gate: PowerShell 7+ is required.  If launched from PS 5.1 or
# any non-PS7 host the script errors out immediately.
# If launched from PS7 but without -STA, it re-invokes itself under
# pwsh -STA, forwarding all bound parameters, then exits the MTA
# instance.

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error ("This application requires PowerShell 7 (pwsh.exe).`n" +
                 "Current host: PowerShell $($PSVersionTable.PSVersion)`n`n" +
                 "Install PowerShell 7 from: https://aka.ms/install-powershell")
    return
}

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne
        [System.Threading.ApartmentState]::STA) {

    Write-Host 'Relaunching under pwsh -STA...' -ForegroundColor Cyan

    $relaunchArgs = @('-STA', '-NoProfile', '-File', $MyInvocation.MyCommand.Path)

    foreach ($key in $PSBoundParameters.Keys) {
        $val = $PSBoundParameters[$key]
        if ($val -is [switch]) {
            if ($val.IsPresent) { $relaunchArgs += "-$key" }
        }
        else {
            $relaunchArgs += "-$key"
            $relaunchArgs += "$val"
        }
    }

    try {
        $pwshPath = (Get-Command pwsh -ErrorAction Stop).Source
        Start-Process -FilePath $pwshPath -ArgumentList $relaunchArgs -ErrorAction Stop
    }
    catch {
        Write-Error ("Failed to relaunch under pwsh -STA.`n`n" +
                     "Ensure pwsh.exe is installed and on your PATH.`n" +
                     "Error: $($_.Exception.Message)")
        return
    }

    return
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#=================================================================
# Section 0:  Assembly Imports
#=================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#=================================================================
# Section 1:  Synchronized State Hashtable
#=================================================================
# Shared between the STA UI thread and the background EXO runspace.
# All cross-thread communication flows exclusively through this object.

$syncHash = [hashtable]::Synchronized(@{
    Window              = $null

    # Connection and run-state flags
    IsConnected         = $false
    IsRunning           = $false
    CancelRequested     = $false

    # Persistent STA runspace for all EXO operations
    EXORunspace         = $null

    # Cached UI element references (populated in Section 4)
    StatusLabel         = $null
    ConnBadge           = $null
    ConnDot             = $null
    ConnStatusLabel     = $null
    LogConsole          = $null
    ProgressBar         = $null
    ProgressLabel       = $null
    ResultsGrid         = $null

    # Connected account identity
    ConnectedUPNLabel   = $null
    ConnectedUPN        = ''

    # Idle timer references
    IdleTimer           = $null
    LastActivityTime    = [datetime]::Now
    IdleTimeoutMinutes  = $IdleTimeoutMinutes
    IdleCountdownLabel  = $null

    # Progress staging (written by runspace, read by Dispatcher)
    CurrentProgress     = 0
    CurrentProgressMsg  = ''
    LastError           = ''

    # Operation results
    ResultObjects       = $null
})

#=================================================================
# Section 2:  XAML Window Definition
#=================================================================

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Shared Mailbox  ·  Permission Manager"
    Width="1100" Height="920"
    MinWidth="900" MinHeight="750"
    WindowStartupLocation="CenterScreen"
    Background="#F0F0F0">

  <Window.Resources>

    <Style TargetType="GroupBox">
      <Setter Property="Margin"      Value="8,3,8,3"/>
      <Setter Property="Padding"     Value="6,4,6,6"/>
      <Setter Property="Background"  Value="White"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="FontSize"    Value="12"/>
    </Style>

    <Style TargetType="Button">
      <Setter Property="Padding"         Value="14,5"/>
      <Setter Property="Margin"          Value="4,2"/>
      <Setter Property="MinWidth"        Value="115"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Background"      Value="#0078D4"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="12"/>
    </Style>

    <Style x:Key="BtnRed" TargetType="Button">
      <Setter Property="Padding"         Value="14,5"/>
      <Setter Property="Margin"          Value="4,2"/>
      <Setter Property="MinWidth"        Value="115"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Background"      Value="#C50F1F"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="12"/>
    </Style>

    <Style x:Key="BtnGray" TargetType="Button">
      <Setter Property="Padding"         Value="14,5"/>
      <Setter Property="Margin"          Value="4,2"/>
      <Setter Property="MinWidth"        Value="100"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Background"      Value="#6B6B6B"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="12"/>
    </Style>

    <Style x:Key="BtnGreen" TargetType="Button">
      <Setter Property="Padding"         Value="14,5"/>
      <Setter Property="Margin"          Value="4,2"/>
      <Setter Property="MinWidth"        Value="115"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Background"      Value="#107C10"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="12"/>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Margin"                   Value="8,4"/>
      <Setter Property="FontSize"                 Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="RadioButton">
      <Setter Property="Margin"                   Value="8,4"/>
      <Setter Property="FontSize"                 Value="13"/>
      <Setter Property="FontWeight"               Value="SemiBold"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Padding"                  Value="5,4"/>
      <Setter Property="Margin"                   Value="4,2"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush"              Value="#AAAAAA"/>
      <Setter Property="FontSize"                 Value="12"/>
    </Style>

    <Style TargetType="Label">
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Padding"                  Value="4,2"/>
      <Setter Property="FontSize"                 Value="12"/>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="AutoGenerateColumns"      Value="False"/>
      <Setter Property="AlternatingRowBackground" Value="#EEF4FF"/>
      <Setter Property="GridLinesVisibility"      Value="Horizontal"/>
      <Setter Property="BorderBrush"              Value="#CCCCCC"/>
      <Setter Property="CanUserResizeRows"        Value="False"/>
      <Setter Property="CanUserAddRows"           Value="False"/>
      <Setter Property="FontSize"                 Value="12"/>
      <Setter Property="RowHeight"                Value="26"/>
      <Setter Property="HeadersVisibility"        Value="Column"/>
      <Setter Property="SelectionMode"            Value="Extended"/>
      <Setter Property="IsReadOnly"               Value="True"/>
      <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
      <Setter Property="ScrollViewer.CanContentScroll"              Value="True"/>
    </Style>

  </Window.Resources>

  <DockPanel LastChildFill="True">

    <!-- Status bar anchored to window bottom -->
    <StatusBar DockPanel.Dock="Bottom" Background="#1A1A1A" Height="28">
      <StatusBarItem>
        <TextBlock x:Name="txtStatus" Text="Ready." Foreground="#DDDDDD" FontSize="12"/>
      </StatusBarItem>
      <Separator Background="#444"/>
      <StatusBarItem>
        <TextBlock x:Name="txtConnBadge" Text="●  Not Connected"
                   Foreground="#FF6B6B" FontWeight="Bold" FontSize="12"/>
      </StatusBarItem>
      <Separator Background="#444"/>
      <StatusBarItem>
        <TextBlock x:Name="txtIdleCountdown" Text="" Foreground="#AAAAAA" FontSize="11"
                   ToolTip="Time remaining before idle auto-disconnect"/>
      </StatusBarItem>
    </StatusBar>

    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
    <Grid Margin="0,4,0,0">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>   <!-- 0: Title -->
        <RowDefinition Height="Auto"/>   <!-- 1: Connection -->
        <RowDefinition Height="Auto"/>   <!-- 2: Mode + Permissions -->
        <RowDefinition Height="Auto"/>   <!-- 3: Mailbox Emails -->
        <RowDefinition Height="Auto"/>   <!-- 4: User UPNs -->
        <RowDefinition Height="Auto"/>   <!-- 5: Execute -->
        <RowDefinition Height="160"/>    <!-- 6: Log Console -->
        <RowDefinition Height="200"/>    <!-- 7: Results Grid -->
        <RowDefinition Height="Auto"/>   <!-- 8: Export -->
      </Grid.RowDefinitions>

      <!-- Row 0: Title Banner -->
      <Border Grid.Row="0" Background="#0078D4" Margin="8,2,8,2" CornerRadius="4">
        <TextBlock Text="Shared Mailbox  ·  Permission Manager"
                   FontSize="17" FontWeight="Bold" Foreground="White" Padding="14,9"/>
      </Border>

      <!-- Row 1: Connection -->
      <GroupBox Grid.Row="1" Header="Exchange Online Connection">
        <StackPanel>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="btnConnect"    Content="Connect"/>
            <Button x:Name="btnDisconnect" Content="Disconnect"
                    Style="{StaticResource BtnRed}" IsEnabled="False"/>
            <Ellipse x:Name="ellConnDot" Width="13" Height="13"
                     Fill="#C50F1F" Margin="12,0,5,0" VerticalAlignment="Center"/>
            <Label x:Name="lblConnStatus" Content="Not Connected" FontWeight="SemiBold"/>
            <Separator Width="1" Background="#CCCCCC" Margin="10,2"/>
            <Label Content="Signed in as:" VerticalAlignment="Center" Foreground="#555555"/>
            <TextBlock x:Name="txtConnectedUPN" Text="—"
                       VerticalAlignment="Center" FontFamily="Consolas"
                       FontSize="12" Margin="4,0,8,0" Foreground="#333333"/>
          </StackPanel>
          <StackPanel Orientation="Horizontal" Margin="4,2,0,0">
            <Label Content="Idle Timeout:" Foreground="#777777" FontSize="11" Padding="4,1"/>
            <TextBlock x:Name="txtIdleStatus" Text="Timer inactive."
                       VerticalAlignment="Center" FontSize="11"
                       Foreground="#888888" FontStyle="Italic"/>
          </StackPanel>
        </StackPanel>
      </GroupBox>

      <!-- Row 2: Operation Mode + Permission Types -->
      <GroupBox Grid.Row="2" Header="Operation Mode &amp; Permission Types">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="0,0,20,0">
            <Label Content="Mode:" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <RadioButton x:Name="rbAdd" Content="Add Permissions"
                         GroupName="Mode" IsChecked="True" Foreground="#107C10"/>
            <RadioButton x:Name="rbRemove" Content="Remove Permissions"
                         GroupName="Mode" Foreground="#C50F1F"/>
          </StackPanel>

          <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Label Content="Permissions:" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkFullAccess"   Content="Full Access"    IsChecked="True"/>
            <CheckBox x:Name="chkSendAs"       Content="Send As"        IsChecked="True"/>
            <CheckBox x:Name="chkSendOnBehalf" Content="Send On Behalf" IsChecked="True"/>
          </StackPanel>
        </Grid>
      </GroupBox>

      <!-- Row 3: Mailbox Emails Input -->
      <GroupBox Grid.Row="3" Header="Shared Mailbox Email Addresses  —  one per line, or load from file">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="80"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal">
            <Label Content="File:" VerticalAlignment="Center"/>
            <TextBox x:Name="txtEmailsFilePath" Width="420"
                     ToolTip="Path to a text file with one mailbox email per line"/>
            <Button x:Name="btnBrowseEmailsFile" Content="Browse"
                    Style="{StaticResource BtnGray}" MinWidth="70"
                    ToolTip="Browse for an existing text file."/>
            <Button x:Name="btnOpenEmailsFile" Content="Open / Create"
                    Style="{StaticResource BtnGray}" MinWidth="110"
                    ToolTip="Opens the file for editing. Creates it if it does not exist."/>
            <Button x:Name="btnLoadEmailsFile" Content="Load from File"
                    Style="{StaticResource BtnGray}" MinWidth="110"
                    ToolTip="Reads the file and replaces the text box contents below."/>
            <Button x:Name="btnClearEmails" Content="Clear"
                    Style="{StaticResource BtnGray}" MinWidth="60"/>
          </StackPanel>

          <TextBox Grid.Row="1" x:Name="txtEmailsManual"
                   AcceptsReturn="True" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto"
                   FontFamily="Consolas" FontSize="11.5"
                   ToolTip="Type or paste mailbox email addresses here, one per line."/>
        </Grid>
      </GroupBox>

      <!-- Row 4: User UPNs Input -->
      <GroupBox Grid.Row="4" Header="User Principal Names (UPNs)  —  one per line, or load from file">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="80"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal">
            <Label Content="File:" VerticalAlignment="Center"/>
            <TextBox x:Name="txtUsersFilePath" Width="420"
                     ToolTip="Path to a text file with one user UPN per line"/>
            <Button x:Name="btnBrowseUsersFile" Content="Browse"
                    Style="{StaticResource BtnGray}" MinWidth="70"
                    ToolTip="Browse for an existing text file."/>
            <Button x:Name="btnOpenUsersFile" Content="Open / Create"
                    Style="{StaticResource BtnGray}" MinWidth="110"
                    ToolTip="Opens the file for editing. Creates it if it does not exist."/>
            <Button x:Name="btnLoadUsersFile" Content="Load from File"
                    Style="{StaticResource BtnGray}" MinWidth="110"
                    ToolTip="Reads the file and replaces the text box contents below."/>
            <Button x:Name="btnClearUsers" Content="Clear"
                    Style="{StaticResource BtnGray}" MinWidth="60"/>
          </StackPanel>

          <TextBox Grid.Row="1" x:Name="txtUsersManual"
                   AcceptsReturn="True" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto"
                   FontFamily="Consolas" FontSize="11.5"
                   ToolTip="Type or paste user UPNs here, one per line."/>
        </Grid>
      </GroupBox>

      <!-- Row 5: Execute -->
      <GroupBox Grid.Row="5" Header="Execution">
        <StackPanel>
          <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
            <Button x:Name="btnExecute" Content="Execute"
                    Style="{StaticResource BtnGreen}" IsEnabled="False"
                    ToolTip="Validate inputs and apply the selected permission changes."/>
            <Button x:Name="btnCancel" Content="Cancel"
                    Style="{StaticResource BtnRed}" IsEnabled="False"/>
            <Label x:Name="lblExecProgress" Content="Idle." VerticalAlignment="Center" Margin="8,0"/>
          </StackPanel>
          <ProgressBar x:Name="progressExec"
                       Height="16" Margin="4,0,4,2"
                       Minimum="0" Maximum="100" Value="0"
                       Foreground="#0078D4" Background="#E0E0E0"/>
        </StackPanel>
      </GroupBox>

      <!-- Row 6: Log Console -->
      <GroupBox Grid.Row="6" Header="Activity Log">
        <TextBox x:Name="txtLogConsole" IsReadOnly="True"
                 AcceptsReturn="True" TextWrapping="Wrap"
                 VerticalScrollBarVisibility="Auto"
                 FontFamily="Consolas" FontSize="11"
                 Background="#1E1E1E" Foreground="#CCCCCC"
                 BorderBrush="#333333"/>
      </GroupBox>

      <!-- Row 7: Results DataGrid -->
      <GroupBox Grid.Row="7" Header="Operation Results">
        <DataGrid x:Name="dgResults">
          <DataGrid.Columns>
            <DataGridTextColumn Header="Mailbox"
                Binding="{Binding Mailbox}"         Width="220"/>
            <DataGridTextColumn Header="User"
                Binding="{Binding User}"            Width="220"/>
            <DataGridTextColumn Header="Operation"
                Binding="{Binding Operation}"       Width="80"/>
            <DataGridTextColumn Header="Full Access"
                Binding="{Binding FullAccessResult}" Width="130"/>
            <DataGridTextColumn Header="Send As"
                Binding="{Binding SendAsResult}"     Width="130"/>
            <DataGridTextColumn Header="Send On Behalf"
                Binding="{Binding SendOnBehalfResult}" Width="130"/>
            <DataGridTextColumn Header="Status"
                Binding="{Binding Status}"           Width="*"/>
          </DataGrid.Columns>
        </DataGrid>
      </GroupBox>

      <!-- Row 8: Export -->
      <GroupBox Grid.Row="8" Header="Export Results to CSV">
        <StackPanel Orientation="Horizontal">
          <Label Content="Output Path:"/>
          <TextBox x:Name="txtOutputPath" Width="420"/>
          <Button x:Name="btnBrowseOutput" Content="Browse"
                  Style="{StaticResource BtnGray}" MinWidth="80"/>
          <Button x:Name="btnExport" Content="Export CSV"
                  Style="{StaticResource BtnGreen}" IsEnabled="False"/>
          <Button x:Name="btnOpenExport" Content="Open File"
                  Style="{StaticResource BtnGray}" IsEnabled="False" MinWidth="80"/>
          <Label x:Name="lblExportStatus" Content="" Foreground="#107C10" FontWeight="SemiBold"/>
        </StackPanel>
      </GroupBox>

    </Grid>
    </ScrollViewer>

  </DockPanel>
</Window>
'@

#=================================================================
# Section 3:  Parse XAML & Obtain Window Reference
#=================================================================

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$syncHash.Window = [System.Windows.Markup.XamlReader]::Load($reader)

function Get-Ctrl {
    param([string]$Name)
    $syncHash.Window.FindName($Name)
}

#=================================================================
# Section 4:  Control References & Default Values
#=================================================================

# Connection controls
$btnConnect          = Get-Ctrl 'btnConnect'
$btnDisconnect       = Get-Ctrl 'btnDisconnect'
$txtIdleStatus       = Get-Ctrl 'txtIdleStatus'

# Mode & permission controls
$rbAdd               = Get-Ctrl 'rbAdd'
$rbRemove            = Get-Ctrl 'rbRemove'
$chkFullAccess       = Get-Ctrl 'chkFullAccess'
$chkSendAs           = Get-Ctrl 'chkSendAs'
$chkSendOnBehalf     = Get-Ctrl 'chkSendOnBehalf'

# Emails input controls
$txtEmailsFilePath   = Get-Ctrl 'txtEmailsFilePath'
$btnBrowseEmailsFile = Get-Ctrl 'btnBrowseEmailsFile'
$btnOpenEmailsFile   = Get-Ctrl 'btnOpenEmailsFile'
$btnLoadEmailsFile   = Get-Ctrl 'btnLoadEmailsFile'
$btnClearEmails      = Get-Ctrl 'btnClearEmails'
$txtEmailsManual     = Get-Ctrl 'txtEmailsManual'

# Users input controls
$txtUsersFilePath    = Get-Ctrl 'txtUsersFilePath'
$btnBrowseUsersFile  = Get-Ctrl 'btnBrowseUsersFile'
$btnOpenUsersFile    = Get-Ctrl 'btnOpenUsersFile'
$btnLoadUsersFile    = Get-Ctrl 'btnLoadUsersFile'
$btnClearUsers       = Get-Ctrl 'btnClearUsers'
$txtUsersManual      = Get-Ctrl 'txtUsersManual'

# Execution controls
$btnExecute          = Get-Ctrl 'btnExecute'
$btnCancel           = Get-Ctrl 'btnCancel'
$lblExecProgress     = Get-Ctrl 'lblExecProgress'
$progressExec        = Get-Ctrl 'progressExec'

# Log console
$txtLogConsole       = Get-Ctrl 'txtLogConsole'

# Results & export controls
$btnBrowseOutput     = Get-Ctrl 'btnBrowseOutput'
$btnExport           = Get-Ctrl 'btnExport'
$btnOpenExport       = Get-Ctrl 'btnOpenExport'
$lblExportStatus     = Get-Ctrl 'lblExportStatus'
$txtOutputPath       = Get-Ctrl 'txtOutputPath'

# Cache frequently-accessed controls in syncHash for background thread access
$syncHash.StatusLabel       = Get-Ctrl 'txtStatus'
$syncHash.ConnBadge         = Get-Ctrl 'txtConnBadge'
$syncHash.ConnDot           = Get-Ctrl 'ellConnDot'
$syncHash.ConnStatusLabel   = Get-Ctrl 'lblConnStatus'
$syncHash.ConnectedUPNLabel = Get-Ctrl 'txtConnectedUPN'
$syncHash.LogConsole        = $txtLogConsole
$syncHash.ProgressBar       = $progressExec
$syncHash.ProgressLabel     = $lblExecProgress
$syncHash.ResultsGrid       = Get-Ctrl 'dgResults'
$syncHash.IdleCountdownLabel = Get-Ctrl 'txtIdleCountdown'

# Set default values
$txtEmailsFilePath.Text = $DefaultEmailsFilePath
$txtUsersFilePath.Text  = $DefaultUsersFilePath
$txtOutputPath.Text     = $DefaultOutputPath

#=================================================================
# Section 5:  Security Helper Functions
#=================================================================

function Get-SanitizedErrorMessage {
    param(
        [Parameter(Mandatory)]
        [string]$RawMessage
    )

    $clean = $RawMessage -replace '[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}', '[ID]'
    $clean = $clean -replace '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', '[UPN]'
    $clean = $clean -replace '[\w\-]+\.ps\.compliance\.protection\.outlook\.com', '[compliance-host]'
    $clean = $clean -replace '[\w\-]+\.outlook\.com', '[exchange-host]'
    $clean = $clean -replace '[\w\-]+\.protection\.outlook\.com', '[protection-host]'
    $clean = $clean -replace '\b(?:\d{1,3}\.){3}\d{1,3}\b', '[IP]'

    if ($clean.Length -gt 350) {
        $clean = $clean.Substring(0, 347) + '...'
    }

    return $clean
}

# Scriptblock copy for background runspace access
$syncHash.SanitizeError = {
    param([string]$RawMessage)
    $clean = $RawMessage -replace '[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}', '[ID]'
    $clean = $clean -replace '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', '[UPN]'
    $clean = $clean -replace '[\w\-]+\.ps\.compliance\.protection\.outlook\.com', '[compliance-host]'
    $clean = $clean -replace '[\w\-]+\.outlook\.com', '[exchange-host]'
    $clean = $clean -replace '[\w\-]+\.protection\.outlook\.com', '[protection-host]'
    $clean = $clean -replace '\b(?:\d{1,3}\.){3}\d{1,3}\b', '[IP]'
    if ($clean.Length -gt 350) { $clean = $clean.Substring(0, 347) + '...' }
    return $clean
}

#=================================================================
# Section 6:  UI Helper Functions (UI thread only)
#=================================================================

function New-EXORunspace {
    if ($syncHash.EXORunspace -and
        $syncHash.EXORunspace.RunspaceStateInfo.State -eq
            [System.Management.Automation.Runspaces.RunspaceState]::Opened) {
        return $syncHash.EXORunspace
    }

    Write-Verbose "Creating new STA EXO runspace."
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('syncHash', $syncHash)
    $syncHash.EXORunspace = $rs
    return $rs
}

function Start-BackgroundWorker {
    param([Parameter(Mandatory)][scriptblock]$Worker)
    $rs = New-EXORunspace
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Worker)
    [void]$ps.BeginInvoke()
}

function Close-EXORunspace {
    if ($syncHash.EXORunspace) {
        try {
            if ($syncHash.EXORunspace.RunspaceStateInfo.State -eq
                    [System.Management.Automation.Runspaces.RunspaceState]::Opened) {
                $syncHash.EXORunspace.Close()
            }
        }
        catch { Write-Verbose "EXORunspace.Close() error (non-fatal): $_" }
        finally {
            try { $syncHash.EXORunspace.Dispose() } catch { }
            $syncHash.EXORunspace = $null
            Write-Verbose "EXO runspace disposed."
        }
    }
}

function Reset-IdleClock {
    $syncHash.LastActivityTime = [datetime]::Now
}

# Appends a timestamped line to the dark log console
function Write-Log {
    param([string]$Message)
    $timestamp = [datetime]::Now.ToString('HH:mm:ss')
    $txtLogConsole.AppendText("[$timestamp]  $Message`r`n")
    $txtLogConsole.ScrollToEnd()
}

# Resets all user-entered data and log — used on idle timeout
function Reset-UserData {
    $txtEmailsManual.Text    = ''
    $txtUsersManual.Text     = ''
    $txtLogConsole.Text      = ''
    $txtEmailsFilePath.Text  = $DefaultEmailsFilePath
    $txtUsersFilePath.Text   = $DefaultUsersFilePath
    $syncHash.ResultsGrid.ItemsSource = $null
    $syncHash.ResultObjects  = $null
    $btnExport.IsEnabled     = $false
    $btnOpenExport.IsEnabled = $false
    $lblExportStatus.Content = ''
    $progressExec.Value      = 0
    $lblExecProgress.Content = 'Idle.'
    $rbAdd.IsChecked         = $true
    $chkFullAccess.IsChecked = $true
    $chkSendAs.IsChecked     = $true
    $chkSendOnBehalf.IsChecked = $true
}

function Invoke-AutoDisconnect {
    param([string]$Reason = 'Idle timeout')

    Write-Verbose "Auto-disconnect triggered. Reason: $Reason"

    if (-not $syncHash.IsConnected) { return }
    if ($syncHash.IsRunning)        { return }

    $syncHash.IsRunning = $true

    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.StatusLabel.Text        = "Auto-disconnecting: $Reason"
        $syncHash.IdleCountdownLabel.Text = ''
        $syncHash.Window.FindName('txtIdleStatus').Text = "Auto-disconnecting: $Reason"
    })

    if ($syncHash.EXORunspace -and
        $syncHash.EXORunspace.RunspaceStateInfo.State -eq
            [System.Management.Automation.Runspaces.RunspaceState]::Opened) {
        Start-BackgroundWorker -Worker $disconnectWorker
    }
    else {
        $syncHash.IsConnected = $false
        $syncHash.IsRunning   = $false
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.ConnDot.Fill                = [System.Windows.Media.Brushes]::Red
            $syncHash.ConnStatusLabel.Content     = 'Not Connected'
            $syncHash.ConnBadge.Text              = '●  Not Connected'
            $syncHash.ConnBadge.Foreground        = [System.Windows.Media.Brushes]::Salmon
            $syncHash.ConnectedUPNLabel.Text      = '—'
            $syncHash.Window.FindName('btnConnect').IsEnabled    = $true
            $syncHash.Window.FindName('btnDisconnect').IsEnabled = $false
            $syncHash.Window.FindName('btnExecute').IsEnabled    = $false
            $syncHash.Window.FindName('txtIdleStatus').Text      = 'Timer inactive.'
        })
    }

    # Reset user data on auto-disconnect
    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.Window.FindName('txtEmailsManual').Text    = ''
        $syncHash.Window.FindName('txtUsersManual').Text     = ''
        $syncHash.Window.FindName('txtLogConsole').Text      = ''
        $syncHash.ResultsGrid.ItemsSource = $null
        $syncHash.ResultObjects  = $null
        $syncHash.Window.FindName('btnExport').IsEnabled     = $false
        $syncHash.Window.FindName('btnOpenExport').IsEnabled = $false
        $syncHash.Window.FindName('lblExportStatus').Content = ''
        $syncHash.ProgressBar.Value      = 0
        $syncHash.ProgressLabel.Content  = 'Idle.'
    })
}

#=================================================================
# Section 7:  Idle Session Timer
#=================================================================

function Start-IdleTimer {
    if ($syncHash.IdleTimer) { return }

    $timer          = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromSeconds(30)
    $syncHash.IdleTimer = $timer

    $timer.Add_Tick({
        if (-not $syncHash.IsConnected) {
            $syncHash.IdleCountdownLabel.Text = ''
            $txtIdleStatus.Text               = 'Timer inactive.'
            return
        }

        $elapsed    = ([datetime]::Now - $syncHash.LastActivityTime).TotalMinutes
        $remaining  = [Math]::Max(0, $syncHash.IdleTimeoutMinutes - $elapsed)

        $syncHash.IdleCountdownLabel.Text = "Auto-disconnect in: $([Math]::Ceiling($remaining)) min"
        $txtIdleStatus.Text               = "Idle for $([Math]::Floor($elapsed)) min  |  " +
                                            "Auto-disconnect in $([Math]::Ceiling($remaining)) min"

        if ($elapsed -ge $syncHash.IdleTimeoutMinutes) {
            Invoke-AutoDisconnect -Reason "No activity for $($syncHash.IdleTimeoutMinutes) minutes"
        }
    })

    $timer.Start()
    Write-Verbose "Idle timer started (timeout: $($syncHash.IdleTimeoutMinutes) min)."
}

function Stop-IdleTimer {
    if ($syncHash.IdleTimer) {
        $syncHash.IdleTimer.Stop()
        $syncHash.IdleTimer = $null
        $syncHash.IdleCountdownLabel.Text = ''
        $txtIdleStatus.Text               = 'Timer inactive.'
        Write-Verbose "Idle timer stopped."
    }
}

#=================================================================
# Section 8:  Background Worker Script Blocks
#=================================================================

# --- Disconnect Worker ---
$disconnectWorker = {
    try {
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text = 'Disconnecting from Exchange Online...'
        })

        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

        $syncHash.IsConnected  = $false
        $syncHash.ConnectedUPN = ''

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text            = 'Disconnected.'
            $syncHash.ConnDot.Fill                = [System.Windows.Media.Brushes]::Red
            $syncHash.ConnStatusLabel.Content     = 'Not Connected'
            $syncHash.ConnBadge.Text              = '●  Not Connected'
            $syncHash.ConnBadge.Foreground        = [System.Windows.Media.Brushes]::Salmon
            $syncHash.ConnectedUPNLabel.Text      = '—'
            $syncHash.Window.FindName('btnConnect').IsEnabled    = $true
            $syncHash.Window.FindName('btnDisconnect').IsEnabled = $false
            $syncHash.Window.FindName('btnExecute').IsEnabled    = $false
        })
    }
    catch {
        $rawMsg = $_.Exception.Message
        Write-Verbose "Disconnect error (full): $rawMsg"
        $syncHash.LastError = & $syncHash.SanitizeError $rawMsg

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text = "Disconnect error: $($syncHash.LastError)"
        })
    }
    finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.IsRunning = $false
            $syncHash.Window.FindName('btnCancel').IsEnabled = $false
        })
    }
}

# --- Connect Worker ---
$connectWorker = {
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text = 'Connecting to Exchange Online — complete sign-in when prompted...'
        })

        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

        # Resolve connected UPN
        $connectedUPN = ''
        try {
            $connInfo     = Get-ConnectionInformation -ErrorAction SilentlyContinue
            $connectedUPN = if ($connInfo) { $connInfo.UserPrincipalName } else { 'Unknown' }
        }
        catch {
            $connectedUPN = 'Unknown'
            Write-Verbose "Could not resolve connected UPN: $($_.Exception.Message)"
        }

        $syncHash.IsConnected  = $true
        $syncHash.ConnectedUPN = $connectedUPN

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text            = 'Connected to Exchange Online.'
            $syncHash.ConnDot.Fill                = [System.Windows.Media.Brushes]::LimeGreen
            $syncHash.ConnStatusLabel.Content     = 'Connected'
            $syncHash.ConnBadge.Text              = '●  Connected'
            $syncHash.ConnBadge.Foreground        = [System.Windows.Media.Brushes]::LightGreen
            $syncHash.ConnectedUPNLabel.Text      = $syncHash.ConnectedUPN
            $syncHash.Window.FindName('btnConnect').IsEnabled    = $false
            $syncHash.Window.FindName('btnDisconnect').IsEnabled = $true
            $syncHash.Window.FindName('btnExecute').IsEnabled    = $true
        })
    }
    catch {
        $syncHash.IsConnected = $false
        $rawMsg = $_.Exception.Message
        Write-Verbose "Connection error (full): $rawMsg"
        $syncHash.LastError = & $syncHash.SanitizeError $rawMsg

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text = "Connection failed: $($syncHash.LastError)"
            [System.Windows.MessageBox]::Show(
                "Connection failed.`n`n$($syncHash.LastError)",
                'Connection Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error) | Out-Null
        })
    }
    finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.IsRunning = $false

            if (-not $syncHash.IsConnected) {
                $syncHash.Window.FindName('btnConnect').IsEnabled    = $true
                $syncHash.Window.FindName('btnDisconnect').IsEnabled = $false
            }
        })
    }
}

# --- Permission Execution Worker ---
# Expects the following keys pre-set in $syncHash before invocation:
#   ExecMailboxes   : [string[]] array of mailbox identifiers
#   ExecUsers       : [string[]] array of user identifiers
#   ExecMode        : 'Add' or 'Remove'
#   ExecFullAccess  : $true/$false
#   ExecSendAs      : $true/$false
#   ExecSendOnBehalf: $true/$false

$executeWorker = {

    # Helper: safely read a property even under strict mode
    function Get-SafeProp {
        param([object]$Obj, [string]$Prop)
        if (-not $Obj) { return $null }
        $p = $Obj.PSObject.Properties[$Prop]
        if ($p) { return $p.Value } else { return $null }
    }

    # Helper: append log line via dispatcher
    function Send-Log {
        param([string]$Msg)
        $syncHash.Window.Dispatcher.Invoke([action]{
            $ts = [datetime]::Now.ToString('HH:mm:ss')
            $syncHash.LogConsole.AppendText("[$ts]  $Msg`r`n")
            $syncHash.LogConsole.ScrollToEnd()
        })
    }

    # Helper: update progress
    function Send-Progress {
        param([int]$Pct, [string]$Msg)
        $syncHash.CurrentProgress    = $Pct
        $syncHash.CurrentProgressMsg = $Msg
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.ProgressBar.Value     = $syncHash.CurrentProgress
            $syncHash.ProgressLabel.Content = $syncHash.CurrentProgressMsg
        })
    }

    try {
        $mode       = $syncHash.ExecMode
        $mailboxes  = $syncHash.ExecMailboxes
        $users      = $syncHash.ExecUsers
        $doFA       = $syncHash.ExecFullAccess
        $doSA       = $syncHash.ExecSendAs
        $doSOB      = $syncHash.ExecSendOnBehalf

        $results = [System.Collections.Generic.List[PSObject]]::new()

        $totalPairs = $mailboxes.Count * $users.Count
        $pairIndex  = 0

        Send-Log "Starting $mode operation: $($mailboxes.Count) mailbox(es) x $($users.Count) user(s) = $totalPairs pair(s)."
        Send-Log "Permissions selected — Full Access: $doFA | Send As: $doSA | Send On Behalf: $doSOB"
        Send-Log '---'

        foreach ($mbxInput in $mailboxes) {

            if ($syncHash.CancelRequested) {
                Send-Log 'Operation cancelled by user.'
                break
            }

            # Resolve mailbox
            $mailbox = $null
            try {
                $mailbox = Get-Mailbox -Identity $mbxInput -ErrorAction Stop
            }
            catch {
                $sanitized = & $syncHash.SanitizeError $_.Exception.Message
                Send-Log "WARNING: Could not resolve mailbox '$mbxInput' — skipping. ($sanitized)"

                # Create a result row for each user with this failed mailbox
                foreach ($usrInput in $users) {
                    $pairIndex++
                    $results.Add([PSCustomObject]@{
                        Mailbox            = $mbxInput
                        User               = $usrInput
                        Operation          = $mode
                        FullAccessResult   = 'N/A'
                        SendAsResult       = 'N/A'
                        SendOnBehalfResult = 'N/A'
                        Status             = "Mailbox not found — skipped"
                    })
                }
                continue
            }

            $mailboxId   = $mailbox.Identity
            $mailboxSmtp = $mailbox.PrimarySmtpAddress

            Send-Log "Mailbox resolved: $mailboxSmtp"

            foreach ($usrInput in $users) {

                $pairIndex++
                if ($syncHash.CancelRequested) {
                    Send-Log 'Operation cancelled by user.'
                    break
                }

                $pct = [Math]::Floor(($pairIndex / $totalPairs) * 100)
                Send-Progress -Pct $pct -Msg "Processing $pairIndex of $totalPairs..."

                # Resolve user
                $userRecipient = $null
                try {
                    $userRecipient = Get-Recipient -Identity $usrInput -ErrorAction Stop
                }
                catch {
                    $sanitized = & $syncHash.SanitizeError $_.Exception.Message
                    Send-Log "  WARNING: Could not resolve user '$usrInput' — skipping. ($sanitized)"
                    $results.Add([PSCustomObject]@{
                        Mailbox            = $mailboxSmtp
                        User               = $usrInput
                        Operation          = $mode
                        FullAccessResult   = 'N/A'
                        SendAsResult       = 'N/A'
                        SendOnBehalfResult = 'N/A'
                        Status             = "User not found — skipped"
                    })
                    continue
                }

                $userId   = $userRecipient.Identity
                $userSmtp = if ($userRecipient.PrimarySmtpAddress) {
                                $userRecipient.PrimarySmtpAddress.ToString()
                            } else { $usrInput }
                $userDn   = Get-SafeProp $userRecipient 'DistinguishedName'

                Send-Log "  User resolved: $userSmtp"

                $faResult  = 'Not selected'
                $saResult  = 'Not selected'
                $sobResult = 'Not selected'
                $status    = 'OK'

                #-----------------------------------------------------
                # ADD MODE
                #-----------------------------------------------------
                if ($mode -eq 'Add') {

                    # --- Full Access ---
                    if ($doFA) {
                        try {
                            $existing = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue
                            if ($existing -and $existing.AccessRights -contains 'FullAccess' -and -not $existing.IsInherited) {
                                $faResult = 'Already present'
                                Send-Log "    Full Access: already present."
                            }
                            else {
                                Add-MailboxPermission -Identity $mailboxId -User $userId `
                                    -AccessRights FullAccess -InheritanceType All `
                                    -AutoMapping:$true -ErrorAction Stop | Out-Null
                                $faResult = 'Granted'
                                Send-Log "    Full Access: granted."
                            }
                        }
                        catch {
                            $faResult = 'FAILED'
                            $status   = 'Partial failure'
                            Send-Log "    Full Access: FAILED — $(& $syncHash.SanitizeError $_.Exception.Message)"
                        }
                    }

                    # --- Send As ---
                    if ($doSA) {
                        try {
                            $existingSA = Get-RecipientPermission -Identity $mailboxId -ErrorAction SilentlyContinue |
                                          Where-Object { $_.Trustee -eq $userId -and $_.AccessRights -contains 'SendAs' }
                            if ($existingSA) {
                                $saResult = 'Already present'
                                Send-Log "    Send As: already present."
                            }
                            else {
                                Add-RecipientPermission -Identity $mailboxId -Trustee $userId `
                                    -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                                $saResult = 'Granted'
                                Send-Log "    Send As: granted."
                            }
                        }
                        catch {
                            $saResult = 'FAILED'
                            $status   = 'Partial failure'
                            Send-Log "    Send As: FAILED — $(& $syncHash.SanitizeError $_.Exception.Message)"
                        }
                    }

                    # --- Send On Behalf ---
                    if ($doSOB) {
                        try {
                            $mbxForSob     = Get-Mailbox -Identity $mailboxId -ErrorAction Stop
                            $currentSobList = @()
                            if ($mbxForSob.GrantSendOnBehalfTo) {
                                $currentSobList = @($mbxForSob.GrantSendOnBehalfTo)
                            }

                            # Robust already-present check using multi-property matching
                            $alreadyPresent = $false
                            foreach ($sobEntry in $currentSobList) {
                                $sobStr  = $sobEntry.ToString()
                                $isMatch = $false

                                # Try resolving the existing SOB entry to compare properties
                                $sobRecip = $null
                                try { $sobRecip = Get-Recipient -Identity $sobEntry -ErrorAction SilentlyContinue } catch { }

                                if ($sobRecip) {
                                    $uExtId = Get-SafeProp $userRecipient 'ExternalDirectoryObjectId'
                                    $sExtId = Get-SafeProp $sobRecip      'ExternalDirectoryObjectId'
                                    $uGuid  = Get-SafeProp $userRecipient 'Guid'
                                    $sGuid  = Get-SafeProp $sobRecip      'Guid'
                                    $uSmtp  = Get-SafeProp $userRecipient 'PrimarySmtpAddress'
                                    $sSmtp  = Get-SafeProp $sobRecip      'PrimarySmtpAddress'
                                    $uUpn   = Get-SafeProp $userRecipient 'UserPrincipalName'
                                    $sUpn   = Get-SafeProp $sobRecip      'UserPrincipalName'

                                    if     ($uExtId -and $sExtId -and $uExtId -eq $sExtId)          { $isMatch = $true }
                                    elseif ($uGuid  -and $sGuid  -and $uGuid  -eq $sGuid)           { $isMatch = $true }
                                    elseif ($uSmtp  -and $sSmtp  -and
                                            $uSmtp.ToString().ToLowerInvariant() -eq
                                            $sSmtp.ToString().ToLowerInvariant())                    { $isMatch = $true }
                                    elseif ($uUpn   -and $sUpn   -and
                                            $uUpn.ToLowerInvariant() -eq $sUpn.ToLowerInvariant())  { $isMatch = $true }
                                }
                                else {
                                    # Fallback string matching against all known user identifiers
                                    $candidates = @($userId, $userSmtp, $userDn,
                                        (Get-SafeProp $userRecipient 'UserPrincipalName'),
                                        (Get-SafeProp $userRecipient 'Alias'),
                                        (Get-SafeProp $userRecipient 'Name')
                                    ) | Where-Object { $_ }
                                    $candidateLower = $candidates | ForEach-Object { $_.ToString().ToLowerInvariant() }
                                    if ($candidateLower -contains $sobStr.ToLowerInvariant()) { $isMatch = $true }
                                }

                                if ($isMatch) { $alreadyPresent = $true; break }
                            }

                            if ($alreadyPresent) {
                                $sobResult = 'Already present'
                                Send-Log "    Send On Behalf: already present."
                            }
                            else {
                                $newSobList = @()
                                if ($currentSobList.Count -gt 0) {
                                    $newSobList += @($currentSobList | ForEach-Object { $_.ToString() })
                                }
                                if ($userDn) { $newSobList += $userDn }
                                else         { $newSobList += $userId }

                                Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                                $sobResult = 'Granted'
                                Send-Log "    Send On Behalf: granted."
                            }
                        }
                        catch {
                            # Check if the error indicates the user is already present
                            $errMsg = $_.Exception.Message
                            if ($errMsg -match 'already present|already exists|duplicate') {
                                $sobResult = 'Already present'
                                Send-Log "    Send On Behalf: already present (confirmed by server)."
                            }
                            else {
                                $sobResult = 'FAILED'
                                $status    = 'Partial failure'
                                Send-Log "    Send On Behalf: FAILED — $(& $syncHash.SanitizeError $errMsg)"
                            }
                        }
                    }
                }
                #-----------------------------------------------------
                # REMOVE MODE
                #-----------------------------------------------------
                elseif ($mode -eq 'Remove') {

                    # --- Full Access ---
                    if ($doFA) {
                        try {
                            $existing = Get-MailboxPermission -Identity $mailboxId -User $userId -ErrorAction SilentlyContinue
                            if ($existing -and $existing.AccessRights -contains 'FullAccess' -and -not $existing.IsInherited) {
                                Remove-MailboxPermission -Identity $mailboxId -User $userId `
                                    -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                                $faResult = 'Removed'
                                Send-Log "    Full Access: removed."
                            }
                            else {
                                $faResult = 'Not present'
                                Send-Log "    Full Access: not present."
                            }
                        }
                        catch {
                            $faResult = 'FAILED'
                            $status   = 'Partial failure'
                            Send-Log "    Full Access: FAILED — $(& $syncHash.SanitizeError $_.Exception.Message)"
                        }
                    }

                    # --- Send As (robust trustee matching) ---
                    if ($doSA) {
                        try {
                            $allSAPerms = @()
                            try {
                                $allSAPerms = @(Get-RecipientPermission -Identity $mailboxId -ErrorAction SilentlyContinue |
                                                Where-Object { $_.AccessRights -contains 'SendAs' })
                            } catch { }

                            if ($allSAPerms.Count -eq 0) {
                                $saResult = 'Not present'
                                Send-Log "    Send As: not present."
                            }
                            else {
                                $matchingPerms = @()
                                $resolvedUserSmtp = $userSmtp

                                foreach ($perm in $allSAPerms) {
                                    if (-not $perm) { continue }
                                    $trusteeStr = $perm.Trustee.ToString()
                                    $isMatch = $false

                                    # Try resolving the trustee
                                    $trusteeRecip = $null
                                    try { $trusteeRecip = Get-Recipient -Identity $perm.Trustee -ErrorAction SilentlyContinue } catch { }

                                    if ($trusteeRecip) {
                                        $uExtId = Get-SafeProp $userRecipient 'ExternalDirectoryObjectId'
                                        $tExtId = Get-SafeProp $trusteeRecip  'ExternalDirectoryObjectId'
                                        $uGuid  = Get-SafeProp $userRecipient 'Guid'
                                        $tGuid  = Get-SafeProp $trusteeRecip  'Guid'
                                        $uSmtp  = Get-SafeProp $userRecipient 'PrimarySmtpAddress'
                                        $tSmtp  = Get-SafeProp $trusteeRecip  'PrimarySmtpAddress'
                                        $uUpn   = Get-SafeProp $userRecipient 'UserPrincipalName'
                                        $tUpn   = Get-SafeProp $trusteeRecip  'UserPrincipalName'

                                        if     ($uExtId -and $tExtId -and $uExtId -eq $tExtId)          { $isMatch = $true }
                                        elseif ($uGuid  -and $tGuid  -and $uGuid  -eq $tGuid)           { $isMatch = $true }
                                        elseif ($uSmtp  -and $tSmtp  -and
                                                $uSmtp.ToString().ToLowerInvariant() -eq
                                                $tSmtp.ToString().ToLowerInvariant())                    { $isMatch = $true }
                                        elseif ($uUpn   -and $tUpn   -and
                                                $uUpn.ToLowerInvariant() -eq $tUpn.ToLowerInvariant())  { $isMatch = $true }
                                    }
                                    else {
                                        # Fallback string matching
                                        $candidates = @($userId, $resolvedUserSmtp,
                                            (Get-SafeProp $userRecipient 'UserPrincipalName'),
                                            (Get-SafeProp $userRecipient 'Alias'),
                                            (Get-SafeProp $userRecipient 'Name'),
                                            (Get-SafeProp $userRecipient 'DistinguishedName')
                                        ) | Where-Object { $_ }
                                        $candidateLower = $candidates | ForEach-Object { $_.ToString().ToLowerInvariant() }
                                        if ($candidateLower -contains $trusteeStr.ToLowerInvariant()) { $isMatch = $true }
                                    }

                                    if ($isMatch) { $matchingPerms += $perm }
                                }

                                if ($matchingPerms.Count -gt 0) {
                                    $trusteesToRemove = $matchingPerms | Select-Object -ExpandProperty Trustee -Unique
                                    foreach ($trustee in $trusteesToRemove) {
                                        Remove-RecipientPermission -Identity $mailboxId -Trustee $trustee `
                                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                                    }
                                    $saResult = 'Removed'
                                    Send-Log "    Send As: removed."
                                }
                                else {
                                    $saResult = 'Not present'
                                    Send-Log "    Send As: not present (no matching trustee)."
                                }
                            }
                        }
                        catch {
                            $saResult = 'FAILED'
                            $status   = 'Partial failure'
                            Send-Log "    Send As: FAILED — $(& $syncHash.SanitizeError $_.Exception.Message)"
                        }
                    }

                    # --- Send On Behalf (robust trustee matching) ---
                    if ($doSOB) {
                        try {
                            $mbxForSob      = Get-Mailbox -Identity $mailboxId -ErrorAction Stop
                            $currentSobList = @()
                            if ($mbxForSob.GrantSendOnBehalfTo) {
                                $currentSobList = @($mbxForSob.GrantSendOnBehalfTo)
                            }

                            if ($currentSobList.Count -eq 0) {
                                $sobResult = 'Not present'
                                Send-Log "    Send On Behalf: not present."
                            }
                            else {
                                $matchingSob = @()

                                foreach ($sobEntry in $currentSobList) {
                                    $sobStr  = $sobEntry.ToString()
                                    $isMatch = $false

                                    $sobRecip = $null
                                    try { $sobRecip = Get-Recipient -Identity $sobEntry -ErrorAction SilentlyContinue } catch { }

                                    if ($sobRecip) {
                                        $uExtId = Get-SafeProp $userRecipient 'ExternalDirectoryObjectId'
                                        $sExtId = Get-SafeProp $sobRecip      'ExternalDirectoryObjectId'
                                        $uGuid  = Get-SafeProp $userRecipient 'Guid'
                                        $sGuid  = Get-SafeProp $sobRecip      'Guid'
                                        $uSmtp  = Get-SafeProp $userRecipient 'PrimarySmtpAddress'
                                        $sSmtp  = Get-SafeProp $sobRecip      'PrimarySmtpAddress'
                                        $uUpn   = Get-SafeProp $userRecipient 'UserPrincipalName'
                                        $sUpn   = Get-SafeProp $sobRecip      'UserPrincipalName'

                                        if     ($uExtId -and $sExtId -and $uExtId -eq $sExtId)          { $isMatch = $true }
                                        elseif ($uGuid  -and $sGuid  -and $uGuid  -eq $sGuid)           { $isMatch = $true }
                                        elseif ($uSmtp  -and $sSmtp  -and
                                                $uSmtp.ToString().ToLowerInvariant() -eq
                                                $sSmtp.ToString().ToLowerInvariant())                    { $isMatch = $true }
                                        elseif ($uUpn   -and $sUpn   -and
                                                $uUpn.ToLowerInvariant() -eq $sUpn.ToLowerInvariant())  { $isMatch = $true }
                                    }
                                    else {
                                        $candidates = @($userId, $userSmtp,
                                            (Get-SafeProp $userRecipient 'UserPrincipalName'),
                                            (Get-SafeProp $userRecipient 'Alias'),
                                            (Get-SafeProp $userRecipient 'Name'),
                                            (Get-SafeProp $userRecipient 'DistinguishedName')
                                        ) | Where-Object { $_ }
                                        $candidateLower = $candidates | ForEach-Object { $_.ToString().ToLowerInvariant() }
                                        if ($candidateLower -contains $sobStr.ToLowerInvariant()) { $isMatch = $true }
                                    }

                                    if ($isMatch) { $matchingSob += $sobEntry }
                                }

                                if ($matchingSob.Count -gt 0) {
                                    $newSobList = $currentSobList | Where-Object { $matchingSob -notcontains $_ }
                                    Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo $newSobList -ErrorAction Stop
                                    $sobResult = 'Removed'
                                    Send-Log "    Send On Behalf: removed."
                                }
                                else {
                                    $sobResult = 'Not present'
                                    Send-Log "    Send On Behalf: not present (no matching entry)."
                                }
                            }
                        }
                        catch {
                            $sobResult = 'FAILED'
                            $status    = 'Partial failure'
                            Send-Log "    Send On Behalf: FAILED — $(& $syncHash.SanitizeError $_.Exception.Message)"
                        }
                    }
                }

                # Build result object
                $results.Add([PSCustomObject]@{
                    Mailbox            = $mailboxSmtp
                    User               = $userSmtp
                    Operation          = $mode
                    FullAccessResult   = $faResult
                    SendAsResult       = $saResult
                    SendOnBehalfResult = $sobResult
                    Status             = $status
                })

                Send-Log "  Result: $status"
            }
        }

        $syncHash.ResultObjects = $results

        Send-Progress -Pct 100 -Msg 'Complete.'
        Send-Log '---'
        Send-Log "Operation complete. $($results.Count) result(s) generated."

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.ResultsGrid.ItemsSource = $syncHash.ResultObjects
            $syncHash.StatusLabel.Text        = "$mode operation complete — $($syncHash.ResultObjects.Count) result(s)."
            $syncHash.Window.FindName('btnExport').IsEnabled = $true
        })
    }
    catch {
        $rawMsg = $_.Exception.Message
        Write-Verbose "Execute worker error (full): $rawMsg"
        $syncHash.LastError = & $syncHash.SanitizeError $rawMsg

        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.StatusLabel.Text = "Execution error: $($syncHash.LastError)"
            $ts = [datetime]::Now.ToString('HH:mm:ss')
            $syncHash.LogConsole.AppendText("[$ts]  FATAL ERROR: $($syncHash.LastError)`r`n")
            $syncHash.LogConsole.ScrollToEnd()
        })
    }
    finally {
        $syncHash.Window.Dispatcher.Invoke([action]{
            $syncHash.IsRunning = $false
            $syncHash.CancelRequested = $false
            $syncHash.Window.FindName('btnExecute').IsEnabled = $true
            $syncHash.Window.FindName('btnCancel').IsEnabled  = $false
        })
    }
}

#=================================================================
# Section 9:  Button Click Handlers
#=================================================================

# --- Connect ---
$btnConnect.Add_Click({
    Reset-IdleClock

    if ($syncHash.IsRunning) { return }
    $syncHash.IsRunning = $true

    $btnConnect.IsEnabled = $false
    $syncHash.StatusLabel.Text = 'Initiating connection...'

    Start-BackgroundWorker -Worker $connectWorker
    Start-IdleTimer
})

# --- Disconnect ---
$btnDisconnect.Add_Click({
    Reset-IdleClock

    if ($syncHash.IsRunning) { return }
    $syncHash.IsRunning = $true

    $btnDisconnect.IsEnabled = $false
    $syncHash.StatusLabel.Text = 'Disconnecting...'
    Stop-IdleTimer

    Start-BackgroundWorker -Worker $disconnectWorker
})

# --- Browse Emails File ---
$btnBrowseEmailsFile.Add_Click({
    Reset-IdleClock

    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
    $dlg.Title  = 'Select Mailbox Emails File'

    if (-not [string]::IsNullOrWhiteSpace($txtEmailsFilePath.Text)) {
        try {
            $dir = Split-Path $txtEmailsFilePath.Text -Parent
            if (Test-Path $dir) { $dlg.InitialDirectory = $dir }
        } catch { }
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtEmailsFilePath.Text = $dlg.FileName
    }
})

# --- Open/Create Emails File ---
$btnOpenEmailsFile.Add_Click({
    Reset-IdleClock

    $filePath = $txtEmailsFilePath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        [System.Windows.MessageBox]::Show(
            'Please specify a file path first.',
            'No File Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
        $dir = Split-Path $resolvedPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            New-Item -Path $resolvedPath -ItemType File -Force | Out-Null
            Write-Log "Created new file: $resolvedPath"
        }
        Start-Process -FilePath 'notepad.exe' -ArgumentList $resolvedPath
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Could not open or create the file.`n`n$(Get-SanitizedErrorMessage -RawMessage $_.Exception.Message)",
            'File Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

# --- Load Emails from File ---
$btnLoadEmailsFile.Add_Click({
    Reset-IdleClock

    $filePath = $txtEmailsFilePath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        [System.Windows.MessageBox]::Show(
            'Please specify a file path first.',
            'No File Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            [System.Windows.MessageBox]::Show(
                "File not found:`n$resolvedPath",
                'File Not Found',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $lines = Get-Content -LiteralPath $resolvedPath -ErrorAction Stop |
                 Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^\s*#' } |
                 ForEach-Object { $_.Trim() }

        if ($lines.Count -eq 0) {
            Write-Log "File loaded but contained no valid entries: $resolvedPath"
            return
        }

        $txtEmailsManual.Text = ($lines -join "`r`n")
        Write-Log "Loaded $($lines.Count) mailbox entry/entries from file."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Could not read the file.`n`n$(Get-SanitizedErrorMessage -RawMessage $_.Exception.Message)",
            'File Read Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

# --- Clear Emails ---
$btnClearEmails.Add_Click({
    Reset-IdleClock
    $txtEmailsManual.Text = ''
})

# --- Browse Users File ---
$btnBrowseUsersFile.Add_Click({
    Reset-IdleClock

    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
    $dlg.Title  = 'Select User UPNs File'

    if (-not [string]::IsNullOrWhiteSpace($txtUsersFilePath.Text)) {
        try {
            $dir = Split-Path $txtUsersFilePath.Text -Parent
            if (Test-Path $dir) { $dlg.InitialDirectory = $dir }
        } catch { }
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtUsersFilePath.Text = $dlg.FileName
    }
})

# --- Open/Create Users File ---
$btnOpenUsersFile.Add_Click({
    Reset-IdleClock

    $filePath = $txtUsersFilePath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        [System.Windows.MessageBox]::Show(
            'Please specify a file path first.',
            'No File Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
        $dir = Split-Path $resolvedPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            New-Item -Path $resolvedPath -ItemType File -Force | Out-Null
            Write-Log "Created new file: $resolvedPath"
        }
        Start-Process -FilePath 'notepad.exe' -ArgumentList $resolvedPath
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Could not open or create the file.`n`n$(Get-SanitizedErrorMessage -RawMessage $_.Exception.Message)",
            'File Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

# --- Load Users from File ---
$btnLoadUsersFile.Add_Click({
    Reset-IdleClock

    $filePath = $txtUsersFilePath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        [System.Windows.MessageBox]::Show(
            'Please specify a file path first.',
            'No File Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            [System.Windows.MessageBox]::Show(
                "File not found:`n$resolvedPath",
                'File Not Found',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $lines = Get-Content -LiteralPath $resolvedPath -ErrorAction Stop |
                 Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^\s*#' } |
                 ForEach-Object { $_.Trim() }

        if ($lines.Count -eq 0) {
            Write-Log "File loaded but contained no valid entries: $resolvedPath"
            return
        }

        $txtUsersManual.Text = ($lines -join "`r`n")
        Write-Log "Loaded $($lines.Count) user entry/entries from file."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Could not read the file.`n`n$(Get-SanitizedErrorMessage -RawMessage $_.Exception.Message)",
            'File Read Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

# --- Clear Users ---
$btnClearUsers.Add_Click({
    Reset-IdleClock
    $txtUsersManual.Text = ''
})

# --- Execute ---
$btnExecute.Add_Click({
    Reset-IdleClock

    if ($syncHash.IsRunning) { return }
    if (-not $syncHash.IsConnected) {
        [System.Windows.MessageBox]::Show(
            'Please connect to Exchange Online first.',
            'Not Connected',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Parse mailbox entries from the text box
    $mailboxEntries = @($txtEmailsManual.Text -split "`r`n|`r|`n" |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^\s*#' } |
                        ForEach-Object { $_.Trim() } |
                        Select-Object -Unique)

    # Parse user entries from the text box
    $userEntries = @($txtUsersManual.Text -split "`r`n|`r|`n" |
                     Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^\s*#' } |
                     ForEach-Object { $_.Trim() } |
                     Select-Object -Unique)

    # Validate inputs
    if ($mailboxEntries.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'Please enter at least one mailbox email address.',
            'No Mailboxes',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if ($userEntries.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'Please enter at least one user UPN.',
            'No Users',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $doFA  = ($chkFullAccess.IsChecked -eq $true)
    $doSA  = ($chkSendAs.IsChecked -eq $true)
    $doSOB = ($chkSendOnBehalf.IsChecked -eq $true)

    if (-not $doFA -and -not $doSA -and -not $doSOB) {
        [System.Windows.MessageBox]::Show(
            'Please select at least one permission type.',
            'No Permissions Selected',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $mode = if ($rbAdd.IsChecked -eq $true) { 'Add' } else { 'Remove' }

    # Confirmation dialog
    $totalPairs = $mailboxEntries.Count * $userEntries.Count
    $permList   = @()
    if ($doFA)  { $permList += 'Full Access' }
    if ($doSA)  { $permList += 'Send As' }
    if ($doSOB) { $permList += 'Send On Behalf' }

    $confirmMsg = "You are about to $($mode.ToUpper()) the following permissions:`n" +
                  "  $($permList -join ', ')`n`n" +
                  "Mailboxes: $($mailboxEntries.Count)`n" +
                  "Users:     $($userEntries.Count)`n" +
                  "Total operations: $totalPairs`n`n" +
                  "Do you want to proceed?"

    $answer = [System.Windows.MessageBox]::Show(
        $confirmMsg,
        "Confirm $mode Operation",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)

    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log "$mode operation cancelled by user."
        return
    }

    # Stage parameters for the background worker
    $syncHash.ExecMailboxes    = $mailboxEntries
    $syncHash.ExecUsers        = $userEntries
    $syncHash.ExecMode         = $mode
    $syncHash.ExecFullAccess   = $doFA
    $syncHash.ExecSendAs       = $doSA
    $syncHash.ExecSendOnBehalf = $doSOB
    $syncHash.CancelRequested  = $false

    $syncHash.IsRunning   = $true
    $btnExecute.IsEnabled = $false
    $btnCancel.IsEnabled  = $true
    $btnExport.IsEnabled  = $false
    $btnOpenExport.IsEnabled = $false
    $progressExec.Value   = 0
    $lblExecProgress.Content = 'Starting...'
    $syncHash.ResultsGrid.ItemsSource = $null
    $syncHash.ResultObjects = $null

    Start-BackgroundWorker -Worker $executeWorker
})

# --- Cancel ---
$btnCancel.Add_Click({
    Reset-IdleClock
    $syncHash.CancelRequested = $true
    Write-Log 'Cancellation requested — finishing current operation...'
    $btnCancel.IsEnabled = $false
})

# --- Browse Export Path ---
$btnBrowseOutput.Add_Click({
    Reset-IdleClock

    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter   = 'CSV Files (*.csv)|*.csv|All Files (*.*)|*.*'
    $dlg.Title    = 'Select Export File Location'
    $dlg.FileName = 'PermissionResults.csv'

    if (-not [string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
        try {
            $dlg.InitialDirectory = Split-Path $txtOutputPath.Text -Parent
            $dlg.FileName         = Split-Path $txtOutputPath.Text -Leaf
        } catch { }
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputPath.Text = $dlg.FileName
    }
})

# --- Export CSV ---
$btnExport.Add_Click({
    Reset-IdleClock

    $outPath = $txtOutputPath.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($outPath)) {
        [System.Windows.MessageBox]::Show(
            'Please specify a valid output CSV path before exporting.',
            'No Output Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if (-not $syncHash.ResultObjects -or $syncHash.ResultObjects.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'There are no results to export. Run an operation first.',
            'No Results',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($outPath)
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "The specified output path could not be resolved.`n`nPlease verify the path and try again.",
            'Invalid Path',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Overwrite confirmation
    if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
        $overwriteAnswer = [System.Windows.MessageBox]::Show(
            "The file already exists:`n$resolvedPath`n`nOverwrite it?",
            'File Exists — Confirm Overwrite',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)

        if ($overwriteAnswer -ne [System.Windows.MessageBoxResult]::Yes) {
            $lblExportStatus.Content   = 'Export cancelled — file not overwritten.'
            $lblExportStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
            return
        }
    }

    try {
        $dir = Split-Path $resolvedPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }

        $count = $syncHash.ResultObjects.Count

        $syncHash.ResultObjects |
            Select-Object Mailbox, User, Operation,
                          FullAccessResult, SendAsResult, SendOnBehalfResult, Status |
            Export-Csv -Path $resolvedPath -NoTypeInformation -Encoding UTF8

        $lblExportStatus.Content   = "Exported $count record(s)."
        $lblExportStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
        $syncHash.StatusLabel.Text  = "Results exported: $resolvedPath"
        $btnOpenExport.IsEnabled    = $true

        Write-Log "Exported $count result(s) to: $resolvedPath"
    }
    catch {
        $rawMsg = $_.Exception.Message
        Write-Verbose "Export error (full): $rawMsg"
        [System.Windows.MessageBox]::Show(
            "Export failed.`n`n$(Get-SanitizedErrorMessage -RawMessage $rawMsg)",
            'Export Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

# --- Open Exported File ---
$btnOpenExport.Add_Click({
    Reset-IdleClock

    $outPath = $txtOutputPath.Text.Trim()
    try {
        $resolvedPath = [System.IO.Path]::GetFullPath($outPath)
        if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
            Start-Process -FilePath $resolvedPath
        }
        else {
            [System.Windows.MessageBox]::Show(
                "File not found:`n$resolvedPath",
                'File Not Found',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Could not open the file.`n`n$(Get-SanitizedErrorMessage -RawMessage $_.Exception.Message)",
            'Open File Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

#=================================================================
# Section 10:  Window Lifecycle Events
#=================================================================

# Minimize guard — disconnect when minimized
$syncHash.Window.Add_StateChanged({
    param($sender, $e)
    if ($syncHash.Window.WindowState -eq [System.Windows.WindowState]::Minimized) {
        Write-Verbose "Window minimized — triggering auto-disconnect."
        Invoke-AutoDisconnect -Reason 'Window minimized'
    }
})

# Workstation lock guard
$sessionSwitchHandler = [Microsoft.Win32.SessionSwitchEventHandler]{
    param($sender, $e)
    if ($e.Reason -eq [Microsoft.Win32.SessionSwitchReason]::SessionLock) {
        Write-Verbose "Workstation locked — queuing auto-disconnect."
        $syncHash.Window.Dispatcher.BeginInvoke(
            [action]{ Invoke-AutoDisconnect -Reason 'Workstation locked' }
        ) | Out-Null
    }
}
[Microsoft.Win32.SystemEvents]::add_SessionSwitch($sessionSwitchHandler)

# Window closing — unconditional cleanup
$syncHash.Window.Add_Closing({
    param($sender, $e)
    try {
        Stop-IdleTimer

        try {
            [Microsoft.Win32.SystemEvents]::remove_SessionSwitch($sessionSwitchHandler)
        }
        catch { Write-Verbose "SessionSwitch handler removal error (non-fatal): $_" }

        if ($syncHash.EXORunspace -and
            $syncHash.EXORunspace.RunspaceStateInfo.State -eq
                [System.Management.Automation.Runspaces.RunspaceState]::Opened) {
            try {
                $ps = [powershell]::Create()
                $ps.Runspace = $syncHash.EXORunspace
                [void]$ps.AddScript({
                    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
                })
                $handle = $ps.BeginInvoke()
                [void]$handle.AsyncWaitHandle.WaitOne(6000)
                $ps.Dispose()
            }
            catch { Write-Verbose "Error during close-time disconnect (non-fatal): $_" }
        }
    }
    catch {
        Write-Verbose "Window.Closing handler error (non-fatal): $_"
    }
    finally {
        Close-EXORunspace
        Write-Verbose "Window closed. Script complete."
    }
})

#=================================================================
# Section 11:  Launch Window
#=================================================================

Write-Verbose "Displaying Shared Mailbox Permission Manager window."
[void]$syncHash.Window.ShowDialog()
