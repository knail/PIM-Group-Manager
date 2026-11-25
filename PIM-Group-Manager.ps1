# ==================================================
# MARK: Script Information
# ==================================================
 
<#
.SYNOPSIS
    Creates Microsoft Entra groups and configures PIM for role-assignable groups.
 
.DESCRIPTION
    Provides a Windows Forms GUI to create Microsoft Entra groups individually
    or from a CSV/XLSX file. Connect, check, and disconnect the Microsoft Graph
    session from the GUI.
 
    Supported group types:
    - PIM-enabled, role-assignable security group (default)
    - Security group with assigned or dynamic membership
    - Microsoft 365 group with assigned or dynamic membership
 
    For role-assignable security groups, the script configures PIM Member
    settings and can optionally create eligible member assignments by UPN.
    The GUI can also apply the selected PIM Member settings to an existing
    role-assignable security group, and submit eligible member requests for
    existing PIM-supported security and Microsoft 365 groups.
 
    Bulk templates support group properties, PIM settings, and optional eligible
    member UPNs per row. Bulk input is validated before group creation.
    A selectable audit log records actions, results, errors, account, and tenant.
 
    PIM MEMBER DEFAULTS
    - Activation maximum: 8 hours; require MFA on activation.
    - Eligible assignments expire after 365 days; permanent eligibility off.
    - Permanent active assignments off; preserve the current maximum duration.
    - Require MFA and justification on active assignment.
    - Notification settings are not changed.
 
    The PIM group description is prefixed with:
    *DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal
 
    REQUIREMENTS
    - Windows PowerShell 5.1 in STA mode and network access to Microsoft Graph.
    - Microsoft.Graph.Authentication PowerShell module.
    - ImportExcel module for XLSX templates, imports, and report exports; CSV does not require it.
    - Microsoft Entra ID P2 or EMS E5 licensing for PIM.
    - Appropriate Entra directory role, generally Privileged Role Administrator
      for creating role-assignable groups and managing their PIM settings.
    - Delegated consent for User.Read, User.Read.All, Group.ReadWrite.All,
      RoleManagementPolicy.ReadWrite.AzureADGroup, and
      PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup. The reports also
      request PrivilegedAssignmentSchedule.Read.AzureADGroup. Active PIM
      removal additionally requests PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup;
      direct removal from role-assignable groups also requires
      RoleManagement.ReadWrite.Directory.
 
    Install modules in the PowerShell environment used to launch the GUI:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module ImportExcel -Scope CurrentUser
 
.EXAMPLE
    powershell.exe -STA -NoProfile -File ".\PIM-Group-Manager.ps1"
 
    Run from the folder containing the script. Select Connect / Login, check
    the connection, enter the group details and settings, then confirm creation.
 
.EXAMPLE
    In the GUI, select Create CSV / XLSX template on the Bulk create tab.
    Fill in the template, import it, and confirm the validated groups.
    Results are saved as PIM-Group-Creation-Results-yyyy-MM-dd_HHmmss.csv
    beside the imported file.
 
.NOTES
    Script Name : PIM-Group-Manager.ps1
    Author      : Keith Nail (KNail)
    Version     : 1.3
    Modified    : 2026-11-21
    Purpose     : Create Entra groups and configure PIM Member settings.
    Repository  : https://github.com/knail/PIM-Group-Manager
 
    A role-assignable group must be a security group with assigned membership.
    Creating it does not assign an Entra role; assign roles separately.
 
    Group creation, PIM policy changes, and eligible member assignments are
    separate operations. A failure may leave a group or partial changes in
    place. Review the log and Entra PIM before retrying.
 
    To retry PIM settings, enter the existing group's exact display name or
    object ID, review the settings, and select Apply PIM to existing group.
    If multiple groups share a name, use the object ID. This action does not
    create a group, assign roles, or add eligible members.
 
    PIM configuration for newly created groups waits 10 seconds for directory
    propagation. Policy settings are read back after updates. Verify the
    resulting settings and eligible assignments in Entra PIM. Policy-related
    404 responses are retried for a limited time; other errors are reported.
 
    The GUI creates an audit log in Documents by default using the name
    yyyy-MM-dd_HHmm-PIM-Group-Manager.log. A log already at that path is
    appended automatically; sessions started in the same minute share the
    file, and each entry includes a session ID. Browse to choose another
    .log file; an existing file is appended only after confirmation.
    Entries can contain group details, membership rules, and user UPNs.
    Restrict access to the log and retain it according to your policies.
    Graph changes are blocked when required audit records cannot be written.
    This local log is not tamper-proof; use Entra audit logs for authoritative
    service-side auditing.
 
    VERSION
    1.3 - Refreshed GUI layout and bulk import feedback.
 
    CHANGELOG
    1.3 - Updated the GUI styling and resizing, and cleared stale bulk previews.
    1.2 - Removed environment-specific documentation references.
    1.1 - Added persistent audit records, log selection, and write safeguards.
    1.0 - Documented the current GUI, bulk creation, PIM policy configuration,
          existing-group recovery, and validation behavior.
 
.LINK
    https://github.com/knail/PIM-Group-Manager
#>
 
# ==================================================
# MARK: Prerequisites and Script Configuration
# ==================================================
 
& {
# Keep the guard local to this invocation so F8 cannot overwrite a running instance.
$InstanceCreated = $false
$InstanceName = 'Local\PIM-Group-Manager-' + [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$InstanceGuard = [System.Threading.Mutex]::new($false, $InstanceName, [ref]$InstanceCreated)
if (-not $InstanceCreated)
{
    try
    {
        Write-Warning 'PIM Group Manager is already running in this Windows session. Use the existing window from the taskbar; this launch was skipped.'
        if (('System.Windows.Forms.Application' -as [type]) -and $Script:MainForm -and -not $Script:MainForm.IsDisposed)
        {
            if ($Script:MainForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized)
            {
                $Script:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            }
            [void]$Script:MainForm.Activate()
        }
    }
    catch
    {
        Write-Warning "Could not bring the existing window forward: $($_.Exception.Message)"
    }
    finally
    {
        $InstanceGuard.Dispose()
    }
    return
}

# The named handle's lifetime is the lock; mutex ownership would allow same-thread reentry.
try
{
$ErrorActionPreference = 'Stop'
$Script:MainForm = $null
$Script:HelpWindow = $null
 
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA')
{
    throw 'Run this GUI in STA mode: powershell.exe -STA -File .\PIM-Group-Manager.ps1'
}
 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
 
$Script:HelpContent = @'
PREREQUISITES

- Run in Windows PowerShell 5.1 with an internet connection.

- Install Microsoft.Graph.Authentication: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

- XLSX templates, imports, and report exports need ImportExcel: Install-Module ImportExcel -Scope CurrentUser. CSV works without it.

- PIM requires Microsoft Entra ID P2 or EMS E5 licensing.

- The signed-in account needs consent for the Graph permissions listed below.

- Creating role-assignable groups and changing their PIM settings generally requires Privileged Role Administrator or equivalent permissions.

- Delegated Graph scopes requested: User.Read, User.Read.All, Group.ReadWrite.All, RoleManagementPolicy.ReadWrite.AzureADGroup, PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup, and PrivilegedAssignmentSchedule.Read.AzureADGroup (for active-access reports).

CONNECTION

1. Review the audit log path at the bottom of the window. Browse to select a different location and .log file name.

2. Click Connect / Login and sign in interactively. Grant/obtain admin consent for the requested Graph scopes.

3. Click Check connection. Group operations remain blocked unless Graph confirms the account and all required scopes.

4. Use Disconnect to sign out of the Graph session.

AUDIT LOG

- By default, the log is stored in the current user's Documents folder as yyyy-MM-dd_HHmm-PIM-Group-Manager.log. The path is chosen when the app starts; sessions launched during the same minute append to the same file, and each entry includes a session ID.

- If the default file already exists, new entries are appended. When selecting an existing file with Browse, the GUI asks for confirmation first.

- Entries include UTC time, session ID, Windows and Graph accounts, tenant, group and PIM actions, outcomes, and errors.

- Entries may contain UPNs, group properties, and dynamic membership rules. Protect the file and follow retention policies.

- If logging fails, required Graph changes are blocked; earlier actions may have completed. Review the log and Entra PIM.

- The file is not tamper-proof. Use Entra audit logs as the authoritative service-side record.

ONE GROUP

1. Enter a group name. The description and mail nickname are optional; if the mail nickname is blank, the app generates one from the group name.

2. Choose a group type. PIM-enabled role-assignable security is selected by default.

3. For dynamic groups, enter an Entra membership rule. Dynamic security groups can target Users or Devices; Microsoft 365 dynamic groups support Users only. Entra validates the rule.

4. For a PIM group, review the Member policy settings. Defaults are 8-hour activation, MFA, 365-day eligible assignments, no permanent active assignments, and MFA plus justification for active assignments.

5. Optionally enter user UPNs separated by commas, semicolons, or new lines to assign as PIM-eligible members. This does not make them active members.

6. Click Create group and confirm. Newly created PIM groups wait 10 seconds for Microsoft Entra propagation before PIM settings are applied. Review the outcome in the log.

RETRY PIM SETTINGS ON AN EXISTING GROUP

1. Select the PIM-enabled role-assignable group type and enter the existing group's exact display name or object ID. The group must be cloud-only, role-assignable security with assigned membership. If multiple groups have that name, use the object ID.

2. Review the Member settings, then select Apply PIM to existing group. Check the returned group name and ID before confirming.

3. This changes policy only: it does not create a group, assign roles, or add eligible members. It can partially update settings on failure.

4. Review the log and verify the settings and any prior eligible assignments in Entra PIM.

BULK GROUPS

1. Click Create CSV / XLSX template. XLSX creation requires ImportExcel; CSV does not.

2. Fill GroupName for each row. GroupType must be PIMRoleAssignable, SecurityAssigned, SecurityDynamic, Microsoft365Assigned, or Microsoft365Dynamic.

3. For dynamic groups, set MembershipType and provide MembershipRule. Security dynamic groups can use User or Device; Microsoft 365 dynamic groups support User only.

4. PIM fields can be set per row. Use True/False values. EligibleMemberUPNs accepts semicolon-separated UPNs.

5. Import the file. The GUI validates all rows before allowing creation. Confirm the group count before changes are submitted.

6. Review the results CSV saved beside the bulk input.

ASSIGNMENTS

1. Connect to Graph, then open Assignments. In One group, enter its exact display name or object ID and select Query one group.

2. In All groups (read-only), select Query all eligible to scan all role-assignable groups for current eligible members and owners. Select Scan permanent active to inspect active member assignments and direct members.

3. Review the results and optionally export them to Excel or CSV. Excel is the default when ImportExcel is installed. Results contain directory information; protect the exported file.

- If the permanent scan needs additional permission, choose whether to reconnect and request consent. An administrator may need to grant it.

- The all-groups eligible report omits principal ID and type. Query a group individually to see them.

- Enter the group name or ID in One group before selecting Add eligible members or Remove members. The popup shows that group as bold, non-editable text; close it to change groups in the main window. Query one group is read-only. Removal requires selecting categories and reviewing the resolved users and group before confirming.

- Removal can request additional consent: PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup for PIM active schedules, and RoleManagement.ReadWrite.Directory for direct membership of role-assignable groups.

- PIM active membership also appears in the group's direct members. Removing either can affect active access; removing eligibility may not end an activation. Owners and nested access are not removed. Check Entra after each request, especially after an error.

- Select Add eligible members to submit eligible member requests for an existing cloud-only, assigned-membership security or Microsoft 365 group. Enter the exact group name or ID in the main window, then enter one or more user UPNs in the popup separated by commas, semicolons, or new lines.

- Start defaults to now and end to one year later. Choose local start/end times or permanent eligibility if the group policy allows it. Review the group, users and schedule before confirming. The action may onboard a group to PIM.

- Role-assignable groups are included even if not onboarded to PIM. Other groups are not included. Direct members are candidates, not proof of permanence; verify in Entra. A member with a temporary PIM assignment may also have separate direct access.

IMPORTANT

- A role-assignable group must be a security group with assigned (not dynamic) membership. Entra does not allow dynamic membership for role-assignable groups.

- Creating a role-assignable group only makes it eligible to receive Entra roles. It does not assign any role. Assign the required role to the group separately in Entra PIM.

- The GUI automatically prefixes PIM group descriptions with "*DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal".

- EligibleMemberUPNs creates eligible membership assignments through PIM. The group policy is configured first. A failed PIM configuration can leave the group created; the log and results file report partial completion.

- Dynamic rules are submitted to Entra as entered. Validate the rule and membership impact before creating the group.

- PIM policy changes use Microsoft Graph role-management policy APIs and are read back after updating. Verify the resulting settings in the Entra PIM portal.
'@

function Show-HelpWindow
{
    if ($Script:HelpWindow -and -not $Script:HelpWindow.IsDisposed)
    {
        if ($Script:HelpWindow.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized)
        {
            $Script:HelpWindow.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        }
        $Script:HelpWindow.Activate()
        return
    }

    $HelpWindow = New-Object System.Windows.Forms.Form
    $HelpWindow.Text = 'PIM Group Manager - How to use'
    $HelpWindow.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $HelpWindow.Size = New-Object System.Drawing.Size(950, 760)
    $HelpWindow.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $HelpWindow.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $HelpWindow.BackColor = [System.Drawing.Color]::White
    $HelpWindow.ShowInTaskbar = $false

    $HelpTextBox = New-Object System.Windows.Forms.TextBox
    $HelpTextBox.Multiline = $true
    $HelpTextBox.ReadOnly = $true
    $HelpTextBox.WordWrap = $true
    $HelpTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $HelpTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $HelpTextBox.Location = New-Object System.Drawing.Point(16, 16)
    $HelpTextBox.Size = New-Object System.Drawing.Size(900, 690)
    $HelpTextBox.Anchor = 'Top, Bottom, Left, Right'
    $HelpTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $HelpTextBox.BackColor = [System.Drawing.Color]::White
    $HelpTextBox.Text = $Script:HelpContent
    $HelpWindow.Controls.Add($HelpTextBox)
    $HelpWindow.Add_FormClosed({ $Script:HelpWindow = $null })
    $Script:HelpWindow = $HelpWindow
    $HelpWindow.Show($Script:MainForm) | Out-Null
}

$Script:GraphScopes = @(
    'User.Read'
    'User.Read.All'
    'Group.ReadWrite.All'
    'RoleManagementPolicy.ReadWrite.AzureADGroup'
    'PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup'
)
$Script:ReportScope = 'PrivilegedAssignmentSchedule.Read.AzureADGroup'
$Script:AssignmentWriteScope = 'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'
$Script:RoleManagementWriteScope = 'RoleManagement.ReadWrite.Directory'
$Script:ReportRows = @()
$Script:ReportMode = $null
$Script:BulkRows = @()
$Script:BulkPath = ''
$Script:ConnectionStatusLabel = $null
$Script:MainForm = $null
$Script:HelpWindow = $null
$Script:GroupTypeBox = $null
$Script:PreviousGroupTypeIndex = 0
$Script:MembershipTypeBox = $null
$Script:MembershipRuleBox = $null
$Script:MembershipTypeLabel = $null
$Script:MembershipRuleLabel = $null
$Script:PimGroupBox = $null
$Script:EligibleMembersBox = $null
$Script:EligibleMembersLabel = $null
$Script:GroupNameBox = $null
$Script:ExistingGroupIdBox = $null
$Script:ExistingGroupIdLabel = $null
$Script:ExistingGroupHint = $null
$Script:ApplyExistingButton = $null
$Script:DescriptionBox = $null
$Script:MailNicknameBox = $null
$Script:ActivationHoursBox = $null
$Script:EligibilityDaysBox = $null
$Script:ActiveAssignmentDaysBox = $null
$Script:RequireMfaActivationBox = $null
$Script:RequireMfaActiveBox = $null
$Script:RequireJustificationActiveBox = $null
$Script:AllowPermanentEligibleBox = $null
$Script:AllowPermanentActiveBox = $null
$Script:BulkPathBox = $null
$Script:BulkGrid = $null
$Script:BulkCreateButton = $null
$Script:LogBox = $null
$Script:AuditPathBox = $null
$Script:AuditStatusLabel = $null
$Script:AuditBrowseButton = $null
$Script:AuditSessionId = [guid]::NewGuid().ToString()
$Script:AuditAccount = 'Not connected'
$Script:AuditTenant = 'Unknown'
$Script:AuditFailure = $null
$AuditDirectory = [Environment]::GetFolderPath('MyDocuments')
$Script:AuditPath = if ([string]::IsNullOrWhiteSpace($AuditDirectory)) { '' } else {
    Join-Path $AuditDirectory ("{0}-PIM-Group-Manager.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmm'))
}
$Script:TabControl = $null
$Script:TemplateColumns = @(
    'GroupName'
    'Description'
    'MailNickname'
    'GroupType'
    'MembershipType'
    'MembershipRule'
    'ActivationMaxHours'
    'RequireMFAOnActivation'
    'RequireMFAOnActiveAssignment'
    'RequireJustificationOnActiveAssignment'
    'AllowPermanentEligibleAssignments'
    'EligibleAssignmentDays'
    'AllowPermanentActiveAssignments'
    'ActiveAssignmentDays'
    'EligibleMemberUPNs'
)
 
$Script:GroupTypeOptions = @(
    'PIMRoleAssignable'
    'SecurityAssigned'
    'SecurityDynamic'
    'Microsoft365Assigned'
    'Microsoft365Dynamic'
)
 
# ==================================================
# ^^^^^ Prerequisites and Script Configuration ^^^^^
# ==================================================
 
# ==================================================
# MARK: Validation and Group Configuration Helpers
# ==================================================
 
function Show-Message
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$Message,
 
        [System.String]$Title = 'PIM Group Manager',
 
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
 
    [void][System.Windows.Forms.MessageBox]::Show(
        $Script:MainForm,
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    )
}
 
function Write-UiLog
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$Message,
 
        [switch]$Required
    )
 
    $Line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    if ($Script:LogBox)
    {
        $Script:LogBox.AppendText($Line + [Environment]::NewLine)
        $Script:LogBox.Refresh()
    }
 
    try
    {
        if ([string]::IsNullOrWhiteSpace($Script:AuditPath) -or
            -not [System.IO.Path]::IsPathRooted($Script:AuditPath) -or
            [System.IO.Path]::GetExtension($Script:AuditPath) -ne '.log')
        {
            throw 'Select an absolute path ending in .log.'
        }
        $Entry = [ordered]@{
            timestampUtc = [DateTime]::UtcNow.ToString('o')
            sessionId = $Script:AuditSessionId
            computer = [Environment]::MachineName
            windowsUser = '{0}\{1}' -f [Environment]::UserDomainName, [Environment]::UserName
            account = $Script:AuditAccount
            tenantId = $Script:AuditTenant
            message = $Message
        } | ConvertTo-Json -Compress -Depth 10
        [System.IO.File]::AppendAllText($Script:AuditPath, $Entry + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
        $Script:AuditFailure = $null
        if ($Script:AuditStatusLabel)
        {
            $Script:AuditStatusLabel.Text = 'Audit log'
            $Script:AuditStatusLabel.ForeColor = [System.Drawing.Color]::Black
        }
    }
    catch
    {
        $Script:AuditFailure = "Cannot write audit log '$($Script:AuditPath)': $($_.Exception.Message)"
        if ($Script:AuditStatusLabel)
        {
            $Script:AuditStatusLabel.Text = 'LOG ERROR'
            $Script:AuditStatusLabel.ForeColor = [System.Drawing.Color]::DarkRed
        }
        if ($Script:LogBox)
        {
            $Script:LogBox.AppendText($Script:AuditFailure + [Environment]::NewLine)
        }
        if ($Required)
        {
            throw "$($Script:AuditFailure) No further changes will be submitted until logging is restored."
        }
    }
}
 
function Get-RowValue
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Object]$Row,
 
        [Parameter(Mandatory = $true)]
        [System.String]$Name
    )
 
    $Property = $Row.PSObject.Properties[$Name]
    if ($null -eq $Property)
    {
        return $null
    }
 
    return $Property.Value
}
 
function ConvertTo-BooleanSetting
{
    param
    (
        [AllowNull()]
        [System.Object]$Value,
 
        [Parameter(Mandatory = $true)]
        [System.String]$FieldName,
 
        [System.Boolean]$DefaultValue
    )
 
    if ($null -eq $Value -or [System.String]::IsNullOrWhiteSpace([System.String]$Value))
    {
        return $DefaultValue
    }
 
    $ParsedValue = $false
    if (-not [System.Boolean]::TryParse(([System.String]$Value).Trim(), [ref]$ParsedValue))
    {
        throw "The '$FieldName' value '$Value' must be True or False."
    }
 
    return $ParsedValue
}
 
function ConvertTo-IntegerSetting
{
    param
    (
        [AllowNull()]
        [System.Object]$Value,
 
        [Parameter(Mandatory = $true)]
        [System.String]$FieldName,
 
        [Parameter(Mandatory = $true)]
        [System.Int32]$Minimum,
 
        [Parameter(Mandatory = $true)]
        [System.Int32]$Maximum,
 
        [Parameter(Mandatory = $true)]
        [System.Int32]$DefaultValue,
 
        [System.Boolean]$AllowBlank = $false
    )
 
    if ($null -eq $Value -or [System.String]::IsNullOrWhiteSpace([System.String]$Value))
    {
        if ($AllowBlank)
        {
            return $null
        }
 
        return $DefaultValue
    }
 
    $ParsedValue = 0
    if (-not [System.Int32]::TryParse(([System.String]$Value).Trim(), [ref]$ParsedValue))
    {
        throw "The '$FieldName' value '$Value' must be a whole number."
    }
 
    if ($ParsedValue -lt $Minimum -or $ParsedValue -gt $Maximum)
    {
        throw "The '$FieldName' value must be between $Minimum and $Maximum."
    }
 
    return $ParsedValue
}
 
function Get-DescriptionForGroupType
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupType,
 
        [AllowEmptyString()]
        [System.String]$Description
    )
 
    $PimPrefix = '*DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal'
    if ($GroupType -ne 'PIMRoleAssignable')
    {
        return $Description.Trim()
    }
 
    if ([System.String]::IsNullOrWhiteSpace($Description))
    {
        return $PimPrefix
    }
 
    if ($Description.StartsWith($PimPrefix, [System.StringComparison]::OrdinalIgnoreCase))
    {
        return $Description.Trim()
    }
 
    return "$PimPrefix | $($Description.Trim())"
}
 
function Get-SafeMailNickname
{
    param([Parameter(Mandatory = $true)][System.String]$DisplayName)
 
    $Nickname = $DisplayName -replace '[^A-Za-z0-9._-]', ''
    $Nickname = $Nickname.Trim('.', '-', '_')
    if ([System.String]::IsNullOrWhiteSpace($Nickname))
    {
        throw 'MailNickname must contain at least one letter or number.'
    }
 
    if ($Nickname.Length -gt 64)
    {
        $Nickname = $Nickname.Substring(0, 64)
    }
 
    return $Nickname
}
 
function ConvertTo-EligibleMemberUPNs
{
    param([AllowEmptyString()][string]$Text)

    $Seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $EligibleMemberUPNs = [System.Collections.Generic.List[string]]::new()
    foreach ($Entry in ($Text -split '[;,\r\n]+'))
    {
        $Upn = $Entry.Trim()
        if ([string]::IsNullOrWhiteSpace($Upn)) { continue }
        if ($Upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$')
        {
            throw "EligibleMemberUPNs contains an invalid UPN: $Upn"
        }
        if ($Seen.Add($Upn)) { $EligibleMemberUPNs.Add($Upn) }
    }
    return $EligibleMemberUPNs.ToArray()
}

function ConvertTo-GroupConfiguration
{
    param([Parameter(Mandatory = $true)][System.Object]$Row)
 
    $GroupName = ([System.String](Get-RowValue -Row $Row -Name 'GroupName')).Trim()
    if ([System.String]::IsNullOrWhiteSpace($GroupName))
    {
        throw 'GroupName is required.'
    }
 
    if ($GroupName -eq '__REPLACE_ME__')
    {
        throw 'Replace the XLSX template placeholder GroupName before creating groups.'
    }
 
    $GroupType = [System.String](Get-RowValue -Row $Row -Name 'GroupType')
    if ([System.String]::IsNullOrWhiteSpace($GroupType))
    {
        throw "GroupType is required. Choose one of: $($Script:GroupTypeOptions -join ', ')."
    }
 
    if ($Script:GroupTypeOptions -notcontains $GroupType)
    {
        throw "Unsupported GroupType '$GroupType'. Use: $($Script:GroupTypeOptions -join ', ')."
    }
 
    $Description = Get-DescriptionForGroupType -GroupType $GroupType -Description ([System.String](Get-RowValue -Row $Row -Name 'Description'))
    $MailNickname = [System.String](Get-RowValue -Row $Row -Name 'MailNickname')
    if ([System.String]::IsNullOrWhiteSpace($MailNickname))
    {
        $MailNickname = Get-SafeMailNickname -DisplayName $GroupName
    }
    elseif ($MailNickname -notmatch '^[A-Za-z0-9._-]{1,64}$')
    {
        throw 'MailNickname must contain 1-64 letters, numbers, periods, underscores, or hyphens.'
    }
 
    $IsDynamic = $GroupType -in @('SecurityDynamic', 'Microsoft365Dynamic')
    $IsMicrosoft365 = $GroupType -in @('Microsoft365Assigned', 'Microsoft365Dynamic')
    $MembershipType = [System.String](Get-RowValue -Row $Row -Name 'MembershipType')
    if ([System.String]::IsNullOrWhiteSpace($MembershipType))
    {
        $MembershipType = 'User'
    }
 
    if ($MembershipType -notin @('User', 'Device'))
    {
        throw "MembershipType '$MembershipType' must be User or Device."
    }
 
    if ($IsMicrosoft365 -and $MembershipType -eq 'Device')
    {
        throw 'Microsoft 365 groups support user membership only.'
    }
 
    $MembershipRule = [System.String](Get-RowValue -Row $Row -Name 'MembershipRule')
    if ($IsDynamic -and [System.String]::IsNullOrWhiteSpace($MembershipRule))
    {
        throw 'MembershipRule is required for a dynamic group.'
    }
 
    $EligibleMemberText = [System.String](Get-RowValue -Row $Row -Name 'EligibleMemberUPNs')
    $EligibleMemberUPNs = @(ConvertTo-EligibleMemberUPNs -Text $EligibleMemberText)
 
    if ($EligibleMemberUPNs.Count -gt 0 -and $GroupType -ne 'PIMRoleAssignable')
    {
        throw 'EligibleMemberUPNs can only be used with a PIM-enabled role-assignable group.'
    }
 
    $ActivationMaxHours = ConvertTo-IntegerSetting `
        -Value (Get-RowValue -Row $Row -Name 'ActivationMaxHours') `
        -FieldName 'ActivationMaxHours' -Minimum 1 -Maximum 24 -DefaultValue 8
    $EligibleAssignmentDays = ConvertTo-IntegerSetting `
        -Value (Get-RowValue -Row $Row -Name 'EligibleAssignmentDays') `
        -FieldName 'EligibleAssignmentDays' -Minimum 1 -Maximum 3650 -DefaultValue 365
    $ActiveAssignmentDays = ConvertTo-IntegerSetting `
        -Value (Get-RowValue -Row $Row -Name 'ActiveAssignmentDays') `
        -FieldName 'ActiveAssignmentDays' -Minimum 1 -Maximum 3650 -DefaultValue 365 -AllowBlank $true
 
    return [PSCustomObject]@{
        GroupName                         = $GroupName
        Description                       = $Description
        MailNickname                      = $MailNickname
        GroupType                         = $GroupType
        IsDynamic                         = $IsDynamic
        IsMicrosoft365                    = $IsMicrosoft365
        MembershipType                    = $MembershipType
        MembershipRule                    = $MembershipRule.Trim()
        ActivationMaxHours                = $ActivationMaxHours
        RequireMFAOnActivation            = ConvertTo-BooleanSetting (Get-RowValue -Row $Row -Name 'RequireMFAOnActivation') 'RequireMFAOnActivation' $true
        RequireMFAOnActiveAssignment      = ConvertTo-BooleanSetting (Get-RowValue -Row $Row -Name 'RequireMFAOnActiveAssignment') 'RequireMFAOnActiveAssignment' $true
        RequireJustificationOnActiveAssignment = ConvertTo-BooleanSetting (Get-RowValue -Row $Row -Name 'RequireJustificationOnActiveAssignment') 'RequireJustificationOnActiveAssignment' $true
        AllowPermanentEligibleAssignments = ConvertTo-BooleanSetting (Get-RowValue -Row $Row -Name 'AllowPermanentEligibleAssignments') 'AllowPermanentEligibleAssignments' $false
        EligibleAssignmentDays            = $EligibleAssignmentDays
        AllowPermanentActiveAssignments   = ConvertTo-BooleanSetting (Get-RowValue -Row $Row -Name 'AllowPermanentActiveAssignments') 'AllowPermanentActiveAssignments' $false
        ActiveAssignmentDays              = $ActiveAssignmentDays
        EligibleMemberUPNs                = $EligibleMemberUPNs
    }
}
 
# ==================================================
# ^^^^^ Validation and Group Configuration Helpers ^^^^^
# ==================================================
 
# ==================================================
# MARK: Microsoft Graph Connection
# ==================================================
 
function Get-GraphContext
{
    if (-not (Get-Module -Name Microsoft.Graph.Authentication))
    {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }
 
    return Get-MgContext
}
 
function Connect-ToGraph
{
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication))
    {
        throw 'Microsoft.Graph.Authentication is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
 
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes ($Script:GraphScopes + $Script:ReportScope) -ContextScope Process -ErrorAction Stop
    Test-GraphConnection
}
 
function Test-GraphConnection
{
    $Context = Get-GraphContext
    if ($null -eq $Context -or [System.String]::IsNullOrWhiteSpace($Context.Account))
    {
        throw 'No Microsoft Graph account is connected.'
    }
 
    $Me = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName' -ErrorAction Stop
    $Script:AuditAccount = $Me.userPrincipalName
    $Script:AuditTenant = $Context.TenantId
    $MissingScopes = @($Script:GraphScopes | Where-Object { $Context.Scopes -notcontains $_ })
    $ScopeNote = if ($MissingScopes.Count -gt 0) { " Missing consented scopes: $($MissingScopes -join ', ')." } else { '' }
    if (-not (Test-ReportAssignmentScope -Context $Context)) { $ScopeNote += ' Permanent scan permission missing.' }
    $Script:ConnectionStatusLabel.Text = "Connected: $($Me.userPrincipalName) | Tenant: $($Context.TenantId)$ScopeNote"
    $Script:ConnectionStatusLabel.ForeColor = if ($MissingScopes.Count -gt 0 -or -not (Test-ReportAssignmentScope -Context $Context)) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DarkGreen }
 
    if ($MissingScopes.Count -gt 0)
    {
        Write-UiLog "Graph connectivity succeeded, but the token is missing: $($MissingScopes -join ', '). Reconnect and consent."
        return $false
    }
 
    Write-UiLog "Graph connection check succeeded for $($Me.userPrincipalName)."
    return $true
}
 
function Assert-GraphReady
{
    $Context = Get-GraphContext
    if ($null -eq $Context -or [System.String]::IsNullOrWhiteSpace($Context.Account))
    {
        throw 'Connect to Microsoft Graph and check the connection before creating groups.'
    }
 
    $Script:AuditAccount = $Context.Account
    $Script:AuditTenant = $Context.TenantId
    $MissingScopes = @($Script:GraphScopes | Where-Object { $Context.Scopes -notcontains $_ })
    if ($MissingScopes.Count -gt 0)
    {
        throw "Reconnect to Microsoft Graph and consent to: $($MissingScopes -join ', ')."
    }
}
 
# ==================================================
# ^^^^^ Microsoft Graph Connection ^^^^^
# ==================================================
 
# ==================================================
# MARK: Entra Group Creation
# ==================================================
 
function New-EntraGroup
{
    param([Parameter(Mandatory = $true)][System.Object]$Configuration)
 
    $Body = [ordered]@{
        displayName     = $Configuration.GroupName
        description     = $Configuration.Description
        mailNickname    = $Configuration.MailNickname
        mailEnabled     = [bool]$Configuration.IsMicrosoft365
        securityEnabled = [bool](-not $Configuration.IsMicrosoft365)
        groupTypes      = @()
    }
 
    if ($Configuration.GroupType -eq 'PIMRoleAssignable')
    {
        $Body.isAssignableToRole = $true
    }
 
    if ($Configuration.IsMicrosoft365)
    {
        $Body.groupTypes += 'Unified'
    }
 
    if ($Configuration.IsDynamic)
    {
        $Body.groupTypes += 'DynamicMembership'
        $Body.membershipRule = $Configuration.MembershipRule
        $Body.membershipRuleProcessingState = 'On'
    }
 
    $ConfigurationJson = $Configuration | ConvertTo-Json -Compress -Depth 8
    Write-UiLog -Message "Submitting group creation: $ConfigurationJson" -Required
    try
    {
        $Response = Invoke-MgGraphRequest `
            -Method POST `
            -Uri 'https://graph.microsoft.com/v1.0/groups' `
            -Body ($Body | ConvertTo-Json -Depth 8) `
            -ContentType 'application/json' `
            -ErrorAction Stop
    }
    catch
    {
        throw "Graph group creation failed: $($_.Exception.Message) $($_.ErrorDetails.Message)"
    }
 
    if ([System.String]::IsNullOrWhiteSpace($Response.id))
    {
        throw "Microsoft Graph returned no group ID for '$($Configuration.GroupName)'."
    }
 
    return $Response
}
 
# ==================================================
# ^^^^^ Entra Group Creation ^^^^^
# ==================================================
 
# ==================================================
# MARK: PIM Member Policy and Existing Group Recovery
# ==================================================
 
function Invoke-PimGraphRequest
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'PATCH')]
        [System.String]$Method,
 
        [Parameter(Mandatory = $true)]
        [System.String]$Uri,
 
        [AllowNull()]
        [System.Collections.IDictionary]$Body
    )
 
    $StatusCode = $null
    $Parameters = @{
        Method = $Method
        Uri = $Uri
        SkipHttpErrorCheck = $true
        StatusCodeVariable = 'StatusCode'
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body)
    {
        $Parameters.Body = $Body | ConvertTo-Json -Depth 10
        $Parameters.ContentType = 'application/json'
    }
    try
    {
        $Response = Invoke-MgGraphRequest @Parameters
    }
    catch
    {
        throw "PIM Graph $Method $Uri failed: $($_.Exception.Message) $($_.ErrorDetails.Message)"
    }
    if ($null -eq $StatusCode)
    {
        throw "PIM Graph $Method $Uri returned no HTTP status; the result cannot be verified."
    }
    if ([int]$StatusCode -lt 200 -or [int]$StatusCode -ge 300)
    {
        $Message = "PIM Graph $Method $Uri returned HTTP $StatusCode"
        if ($Response.error)
        {
            $Message += ': ' + ($Response.error | ConvertTo-Json -Depth 10 -Compress)
        }
        elseif ($Response)
        {
            $Message += ": $Response"
        }
        $Exception = New-Object System.Exception($Message)
        $Exception.Data['HttpStatusCode'] = [int]$StatusCode
        throw $Exception
    }
    return $Response
}
 
function Get-PimMemberPolicyAssignment
{
    param([Parameter(Mandatory = $true)][System.String]$GroupId)
 
    $Filter = "scopeId eq '$GroupId' and scopeType eq 'Group'"
    $Uri = 'https://graph.microsoft.com/v1.0/policies/roleManagementPolicyAssignments?$filter=' + [uri]::EscapeDataString($Filter)
    $Response = Invoke-PimGraphRequest -Method GET -Uri $Uri
    $Assignments = @($Response.value | Where-Object { $_.roleDefinitionId -eq 'member' })
 
    if ($Assignments.Count -ne 1 -or [System.String]::IsNullOrWhiteSpace($Assignments[0].policyId))
    {
        $Exception = New-Object System.Exception("Could not resolve exactly one PIM Member policy assignment for group $GroupId (found $($Assignments.Count)).")
        if ($Assignments.Count -eq 0)
        {
            $Exception.Data['HttpStatusCode'] = 404
        }
        throw $Exception
    }
 
    return $Assignments[0]
}
 
function Get-PimPolicyRule
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$PolicyId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$RuleId
    )
 
    $Uri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/$PolicyId/rules/$RuleId"
    return Invoke-PimGraphRequest -Method GET -Uri $Uri
}
 
function Set-PimPolicyRule
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$PolicyId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$RuleId,
 
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Body
    )
 
    $Uri = "https://graph.microsoft.com/v1.0/policies/roleManagementPolicies/$PolicyId/rules/$RuleId"
    $RuleJson = $Body | ConvertTo-Json -Compress -Depth 10
    Write-UiLog -Message "Submitting PIM rule '$RuleId' for group $GroupId, policy ${PolicyId}: $RuleJson" -Required
    Invoke-PimGraphRequest -Method PATCH -Uri $Uri -Body $Body | Out-Null
    Write-UiLog -Message "PIM rule '$RuleId' updated for group $GroupId, policy $PolicyId." -Required
}
 
function Update-PimEnablementRule
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$RuleId,
 
        [Parameter(Mandatory = $true)]
        [System.String[]]$ManagedRequirements,
 
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$DesiredRequirements
    )
 
    $Assignment = Get-PimMemberPolicyAssignment -GroupId $GroupId
    $CurrentRule = Get-PimPolicyRule -PolicyId $Assignment.policyId -RuleId $RuleId
    $EnabledRules = [System.Collections.Generic.List[System.String]]::new()
    foreach ($EnabledRule in @($CurrentRule.enabledRules))
    {
        if (-not [System.String]::IsNullOrWhiteSpace([System.String]$EnabledRule) -and $ManagedRequirements -notcontains [System.String]$EnabledRule)
        {
            $EnabledRules.Add([System.String]$EnabledRule)
        }
    }
 
    foreach ($Requirement in $ManagedRequirements)
    {
        if ($DesiredRequirements[$Requirement] -and -not $EnabledRules.Contains($Requirement))
        {
            $EnabledRules.Add($Requirement)
        }
    }
 
    $Body = [ordered]@{
        '@odata.type' = '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule'
        id            = $RuleId
        enabledRules  = @($EnabledRules.ToArray())
    }
    Set-PimPolicyRule -GroupId $GroupId -PolicyId $Assignment.policyId -RuleId $RuleId -Body $Body
}
 
function Update-PimExpirationRule
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$RuleId,
 
        [Parameter(Mandatory = $true)]
        [System.Boolean]$ExpirationRequired,
 
        [AllowNull()]
        [System.String]$MaximumDuration
    )
 
    $Assignment = Get-PimMemberPolicyAssignment -GroupId $GroupId
    $CurrentRule = Get-PimPolicyRule -PolicyId $Assignment.policyId -RuleId $RuleId
    if ([System.String]::IsNullOrWhiteSpace($MaximumDuration))
    {
        $MaximumDuration = [System.String]$CurrentRule.maximumDuration
    }
 
    if ($ExpirationRequired -and [System.String]::IsNullOrWhiteSpace($MaximumDuration))
    {
        throw "PIM policy rule '$RuleId' has no maximum duration to preserve."
    }
 
    $Body = [ordered]@{
        '@odata.type'       = '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule'
        id                  = $RuleId
        isExpirationRequired = $ExpirationRequired
    }
    if (-not [System.String]::IsNullOrWhiteSpace($MaximumDuration))
    {
        $Body.maximumDuration = $MaximumDuration
    }
 
    Set-PimPolicyRule -GroupId $GroupId -PolicyId $Assignment.policyId -RuleId $RuleId -Body $Body
}
 
function Invoke-PimRuleUpdate
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$RuleId,
 
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Parameters,
 
        [Parameter(Mandatory = $true)]
        [ValidateSet('Expiration', 'Enablement')]
        [System.String]$Kind
    )
 
    for ($Attempt = 1; $Attempt -le 5; $Attempt++)
    {
        try
        {
            if ($Kind -eq 'Expiration')
            {
                Update-PimExpirationRule -GroupId $GroupId -RuleId $RuleId @Parameters
            }
            else
            {
                Update-PimEnablementRule -GroupId $GroupId -RuleId $RuleId @Parameters
            }
            return
        }
        catch
        {
            if ($_.Exception.Data['HttpStatusCode'] -ne 404 -or $Attempt -eq 5)
            {
                throw "PIM rule '$RuleId' failed for group $GroupId (attempt $Attempt): $($_.Exception.Message)"
            }
            $DelaySeconds = $Attempt * 5
            Write-UiLog "PIM rule '$RuleId' returned 404; refreshing the policy and retrying in $DelaySeconds seconds (attempt $Attempt of 5)."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}
 
function Assert-PimMemberSettings
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.Object]$Configuration
    )
 
    $Rules = @(
        @{ Id = 'Expiration_EndUser_Assignment'; Required = $true; Duration = "PT$($Configuration.ActivationMaxHours)H" }
        @{ Id = 'Enablement_EndUser_Assignment'; Requirements = @{ MultiFactorAuthentication = $Configuration.RequireMFAOnActivation } }
        @{ Id = 'Expiration_Admin_Eligibility'; Required = (-not $Configuration.AllowPermanentEligibleAssignments); Duration = "P$($Configuration.EligibleAssignmentDays)D" }
        @{ Id = 'Expiration_Admin_Assignment'; Required = (-not $Configuration.AllowPermanentActiveAssignments); Duration = $(if ($null -ne $Configuration.ActiveAssignmentDays) { "P$($Configuration.ActiveAssignmentDays)D" } else { $null }) }
        @{ Id = 'Enablement_Admin_Assignment'; Requirements = @{ MultiFactorAuthentication = $Configuration.RequireMFAOnActiveAssignment; Justification = $Configuration.RequireJustificationOnActiveAssignment } }
    )
 
    for ($Attempt = 1; $Attempt -le 5; $Attempt++)
    {
        try
        {
            $Mismatches = [System.Collections.Generic.List[System.String]]::new()
            $Assignment = Get-PimMemberPolicyAssignment -GroupId $GroupId
            foreach ($Expected in $Rules)
            {
                $Actual = Get-PimPolicyRule -PolicyId $Assignment.policyId -RuleId $Expected.Id
                if ($Expected.ContainsKey('Required'))
                {
                    if ($null -eq $Actual.isExpirationRequired -or [bool]$Actual.isExpirationRequired -ne $Expected.Required)
                    {
                        $Mismatches.Add("$($Expected.Id): expiration required expected $($Expected.Required), returned '$($Actual.isExpirationRequired)'")
                    }
                    if ($Expected.Duration -and ([System.String]::IsNullOrWhiteSpace([string]$Actual.maximumDuration) -or
                        [System.Xml.XmlConvert]::ToTimeSpan([string]$Actual.maximumDuration) -ne [System.Xml.XmlConvert]::ToTimeSpan($Expected.Duration)))
                    {
                        $Mismatches.Add("$($Expected.Id): duration expected $($Expected.Duration), returned '$($Actual.maximumDuration)'")
                    }
                }
                else
                {
                    if ($null -eq $Actual.enabledRules)
                    {
                        $Mismatches.Add("$($Expected.Id): Graph returned no enabledRules")
                        continue
                    }
                    foreach ($Requirement in $Expected.Requirements.Keys)
                    {
                        $Enabled = @($Actual.enabledRules) -contains $Requirement
                        if ($Enabled -ne $Expected.Requirements[$Requirement])
                        {
                            $Mismatches.Add("$($Expected.Id): $Requirement expected $($Expected.Requirements[$Requirement]), returned $Enabled")
                        }
                    }
                }
            }
            if ($Mismatches.Count -eq 0) { return }
            $Failure = "PIM settings verification mismatch for group ${GroupId}: $($Mismatches -join '; ')"
        }
        catch
        {
            if ($_.Exception.Data['HttpStatusCode'] -ne 404)
            {
                throw "PIM settings verification failed for group ${GroupId}: $($_.Exception.Message)"
            }
            $Failure = "PIM settings verification failed for group ${GroupId}: $($_.Exception.Message)"
        }
        if ($Attempt -eq 5) { throw $Failure }
        $DelaySeconds = $Attempt * 5
        Write-UiLog "$Failure Retrying verification in $DelaySeconds seconds (attempt $Attempt of 5)."
        Start-Sleep -Seconds $DelaySeconds
    }
}
 
function Set-PimMemberSettings
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.Object]$Configuration
    )
 
    Invoke-PimRuleUpdate -GroupId $GroupId -RuleId 'Expiration_EndUser_Assignment' -Kind Expiration -Parameters @{
        ExpirationRequired = $true
        MaximumDuration = "PT$($Configuration.ActivationMaxHours)H"
    }
    Invoke-PimRuleUpdate -GroupId $GroupId -RuleId 'Enablement_EndUser_Assignment' -Kind Enablement -Parameters @{
        ManagedRequirements = @('MultiFactorAuthentication')
        DesiredRequirements = @{ MultiFactorAuthentication = $Configuration.RequireMFAOnActivation }
    }
    Invoke-PimRuleUpdate -GroupId $GroupId -RuleId 'Expiration_Admin_Eligibility' -Kind Expiration -Parameters @{
        ExpirationRequired = (-not $Configuration.AllowPermanentEligibleAssignments)
        MaximumDuration = "P$($Configuration.EligibleAssignmentDays)D"
    }
 
    $ActiveAssignmentDuration = $null
    if ($null -ne $Configuration.ActiveAssignmentDays)
    {
        $ActiveAssignmentDuration = "P$($Configuration.ActiveAssignmentDays)D"
    }
    Invoke-PimRuleUpdate -GroupId $GroupId -RuleId 'Expiration_Admin_Assignment' -Kind Expiration -Parameters @{
        ExpirationRequired = (-not $Configuration.AllowPermanentActiveAssignments)
        MaximumDuration = $ActiveAssignmentDuration
    }
    Invoke-PimRuleUpdate -GroupId $GroupId -RuleId 'Enablement_Admin_Assignment' -Kind Enablement -Parameters @{
        ManagedRequirements = @('MultiFactorAuthentication', 'Justification')
        DesiredRequirements = @{
            MultiFactorAuthentication = $Configuration.RequireMFAOnActiveAssignment
            Justification             = $Configuration.RequireJustificationOnActiveAssignment
        }
    }
    Assert-PimMemberSettings -GroupId $GroupId -Configuration $Configuration
}
 
function Get-ExistingPimGroup
{
    param([Parameter(Mandatory = $true)][System.String]$GroupNameOrId, [switch]$AllowNonRoleAssignable)
 
    $SearchTerm = $GroupNameOrId.Trim()
    if ([System.String]::IsNullOrWhiteSpace($SearchTerm))
    {
        throw 'Enter an existing group display name or object ID.'
    }

    $ParsedId = [guid]::Empty
    $IsObjectId = [guid]::TryParse($SearchTerm, [ref]$ParsedId)
    if ($IsObjectId)
    {
        $Uri = 'https://graph.microsoft.com/v1.0/groups/' + $ParsedId.ToString() + '?$select=id,displayName,securityEnabled,mailEnabled,isAssignableToRole,groupTypes,onPremisesSyncEnabled'
        $Group = Invoke-PimGraphRequest -Method GET -Uri $Uri
    }
    else
    {
        $EscapedDisplayName = $SearchTerm.Replace("'", "''")
        $Filter = "displayName eq '$EscapedDisplayName'"
        $Uri = 'https://graph.microsoft.com/v1.0/groups?$filter=' + [uri]::EscapeDataString($Filter) + '&$select=id,displayName,securityEnabled,mailEnabled,isAssignableToRole,groupTypes,onPremisesSyncEnabled&$top=2'
        $GroupMatches = @(Get-GraphCollection -Uri $Uri)
        if ($GroupMatches.Count -eq 0)
        {
            throw "No group with the exact display name '$SearchTerm' was found. No settings were changed."
        }
        if ($GroupMatches.Count -gt 1)
        {
            throw "Multiple groups have the display name '$SearchTerm'. Enter the group's object ID instead. No settings were changed."
        }
        $Group = $GroupMatches[0]
    }

    $VerifiedId = [guid]::Empty
    $IsSecurityGroup = $Group.securityEnabled -eq $true -and $Group.mailEnabled -eq $false
    $IsMicrosoft365Group = $Group.mailEnabled -eq $true -and @($Group.groupTypes) -contains 'Unified'
    if (($IsObjectId -and $Group.id -ne $ParsedId.ToString()) -or
        -not [guid]::TryParse([string]$Group.id, [ref]$VerifiedId) -or [string]::IsNullOrWhiteSpace($Group.displayName) -or
        (-not $IsSecurityGroup -and -not ($AllowNonRoleAssignable -and $IsMicrosoft365Group)) -or
        (-not $AllowNonRoleAssignable -and $Group.isAssignableToRole -ne $true) -or
        $Group.onPremisesSyncEnabled -eq $true -or @($Group.groupTypes) -contains 'DynamicMembership')
    {
        $Requirement = if ($AllowNonRoleAssignable) { 'cloud-only security or Microsoft 365 group with assigned membership' } else { 'cloud-only, role-assignable security group with assigned membership' }
        throw "Group '$($Group.displayName)' [$($Group.id)] is not a verified $Requirement. No settings were changed."
    }
    return $Group
}
 
# ==================================================
# ^^^^^ PIM Member Policy and Existing Group Recovery ^^^^^
# ==================================================
 
# ==================================================
# MARK: PIM Eligible Member Assignments
# ==================================================
 
function Get-EntraUserByUpn
{
    param([Parameter(Mandatory = $true)][System.String]$UserPrincipalName)
 
    $EscapedUpn = $UserPrincipalName.Replace("'", "''")
    $Filter = "userPrincipalName eq '$EscapedUpn'"
    $Uri = 'https://graph.microsoft.com/v1.0/users?$filter=' + [uri]::EscapeDataString($Filter) + '&$select=id,userPrincipalName'
    $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
    $Users = @($Response.value)
 
    if ($Users.Count -ne 1 -or [System.String]::IsNullOrWhiteSpace($Users[0].id))
    {
        throw "Could not resolve exactly one Entra user for '$UserPrincipalName'."
    }
 
    return $Users[0]
}
 
function Add-PimEligibleMember
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]$GroupId,
 
        [Parameter(Mandatory = $true)]
        [System.String]$UserPrincipalName,
 
        [Parameter(Mandatory = $true)]
        [System.Object]$Configuration,

        [AllowNull()]
        [System.Collections.IDictionary]$ScheduleInfo,

        [AllowNull()]
        [System.Object]$ResolvedUser
    )
 
    $User = if ($ResolvedUser) { $ResolvedUser } else { Get-EntraUserByUpn -UserPrincipalName $UserPrincipalName }
    if ([string]::IsNullOrWhiteSpace([string]$User.id) -or $User.userPrincipalName -ne $UserPrincipalName)
    {
        throw "Could not verify the resolved user for '$UserPrincipalName'."
    }
    if ($null -eq $ScheduleInfo)
    {
        $ScheduleInfo = @{
            startDateTime = [DateTime]::UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
        }
        if ($Configuration.AllowPermanentEligibleAssignments)
        {
            $ScheduleInfo.expiration = @{ type = 'noExpiration' }
        }
        else
        {
            $ScheduleInfo.expiration = @{
                type        = 'afterDateTime'
                endDateTime = [DateTime]::UtcNow.AddDays($Configuration.EligibleAssignmentDays).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
            }
        }
    }
 
    $Body = @{
        accessId     = 'member'
        principalId  = $User.id
        groupId      = $GroupId
        action       = 'AdminAssign'
        scheduleInfo = $ScheduleInfo
        justification = 'Eligible membership created by PIM Group Manager.'
    }
 
    Write-UiLog -Message "Submitting eligible member assignment for '$UserPrincipalName' [$($User.id)] to group $GroupId; expiration: $($ScheduleInfo.expiration | ConvertTo-Json -Compress)." -Required
    try
    {
        Invoke-MgGraphRequest `
            -Method POST `
            -Uri 'https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilityScheduleRequests' `
            -Body ($Body | ConvertTo-Json -Depth 8) `
            -ContentType 'application/json' `
            -ErrorAction Stop | Out-Null
    }
    catch
    {
        throw "Eligible member request for '$UserPrincipalName' in group $GroupId failed: $($_.Exception.Message) $($_.ErrorDetails.Message)"
    }
    Write-UiLog -Message "Eligible member assignment request accepted for '$UserPrincipalName' [$($User.id)] in group $GroupId." -Required
}

function New-EligibleAssignmentSchedule
{
    param(
        [bool]$StartNow,
        [datetime]$StartLocal,
        [bool]$Permanent,
        [datetime]$EndLocal,
        [datetime]$NowUtc = [datetime]::UtcNow
    )

    $TimeZone = [System.TimeZoneInfo]::Local
    $Start = [datetime]::SpecifyKind($StartLocal, [DateTimeKind]::Unspecified)
    $End = [datetime]::SpecifyKind($EndLocal, [DateTimeKind]::Unspecified)
    if (-not $StartNow -and ($TimeZone.IsInvalidTime($Start) -or $TimeZone.IsAmbiguousTime($Start)))
    {
        throw 'The selected start time is invalid or ambiguous in the local time zone. Choose another time.'
    }
    if (-not $Permanent -and ($TimeZone.IsInvalidTime($End) -or $TimeZone.IsAmbiguousTime($End)))
    {
        throw 'The selected end time is invalid or ambiguous in the local time zone. Choose another time.'
    }
    $StartUtc = if ($StartNow) { $NowUtc } else { [System.TimeZoneInfo]::ConvertTimeToUtc($Start, $TimeZone) }
    if (-not $StartNow -and $StartUtc -lt $NowUtc) { throw 'The start time must not be in the past.' }
    $ScheduleInfo = @{ startDateTime = $StartUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'") }
    if ($Permanent)
    {
        $ScheduleInfo.expiration = @{ type = 'noExpiration' }
    }
    else
    {
        $EndUtc = [System.TimeZoneInfo]::ConvertTimeToUtc($End, $TimeZone)
        if ($EndUtc -le $StartUtc) { throw 'The end time must be after the start time.' }
        $ScheduleInfo.expiration = @{ type = 'afterDateTime'; endDateTime = $EndUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'") }
    }
    return $ScheduleInfo
}
 
# ==================================================
# ^^^^^ PIM Eligible Member Assignments ^^^^^
# ==================================================
 
# ==================================================
# MARK: Read-only Assignment Reports
# ==================================================

function Get-GraphCollection
{
    param([Parameter(Mandatory = $true)][string]$Uri)

    $Seen = @{}
    while ($Uri)
    {
        $ParsedUri = [uri]$Uri
        if ($ParsedUri.Scheme -ne 'https' -or $ParsedUri.Host -ne 'graph.microsoft.com' -or
            -not $ParsedUri.IsDefaultPort -or $Seen.ContainsKey($Uri))
        {
            throw "Invalid or repeated Graph pagination URL: $Uri"
        }
        $Seen[$Uri] = $true
        $Response = Invoke-PimGraphRequest -Method GET -Uri $Uri
        if ($Response -isnot [System.Collections.IDictionary] -or -not $Response.Contains('value') -or $null -eq $Response.value)
        {
            throw "Graph returned an invalid collection for $Uri."
        }
        foreach ($Item in $Response.value) { $Item }
        $Uri = [string]$Response.'@odata.nextLink'
    }
}

function Get-ReportGroups
{
    $Uri = 'https://graph.microsoft.com/v1.0/groups?$filter=isAssignableToRole%20eq%20true&$select=id,displayName,isAssignableToRole'
    foreach ($Group in @(Get-GraphCollection -Uri $Uri))
    {
        $GroupId = [guid]::Empty
        if (-not [guid]::TryParse([string]$Group.id, [ref]$GroupId) -or
            [string]::IsNullOrWhiteSpace($Group.displayName) -or $Group.isAssignableToRole -ne $true)
        {
            throw 'Graph returned a role-assignable group without a valid ID, name, or role-assignable flag.'
        }
        $Group
    }
}

function Get-GroupReportInstances
{
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][ValidateSet('eligibilityScheduleInstances', 'assignmentScheduleInstances')][string]$Kind
    )
    $ParsedId = [guid]::Empty
    if (-not [guid]::TryParse($GroupId, [ref]$ParsedId)) { throw "Invalid group ID: $GroupId" }
    $Filter = [uri]::EscapeDataString("groupId eq '$ParsedId'")
    $Uri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/${Kind}?`$filter=$Filter&`$expand=principal"
    Get-GraphCollection -Uri $Uri
}

function Test-CurrentReportInstance
{
    param([Parameter(Mandatory = $true)]$Instance, [Parameter(Mandatory = $true)][datetimeoffset]$Now)
    if ($Instance -isnot [System.Collections.IDictionary] -or -not $Instance.principalId -or -not $Instance.accessId -or
        -not $Instance.Contains('startDateTime') -or -not $Instance.Contains('endDateTime'))
    {
        throw 'Graph returned an assignment instance without its principal, access type, start, or end date.'
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Instance.startDateTime) -and
        [datetimeoffset]::Parse([string]$Instance.startDateTime) -gt $Now) { return $false }
    if ($null -ne $Instance.endDateTime -and [string]$Instance.endDateTime -ne '')
    {
        return ([datetimeoffset]::Parse([string]$Instance.endDateTime) -gt $Now)
    }
    return $true
}

function New-AssignmentReportRow
{
    param($Group, [string]$Type, [string]$Access, [string]$PrincipalId, $Principal, $Start, $End)
    [PSCustomObject][ordered]@{
        GroupName = [string]$Group.displayName
        GroupId = [string]$Group.id
        Finding = $Type
        Access = $Access
        PrincipalName = [string]$Principal.displayName
        UserPrincipalName = [string]$Principal.userPrincipalName
        PrincipalId = $PrincipalId
        PrincipalType = [string]$Principal.'@odata.type'
        StartUtc = [string]$Start
        EndUtc = [string]$End
    }
}

function Get-EligibleReportRows
{
    param($Group, [datetimeoffset]$Now)
    foreach ($Instance in @(Get-GroupReportInstances -GroupId $Group.id -Kind 'eligibilityScheduleInstances'))
    {
        if (Test-CurrentReportInstance -Instance $Instance -Now $Now)
        {
            New-AssignmentReportRow -Group $Group -Type 'Eligible' -Access $Instance.accessId `
                -PrincipalId $Instance.principalId -Principal $Instance.principal -Start $Instance.startDateTime -End $Instance.endDateTime
        }
    }
}

function Get-PermanentReportRows
{
    param($Group, [datetimeoffset]$Now)
    $ActivePrincipals = @{}
    foreach ($Instance in @(Get-GroupReportInstances -GroupId $Group.id -Kind 'assignmentScheduleInstances'))
    {
        if (-not (Test-CurrentReportInstance -Instance $Instance -Now $Now) -or $Instance.accessId -ne 'member') { continue }
        if ($Instance.assignmentType -notin @('assigned', 'activated'))
        {
            throw "Unknown PIM assignment type '$($Instance.assignmentType)' for group $($Group.id); cannot classify direct members safely."
        }
        if ($Instance.assignmentType -eq 'assigned' -and [string]::IsNullOrWhiteSpace([string]$Instance.endDateTime))
        {
            $ActivePrincipals[[string]$Instance.principalId] = 'Permanent'
            New-AssignmentReportRow -Group $Group -Type 'Permanent PIM active' -Access 'member' `
                -PrincipalId $Instance.principalId -Principal $Instance.principal -Start $Instance.startDateTime -End $null
        }
        elseif (-not $ActivePrincipals.ContainsKey([string]$Instance.principalId))
        {
            $ActivePrincipals[[string]$Instance.principalId] = 'Temporary'
        }
    }

    # Graph v1.0 omits service principals from this collection; Microsoft documents beta as a workaround.
    $Uri = "https://graph.microsoft.com/beta/groups/$($Group.id)/members"
    foreach ($Member in @(Get-GraphCollection -Uri $Uri))
    {
        if ([string]::IsNullOrWhiteSpace([string]$Member.id)) { throw "Graph returned a member without an ID for group $($Group.id)." }
        $Classification = $ActivePrincipals[[string]$Member.id]
        if ($Classification -ne 'Permanent')
        {
            $Finding = if ($Classification -eq 'Temporary') { 'Direct member / temporary PIM - verify' } else { 'Direct member - verify permanence' }
            New-AssignmentReportRow -Group $Group -Type $Finding -Access 'member' `
                -PrincipalId $Member.id -Principal $Member -Start $null -End $null
        }
    }
}

function Test-ReportAssignmentScope
{
    param($Context)
    return ($Context.Scopes -contains $Script:ReportScope -or
        $Context.Scopes -contains 'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup')
}

function Invoke-AssignmentReport
{
    param([ValidateSet('OneEligible', 'AllEligible', 'Permanent')][string]$Mode)
    $Script:ReportRows = @()
    $Script:ReportMode = $null
    $Script:ReportGrid.Rows.Clear()
    $Script:ReportGrid.Columns['PrincipalId'].Visible = $Mode -ne 'AllEligible'
    $Script:ReportGrid.Columns['PrincipalType'].Visible = $Mode -ne 'AllEligible'
    $Script:ReportStatus.Text = 'Scanning...'
    $Script:ReportExportButton.Enabled = $false
    try
    {
        Assert-GraphReady
        $Context = Get-GraphContext
        if ($Mode -eq 'Permanent' -and -not (Test-ReportAssignmentScope -Context $Context))
        {
            $Prompt = "The permanent-access scan needs $($Script:ReportScope). Reconnect to request consent now? An administrator may need to grant it."
            if ([System.Windows.Forms.MessageBox]::Show($Script:MainForm, $Prompt, 'Additional Graph permission',
                [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes)
            {
                throw "Permission was not requested. Reconnect and obtain consent for $($Script:ReportScope) before scanning."
            }
            $OriginalAccount = $Context.Account
            $OriginalTenant = $Context.TenantId
            Write-UiLog "Requesting Graph consent for the permanent-access report."
            Connect-MgGraph -Scopes ($Script:GraphScopes + $Script:ReportScope) -TenantId $OriginalTenant -ContextScope Process -ErrorAction Stop | Out-Null
            $Context = Get-GraphContext
            if ($null -eq $Context -or $Context.Account -ne $OriginalAccount -or $Context.TenantId -ne $OriginalTenant)
            {
                $Script:ConnectionStatusLabel.Text = 'Connection changed; check connection'
                $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
                throw 'The Graph account or tenant changed during sign-in. Check the connection and run the report again.'
            }
            [void](Test-GraphConnection)
            Assert-GraphReady
            if (-not (Test-ReportAssignmentScope -Context $Context))
            {
                throw "The signed-in Graph session still lacks $($Script:ReportScope). Ask an administrator to grant consent, then disconnect and reconnect."
            }
        }
        $Script:TabControl.Enabled = $false
        $Script:AuditBrowseButton.Enabled = $false
        $Script:MainForm.UseWaitCursor = $true
        [System.Windows.Forms.Application]::DoEvents()
        $Groups = if ($Mode -eq 'OneEligible') { @(Get-ExistingPimGroup -GroupNameOrId $Script:ReportGroupBox.Text) } else { @(Get-ReportGroups) }
        $Rows = [System.Collections.Generic.List[object]]::new()
        $Now = [datetimeoffset]::UtcNow
        $Index = 0
        foreach ($Group in $Groups)
        {
            $Index++
            $Script:ReportStatus.Text = "Scanning group $Index of $($Groups.Count): $($Group.displayName)"
            $Script:ReportStatus.Refresh()
            Write-UiLog "Assignment report ${Mode}: scanning $Index of $($Groups.Count): '$($Group.displayName)' [$($Group.id)]."
            $Findings = if ($Mode -eq 'Permanent') { @(Get-PermanentReportRows -Group $Group -Now $Now) } else { @(Get-EligibleReportRows -Group $Group -Now $Now) }
            foreach ($Finding in $Findings) { $Rows.Add($Finding) }
        }
        $Script:ReportRows = if ($Mode -eq 'AllEligible') {
            @($Rows.ToArray() | Select-Object -Property GroupName, GroupId, Finding, Access, PrincipalName, UserPrincipalName, StartUtc, EndUtc)
        } else { $Rows.ToArray() }
        foreach ($Row in $Script:ReportRows)
        {
            [void]$Script:ReportGrid.Rows.Add(@($Row.GroupName, $Row.GroupId, $Row.Finding, $Row.Access, $Row.PrincipalName, $Row.UserPrincipalName, $Row.PrincipalId, $Row.PrincipalType, $Row.StartUtc, $Row.EndUtc))
        }
        $Script:ReportStatus.Text = "Completed: $($Groups.Count) group(s), $($Rows.Count) finding(s)."
        $Script:ReportMode = $Mode
        $Script:ReportExportButton.Enabled = $Rows.Count -gt 0
        Write-UiLog "Assignment report ${Mode} completed: $($Groups.Count) group(s), $($Rows.Count) finding(s)."
    }
    catch
    {
        $Script:ReportRows = @()
        $Script:ReportMode = $null
        $Script:ReportGrid.Rows.Clear()
        $Script:ReportStatus.Text = 'Scan failed; no results displayed.'
        Write-UiLog "Assignment report ${Mode} failed: $($_.Exception.Message)"
        Show-Message -Message "The report could not be completed; no partial results are shown.`r`n`r`n$($_.Exception.Message)" -Title 'Assignment report failed' -Icon Error
    }
    finally
    {
        $Script:TabControl.Enabled = $true
        $Script:AuditBrowseButton.Enabled = $true
        $Script:MainForm.UseWaitCursor = $false
    }
}

function Write-AssignmentReportFile
{
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][object[]]$Rows)

    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($Extension -notin @('.csv', '.xlsx')) { throw 'Choose a .csv or .xlsx report file.' }
    if ($Extension -eq '.xlsx')
    {
        if (-not (Get-Module -ListAvailable -Name ImportExcel))
        {
            throw 'ImportExcel is required for XLSX export. Install it with: Install-Module ImportExcel -Scope CurrentUser'
        }
        Import-Module ImportExcel -ErrorAction Stop
    }
    $SafeRows = foreach ($Row in $Rows)
    {
        $Safe = [ordered]@{}
        foreach ($Property in $Row.PSObject.Properties)
        {
            $Value = [string]$Property.Value
            if ($Value -match '^[\s]*[=+\-@]' -or $Value -match '^[\t\r\n]') { $Value = "'" + $Value }
            $Safe[$Property.Name] = $Value
        }
        [PSCustomObject]$Safe
    }
    $TempPath = Join-Path ([System.IO.Path]::GetDirectoryName($Path)) ('.PIM-Assignment-Report-' + [guid]::NewGuid().ToString() + $Extension)
    try
    {
        if ($Extension -eq '.csv')
        {
            $SafeRows | Export-Csv -LiteralPath $TempPath -Encoding UTF8
        }
        else
        {
            $SafeRows | Export-Excel -Path $TempPath -WorksheetName 'Assignments' -TableName 'PIMAssignments' -TableStyle Light20 -AutoSize -FreezeTopRow -AutoFilter -NoNumberConversion '*' -NoHyperLinkConversion '*'
        }
        Move-Item -LiteralPath $TempPath -Destination $Path -Force -ErrorAction Stop
    }
    finally
    {
        if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -ErrorAction Stop }
    }
}

function Export-AssignmentReport
{
    if ($Script:ReportRows.Count -eq 0) { throw 'Run a report with findings before exporting.' }
    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    try
    {
        $Dialog.Title = 'Export assignment report'
        $ExcelAvailable = [bool](Get-Module -ListAvailable -Name ImportExcel)
        $Dialog.Filter = if ($ExcelAvailable) { 'Excel workbook (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv' } else { 'CSV files (*.csv)|*.csv|Excel workbook (*.xlsx)|*.xlsx' }
        $Dialog.DefaultExt = if ($ExcelAvailable) { 'xlsx' } else { 'csv' }
        $Dialog.AddExtension = $true
        $Dialog.OverwritePrompt = $true
        $ReportName = switch ($Script:ReportMode)
        {
            'OneEligible' { 'one-group-eligible' }
            'AllEligible' { 'all-eligible' }
            'Permanent' { 'all-permanent' }
            default { throw 'The report type is unavailable. Run the report again before exporting.' }
        }
        $Dialog.FileName = "$(Get-Date -Format 'yyyy-MM-dd_HHmm')-$ReportName-PIM-Assignment-Report"
        if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Write-AssignmentReportFile -Path $Dialog.FileName -Rows @($Script:ReportRows)
        Write-UiLog "Assignment report exported to '$($Dialog.FileName)' ($($Script:ReportRows.Count) findings)."
    }
    finally { $Dialog.Dispose() }
}

# ==================================================
# ^^^^^ Read-only Assignment Reports ^^^^^
# ==================================================

# ==================================================
# MARK: Existing Group Eligible Assignments
# ==================================================

function Get-MemberManagementGroupNameOrId
{
    if ([string]::IsNullOrWhiteSpace($Script:ReportGroupBox.Text))
    {
        throw 'Enter a group name or object ID in One group on the Assignments tab before adding or removing members.'
    }
    return $Script:ReportGroupBox.Text.Trim()
}

function Show-AddEligibleMembers
{
    $GroupNameOrId = Get-MemberManagementGroupNameOrId
    $Dialog = New-Object System.Windows.Forms.Form
    try
    {
        $Dialog.Text = 'Add eligible members'
        $Dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $Dialog.Size = New-Object System.Drawing.Size(670, 430)
        $Dialog.MinimumSize = $Dialog.Size
        $Dialog.Font = $Script:MainForm.Font
        $Dialog.BackColor = [System.Drawing.Color]::White
        [void](Add-UiLabel -Parent $Dialog -Text 'Group Name or ID' -X 18 -Y 18 -Width 140)
        $GroupLabel = Add-UiLabel -Parent $Dialog -Text $GroupNameOrId -X 165 -Y 18 -Width 460 -Height 28
        $GroupLabel.Font = New-Object System.Drawing.Font($Dialog.Font.FontFamily, ($Dialog.Font.Size + 1), [System.Drawing.FontStyle]::Bold)
        $GroupLabel.AutoEllipsis = $true
        [void](Add-UiLabel -Parent $Dialog -Text 'User UPNs' -X 18 -Y 57 -Width 140)
        $MembersBox = Add-UiTextBox -Parent $Dialog -X 165 -Y 57 -Width 460 -Height 70 -Multiline $true
        [void](Add-UiLabel -Parent $Dialog -Text 'Separate UPNs with commas, semicolons, or new lines.' -X 165 -Y 130 -Width 460)
        $Now = [datetime]::Now
        $StartNowBox = Add-UiCheckBox -Parent $Dialog -Text 'Start now' -X 18 -Y 169 -Checked $true -Width 130
        $StartPicker = New-Object System.Windows.Forms.DateTimePicker
        $StartPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
        $StartPicker.CustomFormat = 'yyyy-MM-dd HH:mm:ss'
        $StartPicker.Location = New-Object System.Drawing.Point(165, 169)
        $StartPicker.Size = New-Object System.Drawing.Size(220, 24)
        $StartPicker.Value = $Now
        $StartPicker.Enabled = $false
        $Dialog.Controls.Add($StartPicker)
        $StartNowBox.Add_CheckedChanged({ $StartPicker.Enabled = -not $StartNowBox.Checked }.GetNewClosure())
        [void](Add-UiLabel -Parent $Dialog -Text 'End (local time)' -X 18 -Y 207 -Width 140)
        $EndPicker = New-Object System.Windows.Forms.DateTimePicker
        $EndPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
        $EndPicker.CustomFormat = 'yyyy-MM-dd HH:mm:ss'
        $EndPicker.Location = New-Object System.Drawing.Point(165, 207)
        $EndPicker.Size = New-Object System.Drawing.Size(220, 24)
        $EndPicker.Value = $Now.ToUniversalTime().AddDays(365).ToLocalTime()
        $Dialog.Controls.Add($EndPicker)
        $PermanentBox = Add-UiCheckBox -Parent $Dialog -Text 'Permanent eligibility (if allowed by policy)' -X 165 -Y 243 -Checked $false -Width 400
        $PermanentBox.Add_CheckedChanged({ $EndPicker.Enabled = -not $PermanentBox.Checked }.GetNewClosure())
        $Hint = Add-UiLabel -Parent $Dialog -Text 'Dates and times use this computer''s local time. The group''s PIM policy may reject the requested schedule.' -X 18 -Y 278 -Width 600 -Height 38
        $Hint.ForeColor = [System.Drawing.Color]::DimGray
        $SubmitButton = New-Object System.Windows.Forms.Button
        $SubmitButton.Text = 'Review assignments'
        $SubmitButton.Location = New-Object System.Drawing.Point(315, 333)
        $SubmitButton.Size = New-Object System.Drawing.Size(150, 34)
        $SubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        Set-UiButtonStyle -Button $SubmitButton -Primary $true
        $Dialog.Controls.Add($SubmitButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Text = 'Cancel'
        $CancelButton.Location = New-Object System.Drawing.Point(475, 333)
        $CancelButton.Size = New-Object System.Drawing.Size(150, 34)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        Set-UiButtonStyle -Button $CancelButton
        $Dialog.Controls.Add($CancelButton)
        $Dialog.AcceptButton = $SubmitButton
        $Dialog.CancelButton = $CancelButton

        while ($true)
        {
            if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            try
            {
                $Upns = @(ConvertTo-EligibleMemberUPNs -Text $MembersBox.Text)
                if ($Upns.Count -eq 0) { throw 'Enter at least one user UPN.' }
                $ScheduleInfo = New-EligibleAssignmentSchedule -StartNow $StartNowBox.Checked -StartLocal $StartPicker.Value `
                    -Permanent $PermanentBox.Checked -EndLocal $EndPicker.Value
                break
            }
            catch { Show-Message -Message $_.Exception.Message -Title 'Invalid assignment details' -Icon Error }
        }

        $Script:TabControl.Enabled = $false
        $ConnectionGroup.Enabled = $false
        $Script:AuditBrowseButton.Enabled = $false
        $Script:MainForm.UseWaitCursor = $true
        [System.Windows.Forms.Application]::DoEvents()
        try
        {
            Assert-GraphReady
            $Group = Get-ExistingPimGroup -GroupNameOrId $GroupNameOrId -AllowNonRoleAssignable
            $Users = @()
            foreach ($Upn in $Upns) { $Users += Get-EntraUserByUpn -UserPrincipalName $Upn }
            $StartDescription = if ($StartNowBox.Checked) { 'At submission' } else { $StartPicker.Value.ToString('yyyy-MM-dd HH:mm:ss') }
            $EndDescription = if ($PermanentBox.Checked) { 'Permanent' } else { $EndPicker.Value.ToString('yyyy-MM-dd HH:mm:ss') }
            $Summary = "Add PIM-eligible members to:`r`n$($Group.displayName)`r`nGroup ID: $($Group.id)`r`n`r`nUsers: $($Upns -join ', ')`r`nStart: $StartDescription`r`nEnd: $EndDescription`r`nTimes are local.`r`n`r`nThis may onboard the group to PIM if it is not already managed; onboarding cannot be undone. This does not grant active membership. Continue?"
            if ([System.Windows.Forms.MessageBox]::Show($Script:MainForm, $Summary, 'Confirm eligible assignments',
                [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $Script:ReportGroupBox.Text = $Group.id
            $Script:ReportRows = @()
            $Script:ReportMode = $null
            $Script:ReportGrid.Rows.Clear()
            $Script:ReportExportButton.Enabled = $false
            $Script:ReportStatus.Text = 'Assignments may have changed; run a new report.'
            $Accepted = [System.Collections.Generic.List[string]]::new()
            $Configuration = @{ AllowPermanentEligibleAssignments = $PermanentBox.Checked }
            for ($Index = 0; $Index -lt $Upns.Count; $Index++)
            {
                try
                {
                    if ($StartNowBox.Checked)
                    {
                        $ScheduleInfo = New-EligibleAssignmentSchedule -StartNow $true -StartLocal $StartPicker.Value `
                            -Permanent $PermanentBox.Checked -EndLocal $EndPicker.Value
                    }
                    Add-PimEligibleMember -GroupId $Group.id -UserPrincipalName $Upns[$Index] -Configuration $Configuration `
                        -ScheduleInfo $ScheduleInfo -ResolvedUser $Users[$Index]
                    $Accepted.Add($Upns[$Index])
                }
                catch
                {
                    $Remaining = @($Upns | Select-Object -Skip ($Index + 1))
                    $Details = "Accepted requests: $(if ($Accepted.Count) { $Accepted -join ', ' } else { 'none' }).`r`nFailed or uncertain: $($Upns[$Index]).`r`nNot attempted: $(if ($Remaining.Count) { $Remaining -join ', ' } else { 'none' }).`r`n`r`n$($_.Exception.Message)`r`n`r`nA failed request may have been accepted before the error. Check Entra PIM before retrying."
                    Write-UiLog "Eligible assignment batch stopped for group $($Group.id): $Details"
                    Show-Message -Message $Details -Title 'Eligible assignments incomplete' -Icon Error
                    return
                }
            }
            Show-Message -Message "Submitted $($Accepted.Count) eligible-member request(s) for $($Group.displayName) [$($Group.id)].`r`n`r`nCheck Entra PIM for provisioning and assignment status."
        }
        finally
        {
            $Script:TabControl.Enabled = $true
            $ConnectionGroup.Enabled = $true
            $Script:AuditBrowseButton.Enabled = $true
            $Script:MainForm.UseWaitCursor = $false
        }
    }
    finally { $Dialog.Dispose() }
}

# ==================================================
# ^^^^^ Existing Group Eligible Assignments ^^^^^
# ==================================================

# ==================================================
# MARK: Existing Group Member Removal
# ==================================================

function Get-MemberRemovalPlan
{
    param(
        [Parameter(Mandatory = $true)]$Group,
        [Parameter(Mandatory = $true)][object[]]$Users,
        [bool]$Eligible,
        [bool]$Active,
        [bool]$Direct
    )

    $GroupId = [guid]::Empty
    if (-not [guid]::TryParse([string]$Group.id, [ref]$GroupId)) { throw 'The group ID is invalid.' }
    if (-not ($Eligible -or $Active -or $Direct)) { throw 'Select at least one removal category.' }
    $UserById = @{}
    foreach ($User in $Users)
    {
        $UserId = [guid]::Empty
        if (-not [guid]::TryParse([string]$User.id, [ref]$UserId) -or [string]::IsNullOrWhiteSpace([string]$User.userPrincipalName))
        {
            throw 'Graph returned a user without a valid ID and UPN.'
        }
        $UserById[$UserId.ToString()] = $User
    }

    $Plan = [System.Collections.Generic.List[object]]::new()
    $Filter = [uri]::EscapeDataString("groupId eq '$GroupId'")
    foreach ($Kind in @('assignmentSchedules', 'eligibilitySchedules'))
    {
        if (($Kind -eq 'assignmentSchedules' -and -not $Active) -or
            ($Kind -eq 'eligibilitySchedules' -and -not $Eligible)) { continue }
        $Uri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/${Kind}?`$filter=$Filter"
        $Seen = @{}
        foreach ($Schedule in @(Get-GraphCollection -Uri $Uri))
        {
            if ($Schedule.groupId -ne $GroupId.ToString()) { throw "Graph returned a schedule outside group $GroupId." }
            if ($Schedule.accessId -ne 'member' -or -not $UserById.ContainsKey([string]$Schedule.principalId)) { continue }
            if ($Schedule.memberType -eq 'group') { continue }
            if ($Schedule.memberType -ne 'direct' -or [string]::IsNullOrWhiteSpace([string]$Schedule.id))
            {
                throw "Cannot safely classify a membership schedule for $($Schedule.principalId)."
            }
            if ($Schedule.status -in @('Canceled', 'Denied', 'Failed', 'Revoked')) { continue }
            if ($Schedule.status -notin @('Provisioned', 'ScheduleCreated'))
            {
                throw "Membership schedule $($Schedule.id) has status '$($Schedule.status)'. Review it in Entra before removing access."
            }
            if ($Kind -eq 'assignmentSchedules' -and $Schedule.assignmentType -notin @('assigned', 'activated'))
            {
                throw "Cannot safely classify active schedule $($Schedule.id)."
            }
            $User = $UserById[[string]$Schedule.principalId]
            if ($Seen.ContainsKey([string]$User.id))
            {
                throw "Multiple $Kind exist for $($User.userPrincipalName). Review them in Entra before removing access."
            }
            $Seen[[string]$User.id] = $true
            $Category = if ($Kind -eq 'assignmentSchedules') { 'PIM active' } else { 'PIM eligible' }
            $Plan.Add([pscustomobject]@{
                Category = $Category; UserPrincipalName = $User.userPrincipalName; UserId = $User.id
                Detail = "Schedule: $($Schedule.id); status: $($Schedule.status); type: $($Schedule.assignmentType); schedule: $(ConvertTo-Json -InputObject $Schedule.scheduleInfo -Depth 6 -Compress)"
            })
        }
    }
    if ($Direct)
    {
        $Uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id"
        foreach ($Member in @(Get-GraphCollection -Uri $Uri))
        {
            if ([string]::IsNullOrWhiteSpace([string]$Member.id)) { throw "Graph returned a group member without an ID." }
            if ($UserById.ContainsKey([string]$Member.id))
            {
                $User = $UserById[[string]$Member.id]
                $Plan.Add([pscustomobject]@{
                    Category = 'Group membership'; UserPrincipalName = $User.userPrincipalName
                    UserId = $User.id; Detail = 'Direct members relationship (may be PIM-provisioned)'
                })
            }
        }
    }
    return $Plan.ToArray()
}

function Invoke-MemberRemoval
{
    param([Parameter(Mandatory = $true)][string]$GroupId, [Parameter(Mandatory = $true)]$Operation)
    $ParsedGroup = [guid]::Empty
    $ParsedUser = [guid]::Empty
    if (-not [guid]::TryParse($GroupId, [ref]$ParsedGroup) -or
        -not [guid]::TryParse([string]$Operation.UserId, [ref]$ParsedUser))
    {
        throw 'Invalid group or user ID; no removal request was sent.'
    }
    if ($Operation.Category -notin @('PIM active', 'PIM eligible', 'Group membership'))
    {
        throw 'Invalid removal category; no removal request was sent.'
    }
    Write-UiLog -Message "Removing $($Operation.Category) for '$($Operation.UserPrincipalName)' [$ParsedUser] from group $ParsedGroup. $($Operation.Detail)" -Required
    if ($Operation.Category -eq 'Group membership')
    {
        $Uri = "https://graph.microsoft.com/v1.0/groups/$ParsedGroup/members/$ParsedUser/`$ref"
        Invoke-MgGraphRequest -Method DELETE -Uri $Uri -ErrorAction Stop | Out-Null
        $Result = 'Group membership DELETE accepted'
    }
    else
    {
        $Endpoint = if ($Operation.Category -eq 'PIM active') { 'assignmentScheduleRequests' } else { 'eligibilityScheduleRequests' }
        $Body = @{
            action = 'adminRemove'; accessId = 'member'; groupId = $ParsedGroup.ToString(); principalId = $ParsedUser.ToString()
            justification = 'Membership removal requested by PIM Group Manager.'
        }
        $Response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/$Endpoint" `
            -Body ($Body | ConvertTo-Json -Depth 8) -ContentType 'application/json' -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace([string]$Response.id) -or [string]::IsNullOrWhiteSpace([string]$Response.status))
        {
            throw 'Graph returned no request ID or status; the removal outcome is uncertain.'
        }
        if ($Response.status -in @('Canceled', 'Denied', 'Failed', 'Revoked'))
        {
            throw "PIM request $($Response.id) returned status $($Response.status); review it in Entra."
        }
        $Result = "PIM request $($Response.id), status $($Response.status)"
    }
    Write-UiLog -Message "$Result for $($Operation.UserPrincipalName) [$ParsedUser] in group $ParsedGroup." -Required
    return $Result
}

function Invoke-MemberRemovalBatch
{
    param([Parameter(Mandatory = $true)][string]$GroupId, [Parameter(Mandatory = $true)][object[]]$Plan)
    $Accepted = [System.Collections.Generic.List[string]]::new()
    for ($Index = 0; $Index -lt $Plan.Count; $Index++)
    {
        $Operation = $Plan[$Index]
        try
        {
            $Result = Invoke-MemberRemoval -GroupId $GroupId -Operation $Operation
            $Accepted.Add("$($Operation.UserPrincipalName) - $($Operation.Category): $Result")
        }
        catch
        {
            return [pscustomobject]@{
                Accepted = $Accepted.ToArray()
                Failed = "$($Operation.UserPrincipalName) - $($Operation.Category)"
                Remaining = @($Plan | Select-Object -Skip ($Index + 1) | ForEach-Object { "$($_.UserPrincipalName) - $($_.Category)" })
                ErrorMessage = $_.Exception.Message
            }
        }
    }
    return [pscustomobject]@{ Accepted = $Accepted.ToArray(); Failed = $null; Remaining = @(); ErrorMessage = $null }
}

function Request-MemberRemovalScopes
{
    param([string[]]$RequiredScopes)
    $Context = Get-GraphContext
    $Missing = @($RequiredScopes | Where-Object { $Context.Scopes -notcontains $_ })
    if ($Missing.Count -eq 0) { return }
    $Prompt = "Removal needs additional Graph permissions: $($Missing -join ', '). Reconnect to request consent? An administrator may need to grant it."
    if ([System.Windows.Forms.MessageBox]::Show($Script:MainForm, $Prompt, 'Additional Graph permissions',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes)
    {
        throw 'Additional permissions were not requested; no memberships were changed.'
    }
    $OriginalAccount = $Context.Account
    $OriginalTenant = $Context.TenantId
    Write-UiLog "Requesting Graph consent for member removal: $($Missing -join ', ')."
    Connect-MgGraph -Scopes @($Script:GraphScopes + $Script:ReportScope + $RequiredScopes | Select-Object -Unique) `
        -TenantId $OriginalTenant -ContextScope Process -ErrorAction Stop | Out-Null
    $Context = Get-GraphContext
    if ($null -eq $Context -or $Context.Account -ne $OriginalAccount -or $Context.TenantId -ne $OriginalTenant)
    {
        $Script:ConnectionStatusLabel.Text = 'Connection changed; check connection'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        throw 'The Graph account or tenant changed during sign-in. Check the connection and start again.'
    }
    [void](Test-GraphConnection)
    Assert-GraphReady
    $Missing = @($RequiredScopes | Where-Object { $Context.Scopes -notcontains $_ })
    if ($Missing.Count) { throw "Missing Graph permissions: $($Missing -join ', '). No memberships were changed." }
}

function Show-RemoveMembers
{
    $GroupNameOrId = Get-MemberManagementGroupNameOrId
    $Dialog = New-Object System.Windows.Forms.Form
    try
    {
        $Dialog.Text = 'Remove members'
        $Dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $Dialog.Size = New-Object System.Drawing.Size(670, 435)
        $Dialog.MinimumSize = $Dialog.Size
        $Dialog.Font = $Script:MainForm.Font
        $Dialog.BackColor = [System.Drawing.Color]::White
        [void](Add-UiLabel -Parent $Dialog -Text 'Group Name or ID' -X 18 -Y 18 -Width 140)
        $GroupLabel = Add-UiLabel -Parent $Dialog -Text $GroupNameOrId -X 165 -Y 18 -Width 460 -Height 28
        $GroupLabel.Font = New-Object System.Drawing.Font($Dialog.Font.FontFamily, ($Dialog.Font.Size + 1), [System.Drawing.FontStyle]::Bold)
        $GroupLabel.AutoEllipsis = $true
        [void](Add-UiLabel -Parent $Dialog -Text 'User UPNs' -X 18 -Y 57 -Width 140)
        $MembersBox = Add-UiTextBox -Parent $Dialog -X 165 -Y 57 -Width 460 -Height 70 -Multiline $true
        [void](Add-UiLabel -Parent $Dialog -Text 'Separate UPNs with commas, semicolons, or new lines.' -X 165 -Y 130 -Width 460)
        [void](Add-UiLabel -Parent $Dialog -Text 'Select what to remove:' -X 18 -Y 166 -Width 140)
        $ActiveBox = Add-UiCheckBox -Parent $Dialog -Text 'PIM active membership schedules' -X 165 -Y 166 -Checked $false -Width 400
        $EligibleBox = Add-UiCheckBox -Parent $Dialog -Text 'PIM eligible membership schedules' -X 165 -Y 193 -Checked $false -Width 400
        $DirectBox = Add-UiCheckBox -Parent $Dialog -Text 'Direct group members relationship' -X 165 -Y 220 -Checked $false -Width 400
        $Hint = Add-UiLabel -Parent $Dialog -Text 'PIM active membership also appears as a direct group member. Removing either can affect active access; eligibility can still allow reactivation. Nested access and owners are not removed.' -X 18 -Y 255 -Width 610 -Height 55
        $Hint.ForeColor = [System.Drawing.Color]::DimGray
        $SubmitButton = New-Object System.Windows.Forms.Button
        $SubmitButton.Text = 'Review removals'
        $SubmitButton.Location = New-Object System.Drawing.Point(315, 333)
        $SubmitButton.Size = New-Object System.Drawing.Size(150, 34)
        $SubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        Set-UiButtonStyle -Button $SubmitButton -Primary $true
        $Dialog.Controls.Add($SubmitButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Text = 'Cancel'
        $CancelButton.Location = New-Object System.Drawing.Point(475, 333)
        $CancelButton.Size = New-Object System.Drawing.Size(150, 34)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        Set-UiButtonStyle -Button $CancelButton
        $Dialog.Controls.Add($CancelButton)
        $Dialog.AcceptButton = $SubmitButton
        $Dialog.CancelButton = $CancelButton
        while ($true)
        {
            if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            try
            {
                $Upns = @(ConvertTo-EligibleMemberUPNs -Text $MembersBox.Text)
                if ($Upns.Count -eq 0) { throw 'Enter at least one user UPN.' }
                if (-not ($ActiveBox.Checked -or $EligibleBox.Checked -or $DirectBox.Checked)) { throw 'Select at least one removal category.' }
                break
            }
            catch { Show-Message -Message $_.Exception.Message -Title 'Invalid removal details' -Icon Error }
        }
        $Script:TabControl.Enabled = $false
        $ConnectionGroup.Enabled = $false
        $Script:AuditBrowseButton.Enabled = $false
        $Script:MainForm.UseWaitCursor = $true
        [System.Windows.Forms.Application]::DoEvents()
        try
        {
            Assert-GraphReady
            $Group = Get-ExistingPimGroup -GroupNameOrId $GroupNameOrId -AllowNonRoleAssignable
            $Scopes = @()
            if ($ActiveBox.Checked) { $Scopes += $Script:AssignmentWriteScope }
            if ($DirectBox.Checked -and $Group.isAssignableToRole -eq $true) { $Scopes += $Script:RoleManagementWriteScope }
            if ($Scopes.Count) { Request-MemberRemovalScopes -RequiredScopes $Scopes }
            $Users = @()
            foreach ($Upn in $Upns)
            {
                $User = Get-EntraUserByUpn -UserPrincipalName $Upn
                if ($User.userPrincipalName -ne $Upn) { throw "Could not verify the resolved user for '$Upn'." }
                $Users += $User
            }
            $Plan = @(Get-MemberRemovalPlan -Group $Group -Users $Users -Eligible $EligibleBox.Checked -Active $ActiveBox.Checked -Direct $DirectBox.Checked)
            if ($Plan.Count -eq 0)
            {
                Show-Message -Message 'No matching direct memberships or PIM membership schedules were found for the selected categories. Indirect or nested access is not evaluated.' -Title 'No matching removals'
                return
            }
            $Lines = @($Plan | ForEach-Object { "$($_.UserPrincipalName) [$($_.UserId)] - $($_.Category): $($_.Detail)" })
            $Categories = @()
            if ($ActiveBox.Checked) { $Categories += 'PIM active' }
            if ($EligibleBox.Checked) { $Categories += 'PIM eligible' }
            if ($DirectBox.Checked) { $Categories += 'Group membership' }
            $Summary = "Group: $($Group.displayName) [$($Group.id)]`r`nRequested users: $($Upns -join ', ')`r`nSelected categories: $($Categories -join ', ')`r`n`r`nMatching operations:`r`n$($Lines -join "`r`n")`r`n`r`nOnly the matching operations listed above will be submitted. PIM schedules may include future assignments. Removing an active PIM assignment or direct group membership can affect the same access. Removing eligibility may not end an existing activation; users may retain access through other paths. PIM requests can take time to provision. Continue?"
            $Review = New-Object System.Windows.Forms.Form
            try
            {
                $Review.Text = 'Confirm member removals'
                $Review.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
                $Review.Size = New-Object System.Drawing.Size(760, 500)
                $Review.Font = $Script:MainForm.Font
                $Details = Add-UiTextBox -Parent $Review -X 15 -Y 15 -Width 710 -Height 390 -Multiline $true
                $Details.ReadOnly = $true
                $Details.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
                $Details.Anchor = 'Top, Bottom, Left, Right'
                $Details.Text = $Summary
                $Confirm = New-Object System.Windows.Forms.Button
                $Confirm.Text = 'Remove selected access'
                $Confirm.Location = New-Object System.Drawing.Point(410, 415)
                $Confirm.Size = New-Object System.Drawing.Size(190, 34)
                $Confirm.Anchor = 'Bottom, Right'
                $Confirm.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                Set-UiButtonStyle -Button $Confirm -Primary $true
                $Review.Controls.Add($Confirm)
                $Cancel = New-Object System.Windows.Forms.Button
                $Cancel.Text = 'Cancel'
                $Cancel.Location = New-Object System.Drawing.Point(610, 415)
                $Cancel.Size = New-Object System.Drawing.Size(115, 34)
                $Cancel.Anchor = 'Bottom, Right'
                $Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                Set-UiButtonStyle -Button $Cancel
                $Review.Controls.Add($Cancel)
                $Review.CancelButton = $Cancel
                if ($Review.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
            finally { $Review.Dispose() }
            $Script:ReportGroupBox.Text = $Group.id
            $Script:ReportRows = @()
            $Script:ReportMode = $null
            $Script:ReportGrid.Rows.Clear()
            $Script:ReportExportButton.Enabled = $false
            $Script:ReportStatus.Text = 'Assignments may have changed; run a new report.'
            $Outcome = Invoke-MemberRemovalBatch -GroupId $Group.id -Plan $Plan
            if ($Outcome.Failed)
            {
                $Details = "Accepted: $(if ($Outcome.Accepted.Count) { $Outcome.Accepted -join '; ' } else { 'none' }).`r`nFailed or uncertain: $($Outcome.Failed).`r`nNot attempted: $(if ($Outcome.Remaining.Count) { $Outcome.Remaining -join '; ' } else { 'none' }).`r`n`r`n$($Outcome.ErrorMessage)`r`n`r`nA request may have succeeded before the error. Check Entra PIM and group membership before retrying."
                Write-UiLog "Member removal batch stopped for group $($Group.id): $Details"
                Show-Message -Message $Details -Title 'Member removals incomplete' -Icon Error
                return
            }
            Show-Message -Message "Removal requests submitted:`r`n$($Outcome.Accepted -join "`r`n")`r`n`r`nVerify PIM request statuses and group membership in Entra; accepted requests do not guarantee completed removal."
        }
        finally
        {
            $Script:TabControl.Enabled = $true
            $ConnectionGroup.Enabled = $true
            $Script:AuditBrowseButton.Enabled = $true
            $Script:MainForm.UseWaitCursor = $false
        }
    }
    finally { $Dialog.Dispose() }
}

# ==================================================
# ^^^^^ Existing Group Member Removal ^^^^^
# ==================================================

# ==================================================
# MARK: Single Group Workflow and Form Configuration
# ==================================================
 
function Invoke-CreateGroupWorkflow
{
    param([Parameter(Mandatory = $true)][System.Object]$Configuration)
 
    Assert-GraphReady
    $Group = New-EntraGroup -Configuration $Configuration
    $Result = [PSCustomObject]@{
        GroupName = $Configuration.GroupName
        GroupId   = $Group.id
        Status    = 'Created'
        Details   = "Group created. MailNickname: $($Configuration.MailNickname)"
    }
 
    try
    {
        Write-UiLog -Message "Group created: '$($Configuration.GroupName)' [$($Group.id)], type $($Configuration.GroupType)." -Required
    }
    catch
    {
        $Result.Status = 'Group created; audit logging failed'
        $Result.Details += " $($_.Exception.Message) Review this group before retrying."
        return $Result
    }
 
    if ($Configuration.GroupType -eq 'PIMRoleAssignable')
    {
        try
        {
            $PropagationDelaySeconds = 10
            Write-UiLog "Waiting $PropagationDelaySeconds seconds for Microsoft Entra propagation before configuring PIM for '$($Configuration.GroupName)' [$($Group.id)]."
            Start-Sleep -Seconds $PropagationDelaySeconds
            Set-PimMemberSettings -GroupId $Group.id -Configuration $Configuration
            $Result.Details += ' PIM Member settings configured and verified.'
 
            foreach ($Upn in $Configuration.EligibleMemberUPNs)
            {
                Add-PimEligibleMember -GroupId $Group.id -UserPrincipalName $Upn -Configuration $Configuration
                $Result.Details += " Eligible member added: $Upn."
            }
            Write-UiLog -Message "PIM setup completed and verified for '$($Configuration.GroupName)' [$($Group.id)]." -Required
        }
        catch
        {
            $Result.Status = 'Group created; PIM setup failed'
            $Result.Details += " PIM setup error: $($_.Exception.Message) Settings or eligible assignments may be partially applied; inspect before retrying."
            Write-UiLog "ERROR for '$($Configuration.GroupName)': $($_.Exception.Message)"
        }
    }
    elseif ($Configuration.IsDynamic)
    {
        $Result.Details += " Dynamic $($Configuration.MembershipType.ToLowerInvariant()) membership rule submitted to Entra."
    }
 
    return $Result
}
 
function Get-GroupTypeName
{
    param([Parameter(Mandatory = $true)][System.Int32]$Index)
 
    switch ($Index)
    {
        0 { return 'PIMRoleAssignable' }
        1 { return 'SecurityAssigned' }
        2 { return 'SecurityDynamic' }
        3 { return 'Microsoft365Assigned' }
        4 { return 'Microsoft365Dynamic' }
        default { throw 'Select a supported group type.' }
    }
}
 
function Update-GroupTypeVisibility
{
    if ($null -eq $Script:GroupTypeBox -or $Script:GroupTypeBox.SelectedIndex -lt 0)
    {
        return
    }
 
    $GroupType = Get-GroupTypeName -Index $Script:GroupTypeBox.SelectedIndex
    $IsDynamic = $GroupType -in @('SecurityDynamic', 'Microsoft365Dynamic')
    $IsPim = $GroupType -eq 'PIMRoleAssignable'
    $Script:MembershipTypeLabel.Visible = $IsDynamic
    $Script:MembershipTypeBox.Visible = $IsDynamic
    $Script:MembershipRuleLabel.Visible = $IsDynamic
    $Script:MembershipRuleBox.Visible = $IsDynamic
    $Script:PimGroupBox.Visible = $IsPim
    $Script:EligibleMembersLabel.Visible = $IsPim
    $Script:EligibleMembersBox.Visible = $IsPim
    $Script:ExistingGroupIdLabel.Visible = $IsPim
    $Script:ExistingGroupIdBox.Visible = $IsPim
    $Script:ApplyExistingButton.Visible = $IsPim
    $Script:ExistingGroupHint.Visible = $IsPim
}
 
function Handle-GroupTypeChanged
{
    $NewIndex = $Script:GroupTypeBox.SelectedIndex
    if ($NewIndex -lt 0)
    {
        return
    }
 
    $PreviousType = Get-GroupTypeName -Index $Script:PreviousGroupTypeIndex
    $NewType = Get-GroupTypeName -Index $NewIndex
    $PimPrefix = '*DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal'
    if ($PreviousType -eq 'PIMRoleAssignable' -and $NewType -ne 'PIMRoleAssignable')
    {
        if ($Script:DescriptionBox.Text -eq $PimPrefix)
        {
            $Script:DescriptionBox.Clear()
        }
        elseif ($Script:DescriptionBox.Text.StartsWith("$PimPrefix | ", [System.StringComparison]::OrdinalIgnoreCase))
        {
            $Script:DescriptionBox.Text = $Script:DescriptionBox.Text.Substring($PimPrefix.Length + 3)
        }
    }
    elseif ($PreviousType -ne 'PIMRoleAssignable' -and $NewType -eq 'PIMRoleAssignable')
    {
        $Script:DescriptionBox.Text = Get-DescriptionForGroupType -GroupType $NewType -Description $Script:DescriptionBox.Text
    }
 
    $Script:PreviousGroupTypeIndex = $NewIndex
    Update-GroupTypeVisibility
}
 
function Get-SingleGroupConfiguration
{
    param([switch]$PolicyOnly)
 
    $GroupType = Get-GroupTypeName -Index $Script:GroupTypeBox.SelectedIndex
    $Row = [PSCustomObject]@{
        GroupName = $Script:GroupNameBox.Text
        Description = $Script:DescriptionBox.Text
        MailNickname = $Script:MailNicknameBox.Text
        GroupType = $GroupType
        MembershipType = [System.String]$Script:MembershipTypeBox.SelectedItem
        MembershipRule = $Script:MembershipRuleBox.Text
        ActivationMaxHours = [System.String]$Script:ActivationHoursBox.Value
        RequireMFAOnActivation = $Script:RequireMfaActivationBox.Checked.ToString()
        RequireMFAOnActiveAssignment = $Script:RequireMfaActiveBox.Checked.ToString()
        RequireJustificationOnActiveAssignment = $Script:RequireJustificationActiveBox.Checked.ToString()
        AllowPermanentEligibleAssignments = $Script:AllowPermanentEligibleBox.Checked.ToString()
        EligibleAssignmentDays = [System.String]$Script:EligibilityDaysBox.Value
        AllowPermanentActiveAssignments = $Script:AllowPermanentActiveBox.Checked.ToString()
        ActiveAssignmentDays = $Script:ActiveAssignmentDaysBox.Text
        EligibleMemberUPNs = $Script:EligibleMembersBox.Text
    }
 
    if ($PolicyOnly)
    {
        $Row.GroupName = 'existing-group'
        $Row.MailNickname = 'existing-group'
        $Row.EligibleMemberUPNs = ''
    }
    return ConvertTo-GroupConfiguration -Row $Row
}
 
# ==================================================
# ^^^^^ Single Group Workflow and Form Configuration ^^^^^
# ==================================================
 
# ==================================================
# MARK: Bulk Templates, Import, and Creation
# ==================================================
 
function Import-BulkFile
{
    param([Parameter(Mandatory = $true)][System.String]$Path)
 
    $Extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($Extension -eq '.csv')
    {
        $Rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    }
    elseif ($Extension -eq '.xlsx')
    {
        if (-not (Get-Module -ListAvailable -Name ImportExcel))
        {
            throw 'ImportExcel is required for XLSX files. Install it with: Install-Module ImportExcel -Scope CurrentUser'
        }
 
        Import-Module ImportExcel -ErrorAction Stop
        $Rows = @(Import-Excel -Path $Path -ErrorAction Stop)
    }
    else
    {
        throw 'Choose a .csv or .xlsx file.'
    }
 
    if ($Rows.Count -eq 0)
    {
        throw 'The selected file has no data rows.'
    }
 
    $NormalizedRows = [System.Collections.Generic.List[System.Object]]::new()
    $Errors = [System.Collections.Generic.List[System.String]]::new()
    $Index = 0
    foreach ($Row in $Rows)
    {
        $Index++
        $GroupName = Get-RowValue -Row $Row -Name 'GroupName'
        if ([System.String]::IsNullOrWhiteSpace([System.String]$GroupName))
        {
            $HasContent = @($Script:TemplateColumns | Where-Object {
                -not [System.String]::IsNullOrWhiteSpace([System.String](Get-RowValue -Row $Row -Name $_))
            }).Count -gt 0
            if (-not $HasContent)
            {
                continue
            }
        }
 
        try
        {
            $NormalizedRows.Add((ConvertTo-GroupConfiguration -Row $Row))
        }
        catch
        {
            $Errors.Add("Row $Index : $($_.Exception.Message)")
        }
    }
 
    if ($Errors.Count -gt 0)
    {
        throw "Bulk file validation failed. No groups were created.`r`n`r`n$($Errors -join "`r`n")"
    }
 
    if ($NormalizedRows.Count -eq 0)
    {
        throw 'No complete group rows were found in the selected file.'
    }
 
    return $NormalizedRows.ToArray()
}
 
function Set-BulkGrid
{
    param([Parameter(Mandatory = $true)][System.Object[]]$Configurations)
 
    $Table = New-Object System.Data.DataTable
    foreach ($ColumnName in @(
        'GroupName'
        'GroupType'
        'MembershipType'
        'ActivationMaxHours'
        'RequireMFAOnActivation'
        'EligibleAssignmentDays'
        'AllowPermanentEligibleAssignments'
        'RequireMFAOnActiveAssignment'
        'RequireJustificationOnActiveAssignment'
        'AllowPermanentActiveAssignments'
        'EligibleMemberUPNs'
    ))
    {
        [void]$Table.Columns.Add($ColumnName)
    }
 
    foreach ($Configuration in $Configurations)
    {
        $Row = $Table.NewRow()
        $Row['GroupName'] = $Configuration.GroupName
        $Row['GroupType'] = $Configuration.GroupType
        $Row['MembershipType'] = if ($Configuration.IsDynamic) { $Configuration.MembershipType } else { 'Assigned' }
        $Row['ActivationMaxHours'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.ActivationMaxHours } else { '' }
        $Row['RequireMFAOnActivation'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.RequireMFAOnActivation } else { '' }
        $Row['EligibleAssignmentDays'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable' -and -not $Configuration.AllowPermanentEligibleAssignments) { $Configuration.EligibleAssignmentDays } else { '' }
        $Row['AllowPermanentEligibleAssignments'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.AllowPermanentEligibleAssignments } else { '' }
        $Row['RequireMFAOnActiveAssignment'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.RequireMFAOnActiveAssignment } else { '' }
        $Row['RequireJustificationOnActiveAssignment'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.RequireJustificationOnActiveAssignment } else { '' }
        $Row['AllowPermanentActiveAssignments'] = if ($Configuration.GroupType -eq 'PIMRoleAssignable') { $Configuration.AllowPermanentActiveAssignments } else { '' }
        $Row['EligibleMemberUPNs'] = $Configuration.EligibleMemberUPNs -join '; '
        $Table.Rows.Add($Row)
    }
 
    $Script:BulkGrid.DataSource = $Table
    $Script:BulkCreateButton.Enabled = $true
}
 
function New-TemplateFile
{
    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.Title = 'Create bulk group template'
    $Dialog.Filter = 'CSV template (*.csv)|*.csv|Excel workbook (*.xlsx)|*.xlsx'
    $Dialog.FileName = 'PIM-Group-Creation-Template'
    $Dialog.DefaultExt = 'csv'
    $Dialog.AddExtension = $true
    $Dialog.OverwritePrompt = $true
 
    if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK)
    {
        return
    }
 
    $Extension = [System.IO.Path]::GetExtension($Dialog.FileName).ToLowerInvariant()
    if ($Extension -eq '.csv')
    {
        $Header = [ordered]@{}
        foreach ($ColumnName in $Script:TemplateColumns)
        {
            $Header[$ColumnName] = ''
        }
        [PSCustomObject]$Header | Export-Csv -LiteralPath $Dialog.FileName -Encoding UTF8
    }
    elseif ($Extension -eq '.xlsx')
    {
        if (-not (Get-Module -ListAvailable -Name ImportExcel))
        {
            Show-Message -Message 'The ImportExcel module is required for XLSX templates. Install it with: Install-Module ImportExcel -Scope CurrentUser' -Icon Warning
            return
        }
 
        Import-Module ImportExcel -ErrorAction Stop
        $Template = [ordered]@{}
        foreach ($ColumnName in $Script:TemplateColumns)
        {
            $Template[$ColumnName] = ''
        }
        $Template.GroupName = '__REPLACE_ME__'
        $Template.GroupType = 'PIMRoleAssignable'
        $Template.MembershipType = 'User'
        $Template.ActivationMaxHours = 8
        $Template.RequireMFAOnActivation = $true
        $Template.RequireMFAOnActiveAssignment = $true
        $Template.RequireJustificationOnActiveAssignment = $true
        $Template.AllowPermanentEligibleAssignments = $false
        $Template.EligibleAssignmentDays = 365
        $Template.AllowPermanentActiveAssignments = $false
        [PSCustomObject]$Template | Export-Excel -Path $Dialog.FileName -WorksheetName 'Groups' -AutoSize -TableName 'PIMGroupTemplate' -TableStyle Light20
    }
    else
    {
        throw 'Choose a .csv or .xlsx template file.'
    }
 
    Write-UiLog "Created bulk template: $($Dialog.FileName)"
    Show-Message -Message "Template created:`r`n$($Dialog.FileName)`r`n`r`nFor each row, use GroupType values: $($Script:GroupTypeOptions -join ', '). The XLSX placeholder row must be replaced or removed."
}
 
function Select-BulkFile
{
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = 'Select a bulk group file'
    $Dialog.Filter = 'CSV or Excel (*.csv;*.xlsx)|*.csv;*.xlsx|CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx'
    $Dialog.Multiselect = $false
 
    if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK)
    {
        return
    }
 
    try
    {
        $Configurations = @(Import-BulkFile -Path $Dialog.FileName)
        $Script:BulkPath = $Dialog.FileName
        $Script:BulkRows = $Configurations
        $Script:BulkPathBox.Text = $Dialog.FileName
        Set-BulkGrid -Configurations $Configurations
        Write-UiLog "Validated $($Configurations.Count) bulk row(s) from $($Dialog.FileName)."
    }
    catch
    {
        $Script:BulkRows = @()
        $Script:BulkPath = ''
        $Script:BulkPathBox.Clear()
        $Script:BulkGrid.DataSource = $null
        $Script:BulkCreateButton.Enabled = $false
        Show-Message -Message $_.Exception.Message -Title 'Bulk file validation' -Icon Error
        Write-UiLog "Bulk import failed: $($_.Exception.Message)"
    }
}
 
function Export-BulkResults
{
    param([Parameter(Mandatory = $true)][System.Object[]]$Results)
 
    $OutputDirectory = if (-not [System.String]::IsNullOrWhiteSpace($Script:BulkPath)) {
        Split-Path -Parent $Script:BulkPath
    }
    else {
        Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PIM-Group-Manager'
    }
    if (-not (Test-Path -LiteralPath $OutputDirectory))
    {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }
 
    $OutputPath = Join-Path $OutputDirectory ("PIM-Group-Creation-Results-{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    $Results | Export-Csv -LiteralPath $OutputPath -Encoding UTF8
    return $OutputPath
}
 
function Invoke-BulkCreate
{
    if ($Script:BulkRows.Count -eq 0)
    {
        Show-Message -Message 'Import and validate a CSV or XLSX file first.' -Icon Warning
        return
    }
 
    try
    {
        Assert-GraphReady
    }
    catch
    {
        Show-Message -Message $_.Exception.Message -Icon Warning
        return
    }
 
    $PimCount = @($Script:BulkRows | Where-Object GroupType -eq 'PIMRoleAssignable').Count
    $Prompt = "Create $($Script:BulkRows.Count) group(s), including $PimCount role-assignable PIM group(s)?`r`n`r`nFor PIM groups, this applies Member policy settings and optional eligible-member assignments."
    if ([System.Windows.Forms.MessageBox]::Show(
        $Script:MainForm,
        $Prompt,
        'Confirm bulk group creation',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) -ne [System.Windows.Forms.DialogResult]::Yes)
    {
        return
    }
 
    $Results = [System.Collections.Generic.List[System.Object]]::new()
    $Script:BulkCreateButton.Enabled = $false
    $Script:TabControl.Enabled = $false
    $Script:AuditBrowseButton.Enabled = $false
    $Script:MainForm.UseWaitCursor = $true
    try
    {
        $Index = 0
        foreach ($Configuration in $Script:BulkRows)
        {
            $Index++
            [System.Windows.Forms.Application]::DoEvents()
            Write-UiLog "Creating $Index of $($Script:BulkRows.Count): $($Configuration.GroupName)"
            try
            {
                $Result = Invoke-CreateGroupWorkflow -Configuration $Configuration
                Write-UiLog "$($Result.Status): $($Result.GroupName) [$($Result.GroupId)] $($Result.Details)"
                if ($Script:AuditFailure -and $Result.Status -eq 'Created')
                {
                    $Result.Status = 'Group created; audit logging failed'
                    $Result.Details += " $($Script:AuditFailure)"
                }
                $Results.Add($Result)
            }
            catch
            {
                $Results.Add([PSCustomObject]@{
                    GroupName = $Configuration.GroupName
                    GroupId   = ''
                    Status    = 'Failed'
                    Details   = $_.Exception.Message
                })
                Write-UiLog "ERROR creating '$($Configuration.GroupName)': $($_.Exception.Message)"
            }
        }
 
        $ResultsPath = Export-BulkResults -Results $Results.ToArray()
        Write-UiLog "Bulk processing finished. Results: $ResultsPath"
        $FailedCount = @($Results | Where-Object Status -ne 'Created').Count
        $AuditWarning = if ($Script:AuditFailure) { "`r`n`r`nWARNING: $($Script:AuditFailure)" } else { '' }
        Show-Message -Message "Bulk processing finished.`r`nCreated: $($Results.Count - $FailedCount)`r`nFailed or partially configured: $FailedCount`r`n`r`nResults: $ResultsPath$AuditWarning"
    }
    finally
    {
        $Script:TabControl.Enabled = $true
        $Script:AuditBrowseButton.Enabled = $true
        $Script:BulkCreateButton.Enabled = ($Script:BulkRows.Count -gt 0)
        $Script:MainForm.UseWaitCursor = $false
    }
}
 
# ==================================================
# ^^^^^ Bulk Templates, Import, and Creation ^^^^^
# ==================================================
 
# ==================================================
# MARK: GUI Control Helpers
# ==================================================
 
function Add-UiLabel
{
    param
    (
        [System.Windows.Forms.Control]$Parent,
        [System.String]$Text,
        [System.Int32]$X,
        [System.Int32]$Y,
        [System.Int32]$Width = 190,
        [System.Int32]$Height = 24
    )
 
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.Location = New-Object System.Drawing.Point($X, $Y)
    $Label.Size = New-Object System.Drawing.Size($Width, $Height)
    $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Parent.Controls.Add($Label)
    return $Label
}
 
function Add-UiTextBox
{
    param
    (
        [System.Windows.Forms.Control]$Parent,
        [System.Int32]$X,
        [System.Int32]$Y,
        [System.Int32]$Width = 540,
        [System.Int32]$Height = 24,
        [System.Boolean]$Multiline = $false
    )
 
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point($X, $Y)
    $TextBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $TextBox.Multiline = $Multiline
    if ($Multiline)
    {
        $TextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $TextBox.AcceptsReturn = $true
    }
    $Parent.Controls.Add($TextBox)
    return $TextBox
}
 
function Add-UiCheckBox
{
    param
    (
        [System.Windows.Forms.Control]$Parent,
        [System.String]$Text,
        [System.Int32]$X,
        [System.Int32]$Y,
        [System.Boolean]$Checked = $false,
        [System.Int32]$Width = 400
    )
 
    $CheckBox = New-Object System.Windows.Forms.CheckBox
    $CheckBox.Text = $Text
    $CheckBox.Location = New-Object System.Drawing.Point($X, $Y)
    $CheckBox.Size = New-Object System.Drawing.Size($Width, 24)
    $CheckBox.Checked = $Checked
    $Parent.Controls.Add($CheckBox)
    return $CheckBox
}
 
function Add-UiNumeric
{
    param
    (
        [System.Windows.Forms.Control]$Parent,
        [System.Int32]$X,
        [System.Int32]$Y,
        [System.Int32]$Minimum,
        [System.Int32]$Maximum,
        [System.Int32]$Value
    )
 
    $Numeric = New-Object System.Windows.Forms.NumericUpDown
    $Numeric.Location = New-Object System.Drawing.Point($X, $Y)
    $Numeric.Size = New-Object System.Drawing.Size(90, 24)
    $Numeric.Minimum = $Minimum
    $Numeric.Maximum = $Maximum
    $Numeric.Value = $Value
    $Parent.Controls.Add($Numeric)
    return $Numeric
}
 
function Set-UiButtonStyle
{
    param
    (
        [System.Windows.Forms.Button]$Button,
        [System.Boolean]$Primary = $false
    )
 
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.FlatAppearance.BorderSize = 1
    if ($Primary)
    {
        $Button.BackColor = [System.Drawing.Color]::FromArgb(36, 92, 176)
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatAppearance.BorderColor = $Button.BackColor
    }
    else
    {
        $Button.BackColor = [System.Drawing.Color]::White
        $Button.ForeColor = [System.Drawing.Color]::FromArgb(37, 52, 74)
        $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(194, 205, 220)
    }
}
 
# ==================================================
# ^^^^^ GUI Control Helpers ^^^^^
# ==================================================
 
# ==================================================
# MARK: Main Window and Connection Controls
# ==================================================
 
$Script:MainForm = New-Object System.Windows.Forms.Form
$Script:MainForm.Text = 'Microsoft Entra PIM Group Manager'
$Script:MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Script:MainForm.TopMost = $true
$Script:MainForm.Size = New-Object System.Drawing.Size(1040, 970)
$Script:MainForm.MinimumSize = New-Object System.Drawing.Size(840, 680)
$Script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$Script:MainForm.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 250)
 
$HeaderPanel = New-Object System.Windows.Forms.Panel
$HeaderPanel.Location = New-Object System.Drawing.Point(0, 0)
$HeaderPanel.Size = New-Object System.Drawing.Size(1024, 60)
$HeaderPanel.Anchor = 'Top, Left, Right'
$HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(23, 38, 62)
$Script:MainForm.Controls.Add($HeaderPanel)
$HeaderTitle = Add-UiLabel -Parent $HeaderPanel -Text 'PIM Group Manager' -X 18 -Y 7 -Width 500 -Height 27
$HeaderTitle.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$HeaderTitle.ForeColor = [System.Drawing.Color]::White
$HeaderSubtitle = Add-UiLabel -Parent $HeaderPanel -Text 'Microsoft Entra groups and privileged access' -X 20 -Y 34 -Width 500 -Height 18
$HeaderSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(194, 208, 228)
$HelpButton = New-Object System.Windows.Forms.Button
$HelpButton.Text = 'How to use'
$HelpButton.Location = New-Object System.Drawing.Point(860, 15)
$HelpButton.Size = New-Object System.Drawing.Size(145, 30)
$HelpButton.Anchor = 'Top, Right'
$HeaderPanel.Controls.Add($HelpButton)
 
$ConnectionGroup = New-Object System.Windows.Forms.GroupBox
$ConnectionGroup.Text = 'Microsoft Graph connection'
$ConnectionGroup.Location = New-Object System.Drawing.Point(12, 68)
$ConnectionGroup.Size = New-Object System.Drawing.Size(1000, 66)
$ConnectionGroup.Anchor = 'Top, Left, Right'
$ConnectionGroup.BackColor = [System.Drawing.Color]::White
$Script:MainForm.Controls.Add($ConnectionGroup)
 
$Script:ConnectionStatusLabel = Add-UiLabel -Parent $ConnectionGroup -Text 'Not connected' -X 12 -Y 24 -Width 625 -Height 28
$Script:ConnectionStatusLabel.Anchor = 'Top, Left, Right'
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.Text = 'Connect / Login'
$ConnectButton.Location = New-Object System.Drawing.Point(650, 22)
$ConnectButton.Size = New-Object System.Drawing.Size(112, 30)
$ConnectButton.Anchor = 'Top, Right'
$ConnectionGroup.Controls.Add($ConnectButton)
$CheckConnectionButton = New-Object System.Windows.Forms.Button
$CheckConnectionButton.Text = 'Check connection'
$CheckConnectionButton.Location = New-Object System.Drawing.Point(770, 22)
$CheckConnectionButton.Size = New-Object System.Drawing.Size(120, 30)
$CheckConnectionButton.Anchor = 'Top, Right'
$ConnectionGroup.Controls.Add($CheckConnectionButton)
$DisconnectButton = New-Object System.Windows.Forms.Button
$DisconnectButton.Text = 'Disconnect'
$DisconnectButton.Location = New-Object System.Drawing.Point(898, 22)
$DisconnectButton.Size = New-Object System.Drawing.Size(90, 30)
$DisconnectButton.Anchor = 'Top, Right'
$ConnectionGroup.Controls.Add($DisconnectButton)
 
$Script:TabControl = New-Object System.Windows.Forms.TabControl
$Script:TabControl.Location = New-Object System.Drawing.Point(12, 142)
$Script:TabControl.Size = New-Object System.Drawing.Size(1000, 630)
$Script:TabControl.Anchor = 'Top, Bottom, Left, Right'
$Script:TabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$Script:TabControl.ItemSize = New-Object System.Drawing.Size(180, 40)
$Script:TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$Script:TabControl.Add_DrawItem({
    param($Sender, $EventArgs)
    $Selected = $EventArgs.Index -eq $Sender.SelectedIndex
    $Background = if ($Selected) { [System.Drawing.Color]::FromArgb(36, 92, 176) } else { [System.Drawing.Color]::FromArgb(231, 237, 246) }
    $Foreground = if ($Selected) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::FromArgb(37, 52, 74) }
    $Brush = New-Object System.Drawing.SolidBrush($Background)
    $Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    try
    {
        $EventArgs.Graphics.FillRectangle($Brush, $EventArgs.Bounds)
        [System.Windows.Forms.TextRenderer]::DrawText($EventArgs.Graphics, $Sender.TabPages[$EventArgs.Index].Text,
            $Font, $EventArgs.Bounds, $Foreground,
            ([System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter))
        if ($Selected -and $Sender.ContainsFocus)
        {
            $FocusBounds = $EventArgs.Bounds
            $FocusBounds.Inflate(-3, -3)
            [System.Windows.Forms.ControlPaint]::DrawFocusRectangle($EventArgs.Graphics, $FocusBounds, $Foreground, $Background)
        }
    }
    finally
    {
        $Font.Dispose()
        $Brush.Dispose()
    }
})
$Script:TabControl.Add_SelectedIndexChanged({ $Script:TabControl.Invalidate() })
$Script:MainForm.Controls.Add($Script:TabControl)
 
# ==================================================
# ^^^^^ Main Window and Connection Controls ^^^^^
# ==================================================
 
# ==================================================
# MARK: Create One Group Tab
# ==================================================
 
$CreateTab = New-Object System.Windows.Forms.TabPage
$CreateTab.Text = 'Create one group'
$CreateTab.BackColor = [System.Drawing.Color]::White
$CreateTab.AutoScroll = $true
$CreateTab.AutoScrollMinSize = New-Object System.Drawing.Size(960, 622)
$Script:TabControl.TabPages.Add($CreateTab)
 
[void](Add-UiLabel -Parent $CreateTab -Text 'Group name *' -X 18 -Y 16)
$Script:GroupNameBox = Add-UiTextBox -Parent $CreateTab -X 215 -Y 16 -Width 735
[void](Add-UiLabel -Parent $CreateTab -Text 'Description' -X 18 -Y 50)
$Script:DescriptionBox = Add-UiTextBox -Parent $CreateTab -X 215 -Y 50 -Width 735 -Height 44 -Multiline $true
$Script:DescriptionBox.Text = '*DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal'
[void](Add-UiLabel -Parent $CreateTab -Text 'Mail nickname *' -X 18 -Y 105)
$Script:MailNicknameBox = Add-UiTextBox -Parent $CreateTab -X 215 -Y 105 -Width 350
$Script:GroupNameBox.Add_TextChanged({
    if ([System.String]::IsNullOrWhiteSpace($Script:MailNicknameBox.Text) -or $Script:MailNicknameBox.Tag -eq 'Auto')
    {
        try
        {
            $Script:MailNicknameBox.Text = Get-SafeMailNickname -DisplayName $Script:GroupNameBox.Text
            $Script:MailNicknameBox.Tag = 'Auto'
        }
        catch
        {
            $Script:MailNicknameBox.Clear()
        }
    }
})
$Script:MailNicknameBox.Add_TextChanged({
    if ($Script:MailNicknameBox.Focused) { $Script:MailNicknameBox.Tag = 'Manual' }
})
 
[void](Add-UiLabel -Parent $CreateTab -Text 'Group type' -X 18 -Y 142)
$Script:GroupTypeBox = New-Object System.Windows.Forms.ComboBox
$Script:GroupTypeBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Script:GroupTypeBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Script:GroupTypeBox.BackColor = [System.Drawing.Color]::FromArgb(231, 237, 246)
$Script:GroupTypeBox.Items.AddRange(@(
    'PIM-enabled security (role-assignable)'
    'Security group (assigned membership)'
    'Security group (dynamic membership)'
    'Microsoft 365 group (assigned membership)'
    'Microsoft 365 group (dynamic membership)'
))
$Script:GroupTypeBox.SelectedIndex = 0
$Script:GroupTypeBox.Location = New-Object System.Drawing.Point(215, 142)
$Script:GroupTypeBox.Size = New-Object System.Drawing.Size(420, 26)
$CreateTab.Controls.Add($Script:GroupTypeBox)
 
$Script:MembershipTypeLabel = Add-UiLabel -Parent $CreateTab -Text 'Dynamic member type' -X 18 -Y 178
$Script:MembershipTypeBox = New-Object System.Windows.Forms.ComboBox
$Script:MembershipTypeBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Script:MembershipTypeBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Script:MembershipTypeBox.Items.AddRange(@('User', 'Device'))
$Script:MembershipTypeBox.SelectedIndex = 0
$Script:MembershipTypeBox.Location = New-Object System.Drawing.Point(215, 178)
$Script:MembershipTypeBox.Size = New-Object System.Drawing.Size(180, 26)
$CreateTab.Controls.Add($Script:MembershipTypeBox)
$Script:MembershipRuleLabel = Add-UiLabel -Parent $CreateTab -Text 'Membership rule' -X 18 -Y 212
$Script:MembershipRuleBox = Add-UiTextBox -Parent $CreateTab -X 215 -Y 212 -Width 735 -Height 54 -Multiline $true
$Script:MembershipRuleBox.Visible = $false
$Script:MembershipRuleLabel.Visible = $false
$Script:MembershipTypeBox.Visible = $false
$Script:MembershipTypeLabel.Visible = $false
 
$Script:ExistingGroupIdLabel = Add-UiLabel -Parent $CreateTab -Text 'Existing Group Name or ID' -X 18 -Y 178
$Script:ExistingGroupIdBox = Add-UiTextBox -Parent $CreateTab -X 215 -Y 178 -Width 400
$Script:ApplyExistingButton = New-Object System.Windows.Forms.Button
$Script:ApplyExistingButton.Text = 'Apply PIM to existing group'
$Script:ApplyExistingButton.Location = New-Object System.Drawing.Point(630, 175)
$Script:ApplyExistingButton.Size = New-Object System.Drawing.Size(220, 30)
$CreateTab.Controls.Add($Script:ApplyExistingButton)
$Script:ExistingGroupHint = Add-UiLabel -Parent $CreateTab -Text 'Enter the exact display name or object ID. If multiple groups match, use the object ID. Only PIM settings are changed; no eligible members are added.' -X 215 -Y 212 -Width 700 -Height 38
 
$Script:PimGroupBox = New-Object System.Windows.Forms.GroupBox
$Script:PimGroupBox.Text = 'PIM Member settings'
$Script:PimGroupBox.Location = New-Object System.Drawing.Point(18, 276)
$Script:PimGroupBox.Size = New-Object System.Drawing.Size(932, 230)
$Script:PimGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(37, 52, 74)
$CreateTab.Controls.Add($Script:PimGroupBox)
[void](Add-UiLabel -Parent $Script:PimGroupBox -Text 'Activation maximum (hours)' -X 14 -Y 28 -Width 240)
$Script:ActivationHoursBox = Add-UiNumeric -Parent $Script:PimGroupBox -X 260 -Y 28 -Minimum 1 -Maximum 24 -Value 8
$Script:RequireMfaActivationBox = Add-UiCheckBox -Parent $Script:PimGroupBox -Text 'Require MFA on activation' -X 375 -Y 28 -Checked $true
$EligiblePolicySection = New-Object System.Windows.Forms.GroupBox
$EligiblePolicySection.Text = 'Eligible Assignments'
$EligiblePolicySection.Location = New-Object System.Drawing.Point(14, 62)
$EligiblePolicySection.Size = New-Object System.Drawing.Size(430, 154)
$Script:PimGroupBox.Controls.Add($EligiblePolicySection)
[void](Add-UiLabel -Parent $EligiblePolicySection -Text 'Eligible assignment expiry (days)' -X 12 -Y 26 -Width 245)
$Script:EligibilityDaysBox = Add-UiNumeric -Parent $EligiblePolicySection -X 270 -Y 26 -Minimum 1 -Maximum 3650 -Value 365
$Script:AllowPermanentEligibleBox = Add-UiCheckBox -Parent $EligiblePolicySection -Text 'Allow permanent eligible assignments' -X 12 -Y 58 -Checked $false -Width 400
$ActivePolicySection = New-Object System.Windows.Forms.GroupBox
$ActivePolicySection.Text = 'Active Assignments'
$ActivePolicySection.Font = New-Object System.Drawing.Font($Script:MainForm.Font.FontFamily, $Script:MainForm.Font.Size, [System.Drawing.FontStyle]::Bold)
$ActivePolicySection.ForeColor = [System.Drawing.Color]::DarkRed
$ActivePolicySection.Location = New-Object System.Drawing.Point(456, 62)
$ActivePolicySection.Size = New-Object System.Drawing.Size(462, 154)
$Script:PimGroupBox.Controls.Add($ActivePolicySection)
$ActivePolicyNote = Add-UiLabel -Parent $ActivePolicySection -Text 'permanent assignments allow always-active access' -X 145 -Y 0 -Width 305 -Height 18
$ActivePolicyNote.Font = New-Object System.Drawing.Font($Script:MainForm.Font.FontFamily, ($Script:MainForm.Font.Size - 1), [System.Drawing.FontStyle]::Regular)
$ActivePolicyNote.ForeColor = $Script:PimGroupBox.ForeColor
$ActivePolicyBody = New-Object System.Windows.Forms.Panel
$ActivePolicyBody.Location = New-Object System.Drawing.Point(6, 20)
$ActivePolicyBody.Size = New-Object System.Drawing.Size(450, 128)
$ActivePolicyBody.Anchor = 'Top, Bottom, Left, Right'
$ActivePolicyBody.BackColor = [System.Drawing.Color]::FromArgb(255, 242, 242)
$ActivePolicyBody.Font = $Script:MainForm.Font
$ActivePolicyBody.ForeColor = $Script:PimGroupBox.ForeColor
$ActivePolicySection.Controls.Add($ActivePolicyBody)
[void](Add-UiLabel -Parent $ActivePolicyBody -Text 'Expiry days (blank keeps current default)' -X 6 -Y 6 -Width 285)
$Script:ActiveAssignmentDaysBox = Add-UiTextBox -Parent $ActivePolicyBody -X 304 -Y 6 -Width 90
$Script:AllowPermanentActiveBox = Add-UiCheckBox -Parent $ActivePolicyBody -Text 'Allow permanent active assignments (NOT RECOMMENDED)' -X 6 -Y 38 -Checked $false -Width 438
$Script:AllowPermanentActiveBox.ForeColor = [System.Drawing.Color]::DarkRed
$Script:RequireMfaActiveBox = Add-UiCheckBox -Parent $ActivePolicyBody -Text 'Require MFA on active assignment' -X 6 -Y 70 -Checked $true -Width 438
$Script:RequireJustificationActiveBox = Add-UiCheckBox -Parent $ActivePolicyBody -Text 'Require justification on active assignment' -X 6 -Y 100 -Checked $true -Width 438

$Script:EligibleMembersLabel = Add-UiLabel -Parent $CreateTab -Text "Eligible member UPNs (optional)`r`nSeparate with commas, semicolons, or new lines." -X 18 -Y 518 -Width 330 -Height 50
$Script:EligibleMembersBox = Add-UiTextBox -Parent $CreateTab -X 350 -Y 518 -Width 600 -Height 50 -Multiline $true
$CreateHint = Add-UiLabel -Parent $CreateTab -Text 'This assigns listed users as PIM-eligible members, not active members. Entra role assignment to the group is a separate step.' -X 18 -Y 574 -Width 745 -Height 42
$CreateHint.ForeColor = [System.Drawing.Color]::DimGray
$CreateButton = New-Object System.Windows.Forms.Button
$CreateButton.Text = 'Create group'
$CreateButton.Location = New-Object System.Drawing.Point(780, 574)
$CreateButton.Size = New-Object System.Drawing.Size(170, 34)
$CreateTab.Controls.Add($CreateButton)
 
# ==================================================
# ^^^^^ Create One Group Tab ^^^^^
# ==================================================
 
# ==================================================
# MARK: Bulk Create Tab
# ==================================================
 
$BulkTab = New-Object System.Windows.Forms.TabPage
$BulkTab.Text = 'Bulk create'
$BulkTab.BackColor = [System.Drawing.Color]::White
$Script:TabControl.TabPages.Add($BulkTab)
$TemplateButton = New-Object System.Windows.Forms.Button
$TemplateButton.Text = 'Create CSV / XLSX template'
$TemplateButton.Location = New-Object System.Drawing.Point(18, 16)
$TemplateButton.Size = New-Object System.Drawing.Size(210, 32)
$BulkTab.Controls.Add($TemplateButton)
$ImportBulkButton = New-Object System.Windows.Forms.Button
$ImportBulkButton.Text = 'Import CSV / XLSX'
$ImportBulkButton.Location = New-Object System.Drawing.Point(240, 16)
$ImportBulkButton.Size = New-Object System.Drawing.Size(160, 32)
$BulkTab.Controls.Add($ImportBulkButton)
$Script:BulkPathBox = Add-UiTextBox -Parent $BulkTab -X 410 -Y 20 -Width 540
$Script:BulkPathBox.ReadOnly = $true
$Script:BulkPathBox.Anchor = 'Top, Left, Right'
$BulkHint = Add-UiLabel -Parent $BulkTab -Text 'Bulk rows are fully validated before any group is created. For PIM groups, the row can include optional EligibleMemberUPNs separated by semicolons.' -X 18 -Y 58 -Width 930 -Height 40
$BulkHint.ForeColor = [System.Drawing.Color]::DimGray
$BulkHint.Anchor = 'Top, Left, Right'
$Script:BulkGrid = New-Object System.Windows.Forms.DataGridView
$Script:BulkGrid.Location = New-Object System.Drawing.Point(18, 105)
$Script:BulkGrid.Size = New-Object System.Drawing.Size(930, 425)
$Script:BulkGrid.Anchor = 'Top, Bottom, Left, Right'
$Script:BulkGrid.ReadOnly = $true
$Script:BulkGrid.AllowUserToAddRows = $false
$Script:BulkGrid.AllowUserToDeleteRows = $false
$Script:BulkGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
$Script:BulkGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$Script:BulkGrid.RowHeadersVisible = $false
$Script:BulkGrid.BackgroundColor = [System.Drawing.Color]::White
$Script:BulkGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Script:BulkGrid.EnableHeadersVisualStyles = $false
$Script:BulkGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(231, 237, 246)
$Script:BulkGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(37, 52, 74)
$Script:BulkGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(214, 230, 250)
$Script:BulkGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(23, 38, 62)
$Script:BulkGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(247, 249, 252)
$BulkTab.Controls.Add($Script:BulkGrid)
$Script:BulkCreateButton = New-Object System.Windows.Forms.Button
$Script:BulkCreateButton.Text = 'Create all validated groups'
$Script:BulkCreateButton.Location = New-Object System.Drawing.Point(730, 548)
$Script:BulkCreateButton.Size = New-Object System.Drawing.Size(220, 36)
$Script:BulkCreateButton.Anchor = 'Bottom, Right'
$Script:BulkCreateButton.Enabled = $false
$BulkTab.Controls.Add($Script:BulkCreateButton)
 
# ==================================================
# ^^^^^ Bulk Create Tab ^^^^^
# ==================================================
 
# ==================================================
# MARK: Assignment Reports Tab
# ==================================================

$ReportTab = New-Object System.Windows.Forms.TabPage
$ReportTab.Text = 'Assignments'
$ReportTab.BackColor = [System.Drawing.Color]::White
$Script:TabControl.TabPages.Add($ReportTab)
$ReportLayout = New-Object System.Windows.Forms.TableLayoutPanel
$ReportLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$ReportLayout.ColumnCount = 1
$ReportLayout.RowCount = 4
[void]$ReportLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
foreach ($Height in @(218, 28))
{
    [void]$ReportLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $Height)))
}
[void]$ReportLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$ReportLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$ReportLayout.Add_SizeChanged({
    param($Sender, $EventArgs)
    $Sender.RowStyles[0].Height = [Math]::Min(218, [Math]::Max(100, $Sender.ClientSize.Height - 180))
})
$ReportTab.Controls.Add($ReportLayout)
$ReportHeader = New-Object System.Windows.Forms.Panel
$ReportHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
$ReportHeader.Margin = New-Object System.Windows.Forms.Padding(0)
$ReportHeader.AutoScroll = $true
$ReportHeader.AutoScrollMinSize = New-Object System.Drawing.Size(970, 210)
$ReportLayout.Controls.Add($ReportHeader, 0, 0)
$SingleGroupSection = New-Object System.Windows.Forms.GroupBox
$SingleGroupSection.Text = 'One group'
$SingleGroupSection.Location = New-Object System.Drawing.Point(18, 10)
$SingleGroupSection.Size = New-Object System.Drawing.Size(940, 110)
$ReportHeader.Controls.Add($SingleGroupSection)
[void](Add-UiLabel -Parent $SingleGroupSection -Text 'Group Name or ID' -X 12 -Y 29 -Width 128)
$Script:ReportGroupBox = Add-UiTextBox -Parent $SingleGroupSection -X 150 -Y 27 -Width 280
$Script:ReportGroupBox.BackColor = [System.Drawing.Color]::WhiteSmoke
$ReportOneButton = New-Object System.Windows.Forms.Button
$ReportOneButton.Text = 'Query one group'
$ReportOneButton.Location = New-Object System.Drawing.Point(445, 24)
$ReportOneButton.Size = New-Object System.Drawing.Size(150, 32)
$SingleGroupSection.Controls.Add($ReportOneButton)
$AddEligibleButton = New-Object System.Windows.Forms.Button
$AddEligibleButton.Text = 'Add eligible members'
$AddEligibleButton.Location = New-Object System.Drawing.Point(605, 24)
$AddEligibleButton.Size = New-Object System.Drawing.Size(180, 32)
$SingleGroupSection.Controls.Add($AddEligibleButton)
$RemoveMembersButton = New-Object System.Windows.Forms.Button
$RemoveMembersButton.Text = 'Remove members'
$RemoveMembersButton.Location = New-Object System.Drawing.Point(795, 24)
$RemoveMembersButton.Size = New-Object System.Drawing.Size(135, 32)
$SingleGroupSection.Controls.Add($RemoveMembersButton)
$SingleGroupHint = Add-UiLabel -Parent $SingleGroupSection -Text 'Query is read-only. Add/remove changes access; review the group, users and changes before confirming.' -X 12 -Y 68 -Width 900 -Height 30
$SingleGroupHint.ForeColor = [System.Drawing.Color]::DimGray
$AllGroupsSection = New-Object System.Windows.Forms.GroupBox
$AllGroupsSection.Text = 'All groups (read-only)'
$AllGroupsSection.Location = New-Object System.Drawing.Point(18, 128)
$AllGroupsSection.Size = New-Object System.Drawing.Size(940, 65)
$ReportHeader.Controls.Add($AllGroupsSection)
$ReportAllButton = New-Object System.Windows.Forms.Button
$ReportAllButton.Text = 'Query all eligible'
$ReportAllButton.Location = New-Object System.Drawing.Point(12, 24)
$ReportAllButton.Size = New-Object System.Drawing.Size(150, 32)
$AllGroupsSection.Controls.Add($ReportAllButton)
$ReportPermanentButton = New-Object System.Windows.Forms.Button
$ReportPermanentButton.Text = 'Scan permanent active'
$ReportPermanentButton.Location = New-Object System.Drawing.Point(172, 24)
$ReportPermanentButton.Size = New-Object System.Drawing.Size(165, 32)
$AllGroupsSection.Controls.Add($ReportPermanentButton)
$ReportHint = Add-UiLabel -Parent $AllGroupsSection -Text 'Scans include role-assignable groups. Verify direct members, including temporary PIM access.' -X 350 -Y 24 -Width 575 -Height 32
$ReportHint.ForeColor = [System.Drawing.Color]::DimGray
$Script:ReportStatus = Add-UiLabel -Parent $ReportTab -Text 'No report run.' -X 18 -Y 202 -Width 930 -Height 28
$Script:ReportStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
$Script:ReportStatus.Margin = New-Object System.Windows.Forms.Padding(18, 0, 18, 0)
$Script:ReportStatus.AutoEllipsis = $true
$ReportLayout.Controls.Add($Script:ReportStatus, 0, 1)
$Script:ReportGrid = New-Object System.Windows.Forms.DataGridView
$Script:ReportGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$Script:ReportGrid.Margin = New-Object System.Windows.Forms.Padding(18, 0, 18, 0)
$Script:ReportGrid.ReadOnly = $true
$Script:ReportGrid.AllowUserToAddRows = $false
$Script:ReportGrid.AllowUserToDeleteRows = $false
$Script:ReportGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
$Script:ReportGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$Script:ReportGrid.RowHeadersVisible = $false
$Script:ReportGrid.BackgroundColor = [System.Drawing.Color]::White
$Script:ReportGrid.EnableHeadersVisualStyles = $false
$Script:ReportGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(231, 237, 246)
foreach ($Column in @('GroupName', 'GroupId', 'Finding', 'Access', 'PrincipalName', 'UserPrincipalName', 'PrincipalId', 'PrincipalType', 'StartUtc', 'EndUtc'))
{
    [void]$Script:ReportGrid.Columns.Add($Column, $Column)
}
$ReportLayout.Controls.Add($Script:ReportGrid, 0, 2)
$ReportFooter = New-Object System.Windows.Forms.FlowLayoutPanel
$ReportFooter.Dock = [System.Windows.Forms.DockStyle]::Fill
$ReportFooter.Margin = New-Object System.Windows.Forms.Padding(0)
$ReportFooter.Padding = New-Object System.Windows.Forms.Padding(18, 8, 18, 8)
$ReportFooter.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$ReportFooter.WrapContents = $false
$ReportLayout.Controls.Add($ReportFooter, 0, 3)
$Script:ReportExportButton = New-Object System.Windows.Forms.Button
$Script:ReportExportButton.Text = 'Export results'
$Script:ReportExportButton.Size = New-Object System.Drawing.Size(200, 36)
$Script:ReportExportButton.Margin = New-Object System.Windows.Forms.Padding(0)
$Script:ReportExportButton.Enabled = $false
$ReportFooter.Controls.Add($Script:ReportExportButton)

# ==================================================
# ^^^^^ Assignment Reports Tab ^^^^^
# ==================================================

# ==================================================
# MARK: Activity Log
# ==================================================
 
$Script:AuditStatusLabel = Add-UiLabel -Parent $Script:MainForm -Text 'Audit log' -X 12 -Y 784 -Width 85 -Height 30
$Script:AuditStatusLabel.Anchor = 'Bottom, Left'
$Script:AuditPathBox = Add-UiTextBox -Parent $Script:MainForm -X 100 -Y 787 -Width 760
$Script:AuditPathBox.ReadOnly = $true
$Script:AuditPathBox.Anchor = 'Bottom, Left, Right'
$Script:AuditPathBox.Text = $Script:AuditPath
$Script:AuditBrowseButton = New-Object System.Windows.Forms.Button
$Script:AuditBrowseButton.Text = 'Browse...'
$Script:AuditBrowseButton.Location = New-Object System.Drawing.Point(870, 783)
$Script:AuditBrowseButton.Size = New-Object System.Drawing.Size(140, 30)
$Script:AuditBrowseButton.Anchor = 'Bottom, Right'
$Script:MainForm.Controls.Add($Script:AuditBrowseButton)
 
$Script:LogBox = New-Object System.Windows.Forms.TextBox
$Script:LogBox.Multiline = $true
$Script:LogBox.ReadOnly = $true
$Script:LogBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$Script:LogBox.Location = New-Object System.Drawing.Point(12, 822)
$Script:LogBox.Size = New-Object System.Drawing.Size(1000, 70)
$Script:LogBox.Anchor = 'Bottom, Left, Right'
$Script:LogBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Script:LogBox.BackColor = [System.Drawing.Color]::White
$Script:MainForm.Controls.Add($Script:LogBox)
 
foreach ($Button in @($CheckConnectionButton, $DisconnectButton, $Script:ApplyExistingButton, $TemplateButton, $ImportBulkButton, $Script:AuditBrowseButton, $HelpButton, $ReportOneButton, $ReportAllButton, $ReportPermanentButton, $Script:ReportExportButton, $RemoveMembersButton))
{
    Set-UiButtonStyle -Button $Button
}
$RemoveMembersButton.ForeColor = [System.Drawing.Color]::DarkRed
$RemoveMembersButton.FlatAppearance.BorderColor = [System.Drawing.Color]::IndianRed
foreach ($Button in @($ConnectButton, $CreateButton, $Script:BulkCreateButton, $AddEligibleButton))
{
    Set-UiButtonStyle -Button $Button -Primary $true
}
 
# ==================================================
# ^^^^^ Activity Log ^^^^^
# ==================================================
 
# ==================================================
# MARK: Form and Graph Connection Events
# ==================================================
 
$Script:GroupTypeBox.Add_SelectedIndexChanged({ Handle-GroupTypeChanged })
$HelpButton.Add_Click({
    try
    {
        Show-HelpWindow
    }
    catch
    {
        Write-UiLog "Could not open the help window: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Help window failed to open' -Icon Error
    }
})
$Script:AuditBrowseButton.Add_Click({
    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    try
    {
        $Dialog.Title = 'Select audit log location and file name'
        $Dialog.Filter = 'Log files (*.log)|*.log'
        $Dialog.DefaultExt = 'log'
        $Dialog.AddExtension = $true
        $Dialog.CheckPathExists = $true
        $Dialog.OverwritePrompt = $false
        if (-not [string]::IsNullOrWhiteSpace($Script:AuditPath))
        {
            $Dialog.FileName = [System.IO.Path]::GetFileName($Script:AuditPath)
            $Dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($Script:AuditPath)
        }
        if ($Dialog.ShowDialog($Script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        if ([System.IO.Path]::GetExtension($Dialog.FileName) -ne '.log') { throw 'The audit file must have a .log extension.' }
        if ([System.IO.File]::Exists($Dialog.FileName) -and [System.Windows.Forms.MessageBox]::Show(
            $Script:MainForm, 'This log file already exists. Append audit records to it without overwriting?',
            'Confirm audit log', [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
 
        $PreviousPath = $Script:AuditPath
        Write-UiLog "Audit destination changed to '$($Dialog.FileName)'."
        $Script:AuditPath = $Dialog.FileName
        $Script:AuditPathBox.Text = $Script:AuditPath
        Write-UiLog -Message "Audit destination selected. Earlier events remain in '$PreviousPath'." -Required
    }
    catch
    {
        Show-Message -Message $_.Exception.Message -Title 'Audit log selection failed' -Icon Error
    }
    finally
    {
        $Dialog.Dispose()
    }
})
$Script:AllowPermanentEligibleBox.Add_CheckedChanged({
    $Script:EligibilityDaysBox.Enabled = -not $Script:AllowPermanentEligibleBox.Checked
})
$Script:AllowPermanentActiveBox.Add_CheckedChanged({
    $Script:ActiveAssignmentDaysBox.Enabled = -not $Script:AllowPermanentActiveBox.Checked
})
$ConnectButton.Add_Click({
    try
    {
        $Script:ReportRows = @()
        $Script:ReportMode = $null
        $Script:ReportGrid.Rows.Clear()
        $Script:ReportExportButton.Enabled = $false
        $Script:ReportStatus.Text = 'No report run.'
        $ConnectButton.Enabled = $false
        $Script:ConnectionStatusLabel.Text = 'Signing in to Microsoft Graph...'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        [System.Windows.Forms.Application]::DoEvents()
        Connect-ToGraph | Out-Null
    }
    catch
    {
        $Script:ConnectionStatusLabel.Text = 'Connection failed'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkRed
        Write-UiLog "Graph sign-in failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Microsoft Graph sign-in failed' -Icon Error
    }
    finally
    {
        $ConnectButton.Enabled = $true
    }
})
$CheckConnectionButton.Add_Click({
    try
    {
        if (-not (Test-GraphConnection))
        {
            Show-Message -Message 'The Graph connection works, but required scopes are missing. Disconnect and reconnect to consent to all requested scopes.' -Icon Warning
        }
    }
    catch
    {
        $Script:ConnectionStatusLabel.Text = 'Not connected or check failed'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::DarkRed
        Write-UiLog "Graph connection check failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Connection check failed' -Icon Error
    }
})
$DisconnectButton.Add_Click({
    try
    {
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)
        {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
        }
        $Script:ConnectionStatusLabel.Text = 'Not connected'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Black
        $Script:ReportRows = @()
        $Script:ReportMode = $null
        $Script:ReportGrid.Rows.Clear()
        $Script:ReportExportButton.Enabled = $false
        $Script:ReportStatus.Text = 'No report run.'
        Write-UiLog 'Disconnected from Microsoft Graph.'
        $Script:AuditAccount = 'Not connected'
        $Script:AuditTenant = 'Unknown'
    }
    catch
    {
        Write-UiLog "Graph disconnect failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Disconnect failed' -Icon Error
    }
})
# ==================================================
# ^^^^^ Form and Graph Connection Events ^^^^^
# ==================================================
 
# ==================================================
# MARK: Single Group Creation and PIM Recovery Events
# ==================================================
 
$CreateButton.Add_Click({
    try
    {
        $Configuration = Get-SingleGroupConfiguration
        Assert-GraphReady
        $Summary = "Create '$($Configuration.GroupName)' as $($Configuration.GroupType)?"
        if ($Configuration.GroupType -eq 'PIMRoleAssignable')
        {
            $Summary += "`r`n`r`nPIM activation: up to $($Configuration.ActivationMaxHours) hours; activation MFA: $($Configuration.RequireMFAOnActivation); eligible expiry: $(if ($Configuration.AllowPermanentEligibleAssignments) { 'permanent allowed' } else { "$($Configuration.EligibleAssignmentDays) days" }); active-assignment MFA: $($Configuration.RequireMFAOnActiveAssignment); active-assignment justification: $($Configuration.RequireJustificationOnActiveAssignment); permanent active assignments: $($Configuration.AllowPermanentActiveAssignments)."
            if ($Configuration.EligibleMemberUPNs.Count -gt 0)
            {
                $Summary += "`r`nEligible member UPNs: $($Configuration.EligibleMemberUPNs -join ', ')"
            }
        }
 
        if ([System.Windows.Forms.MessageBox]::Show(
            $Script:MainForm,
            $Summary,
            'Confirm group creation',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) -ne [System.Windows.Forms.DialogResult]::Yes)
        {
            return
        }
 
        $CreateButton.Enabled = $false
        $Script:TabControl.Enabled = $false
        $Script:AuditBrowseButton.Enabled = $false
        $Script:MainForm.UseWaitCursor = $true
        [System.Windows.Forms.Application]::DoEvents()
        $Result = Invoke-CreateGroupWorkflow -Configuration $Configuration
        if ($Configuration.GroupType -eq 'PIMRoleAssignable')
        {
            $Script:ExistingGroupIdBox.Text = $Result.GroupId
        }
        Write-UiLog "$($Result.Status): $($Result.GroupName) [$($Result.GroupId)] $($Result.Details)"
        if ($Script:AuditFailure)
        {
            $Result.Details += "`r`nWARNING: $($Script:AuditFailure)"
        }
        Show-Message -Message "$($Result.Status)`r`n$($Result.GroupName)`r`nGroup ID: $($Result.GroupId)`r`n`r`n$($Result.Details)"
    }
    catch
    {
        Write-UiLog "Create failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Group creation failed' -Icon Error
    }
    finally
    {
        $CreateButton.Enabled = $true
        $Script:TabControl.Enabled = $true
        $Script:AuditBrowseButton.Enabled = $true
        $Script:MainForm.UseWaitCursor = $false
    }
})
$Script:ApplyExistingButton.Add_Click({
    try
    {
        Assert-GraphReady
        $Configuration = Get-SingleGroupConfiguration -PolicyOnly
        if ($Configuration.GroupType -ne 'PIMRoleAssignable')
        {
            throw 'Select the PIM-enabled role-assignable group type.'
        }
        $Group = Get-ExistingPimGroup -GroupNameOrId $Script:ExistingGroupIdBox.Text
        $Summary = "Apply the selected PIM Member settings to:`r`n$($Group.displayName)`r`nGroup ID: $($Group.id)`r`n`r`nThis changes the existing policy, can partially complete if an error occurs, and does not create a group, assign roles, or add eligible members. Continue?"
        if ([System.Windows.Forms.MessageBox]::Show(
            $Script:MainForm, $Summary, 'Confirm existing group PIM changes',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) -ne [System.Windows.Forms.DialogResult]::Yes) { return }
 
        $Script:TabControl.Enabled = $false
        $Script:AuditBrowseButton.Enabled = $false
        $Script:MainForm.UseWaitCursor = $true
        [System.Windows.Forms.Application]::DoEvents()
        $Settings = $Configuration | Select-Object ActivationMaxHours, RequireMFAOnActivation, AllowPermanentEligibleAssignments, EligibleAssignmentDays, AllowPermanentActiveAssignments, ActiveAssignmentDays, RequireMFAOnActiveAssignment, RequireJustificationOnActiveAssignment | ConvertTo-Json -Compress
        Write-UiLog -Message "Applying PIM Member settings to existing group '$($Group.displayName)' [$($Group.id)]: $Settings" -Required
        Set-PimMemberSettings -GroupId $Group.id -Configuration $Configuration
        Write-UiLog -Message "PIM Member settings applied and verified for '$($Group.displayName)' [$($Group.id)]. No eligible members were added." -Required
        Show-Message -Message "PIM Member settings applied and verified for $($Group.displayName).`r`nGroup ID: $($Group.id)`r`n`r`nNo eligible members were added. Verify the settings in Entra PIM."
    }
    catch
    {
        Write-UiLog "Existing group PIM update failed; settings may be partially applied: $($_.Exception.Message)"
        Show-Message -Message "PIM update failed. Settings may be partially applied; review the group in Entra PIM.`r`n`r`n$($_.Exception.Message)" -Title 'PIM update failed' -Icon Error
    }
    finally
    {
        $Script:TabControl.Enabled = $true
        $Script:AuditBrowseButton.Enabled = $true
        $Script:MainForm.UseWaitCursor = $false
    }
})
# ==================================================
# ^^^^^ Single Group Creation and PIM Recovery Events ^^^^^
# ==================================================
 
# ==================================================
# MARK: Bulk Creation Events
# ==================================================
 
$TemplateButton.Add_Click({
    try
    {
        New-TemplateFile
    }
    catch
    {
        Show-Message -Message $_.Exception.Message -Title 'Template creation failed' -Icon Error
        Write-UiLog "Template creation failed: $($_.Exception.Message)"
    }
})
$ImportBulkButton.Add_Click({
    try
    {
        Select-BulkFile
    }
    catch
    {
        Show-Message -Message $_.Exception.Message -Title 'Bulk import failed' -Icon Error
        Write-UiLog "Bulk import failed: $($_.Exception.Message)"
    }
})
$Script:BulkCreateButton.Add_Click({ Invoke-BulkCreate })
 
# ==================================================
# ^^^^^ Bulk Creation Events ^^^^^
# ==================================================

# ==================================================
# MARK: Assignment Events
# ==================================================

$ReportOneButton.Add_Click({ Invoke-AssignmentReport -Mode 'OneEligible' })
$ReportAllButton.Add_Click({ Invoke-AssignmentReport -Mode 'AllEligible' })
$ReportPermanentButton.Add_Click({ Invoke-AssignmentReport -Mode 'Permanent' })
$AddEligibleButton.Add_Click({
    try { Show-AddEligibleMembers }
    catch
    {
        Write-UiLog "Eligible assignment preparation failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Eligible assignment failed' -Icon Error
    }
})
$RemoveMembersButton.Add_Click({
    try { Show-RemoveMembers }
    catch
    {
        Write-UiLog "Member removal preparation failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Member removal failed' -Icon Error
    }
})
$Script:ReportExportButton.Add_Click({
    try { Export-AssignmentReport }
    catch
    {
        Write-UiLog "Assignment report export failed: $($_.Exception.Message)"
        Show-Message -Message $_.Exception.Message -Title 'Export failed' -Icon Error
    }
})

# ==================================================
# ^^^^^ Assignment Events ^^^^^
# ==================================================
 
# ==================================================
# MARK: Initialize and Display the GUI
# ==================================================
 
Update-GroupTypeVisibility
$Script:MainForm.Add_Shown({
    $Script:MainForm.BringToFront()
    [void]$Script:MainForm.Activate()
    $Script:MainForm.TopMost = $false
    Write-UiLog "Session started. Audit file: $($Script:AuditPath)"
    try
    {
        $Context = Get-GraphContext
        if ($Context -and -not [System.String]::IsNullOrWhiteSpace($Context.Account))
        {
            [void](Test-GraphConnection)
        }
    }
    catch
    {
        Write-UiLog "Startup connection check failed: $($_.Exception.Message)"
        $Script:ConnectionStatusLabel.Text = 'Not connected'
        $Script:ConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Black
    }
})
$Script:MainForm.Add_FormClosed({ Write-UiLog 'Session closed.' })
$Script:MainForm.Add_FormClosing({
    param($Sender, $EventArgs)
    if ($Sender.UseWaitCursor -or -not $ConnectButton.Enabled)
    {
        $EventArgs.Cancel = $true
        Write-UiLog 'Close blocked: wait for the current operation to finish before closing PIM Group Manager.'
    }
})
 
$WorkingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$Script:MainForm.Size = New-Object System.Drawing.Size(
    [Math]::Min($Script:MainForm.Width, $WorkingArea.Width),
    [Math]::Min($Script:MainForm.Height, $WorkingArea.Height)
)
[void]$Script:MainForm.ShowDialog()
 
# ==================================================
# ^^^^^ Initialize and Display the GUI ^^^^^
# ==================================================
}
finally
{
    try
    {
        if ($Script:HelpWindow -and -not $Script:HelpWindow.IsDisposed)
        {
            $Script:HelpWindow.Dispose()
        }
        if ($Script:MainForm -and -not $Script:MainForm.IsDisposed)
        {
            $Script:MainForm.Dispose()
        }
    }
    finally
    {
        $InstanceGuard.Dispose()
    }
}
}
