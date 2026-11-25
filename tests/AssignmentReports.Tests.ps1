$ErrorActionPreference = 'Stop'
$ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'PIM-Group-Manager.ps1'
$ParseErrors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$ParseErrors)
if ($ParseErrors) { throw ($ParseErrors | Out-String) }
$Names = @('Get-GraphCollection', 'Get-ReportGroups', 'Get-GroupReportInstances',
    'Test-CurrentReportInstance', 'New-AssignmentReportRow', 'Get-EligibleReportRows', 'Get-PermanentReportRows',
    'Test-ReportAssignmentScope', 'Write-AssignmentReportFile', 'ConvertTo-EligibleMemberUPNs',
    'Get-ExistingPimGroup', 'Get-EntraUserByUpn', 'Add-PimEligibleMember', 'New-EligibleAssignmentSchedule',
    'Get-MemberRemovalPlan', 'Invoke-MemberRemoval', 'Invoke-MemberRemovalBatch')
foreach ($Name in $Names)
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    if (-not $Function) { throw "Missing function $Name" }
    . ([scriptblock]::Create($Function.Extent.Text))
}

function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
foreach ($Name in @('Get-MemberManagementGroupNameOrId', 'Show-AddEligibleMembers', 'Show-RemoveMembers',
    'Add-UiLabel', 'Add-UiTextBox', 'Add-UiCheckBox', 'Set-UiButtonStyle'))
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    $Definition = $Function.Extent.Text.Replace('$Dialog.ShowDialog($Script:MainForm)', '(Test-MemberManagementDialog -Dialog $Dialog)')
    . ([scriptblock]::Create($Definition))
}
function Test-MemberManagementDialog($Dialog)
{
    $GroupField = @($Dialog.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Text -eq 'Example group' })[0]
    Assert-True ($GroupField -and $GroupField.BorderStyle -eq [System.Windows.Forms.BorderStyle]::None -and
        $GroupField.Font.Bold -and $GroupField.Font.Size -eq $Dialog.Font.Size + 1) 'Member popup must show the main group as borderless, bold text one font size larger.'
    if ($Dialog.Text -eq 'Remove members')
    {
        $Heading = @($Dialog.Controls | Where-Object { $_.Text -eq 'Select what to remove:' })[0]
        $ActiveOption = @($Dialog.Controls | Where-Object { $_.Text -eq 'PIM active membership schedules' })[0]
        Assert-True (-not $Heading.Bounds.IntersectsWith($ActiveOption.Bounds)) 'Removal heading must not hide the first checkbox.'
    }
    $Script:CheckedMemberDialogs++
    return [System.Windows.Forms.DialogResult]::Cancel
}
$Script:MainForm = [pscustomobject]@{ Font = New-Object System.Drawing.Font('Segoe UI', 9) }
$Script:ReportGroupBox = [pscustomobject]@{ Text = '' }
$Script:CheckedMemberDialogs = 0
try
{
    foreach ($Action in @('Show-AddEligibleMembers', 'Show-RemoveMembers'))
    {
        foreach ($Blank in @('', '   '))
        {
            $Script:ReportGroupBox.Text = $Blank
            try { & $Action; throw 'Empty main group field was accepted.' }
            catch { if ($_.Exception.Message -notmatch 'Enter a group name or object ID in One group') { throw } }
        }
        $Script:ReportGroupBox.Text = '  Example group  '
        & $Action
    }
    Assert-True ($Script:CheckedMemberDialogs -eq 2) 'Both member popups must validate and copy the main group.'
}
finally
{
    $Script:MainForm.Font.Dispose()
    $Script:MainForm = $null
    $Script:ReportGroupBox = $null
}

$Group = @{ id = '11111111-1111-1111-1111-111111111111'; displayName = 'Example group'; isAssignableToRole = $true }
$Now = [datetimeoffset]'2026-01-03T00:00:00Z'
$Script:Responses = @{}
$Script:Calls = @()
function Invoke-PimGraphRequest
{
    param($Method, $Uri)
    $Script:Calls += $Uri
    if (-not $Script:Responses.ContainsKey($Uri)) { throw "Unexpected Graph request: $Uri" }
    return $Script:Responses[$Uri]
}

$GroupList = 'https://graph.microsoft.com/v1.0/groups?$filter=isAssignableToRole%20eq%20true&$select=id,displayName,isAssignableToRole'
$Page2 = 'https://graph.microsoft.com/v1.0/groups?next=2'
$Script:Responses[$GroupList] = @{ value = @($Group); '@odata.nextLink' = $Page2 }
$Script:Responses[$Page2] = @{ value = @(@{ id = '22222222-2222-2222-2222-222222222222'; displayName = 'Second'; isAssignableToRole = $true }) }
$Groups = @(Get-ReportGroups)
Assert-True ($Groups.Count -eq 2) 'All group pages should be read.'
$Script:Responses[$Page2] = @{ value = @(); '@odata.nextLink' = $Page2 }
try { @(Get-ReportGroups) | Out-Null; throw 'Repeated page was accepted.' }
catch { if ($_.Exception.Message -notmatch 'repeated Graph pagination') { throw } }
$Script:Responses[$Page2] = @{}
try { @(Get-ReportGroups) | Out-Null; throw 'Invalid page was accepted.' }
catch { if ($_.Exception.Message -notmatch 'invalid collection') { throw } }

$Filter = [uri]::EscapeDataString("groupId eq '$($Group.id)'")
$EligibilityUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilityScheduleInstances?`$filter=$Filter&`$expand=principal"
$AssignmentUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleInstances?`$filter=$Filter&`$expand=principal"
$MembersUri = "https://graph.microsoft.com/beta/groups/$($Group.id)/members"
$Principal = @{ displayName = 'Example User'; userPrincipalName = 'example@example.invalid'; '@odata.type' = '#microsoft.graph.user' }
$Script:Responses[$EligibilityUri] = @{ value = @(
    @{ principalId = 'one'; accessId = 'member'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = $null; principal = $Principal },
    @{ principalId = 'two'; accessId = 'owner'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = '2026-01-04T00:00:00Z'; principal = $Principal },
    @{ principalId = 'future'; accessId = 'member'; startDateTime = '2026-01-04T00:00:00Z'; endDateTime = $null; principal = $Principal },
    @{ principalId = 'expired'; accessId = 'member'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = '2026-01-03T00:00:00Z'; principal = $Principal }
) }
$Eligible = @(Get-EligibleReportRows -Group $Group -Now $Now)
Assert-True ($Eligible.Count -eq 2 -and $Eligible[1].Access -eq 'owner' -and $Eligible[0].UserPrincipalName -eq 'example@example.invalid') 'Eligible report should include current member and owner only.'

$Script:Responses[$AssignmentUri] = @{ value = @(
    @{ principalId = 'perm'; accessId = 'member'; startDateTime = $null; endDateTime = $null; assignmentType = 'assigned'; principal = $Principal },
    @{ principalId = 'temp'; accessId = 'member'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = '2026-01-04T00:00:00Z'; assignmentType = 'activated'; principal = $Principal },
    @{ principalId = 'owner'; accessId = 'owner'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = $null; assignmentType = 'assigned'; principal = $Principal }
) }
$Script:Responses[$MembersUri] = @{ value = @(
    @{ id = 'perm'; displayName = 'Permanent' },
    @{ id = 'temp'; displayName = 'Temporary' },
    @{ id = 'direct'; displayName = 'Direct' }
) }
$Active = @(Get-PermanentReportRows -Group $Group -Now $Now)
Assert-True ($Active.Count -eq 3) 'Expected permanent PIM and both direct-member candidates without duplication.'
Assert-True (@($Active | Where-Object { $_.Finding -eq 'Permanent PIM active' -and $_.PrincipalId -eq 'perm' }).Count -eq 1) 'Permanent PIM assignment missing.'
Assert-True (@($Active | Where-Object { $_.Finding -eq 'Direct member / temporary PIM - verify' -and $_.PrincipalId -eq 'temp' }).Count -eq 1) 'Temporary PIM member must not be presented as confirmed permanent.'
Assert-True (@($Active | Where-Object { $_.Finding -eq 'Direct member - verify permanence' -and $_.PrincipalId -eq 'direct' }).Count -eq 1) 'Direct member candidate missing.'
$Script:Responses[$AssignmentUri].value[1].assignmentType = 'unknownFutureValue'
try { @(Get-PermanentReportRows -Group $Group -Now $Now) | Out-Null; throw 'Unknown assignment type was accepted.' }
catch { if ($_.Exception.Message -notmatch 'Unknown PIM assignment type') { throw } }
$Script:Responses[$EligibilityUri].value[0].Remove('endDateTime')
try { @(Get-EligibleReportRows -Group $Group -Now $Now) | Out-Null; throw 'Missing expiration was accepted.' }
catch { if ($_.Exception.Message -notmatch 'without its principal') { throw } }

$Script:ReportScope = 'PrivilegedAssignmentSchedule.Read.AzureADGroup'
Assert-True (Test-ReportAssignmentScope @{ Scopes = @($Script:ReportScope) }) 'Read permission should permit the scan.'
Assert-True (Test-ReportAssignmentScope @{ Scopes = @('PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup') }) 'ReadWrite permission should permit the scan.'
Assert-True (-not (Test-ReportAssignmentScope @{ Scopes = @('User.Read') })) 'Missing permission should not permit the scan.'

$Upns = @(ConvertTo-EligibleMemberUPNs -Text "one@example.invalid,two@example.invalid; three@example.invalid`r`nfour@example.invalid,ONE@example.invalid")
Assert-True ($Upns.Count -eq 4 -and $Upns[1] -eq 'two@example.invalid') 'UPN separators should work with or without spaces.'
try { ConvertTo-EligibleMemberUPNs -Text 'valid@example.invalid;bad' | Out-Null; throw 'Invalid UPN was accepted.' }
catch { if ($_.Exception.Message -notmatch 'invalid UPN') { throw } }

$LookupFilter = [uri]::EscapeDataString("displayName eq 'Example group'")
$LookupUri = 'https://graph.microsoft.com/v1.0/groups?$filter=' + $LookupFilter + '&$select=id,displayName,securityEnabled,mailEnabled,isAssignableToRole,groupTypes,onPremisesSyncEnabled&$top=2'
$Group.securityEnabled = $true
$Group.mailEnabled = $false
$Group.groupTypes = @()
$Group.onPremisesSyncEnabled = $false
$Script:Responses[$LookupUri] = @{ value = @($Group); '@odata.nextLink' = $Page2 }
$Script:Responses[$Page2] = @{ value = @(@{ id = '22222222-2222-2222-2222-222222222222' }) }
try { Get-ExistingPimGroup -GroupNameOrId 'Example group' | Out-Null; throw 'Duplicate group on a later page was accepted.' }
catch { if ($_.Exception.Message -notmatch 'Multiple groups') { throw } }
$Script:Responses[$Page2] = @{ value = @() }
Assert-True ((Get-ExistingPimGroup -GroupNameOrId 'Example group').id -eq $Group.id) 'Unique role-assignable group should be resolved.'
$Group.isAssignableToRole = $false
$Group.securityEnabled = $false
$Group.mailEnabled = $true
$Group.groupTypes = @('Unified')
Assert-True ((Get-ExistingPimGroup -GroupNameOrId 'Example group' -AllowNonRoleAssignable).id -eq $Group.id) 'Microsoft 365 groups should be accepted for eligible assignments.'
try { Get-ExistingPimGroup -GroupNameOrId 'Example group' | Out-Null; throw 'Non-role-assignable group was accepted for policy changes.' }
catch { if ($_.Exception.Message -notmatch 'not a verified') { throw } }
$Group.groupTypes = @('Unified', 'DynamicMembership')
try { Get-ExistingPimGroup -GroupNameOrId 'Example group' -AllowNonRoleAssignable | Out-Null; throw 'Dynamic group was accepted.' }
catch { if ($_.Exception.Message -notmatch 'not a verified') { throw } }
$Group.groupTypes = @('Unified')
$Group.onPremisesSyncEnabled = $true
try { Get-ExistingPimGroup -GroupNameOrId 'Example group' -AllowNonRoleAssignable | Out-Null; throw 'Synced group was accepted.' }
catch { if ($_.Exception.Message -notmatch 'not a verified') { throw } }

$Start = [datetime]::Now.AddDays(5)
$End = $Start.AddDays(30)
$Schedule = New-EligibleAssignmentSchedule -StartNow $false -StartLocal $Start -Permanent $false -EndLocal $End
$ExpectedStart = [System.TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($Start, [DateTimeKind]::Unspecified)).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
Assert-True ($Schedule.startDateTime -eq $ExpectedStart -and $Schedule.expiration.type -eq 'afterDateTime') 'Custom local dates must be converted to UTC.'
$PermanentSchedule = New-EligibleAssignmentSchedule -StartNow $true -StartLocal $Start -Permanent $true -EndLocal $End
Assert-True ($PermanentSchedule.expiration.type -eq 'noExpiration') 'Permanent eligibility should use noExpiration.'
try { New-EligibleAssignmentSchedule -StartNow $false -StartLocal $Start -Permanent $false -EndLocal $Start | Out-Null; throw 'Invalid end was accepted.' }
catch { if ($_.Exception.Message -notmatch 'end time must be after') { throw } }

$Script:SubmittedBody = $null
function Write-UiLog { param($Message, [switch]$Required) }
function Invoke-MgGraphRequest
{
    param($Method, $Uri, $Body, $ContentType)
    if ($Method -ne 'POST') { throw 'Unexpected Graph request.' }
    $Script:SubmittedBody = $Body | ConvertFrom-Json
}
$Config = @{ AllowPermanentEligibleAssignments = $false; EligibleAssignmentDays = 365 }
$User = @{ id = '33333333-3333-3333-3333-333333333333'; userPrincipalName = 'one@example.invalid' }
Add-PimEligibleMember -GroupId $Group.id -UserPrincipalName 'one@example.invalid' -Configuration $Config -ScheduleInfo $Schedule -ResolvedUser $User
Assert-True ($Script:SubmittedBody.accessId -eq 'member' -and $Script:SubmittedBody.action -eq 'AdminAssign' -and
    $Script:SubmittedBody.scheduleInfo.startDateTime -eq $Schedule.startDateTime -and $Script:SubmittedBody.scheduleInfo.expiration.endDateTime -eq $Schedule.expiration.endDateTime) 'Eligible request must include the selected schedule and membership.'
Add-PimEligibleMember -GroupId $Group.id -UserPrincipalName 'one@example.invalid' -Configuration $Config -ResolvedUser $User
Assert-True ($Script:SubmittedBody.scheduleInfo.expiration.type -eq 'afterDateTime') 'Existing group creation must retain its default eligibility schedule.'

$EligibleSchedulesUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=$Filter"
$ActiveSchedulesUri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentSchedules?`$filter=$Filter"
$DirectMembersUri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/members?`$select=id"
$EligibleSchedule = @{ id = 'eligible-schedule'; groupId = $Group.id; principalId = $User.id; accessId = 'member'; memberType = 'direct'; status = 'Provisioned'; scheduleInfo = @{ startDateTime = '2026-01-01T00:00:00Z' } }
$ActiveSchedule = @{ id = 'active-schedule'; groupId = $Group.id; principalId = $User.id; accessId = 'member'; memberType = 'direct'; status = 'Provisioned'; assignmentType = 'activated'; scheduleInfo = @{ startDateTime = '2026-01-01T00:00:00Z' } }
$Script:Responses[$EligibleSchedulesUri] = @{ value = @($EligibleSchedule); '@odata.nextLink' = $Page2 }
$Script:Responses[$Page2] = @{ value = @(@{ id = 'owner'; groupId = $Group.id; principalId = $User.id; accessId = 'owner'; memberType = 'direct'; status = 'Provisioned' }) }
$Script:Responses[$ActiveSchedulesUri] = @{ value = @($ActiveSchedule) }
$Script:Responses[$DirectMembersUri] = @{ value = @(@{ id = $User.id }) }
$Script:Calls = @()
$Plan = @(Get-MemberRemovalPlan -Group $Group -Users @($User) -Eligible $true -Active $true -Direct $true)
Assert-True ($Plan.Count -eq 3 -and $Plan[0].Category -eq 'PIM active' -and $Plan[1].Category -eq 'PIM eligible' -and $Plan[2].Category -eq 'Group membership') 'Removal plan should include selected matching categories, not owner assignments.'
Assert-True ($Script:Calls -contains $Page2) 'Removal planning must follow schedule pagination.'
$Script:Calls = @()
$Plan = @(Get-MemberRemovalPlan -Group $Group -Users @($User) -Eligible $true -Active $false -Direct $false)
Assert-True ($Plan.Count -eq 1 -and $Plan[0].Category -eq 'PIM eligible' -and $Script:Calls -notcontains $ActiveSchedulesUri -and $Script:Calls -notcontains $DirectMembersUri) 'Unselected categories must not be queried or included.'
$EligibleSchedule.status = 'PendingProvisioning'
try { Get-MemberRemovalPlan -Group $Group -Users @($User) -Eligible $true | Out-Null; throw 'Pending schedule was accepted.' }
catch { if ($_.Exception.Message -notmatch 'Review it in Entra') { throw } }
$EligibleSchedule.status = 'Provisioned'
$Script:Responses[$Page2] = @{ value = @($EligibleSchedule) }
try { Get-MemberRemovalPlan -Group $Group -Users @($User) -Eligible $true | Out-Null; throw 'Ambiguous schedules were accepted.' }
catch { if ($_.Exception.Message -notmatch 'Multiple eligibilitySchedules') { throw } }
$Script:Responses[$Page2] = @{ value = @() }

$Script:RemovalCalls = @()
function Invoke-MgGraphRequest
{
    param($Method, $Uri, $Body, $ContentType)
    $Script:RemovalCalls += [pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body }
    if ($Method -eq 'POST') { return @{ id = 'request-1'; status = 'PendingProvisioning' } }
}
$Plan = @(Get-MemberRemovalPlan -Group $Group -Users @($User) -Eligible $true -Active $true -Direct $true)
foreach ($Operation in $Plan) { Invoke-MemberRemoval -GroupId $Group.id -Operation $Operation | Out-Null }
Assert-True ($Script:RemovalCalls.Count -eq 3) 'One request per planned operation should be submitted.'
$ActiveBody = $Script:RemovalCalls[0].Body | ConvertFrom-Json
$EligibleBody = $Script:RemovalCalls[1].Body | ConvertFrom-Json
Assert-True ($Script:RemovalCalls[0].Uri -match '/assignmentScheduleRequests$' -and $ActiveBody.action -eq 'adminRemove' -and $ActiveBody.accessId -eq 'member' -and $ActiveBody.principalId -eq $User.id -and $ActiveBody.groupId -eq $Group.id) 'Active request must remove the selected user membership.'
Assert-True ($Script:RemovalCalls[1].Uri -match '/eligibilityScheduleRequests$' -and $EligibleBody.action -eq 'adminRemove' -and $EligibleBody.principalId -eq $User.id) 'Eligible request must remove the selected user membership.'
Assert-True ($Script:RemovalCalls[2].Method -eq 'DELETE' -and $Script:RemovalCalls[2].Uri -eq "https://graph.microsoft.com/v1.0/groups/$($Group.id)/members/$($User.id)/`$ref") 'Direct removal must delete only the membership reference.'
try { Invoke-MemberRemoval -GroupId 'invalid' -Operation $Plan[2] | Out-Null; throw 'Invalid group ID was accepted.' }
catch { if ($_.Exception.Message -notmatch 'Invalid group or user ID') { throw } }
Assert-True ($Script:RemovalCalls.Count -eq 3) 'Invalid IDs must not trigger Graph changes.'
function Invoke-MgGraphRequest
{
    param($Method, $Uri, $Body, $ContentType)
    $Script:RemovalCalls += [pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body }
    if ($Uri -match '/eligibilityScheduleRequests$') { throw 'Simulated network failure' }
    if ($Method -eq 'POST') { return @{ id = 'request-1'; status = 'PendingProvisioning' } }
}
$Script:RemovalCalls = @()
$Outcome = Invoke-MemberRemovalBatch -GroupId $Group.id -Plan $Plan
Assert-True ($Script:RemovalCalls.Count -eq 2 -and $Outcome.Accepted.Count -eq 1 -and $Outcome.Failed -match 'PIM eligible' -and $Outcome.Remaining.Count -eq 1 -and $Outcome.Remaining[0] -match 'Group membership' -and $Outcome.ErrorMessage -match 'Simulated network failure') 'Batch must stop on error and distinguish accepted, uncertain and unattempted operations.'

$TestDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('PimReportTest-' + [guid]::NewGuid().ToString())
[void](New-Item -ItemType Directory -Path $TestDirectory)
$CsvPath = Join-Path $TestDirectory 'report.csv'
$ExcelPath = Join-Path $TestDirectory 'report.xlsx'
try
{
    $Rows = @([pscustomobject][ordered]@{ GroupName = '=danger'; PrincipalName = 'Sample'; GroupId = '00123' })
    Write-AssignmentReportFile -Path $CsvPath -Rows $Rows
    $Csv = @(Import-Csv -LiteralPath $CsvPath)
    Assert-True ($Csv.Count -eq 1 -and $Csv[0].GroupName -eq "'=danger" -and $Csv[0].GroupId -eq '00123') 'CSV export must protect formulas and preserve text.'
    $Rows[0].PrincipalName = 'Updated'
    Write-AssignmentReportFile -Path $CsvPath -Rows $Rows
    Assert-True ((Import-Csv -LiteralPath $CsvPath).PrincipalName -eq 'Updated') 'CSV overwrite should replace the previous report.'
    if (Get-Module -ListAvailable -Name ImportExcel)
    {
        Write-AssignmentReportFile -Path $ExcelPath -Rows $Rows
        $Excel = @(Import-Excel -Path $ExcelPath -WorksheetName 'Assignments')
        Assert-True ($Excel.Count -eq 1 -and $Excel[0].GroupName -eq "'=danger" -and $Excel[0].GroupId -eq '00123') 'Excel export must protect formulas and preserve text.'
        $Package = Open-ExcelPackage -Path $ExcelPath
        try
        {
            $Worksheet = $Package.Workbook.Worksheets['Assignments']
            Assert-True ([string]::IsNullOrWhiteSpace($Worksheet.Cells['A2'].Formula)) 'Excel cell must not be a formula.'
            Assert-True ($Worksheet.Tables.Count -eq 1 -and $Worksheet.Tables['PIMAssignments'].StyleName -eq 'TableStyleLight20') 'Excel report must contain a Light20 table.'
        }
        finally { Close-ExcelPackage -ExcelPackage $Package -NoSave }
        $Rows[0].PrincipalName = 'Replaced'
        Write-AssignmentReportFile -Path $ExcelPath -Rows $Rows
        Assert-True ((Import-Excel -Path $ExcelPath -WorksheetName 'Assignments').PrincipalName -eq 'Replaced') 'Excel overwrite should replace the previous report.'
    }
    function Get-Module { param($Name, [switch]$ListAvailable) return $null }
    try
    {
        try { Write-AssignmentReportFile -Path $ExcelPath -Rows $Rows; throw 'Missing ImportExcel was accepted.' }
        catch { if ($_.Exception.Message -notmatch 'ImportExcel is required') { throw } }
    }
    finally { Remove-Item -LiteralPath 'Function:\Get-Module' }
}
finally
{
    if (Test-Path -LiteralPath $CsvPath) { Remove-Item -LiteralPath $CsvPath }
    if (Test-Path -LiteralPath $ExcelPath) { Remove-Item -LiteralPath $ExcelPath }
    Remove-Item -LiteralPath $TestDirectory
}
foreach ($Name in @('Invoke-AssignmentReport', 'Export-AssignmentReport'))
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    $Definition = $Function.Extent.Text.Replace('$Dialog.ShowDialog($Script:MainForm)', '(Test-ReportExportDialog -Dialog $Dialog)')
    . ([scriptblock]::Create($Definition))
}
function Test-ReportExportDialog($Dialog)
{
    Assert-True ($Dialog.FileName -match ('^\d{4}-\d{2}-\d{2}_\d{4}-' + $Script:ExpectedReportSuffix + '-PIM-Assignment-Report$')) 'Export filename must identify the completed report after its timestamp.'
    return [System.Windows.Forms.DialogResult]::Cancel
}
function Assert-GraphReady { if ($Script:FailReport) { throw 'Simulated report failure.' } }
function Get-GraphContext { return @{ Scopes = @($Script:ReportScope) } }
function Get-ReportGroups { return @{ id = 'test-group'; displayName = 'Test group' } }
function Get-ExistingPimGroup { param($GroupNameOrId) Get-ReportGroups }
function Get-EligibleReportRows { param($Group, $Now) return $Script:TestReportRow }
function Get-PermanentReportRows { param($Group, $Now) return $Script:TestReportRow }
function Show-Message { param($Message, $Title, $Icon) $Script:ReportError = $Message }
$Script:TestReportRow = [pscustomobject]@{ GroupName = 'Test group'; GroupId = 'test-group'; Finding = 'Example'; Access = 'member'; PrincipalName = 'Example user'; UserPrincipalName = 'example@example.invalid'; PrincipalId = 'test-user'; PrincipalType = 'User'; StartUtc = ''; EndUtc = '' }
$Script:ReportGrid = New-Object System.Windows.Forms.DataGridView
$Script:ReportGrid.AllowUserToAddRows = $false
foreach ($Column in $Script:TestReportRow.PSObject.Properties.Name) { [void]$Script:ReportGrid.Columns.Add($Column, $Column) }
$Script:MainForm = [pscustomobject]@{ UseWaitCursor = $false }
$Script:TabControl = [pscustomobject]@{ Enabled = $true }
$Script:AuditBrowseButton = [pscustomobject]@{ Enabled = $true }
$Script:ReportExportButton = [pscustomobject]@{ Enabled = $false }
$Script:ReportGroupBox = [pscustomobject]@{ Text = 'Test group' }
$Script:ReportStatus = [pscustomobject]@{ Text = '' }
$Script:ReportStatus | Add-Member ScriptMethod Refresh {}
$Script:FailReport = $false
try
{
    foreach ($Case in @(
        @{ Mode = 'AllEligible'; Suffix = 'all-eligible' },
        @{ Mode = 'Permanent'; Suffix = 'all-permanent' },
        @{ Mode = 'OneEligible'; Suffix = 'one-group-eligible' }))
    {
        Invoke-AssignmentReport -Mode $Case.Mode
        Assert-True (@($Script:ReportRows).Count -eq 1 -and $Script:ReportMode -eq $Case.Mode) "Completed report must retain its type with its results. $Script:ReportError"
        $Script:ExpectedReportSuffix = $Case.Suffix
        Export-AssignmentReport
    }
    $Script:FailReport = $true
    Invoke-AssignmentReport -Mode 'AllEligible'
    Assert-True ($Script:ReportMode -eq $null -and $Script:ReportRows.Count -eq 0 -and
        $Script:ReportError -match 'Simulated report failure') 'Failed reports must discard both their results and filename metadata.'
    $Script:ReportRows = @($Script:TestReportRow)
    try { Export-AssignmentReport; throw 'Missing report type was accepted.' }
    catch { if ($_.Exception.Message -notmatch 'report type is unavailable') { throw } }
}
finally
{
    $Script:ReportGrid.Dispose()
    $Script:MainForm = $null
}
Write-Output 'Assignment report tests passed.'
