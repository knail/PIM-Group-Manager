$ErrorActionPreference = 'Stop'
$ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'PIM-Group-Manager.ps1'
$ParseErrors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$ParseErrors)
if ($ParseErrors) { throw ($ParseErrors | Out-String) }
$Names = @('Get-GraphCollection', 'Get-ReportGroups', 'Get-GroupReportInstances',
    'Test-CurrentReportInstance', 'New-AssignmentReportRow', 'Get-EligibleReportRows', 'Get-PermanentReportRows',
    'Test-ReportAssignmentScope', 'Write-AssignmentReportFile')
foreach ($Name in $Names)
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    if (-not $Function) { throw "Missing function $Name" }
    . ([scriptblock]::Create($Function.Extent.Text))
}

function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

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
        try { Assert-True ([string]::IsNullOrWhiteSpace($Package.Workbook.Worksheets['Assignments'].Cells['A2'].Formula)) 'Excel cell must not be a formula.' }
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
Write-Output 'Assignment report tests passed.'
