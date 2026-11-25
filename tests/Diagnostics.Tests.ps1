$ErrorActionPreference = 'Stop'
$ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'PIM-Group-Manager.ps1'
$Errors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$Errors)
if ($Errors) { throw ($Errors | Out-String) }
foreach ($Name in @('Get-GraphConnectionDetails', 'Get-GraphRequestLogUri', 'Invoke-LoggedGraphRequest', 'Invoke-PimGraphRequest'))
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    . ([scriptblock]::Create($Function.Extent.Text))
}
function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Get-Module { param($Name) return [pscustomobject]@{ Version = [version]'2.0.0' } }
$Context = [pscustomobject]@{
    Account = 'example@example.invalid'
    TenantId = 'test-tenant'
    AuthType = 'Delegated'
    Environment = 'Global'
    ClientId = 'public-client-id'
    ContextScope = 'Process'
    Scopes = @('User.Read', 'Group.ReadWrite.All', 'User.Read')
    AccessToken = 'secret-token-value'
    ClientSecret = 'secret-client-value'
}
$Details = Get-GraphConnectionDetails -Context $Context
$Json = $Details | ConvertTo-Json -Depth 5
Assert-True ($Details.Account -eq $Context.Account -and $Details.TenantId -eq 'test-tenant' -and
    $Details.GraphAuthenticationVersion -eq '2.0.0' -and $Details.Scopes.Count -eq 2) 'Connection details must accurately report the selected context and loaded module.'
Assert-True ($Json -notmatch 'secret-token-value|secret-client-value|AccessToken|ClientSecret') 'Connection details must use an explicit allowlist without credentials.'

$Script:LogMessages = @()
$Script:CaptureReportApi = $true
$Script:ReportApiRequests = [System.Collections.Generic.List[object]]::new()
$Script:Status = 200
$Script:FailRequest = $false
$Script:FailAudit = $false
$Script:NativeCalls = 0
function Write-UiLog
{
    param($Message, [switch]$Required)
    if ($Required -and $Script:FailAudit) { throw 'Simulated audit failure.' }
    $Script:LogMessages += $Message
}
function Invoke-MgGraphRequest
{
    param($Method, $Uri, $Body, $ContentType, $StatusCodeVariable, $SkipHttpErrorCheck, $ErrorAction)
    $Script:NativeCalls++
    $Script:ReceivedUri = $Uri
    $Script:ReceivedBody = $Body
    Set-Variable -Name $StatusCodeVariable -Value $Script:Status -Scope 1
    if ($Script:FailRequest) { throw 'Simulated transport failure.' }
    return @{ value = @('example'); error = if ($Script:Status -ge 400) { @{ code = 'ExampleError' } } else { $null } }
}

$Uri = 'https://graph.microsoft.com/v1.0/groups?$skiptoken=secret-page-token&$select=id'
$Response = Invoke-PimGraphRequest -Method GET -Uri $Uri
Assert-True ($Response.value[0] -eq 'example' -and $Script:ReportApiRequests[0].HttpStatusCode -eq 200) 'Logging wrapper must preserve the response and pass HTTP status back to the PIM caller.'
Assert-True ($Script:ReceivedUri -eq $Uri -and $Script:ReportApiRequests[0].Uri -match '\[REDACTED\]' -and
    ($Script:LogMessages -join ' ') -notmatch 'secret-page-token') 'Pagination tokens must be redacted only in diagnostics, not in the request sent to Graph.'
$Encoded = Get-GraphRequestLogUri -Uri 'https://graph.microsoft.com/v1.0/groups?%24skiptoken=encoded-secret&access_token=access-secret'
Assert-True ($Encoded -notmatch 'encoded-secret|access-secret') 'Encoded pagination parameter names and access tokens must also be redacted.'

Invoke-LoggedGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groups' -Body '{"private":"secret-body-value"}' | Out-Null
Assert-True ($Script:ReceivedBody -match 'secret-body-value' -and
    ($Script:LogMessages -join ' ') -notmatch 'secret-body-value' -and
    ($Script:ReportApiRequests | ConvertTo-Json -Depth 5) -notmatch 'secret-body-value') 'Request bodies must reach Graph unchanged but never enter API diagnostics.'

$Script:Status = 404
try { Invoke-PimGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/groups/missing'; throw 'HTTP failure was swallowed.' }
catch { if ($_.Exception.Data['HttpStatusCode'] -ne 404) { throw } }
Assert-True ($Script:ReportApiRequests[2].Outcome -eq 'HTTP error') 'HTTP failures must be recorded without being reported as successful requests.'
$Script:FailRequest = $true
try { Invoke-LoggedGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/groups'; throw 'Transport failure was swallowed.' }
catch { if ($_.Exception.Message -ne 'Simulated transport failure.') { throw } }
Assert-True ($Script:ReportApiRequests[3].Outcome -eq 'Failed') 'Transport failures must be recorded and propagated.'
$Script:FailRequest = $false
$Script:Status = 200
$Script:FailAudit = $true
$Before = $Script:NativeCalls
try { Invoke-LoggedGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/groups/example' -Body '{}'; throw 'Audit failure was ignored.' }
catch { if ($_.Exception.Message -ne 'Simulated audit failure.') { throw } }
Assert-True ($Script:NativeCalls -eq $Before) 'Mutation requests must not be sent when their required API audit record cannot be written.'
$Script:FailAudit = $false
$Script:CaptureReportApi = $false
$Before = $Script:ReportApiRequests.Count
Invoke-LoggedGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me' | Out-Null
Assert-True ($Script:ReportApiRequests.Count -eq $Before) 'Requests after a completed report must not alter that report trace.'
'Diagnostics tests passed.'
