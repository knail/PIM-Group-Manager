$ErrorActionPreference = 'Stop'
$ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'PIM-Group-Manager.ps1'
$ParseErrors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$ParseErrors)
if ($ParseErrors) { throw ($ParseErrors | Out-String) }

function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

# Exercise the actual startup and cleanup code without opening a GUI or contacting Graph.
$Wrapper = $Ast.EndBlock.Statements[0].PipelineElements[0].CommandElements[0].ScriptBlock
$LaunchTry = @($Wrapper.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] })[0]
$Guard = $Wrapper.Extent.Text.Substring(1, $LaunchTry.Extent.StartOffset - $Wrapper.Extent.StartOffset - 1)
$Guard = $Guard.Replace('Local\PIM-Group-Manager-', ('Local\PIM-Group-Manager-Test-' + [guid]::NewGuid() + '-'))
$Cleanup = $LaunchTry.Finally.Extent.Text
function New-TestLaunch([string]$Body)
{
    [scriptblock]::Create("& {`n$Guard`ntry {`n$Body`n} finally $Cleanup`n}")
}

Add-Type -AssemblyName System.Windows.Forms
$Script:MainForm = $null
$Script:HelpWindow = $null
$Script:LaunchCount = 0
$Duplicate = New-TestLaunch '$Script:LaunchCount++; throw "Duplicate launch entered initialization."'
$Launch = New-TestLaunch @'
$Script:LaunchCount++
$Script:MainForm = [pscustomobject]@{
    IsDisposed = $false
    WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    Activated = $false
}
$Script:MainForm | Add-Member ScriptMethod Activate { $this.Activated = $true }
$Script:MainForm | Add-Member ScriptMethod Dispose { $this.IsDisposed = $true }
$OriginalForm = $Script:MainForm
$DuplicateWarnings = @(& $Duplicate 3>&1)
Assert-True ($Script:LaunchCount -eq 1) 'Same-thread duplicate must not enter initialization.'
Assert-True ([object]::ReferenceEquals($OriginalForm, $Script:MainForm) -and -not $OriginalForm.IsDisposed) 'Duplicate must preserve the running form.'
Assert-True ($OriginalForm.Activated -and $OriginalForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) 'Duplicate should restore and activate the existing form.'
Assert-True ("$DuplicateWarnings" -match 'already running') 'Duplicate must report why it was skipped.'

$Job = Start-Job -ScriptBlock {
    param($Text)
    & ([scriptblock]::Create($Text))
} -ArgumentList $Duplicate.ToString()
try
{
    if (-not (Wait-Job $Job -Timeout 30)) { throw 'Cross-process duplicate test timed out.' }
    $Result = @(Receive-Job $Job -WarningVariable ProcessWarnings -WarningAction SilentlyContinue -ErrorAction Stop)
    Assert-True ($Job.State -eq 'Completed' -and $Result.Count -eq 0) 'Cross-process duplicate must not enter initialization.'
    Assert-True ("$ProcessWarnings" -match 'already running') 'Cross-process duplicate must report why it was skipped.'
}
finally
{
    if ($Job.State -eq 'Running') { Stop-Job $Job }
    Remove-Job $Job
}
'@
& $Launch
Assert-True $Script:MainForm.IsDisposed 'Normal exit must dispose the main form.'

$Relaunch = New-TestLaunch '$Script:MainForm = $null; $Script:HelpWindow = $null; $Script:LaunchCount++'
& $Relaunch
Assert-True ($Script:LaunchCount -eq 2) 'Normal exit must release the guard for a new launch.'

$FailedLaunch = New-TestLaunch '$Script:MainForm = $null; $Script:HelpWindow = $null; throw "Simulated startup failure."'
try { & $FailedLaunch; throw 'Startup failure was swallowed.' }
catch { if ($_.Exception.Message -ne 'Simulated startup failure.') { throw } }
& $Relaunch
Assert-True ($Script:LaunchCount -eq 3) 'Startup failure must release the guard.'

$FailedCleanup = New-TestLaunch @'
$Script:MainForm = [pscustomobject]@{ IsDisposed = $false }
$Script:MainForm | Add-Member ScriptMethod Dispose { throw 'Simulated disposal failure.' }
'@
try { & $FailedCleanup; throw 'Disposal failure was swallowed.' }
catch { if ($_.Exception.Message -notmatch 'Simulated disposal failure') { throw } }
& $Relaunch
Assert-True ($Script:LaunchCount -eq 4) 'Disposal failure must still release the guard.'

$ClosingCall = $Ast.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $Node.Member.Value -eq 'Add_FormClosing'
}, $true)
$ClosingHandler = [scriptblock]::Create($ClosingCall.Arguments[0].ScriptBlock.Extent.Text.TrimStart('{').TrimEnd('}'))
function Write-UiLog { param($Message) $Script:LastLog = $Message }
$ConnectButton = [pscustomobject]@{ Enabled = $true }
$Sender = [pscustomobject]@{ UseWaitCursor = $true }
$EventArgs = [pscustomobject]@{ Cancel = $false }
& $ClosingHandler $Sender $EventArgs
Assert-True ($EventArgs.Cancel -and $Script:LastLog -match 'Close blocked') 'Closing during work must be cancelled and logged.'
$Sender.UseWaitCursor = $false
$ConnectButton.Enabled = $false
$EventArgs.Cancel = $false
& $ClosingHandler $Sender $EventArgs
Assert-True $EventArgs.Cancel 'Closing during sign-in must be cancelled.'
$ConnectButton.Enabled = $true
$EventArgs.Cancel = $false
& $ClosingHandler $Sender $EventArgs
Assert-True (-not $EventArgs.Cancel) 'Closing while idle must remain allowed.'

$ShownCall = $Ast.Find({
    param($Node)
    $Node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $Node.Member.Value -eq 'Add_Shown'
}, $true)
$ShownHandler = [scriptblock]::Create($ShownCall.Arguments[0].ScriptBlock.Extent.Text.TrimStart('{').TrimEnd('}'))
$Script:StartupEvents = @()
$Script:MainForm = [pscustomobject]@{ TopMost = $true }
$Script:MainForm | Add-Member ScriptMethod BringToFront { $Script:StartupEvents += 'Front' }
$Script:MainForm | Add-Member ScriptMethod Activate { $Script:StartupEvents += 'Activate' }
function Get-GraphContext
{
    Assert-True (($Script:StartupEvents -join ',') -eq 'Front,Activate') 'The window must be brought forward and activated before the startup Graph check.'
    Assert-True (-not $Script:MainForm.TopMost) 'Always-on-top must be released before the startup Graph check.'
    $Script:StartupEvents += 'Graph'
    return $null
}
& $ShownHandler
Assert-True (($Script:StartupEvents -join ',') -eq 'Front,Activate,Graph') 'Startup must still check Graph after activating the window.'
$Script:MainForm = $null

'Startup tests passed.'
