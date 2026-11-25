$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'PIM-Group-Manager.ps1'
$Errors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$Errors)
if ($Errors) { throw ($Errors | Out-String) }
function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
foreach ($Name in @('Add-UiLabel', 'Add-UiTextBox', 'Add-UiNumeric', 'Add-UiCheckBox'))
{
    $Function = $Ast.Find({ param($Node) $Node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Node.Name -eq $Name }, $true)
    . ([scriptblock]::Create($Function.Extent.Text))
}
$Source = [System.IO.File]::ReadAllText($ScriptPath)
$Start = $Source.IndexOf('$Script:MainForm = New-Object System.Windows.Forms.Form')
$End = $Source.IndexOf('# ^^^^^ Main Window and Connection Controls ^^^^^', $Start)
. ([scriptblock]::Create($Source.Substring($Start, $End - $Start)))
try
{
    $Start = $Source.IndexOf('$CreateTab = New-Object System.Windows.Forms.TabPage')
    $End = $Source.IndexOf('# ^^^^^ Create One Group Tab ^^^^^', $Start)
    . ([scriptblock]::Create($Source.Substring($Start, $End - $Start)))
    Assert-True (-not $EligiblePolicySection.Bounds.IntersectsWith($ActivePolicySection.Bounds)) 'Eligible and active policy sections must be separate.'
    foreach ($Section in @($EligiblePolicySection, $ActivePolicySection, $ActivePolicyBody))
    {
        Assert-True ($Section.Parent.ClientRectangle.Contains($Section.Bounds)) 'Policy section exceeds its container.'
        $Controls = @($Section.Controls)
        for ($Index = 0; $Index -lt $Controls.Count; $Index++)
        {
            $Control = $Controls[$Index]
            Assert-True ($Section.ClientRectangle.Contains($Control.Bounds)) "Policy control is clipped: $($Control.Text)."
            for ($Other = $Index + 1; $Other -lt $Controls.Count; $Other++)
            {
                Assert-True (-not $Control.Bounds.IntersectsWith($Controls[$Other].Bounds)) 'Policy controls must not overlap.'
            }
            if ($Control -is [System.Windows.Forms.CheckBox])
            {
                $TextSize = [System.Windows.Forms.TextRenderer]::MeasureText($Control.Text, $Control.Font)
                Assert-True ($TextSize.Width + 24 -le $Control.Width) "Policy checkbox text is clipped: $($Control.Text)."
            }
        }
    }
    Assert-True ($ActivePolicyBody.BackColor.R -gt $ActivePolicyBody.BackColor.G -and
        $Script:AllowPermanentActiveBox.Text -match 'NOT RECOMMENDED') 'Active policy must have a red tint and a permanent-assignment warning.'
    Assert-True ($ActivePolicySection.DisplayRectangle.Contains($ActivePolicyBody.Bounds) -and
        $ActivePolicySection.BackColor.ToArgb() -eq $Script:PimGroupBox.BackColor.ToArgb()) 'Active tint must stay inside the outline without tinting the title area.'
    Assert-True ($CreateTab.AutoScrollMinSize.Height -ge $CreateHint.Bottom -and
        $Script:EligibleMembersBox.Top -gt $Script:PimGroupBox.Bottom) 'Eligible-member input and creation controls must remain accessible below policy settings.'
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = 'Bulk create'
    $Script:TabControl.TabPages.Add($Tab)
    $Start = $Source.IndexOf('$ReportTab = New-Object System.Windows.Forms.TabPage')
    $End = $Source.IndexOf('# ^^^^^ Assignment Reports Tab ^^^^^', $Start)
    . ([scriptblock]::Create($Source.Substring($Start, $End - $Start)))
    $Script:TabControl.SelectedTab = $ReportTab
    [void]$Script:MainForm.Handle
    [void]$Script:TabControl.Handle
    [void]$ReportTab.Handle
    foreach ($Size in @(
        (New-Object System.Drawing.Size(1040, 970)),
        $Script:MainForm.MinimumSize,
        (New-Object System.Drawing.Size(1400, 1000))))
    {
        $Script:MainForm.Size = $Size
        $Script:MainForm.PerformLayout()
        $Script:TabControl.PerformLayout()
        $ReportTab.PerformLayout()
        $ReportLayout.PerformLayout()
        $ReportHeader.PerformLayout()
        $ReportFooter.PerformLayout()
        Assert-True ($ReportTab.ClientRectangle.Contains($ReportLayout.Bounds)) "Report layout exceeds the tab at $Size."
        Assert-True ($ReportLayout.ClientRectangle.Contains($ReportFooter.Bounds)) "Export footer is clipped at $Size."
        Assert-True ($ReportFooter.ClientRectangle.Contains($Script:ReportExportButton.Bounds)) "Export button is clipped at $Size."
        Assert-True (-not $ReportFooter.Bounds.IntersectsWith($Script:ReportGrid.Bounds)) "Results overlap the footer at $Size."
        Assert-True ($Script:ReportGrid.Height -ge 100) "Results have insufficient height at $Size."
        $LastTabBounds = $Script:TabControl.GetTabRect(2)
        Assert-True ($LastTabBounds.Right -le $Script:TabControl.Width) "Larger tabs exceed the available width at $Size."
    }

    $DrawCall = $Ast.Find({
        param($Node)
        $Node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $Node.Member.Value -eq 'Add_DrawItem'
    }, $true)
    $DrawHandler = [scriptblock]::Create($DrawCall.Arguments[0].ScriptBlock.Extent.Text.TrimStart('{').TrimEnd('}'))
    $Script:TabControl.SelectedIndex = 0
    $Colors = @()
    foreach ($Index in @(0, 1))
    {
        $Bitmap = New-Object System.Drawing.Bitmap(180, 40)
        $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
        try
        {
            $Bounds = New-Object System.Drawing.Rectangle(0, 0, 180, 40)
            $EventArgs = New-Object System.Windows.Forms.DrawItemEventArgs($Graphics, $Script:MainForm.Font, $Bounds, $Index, [System.Windows.Forms.DrawItemState]::None)
            & $DrawHandler $Script:TabControl $EventArgs
            $Colors += $Bitmap.GetPixel(2, 2)
        }
        finally
        {
            $Graphics.Dispose()
            $Bitmap.Dispose()
        }
    }
    Assert-True ($Colors[0].ToArgb() -ne $Colors[1].ToArgb() -and $Colors[0].B -gt $Colors[0].R) 'Selected tab must render a distinct blue background.'
}
finally { $Script:MainForm.Dispose(); $Script:MainForm = $null }
'GUI layout tests passed.'
