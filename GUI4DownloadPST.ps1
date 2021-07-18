[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$null = Import-Module "$PSScriptRoot\MainLogic.psm1" -DisableNameChecking | Out-Null

#main form
$form=New-Object Windows.Forms.Form
$form.Width = 408
$form.Height = 150
$form.MaximizeBox=$false
$form.MinimizeBox=$true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.Text = "Download PST 2.0"
$form.Icon = [System.Drawing.Icon]::new("$PSScriptRoot\Download_PST_icon.ico")
$form.BackColor = [System.Drawing.Color]::White

#information label
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object Drawing.Point(10, 10)
$infoLabel.BackColor = [System.Drawing.Color]::Transparent
$infoLabel.Text=""
$infoLabel.AutoSize = $true
$form.Controls.Add($infoLabel)

#start button
$startButton = New-Object System.Windows.Forms.Button
$startButton.AutoSize = $true
$startButton.Text = "Start Download"
$startButton.Font = [System.Drawing.Font]::new("Arial", 12)
$startButton.Size = [System.Drawing.Size]::new(100,30)
$startButton.UseVisualStyleBackColor = $true
$startButton.Location = [System.Drawing.Point]::new(130,35)
$startButton.Anchor = [System.Windows.Forms.AnchorStyles]::none
$startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$startButton.FlatAppearance.BorderSize = 1
$startButton.BackColor = [System.Drawing.Color]::AliceBlue
$startButton.Visible = $true

#progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Step=1
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Location = [System.Drawing.Point]::new(10,70)
$progressBar.Size = [System.Drawing.Point]::new(373,25)

#add controls to form
$form.Controls.Add($startButton)


$startButton.Add_Click({
    Start-Program
})

$User = "$env:USERNAME"
$Path = "$env:LOCALAPPDATA\Microsoft\Outlook"

$syncHash = [hashtable]::Synchronized(@{})
$syncHash.User = $User
$syncHash.Path = $Path
$syncHash.PathToModule = $PSScriptRoot
$syncHash.form = $form
$syncHash.infoLabel = $infoLabel
$syncHash.progressBar = $progressBar

function Start-Program {

$form.controls.Remove($startButton)
$form.Controls.Add($progressBar)

$Global:runspace =[runspacefactory]::CreateRunspace()
$Global:runspace.Open()
$Global:runspace.SessionStateProxy.SetVariable('syncHash',$syncHash)
$Global:powerShell = [powershell]::Create()
$Global:powerShell.runspace = $Global:runspace

$Global:powerShell.AddScript({
    #Wait-Debugger
    $Module = "$($syncHash.PathToModule)" + "\" + "MainLogic.psm1"
    Import-Module $Module
    Go
})

$Global:AsyncObject = $Global:powerShell.BeginInvoke()
}

$form.Add_Closing({
    $Global:powershell.runspace.dispose()
})

[Windows.Forms.Application]::Run($form) | Out-Null