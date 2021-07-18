function Find-PST ($User) {

$ServerAddress = 'http://mail.examlpe.com'
$CurrentYear = (Get-Date).Year
$PreviousYear = $CurrentYear - 1
$Parts = @('p1', 'p2', 'p3', 'full')

for ($i = $PreviousYear; $i -le $CurrentYear; $i++) {
    foreach ($part in $Parts) {
        $url = "$ServerAddress" + "$($i)_" + "$User" + "_$($part).pst"
        $httpRequest = [System.Net.WebRequest]::Create($url)
        try {
            $httpResponse = $httpRequest.GetResponse()
            [int]$statusCode = $httpResponse.StatusCode
            if ($statusCode -eq 200) {
                $httpResponse.Close()
                "$($i)_" + "$User" + "_$($part).pst"
                return $url} else { return $statusCode }
            } catch {
                if ($_.Exception -like "*404*") {Continue}
            }
        }
    }
    
    return $null
}

function Check-Result ($PST) {

if ($null -eq $PST) {
    [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Exclamation
    [System.Windows.Forms.MessageBoxButtons]$button = [System.Windows.Forms.MessageBoxButtons]::OK
    [System.Windows.Forms.MessageBox]::Show("Archive not found. Contact HelpDesk for support.", "Warning",$button,$icon)
    $syncHash.form.BeginInvoke([action]{$syncHash.form.Close})
    exit
}

}

function Check-OutlookAPP {

if ((Get-Process OUTLOOK -ErrorAction SilentlyContinue).Count -eq 0) {
        [string]$message = "Would you like to open Outlook in order to connect your archive?"
        [string]$caption = "Outlook"
        [System.Windows.Forms.MessageBoxButtons]$button = [System.Windows.Forms.MessageBoxButtons]::YesNo
        [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Question
        $userResponse = [System.Windows.Forms.MessageBox]::Show($message, $caption, $button, $icon)

        if ($userResponse -eq "YES") {
            Start-Process Outlook.exe
            Start-Sleep -Seconds 30
            return $true
        } else {return $false}
    }
else { return $true }
}

function Is-Downloaded ($Archive, $Path) {
    $Fullpath = "$Path" + '\' + "$Archive"
    if ((Test-Path $Fullpath)) {return $true}
    else {return $false}
}

function Add-PST ($Archive, $Path) {
    
    $PST = Get-ChildItem -Path $Path -Recurse | Where-Object { $_.Name -like "$Archive" }
    [string]$PSTPath = $PST.FullName
    [int]$Year = $($PST.Name.Remove(4))
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Namespace.AddStore($PSTPath)
    [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Information
    [System.Windows.Forms.MessageBoxButtons]$button = [System.Windows.Forms.MessageBoxButtons]::OK
    [System.Windows.Forms.MessageBox]::Show("$Archive has been added to Outlook", "Download PST 2.0", $button, $icon)
    if ($PSTPath -like "*p2.pst" -or $PSTPath -like "*p3.pst") {exit}
    $Stores = $Namespace.Stores | Where-Object { $_.ExchangeStoreType -eq 3 -and $_.IsDataFileStore -eq $true -and $_.Filepath -notlike $PSTPath -and $_.displayname.Remove(4) -like "$Year*" }
    if ($null -eq $Stores) {return}
    else {
        foreach ($Store in $Stores) {
        if ($Store.FilePath -like "*p1.pst" -or $Store.FilePath -like "*p2.pst" -or $Store.FilePath -like "*p3.pst" -or $Store.FilePath -like "*$($syncHash.User).pst") {
            $Namespace.RemoveStore($Store.GetRootFolder())
            Start-Sleep -Seconds 10
            }
        }
    }
}


function Start-BITSJob ($Link, $Path) {

    Start-BitsTransfer -Source $Link -Destination $Path -Priority High -RetryInterval 60 -Asynchronous -DisplayName "Download PST"

}

function Get-BITSJob {
    $Job = Get-BitsTransfer | where {$_.Displayname -eq "Download PST"}
    return $Job
}

function Remove-BITSJob ($JobObj) {

    Remove-BitsTransfer -BitsJob $($JobObj.JobId)
}


function Get-DownloadStatus ($Transfered, $Total) {

    return [int]$pct = ($Transfered / $Total) * 100
}

function Show-Progress ($Job) {
    [int]$currentPerc = 0
   
    while (($Job.JobState -eq "Transferring") -or ($Job.JobState -eq "Connecting")) {
        Start-Sleep -Seconds 5
        [int]$step = Get-DownloadStatus -Transfered $Job.BytesTransferred -Total $Job.BytesTotal
        $syncHash.infoLabel.Text = "Downloading $global:pstname`n`n$step"+"% completed"
        if ($step -gt $currentPerc) {
            [int]$progress = $step - $currentPerc
            [int]$currentPerc = $step
            $syncHash.progressBar.BeginInvoke([action]{$syncHash.progressBar.Increment($progress)})
        }
        
    }
    if ($Job.JobState -eq "Transferred") {
        Complete-BitsTransfer -BitsJob $Job
        }
    elseif ($Job.JobState -eq "Error") {
        [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Error
        [System.Windows.Forms.MessageBoxButtons]$button = [System.Windows.Forms.MessageBoxButtons]::OK
        [System.Windows.Forms.MessageBox]::Show("$($MyBITSjob.Jobstate) has occured! $($MyBITSjob.ErrorDescription)", "Error occured",$button,$icon)
        $syncHash.form.BeginInvoke([action]{$syncHash.form.Close})
        Remove-BITSJob -JobObj $MyBITSjob
        exit
    }

}

function Connect-PST ($PST) {

if(Check-OutlookAPP) {      
        Add-PST -Archive $PST -Path $syncHash.Path
        $syncHash.form.BeginInvoke([action]{$syncHash.form.Close})
        exit
    }
}

function Go {
    
    $PST = Find-PST -User $syncHash.User
    $MyBITSjob = Get-BITSJob
    $global:pstname = $PST[0]

    if (($MyBITSjob.JobState -eq "Transferring") -or ($MyBITSjob.JobState -eq "Connecting")) {
        Show-Progress -Job $MyBITSjob
        Connect-PST -PST $PST[0]
        }
    elseif ($MyBITSjob.JobState -eq "Transferred") {
        Complete-BitsTransfer -BitsJob $MyBITSjob
        $syncHash.progressBar.Increment(100)
        $syncHash.infoLabel.Text = "Download $global:pstname completed"
        Connect-PST -PST $PST[0]
    }
    elseif ($MyBITSjob.JobState -eq "Error") {
        [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Error
        [System.Windows.Forms.MessageBoxButtons]$button = [System.Windows.Forms.MessageBoxButtons]::OK
        [System.Windows.Forms.MessageBox]::Show("$($MyBITSjob.Jobstate) has occured! $($MyBITSjob.ErrorDescription)","Error occured",$button,$icon)
        $syncHash.form.BeginInvoke([action]{$syncHash.form.Close})
        Remove-BITSJob -JobObj $MyBITSjob
        exit
        }
    elseif ($MyBITSjob -eq $null) {
        Check-Result -PST $PST
        if ((Is-Downloaded -Archive $PST[0] -Path $syncHash.Path)) {
            $syncHash.progressBar.Increment(100)
            $syncHash.infoLabel.Text = "Download $global:pstname completed"
            Connect-PST -PST $PST[0]
        } else {
            Start-BITSJob -Link $PST[1] -Path $syncHash.Path
            $job = Get-BITSJob
            Show-Progress -Job $job
            Connect-PST -PST $PST[0]
        }
}
}