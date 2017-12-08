# Email_Using_Outlook.ps1

param(
    [string]$em,
    [string]$body)

# Log Writing Function
Function LogWrite
{
    Param ([string]$logstring)

    $now = Get-Date -Format "MM-dd-yyyy-hh-mm"
    $logstring = "$($now) :   $logstring`n"
    Add-content $Logfile -value $logstring
}

#Clear our existing errors for diagnostic purposes
$error.clear()

# Close Outlook to avoid COM Object timeouts
get-process outlook | stop-process -EA SilentlyContinue

# Define our script and approots for logging
$scriptroot = split-path -Parent -Path $myInvocation.MyCommand.definition
$approot = split-path -Parent -Path $scriptroot

# Check to see if we're running in a Task Sequence.
try{$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment}
catch{$tsenv = $null}

if($tsenv)
{
    #Hide the progress dialog
    $TSProgressUI = new-object -comobject Microsoft.SMS.TSProgressUI
    $TSProgressUI.CloseProgressDialog()
}
else {
    
}

# Define our logging path and filename
# Because we're running this in the user context we have to store the log somewhere that the user has access to.  Like their desktop.
$logpath = [Environment]::GetFolderPath("MyDocuments")
# Set our logfile name properties
$now = Get-Date -Format "MM-dd-yyyy-hh-mm"
$filename = ("$($myInvocation.MyCommand)").split('.')[0]
$filename = "$($filename)-$now.log"
$logFile = "$logPath\$filename"
Logwrite "LogPath is set to $logpath"
Logwrite "scriptroot is set to $scriptroot"
Logwrite "approot is set to $approot"
Logwrite "Logfile is set to $logfile"

# Define our email options.
$emailTo = $em
LogWrite "Our email recipient is $emailTo"
<#
if($bad)
{
    $body = "An issue was encountered during your data backup transaction, or it was canceled, and the data backup process was unable to complete.  Please contact your Site Support Technician.`n`nThe log files for your data backup transaction are located at $($lastlogpath).`n`nThank you for your patience.  Have a great day!"
}
else
{
    $body = "The process of backing up your User Data is complete.  Your data was backed up to the following location:`n`n`t$($USMTPath)`n`nThe log files for your data backup transaction are located at $($lastlogpath).`n`nThank you for your patience.  Have a great day!"
}
#>

$subject = "USMT User Data Capture Complete"
$jobname = "TimerJob"

# Create the scriptblock for our Timer Job
$jobsb = {
    param($toval,$bodval,$subval)
    # Create the Outlook COM Object
    try{$Outlook = New-Object -ComObject Outlook.Application}
    catch{
        $Outlook = $null
        $olhresult = $_.Exception.HResult
        $olerr = $_.Exception.Message
    }
    if($Outlook)
    {
        # Create our mail object and send
        try{$Mail = $Outlook.CreateItem(0)}
        catch{
            $Mail = $null
            $mailhresult = $_.Exception.HResult
            $mailerr = $_.Exception.Message
        }
        if($Mail)
        {
            $Mail.To = $toval
            $Mail.Subject = $subval
            $Mail.Body = $bodval
            try{$Mail.Send()}
            catch{
                $Mail = $null
                $mailhresult = $_.Exception.HResult
                $mailerr = $_.Exception.Message
            }
            if($mailerr)
            {
                $returncode = $mailhresult
                $returnobj = new-object PSObject -Prop @{'ResultMessage'=$mailerr;'ReturnCode'=$returncode}
            }
            else
            {
                $Outlook.Session.SendAndReceive($false)
                $returncode = "0"
                $returnobj = new-object PSObject -Prop @{'ResultMessage'="Email Successfully Sent";'ReturnCode'=$returncode}    
            }
        }
        else
        {
            $returncode = $mailhresult
            $returnobj = new-object PSObject -Prop @{'ResultMessage'=$mailerr;'ReturnCode'=$returncode}
        }
    }
    else {
        $returncode = $olhresult
        $returnobj = new-object PSObject -Prop @{'ResultMessage'=$olerr;'ReturnCode'=$returncode}
    }
    return $returnobj
}

# Set up the parameters for our loop.
$timeoutval = $false
$jobdone = $false
$JobLoopCount = 1
$jobstarttime = get-date
$j = Start-job -Name $jobname -ScriptBlock $jobsb -ArgumentList $em,$body,$subject
$jobtimeout = $jobstarttime.AddMinutes(1)

while ((!$timeoutval) -and (!$jobdone)) # Begin MailLoop
{
    LogWrite "JobLoop begins run $JobLoopCount"
    $jobchecktime = get-date
    $jobstat = $j.state
    $olpids = @(get-process outlook -EA SilentlyContinue | sort StartTime | select Id,StartTime)
    $lastpid = $olpids[-1].Id
    
    if($jobchecktime -lt $jobtimeout)
    {
        if($jobstat -eq "Completed")
        {
            # Either the email successfully sent, or there was an error.  Checking.
            $result = receive-job $jobname -keep
            $timeoutval=$false
            $jobdone = $true
            get-job $jobname | remove-job
            $j = $null
        }
        else
        {
            # Timer still going.
            LogWrite "The job is still running.  Waiting for 10 seconds."
        }
    }
    else
    {
        LogWrite "The job has timed out."
        # Times up
        $timeoutval=$true
        $jobdone = $true
        stop-process $lastpid -EA SilentlyContinue
        $result = receive-job $jobname -keep
        get-job $jobname | remove-job
        $j = $null
    }
    # Wait 10 seconds and check again.
    start-sleep -seconds 10
    $JobLoopCount++
}
$statmessage = $result.ResultMessage
$statcode = $result.ReturnCode
LogWrite "The job finished with Status Message: $statmessage"
Logwrite "The job finished with Result Code: $statcode"
exit $statcode