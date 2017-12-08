# Log Writing Function
Function LogWrite
{
    Param ([string]$logstring)

    $now = Get-Date -Format "MM-dd-yyyy-hh-mm"
    $logstring = "$($now) :   $logstring"
    Add-content $Logfile -value $logstring
}

function Get-RegistryValue {
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Value
    )

        $valpath = Get-Item -Path $Path
        try{$val = $valpath.GetValue($Value)}
        catch{$val=$null}
        if($val -ne $null)
        {return $val}
        else
        {return $null}
}

# Define our script and approots for logging
$scriptroot = split-path -Parent -Path $myInvocation.MyCommand.definition
$approot = split-path -Parent -Path $scriptroot

# Check to see if we're running in a Task Sequence.
try{$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment}
catch{$tsenv = $null}

# Define our logging path and filename
if($tsenv)
{
    # Set our initial logpath to the default logpath
    $logPath = $tsenv.Value("_SMSTSLogPath")
    # Set our logfile name properties
    $now = Get-Date -Format "MM-dd-yyyy-hh-mm"
    $filename = ("$($myInvocation.MyCommand)").split('.')[0]
    $filename = "$($filename)-$now.log"
    $logFile = "$logPath\$filename"
    Logwrite "Running from Task Sequence."
    Logwrite "LogPath is set to $logpath"
    #Hide the progress dialog
    $TSProgressUI = new-object -comobject Microsoft.SMS.TSProgressUI
    $TSProgressUI.CloseProgressDialog()
}
else {
    # Create a logpath one level above the scriptroot, and define it as the logpath
    $logPath = "$approot\LogFiles"
    if(!(test-path $logPath)){new-item $logpath -type Directory -Force -ErrorAction SilentlyContinue}
    # Set our logfile name properties
    $now = Get-Date -Format "MM-dd-yyyy-hh-mm"
    $filename = ("$($myInvocation.MyCommand)").split('.')[0]
    $filename = "$($filename)-$now.log"
    $logFile = "$logPath\$filename"
    Logwrite "Not running from Task Sequence"
    Logwrite "LogPath is set to $logpath"
}

$rempol = "HKLM:\Software\Policies\Microsoft\Windows\RemovableStorageDevices"
$scope = @(gci $rempol)
$denye = "Deny_Execute"
$denyr = "Deny_Read"
$denyw = "Deny_Write"

foreach($pol in $scope)
{
    $polpath = $pol.PSPath
    $keypath = $polpath.split('::')[-1]
    $keyname = $keypath.split('\')[-1]
    Logwrite "Attempting to reset values in $keyname"

    $deny_execute = Get-RegistryValue -Path $polpath -Value $denye
    if($deny_execute -eq 1)
    {
        LogWrite "Resetting $denye in $keyname"
        try{
            set-itemproperty $polpath -Name $denye -Value 0
            $taskstat="Successfully reset $denye in $keyname"
            $returncode = 0
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
        catch{
            $taskstat=$_.Exception.Message
            $returncode = $_.Exception.HResultLogwrite
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
    }

    $deny_read = Get-RegistryValue -Path $polpath -Value $denyr
    if($deny_read -eq 1)
    {
        LogWrite "Resetting $denyr in $keyname"
        try{
            set-itemproperty $polpath -Name $denyr -Value 0
            $taskstat="Successfully reset $denyr in $keyname"
            $returncode = 0
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
        catch{
            $taskstat=$_.Exception.Message
            $returncode = $_.Exception.HResultLogwrite
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
    }
    $deny_write = Get-RegistryValue -Path $polpath -Value $denyw
    if($deny_write -eq 1)
    {
        
        LogWrite "Resetting $denyw in $keyname"
        try{
            set-itemproperty $polpath -Name $denyw -Value 0
            $taskstat="Successfully reset $denyw in $keyname"
            $returncode = 0
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
        catch{
            $taskstat=$_.Exception.Message
            $returncode = $_.Exception.HResultLogwrite
            $returnobj = new-object PSObject -Prop @{'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
    }
}

$storpol = "HKLM:\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies"
$protkey = "WriteProtect"
$exists = get-itemproperty $storpol -EA "SilentlyContinue"
if($exists)
{
    if($exists.$protkey -eq "1")
    {
        try{
            set-itemproperty $storpol  -Name $protkey -Value 0
            $taskstat="Successfully reset $protkey value"
            $returncode = 0
            $returnobj = new-object PSObject -Prop @{'KeyName'=$protkey;'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "KeyName: $($returnobj.KeyName)   TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
        catch{
            $taskstat=$_.Exception.Message
            $returncode = $_.Exception.HResultLogwrite
            $returnobj = new-object PSObject -Prop @{'KeyName'=$keyname;'TaskStatus'=$taskstat;'TaskResult'=$returncode}
            Logwrite "KeyName: $($returnobj.KeyName)   TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
        }
    }
    else
    {
        $taskstat="$protkey value was not 1"
        $returncode = 0
        $returnobj = new-object PSObject -Prop @{'KeyName'=$protkey;'TaskStatus'=$taskstat;'TaskResult'=$returncode}
        Logwrite "KeyName: $($returnobj.KeyName)   TaskStatus: $($returnobj.TaskStatus)   TaskResult: $($returnobj.TaskResult)"
    }
}

    # Bounce the storage service and stop the Portable Device Enum Service
    $storsvc = get-service storsvc
    $wpdsvc = get-service wpdbusenum
    $storsvc | Start-Service
    start-sleep -seconds 5
    $storsvc | Stop-Service
    $wdpsvc | stop-service