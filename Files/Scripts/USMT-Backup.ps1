# USMT-Backup.ps1

<#
.SYNOPSIS
    This script performs the silent or interactive capture of a PC through USMT.
.DESCRIPTION
    This script can be run standalone with parameters, or be run as a part of a Deployment via SCCM/PSAppDeploy Toolkit.
    This script saves the captured data to the indicated backup path, creating a new folder within named "<username>_<computername>"
    If run standalone without the Username parameter, the script will attempt to resolve the currently logged on user.
    When run via the PSAppDeploy Toolkit Executable or via SCCM, this script is run in System context by PSExec.
    If run interactively, this script requires admin rights.
    If run interactively, this script asks the user if they would like to run checkdisk, which forces a reboot.
    If run interactively, this script will prompt the user to answer the following questions:
        Is the machine shared by multiple users on a regular basis?
        Will the backup be stored on a USB or other removable drive, directly connected to this PC?
        Do you want to run CheckDisk to check the disk for errors?
    The answers to these questions will determine reboot conditions and the fields that are presented on the WinForm.
    This script also considers the possibility that the machine that the user is running on may have the USB ports restricted.
    This script runs a check for USB Removable drives, and if it detects any, it will run an accompanying script to unlock the USB ports.
    This script will prompt for reboot if the removable drives are not removed when choosing the USB drive option.
    This script presumes that your organization has email aliases set up for "username@domain.com".  If this is not the case, you will want to modify the
    lines of this script relevant to email address formatting.
    If run with the silent switch, this script presumes that the backup type is "Network" and that the user does not wish to run Checkdisk.
    This script was written to accomodate Windows 7 PC's running PowerShell v2.0
.EXAMPLE
    
    PS C:\>.\USMT-Backup.exe
    This will launch the interactive version of the script via the executable.

    PS C:\>.\USMT-Backup.exe -deploymode "Silent" -usmtpath "C:\Temp\backup"
    This run of the executable will determine the currently logged on user and back up their data to C:\Temp\Backup\<username>_<computername>, sending an email to that user.
    
    PS C:\>.\USMT-Backup.exe -deploymode "Silent" -username contoso\usera -usmtpath "C:\Temp\backup"
    This run of the executable will determine the user running the executable, back up the data for the specified user to C:\Temp\Backup\<username>_<computername>, and 
    send an email to the running user as well as the backed up user.
    
    PS C:\>.\USMT-Backup.exe -deploymode "Silent" -usmtpath "C:\Temp\backup" -emailaddress "usera@contoso.com,userb@contoso.com"
    This run of the executable will determine the currently logged on user and back up their data to C:\Temp\Backup\<username>_<computername>, sending an email to the running user as well as 
    to the email addresses provided.
    
    PS C:\>.\USMT-Backup.exe -deploymode "Silent" -shared -usmtpath "C:\Temp\backup"
    This run of the executable will determine the currently logged on user or last logged on user, back up the data for all users of the machine to C:\Temp\Backup\<computername>, sending an email to the user.
    
    PS C:\>.\USMT-Backup.exe -deploymode "Silent" -shared -usmtpath "C:\Temp\backup" -emailaddress "usera@contoso.com;userb@contoso.com" 
    This run of the executable will determine the currently logged on user or last logged on user, back up the data for all users of the machine to C:\Temp\Backup\<computername>, sending an email to the user
    as well as to the email addresses provided.
    
    PS C:\>.\USMT-Backup.ps1 -silent -usmtpath "C:\Temp\Backup"
    This run of the script will determine the currently logged on user and back up their data to C:\Temp\Backup\<username>_<computername>, sending an email to that user.

    PS C:\>.\USMT-Backup.ps1 -silent -usmtpath "C:\Temp\Backup" -emailaddress "usera@contoso.com,userb@contoso.com"
    This run of the script will determine the currently logged on user and back up their data to C:\Temp\Backup\<username>_<computername>, sending an email to that user
    as well as to the email addresses provided.

    PS C:\>.\USMT-Backup.ps1 -silent -shared -usmtpath "C:\Temp\Backup" 
    This run of the script will determine the currently logged on user or last logged on user, back up the data for all users of the machine to C:\Temp\Backup\<computername>, sending an email to the user.
    
    PS C:\>.\USMT-Backup.ps1 -silent -shared -usmtpath "C:\Temp\Backup" -emailaddress "usera@contoso.com,userb@contoso.com"
    This run of the script will determine the currently logged on user or last logged on user, back up the data for all users of the machine to C:\Temp\Backup\<computername>, sending an email to the logged
    on user as well as to the email addresses provided.
    
    PS C:\>.\USMT-Backup.ps1 -silent -usmtpath "C:\Temp\Backup" -username "contoso\usera"
    This run of the script will determine the user running the script, back up the data for the specified user to C:\Temp\Backup\<username>_<computername>, and 
    send an email to the running user as well as the backed up user.

    PS C:\>.\USMT-Backup.ps1 -silent -usmtpath "C:\Temp\Backup" -username "contoso\usera" -emailaddress "usera@contoso.com,userb@contoso.com"
    This run of the script will determine the user running the executable, back up the data for the specified user to C:\Temp\Backup\<username>_<computername>, sending an email to the running user as well as 
    to the email addresses provided.
    

.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>

param(    
    [string]$USMTPath,
    [string]$Username,
    [switch]$Shared,
    [switch]$Silent,
    [string]$emailaddress
    )

# Log Writing Function
function Start-Log {
[CmdletBinding()]
    param (
        [ValidateScript({ Split-Path $_ -Parent | Test-Path })]
    [string]$FilePath
    )
    
    try
    {
        if (!(Test-Path $FilePath))
    {
        ## Create the log file
        New-Item $FilePath -Type File | Out-Null
    }
        
    ## Set the global variable to be used as the FilePath for all subsequent Write-Log
    ## calls in this session
    $global:ScriptLogFilePath = $FilePath
    }
    catch
    {
        Write-Error $_.Exception.Message
    }
}
    
Function Write-Log {
    param (
    [Parameter(Mandatory = $true)]
    [string]$Message,
        
    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [int]$LogLevel = 1
    )

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($script:scriptname):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat
    Add-Content -Value $Line -Path $ScriptLogFilePath
}

# Get Script Directory
function Get-ScriptDirectory
{
    <#
    .SYNOPSIS
    Get-ScriptDirectory returns the proper location of the script.

    .OUTPUTS
    System.String

    .NOTES
    Returns the correct path within a packaged executable.
    #>
    [OutputType([string])]
    param ()
    if ($null -ne $hostinvocation)
    {
        Split-Path $hostinvocation.MyCommand.path
        $fnameval = $hostinvocation.MyCommand.ToString()
        $fnameval = $fnameval.split('\')[-1]
		$fnameval = $fnameval.split('.')[0]
		$script:filename = $fnameval
    }
    else
    {
        Split-Path $script:MyInvocation.MyCommand.Path
        $fnameval = $script:myinvocation.MyCommand.ToString()
		$fnameval = $fnameval.split('.')[0]
		$script:filename = $fnameval
    }
}

# Message Pop Up Function
# Function Trigger-PopUp (FTP)
function Trigger-PopUp {
    param(
        [string]$title,
        [string]$msg,
        [string]$msg2= $null,
        [string]$options= $null,
        [string]$style= $null,
        [switch]$poptimeout
        )

    # Add functions for form maintenance and timeout
    Function ClearAndClose()
    {
            $Timer.Stop(); 
            $popform.Close(); 
            $popform.Dispose();
            $Timer.Dispose();
            $Script:CountDown=30
    }
    
    Function Timer_Tick()
    {
            --$Script:CountDown
            $countnum = $script:Countdown
            if ($countnum -le 0)
            {
                    write-log "Closing Idle form to finish script processing."
                    $popbutton3.PerformClick()
                    # try{[System.Windows.Forms.SendKeys]::SendWait('~')}
                    # catch{write-log "There was an error while trying to close the Trigger Pop-up" -Loglevel 2}
            }
    }

    #----------------------------------------------
    #region Import the Assemblies
    #----------------------------------------------
    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    #endregion Import Assemblies

    #----------------------------------------------
    #region Generated Form Objects
    #----------------------------------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $popForm = New-Object 'System.Windows.Forms.Form'
    $poppic = New-Object 'System.Windows.Forms.PictureBox'
    $popLabel1 = New-Object 'System.Windows.Forms.Label'
    $popLabel2 = New-Object 'System.Windows.Forms.Label'
    $popButton1 = New-Object 'System.Windows.Forms.Button'
    $popButton2 = New-Object 'System.Windows.Forms.Button'
    $popButton3 = New-Object 'System.Windows.Forms.Button'
    $Timer = New-Object 'System.Windows.Forms.Timer'
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
    #endregion Generated Form Objects

    #----------------------------------------------
    # User Generated Script
    #----------------------------------------------
    
    # Grab the properties for our primary screen
    $primaryscreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $maxscreenheight = $primaryscreen.workingarea.height

    $popForm_Load={
        #TODO: Initialize Form Controls here
        $formmaxh = $popform.MaximumSize | select -ExpandProperty Height
        $clienth = $popform.ClientSize | select -ExpandProperty Height
        $clientw = $popform.ClientSize | select -ExpandProperty Width
        $formh = $popform.Size | select -ExpandProperty Height
        $buttonw = $popButton1.Size | select -ExpandProperty Width
        $buttonh = $popButton1.Size | select -ExpandProperty Height
        $label1h = $poplabel1.height
        $label2h = $poplabel2.height
        $popButton1.visible = $false
        $popButton2.Visible = $False
        $popButton3.Visible = $False
        $label1hpad = $label1h + 5
        $label2hpad = $label2h + 5
        $buttonwpad = $buttonw + 10
        $buttonhpad = $buttonh + 10
        $buttonx = $clientw - $buttonwpad
        $buttony = $clienth - $buttonhpad

        # Determine button captions and visibility based on option value
        switch ($options)
        {
            "AbortRetryIgnore" {
                $popButton3.Text = "Ignore"
                $popButton3.Visible = $true
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
                $popButton2.Text = "Retry"
                $popButton2.Visible = $true
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::Retry
                $popButton1.Text = "Abort"
                $popButton1.visible = $true
                $popButton1.DialogResult = [System.Windows.Forms.DialogResult]::Abort
            }
            "OK"  {
                $popButton3.Text = "OK"
                $popButton3.Visible = $true
                $popForm.AcceptButton = $popButton3
                $popButton3.TabIndex = 0
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $okbutton = $popButton3
            }
            "OKCancel" {
                $popButton3.Text = "Cancel"
                $popButton3.Visible = $true
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $popButton3.TabIndex = 1
                $popButton2.Text = "OK"
                $popButton2.Visible = $true
                $popForm.AcceptButton = $popButton2
                $popButton2.TabIndex = 0
                $okbutton = $popButton2
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
            "RetryCancel" {
                $popButton3.Text = "Cancel"
                $popButton3.Visible = $true
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $popButton3.TabIndex = 1
                $popButton2.Text = "Retry"
                $popButton2.Visible = $true
                $popForm.AcceptButton = $popButton2
                $popButton2.TabIndex = 0
                $okbutton = $popButton2
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::Retry
            }
            "YesNo" {
                $popButton3.Text = "No"
                $popButton3.Visible = $true
                $popButton3.TabIndex = 1
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::No
                $popButton2.Text = "Yes"
                $popButton2.Visible = $true
                $popButton2.TabIndex = 0
                $popForm.AcceptButton = $popButton2
                $okbutton = $popButton2
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            }
            "YesNoCancel" {
                $popButton3.Text = "Cancel"
                $popButton3.Visible = $true
                $popButton3.TabIndex = 2
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $popButton2.Text = "No"
                $popButton2.Visible = $true
                $popButton2.TabIndex = 1
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::No
                $popButton1.Text = "Yes"
                $popButton1.visible = $true
                $popButton1.TabIndex = 0
                $popForm.AcceptButton = $popButton1
                $okbutton = $popButton1
                $popButton1.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            }
            "RebootRetryCancel" {
                $popButton3.Text = "Cancel"
                $popButton3.Visible = $true
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $popButton2.Text = "Retry"
                $popButton2.Visible = $true
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::Retry
                $popButton1.Text = "Reboot"
                $popButton1.visible = $true
                $popButton1.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
            "RebootCancel" {
                $popButton3.Text = "Cancel"
                $popButton3.Visible = $true
                $popButton3.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $popButton2.Text = "Reboot"
                $popButton2.Visible = $true
                $popButton2.TabIndex = 0
                $popForm.AcceptButton = $popButton2
                # $okbutton = $popButton2
                $popButton2.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
        }

        if($msg2)
        {
            $label2top = $($poplabel1.bottom) + 5
            $poplabel2.Top = $label2top
            $label2bottom = $($popLabel2.bottom)
            $labelbottom = $label2bottom + $buttonhpad + 5
        }
        else
        {
            $label1bottom =$($poplabel1.bottom) 
            $labelbottom = $label1bottom + $buttonhpad + 5
        }

        switch ($style)
        {
            "Information" {$poppic.Image = [System.Drawing.SystemIcons]::Information}
            "Question" {$poppic.Image = [System.Drawing.SystemIcons]::Question}
            "Error" {$poppic.Image = [System.Drawing.SystemIcons]::Error}
            "Warning" {$poppic.Image = [System.Drawing.SystemIcons]::Warning}
            default {$poppic.Image = [System.Drawing.SystemIcons]::Information}
        }
        
        # We want to maintain a padding of 10 on the form.
        # If max form height is 250, and the label height is unlimited, then the form would need
        # to grow if the label height + button height is going to be greater than the client size.

        if($labelbottom -ge $clienth)
        {
            if($labelbottom -ge $maxscreenheight)
            {
                $popform.Autoscroll = $true
                $newformmaxh = $maxscreenheight
            }
            else {
                $newformmaxh = $labelbottom + $buttonhpad + 5
                if($formmaxh -lt $newformmaxh)
                {
                    $popform.Maximumsize = "450,$newformmaxh"
                    $popform.Size = "450,$newformmaxh"
                }
            }

            $popForm.Clientsize = "$clientw,$labelbottom"
            $newbuttonx = $clientw - $buttonwpad
            $newbuttony = $labelbottom - $buttonhpad
            # $popButton1.Location = "$newbuttonx,$newbuttony"
            # Position buttons based on visibility
            # popButton3 will always be visible
            $popButton3.Location = "$newbuttonx,$newbuttony"
            # popButton2 will be visible in every option set except for "OK"
            $button2x = $popButton3.Left - $buttonwpad
            if($popButton2.visible) # Needs to be equally spaced from left edge of right most button
            {
                $popButton2.Location = "$button2x,$newbuttony"
                $button1x = $popButton2.Left - $buttonwpad
                if($popButton1.visible)
                {
                    $popButton1.Location = "$button1x,$newbuttony"            
                }
            }
        }
        else
        {
            # Form contents don't require stretching, which means default positions are okay.
            $popButton3.Location = "$buttonx,$buttony"
            # popButton2 will be visible in every option set except for "OK"
            $button2x = $popButton3.Left - $buttonwpad
            if($popButton2.visible) # Needs to be equally spaced from left edge of right most button
            {
                $popButton2.Location = "$button2x,$buttony"
                $button1x = $popButton2.Left - $buttonwpad
                if($popButton1.visible)
                {
                    $popButton1.Location = "$button1x,$buttony"       
                }
            }
        }

        # Start our timer, if we asked for one.
        if($poptimeout)
        {
            $Timer.Start()
        }
    }
    
    # --End User Generated Script--
    #----------------------------------------------
    #region Generated Events
    #----------------------------------------------
    
    $Form_StateCorrection_Load=
    {
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $popForm.WindowState = $InitialFormWindowState
    }
    
    $Form_Cleanup_FormClosed=
    {
        #Remove all event handlers from the controls
        try
        {
            $popForm.remove_Load($popForm_Load)
            $popForm.remove_Load($Form_StateCorrection_Load)
            $popForm.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
        ClearAndClose
    }

    $popForm.SuspendLayout()
    #
    # popForm
    #
    $popForm.Controls.Add($poppic)
    $popForm.Controls.Add($popLabel1)
    if($msg2)
    {
        $popForm.Controls.Add($popLabel2)
    }
    $popForm.Controls.Add($popButton1)
    $popForm.Controls.Add($popButton2)
    $popForm.Controls.Add($popButton3)
    $popForm.AutoScaleDimensions = '6, 13'
    $popForm.AutoScaleMode = 'Font'
    $popForm.AutoSize = $True
    $popform.AutosizeMode = "GrowOnly"
    $popform.ClientSize = '425,138'
    $popForm.FormBorderStyle = 'FixedToolWindow'
    $popform.MinimumSize = '425,138'
    $popForm.MaximumSize = '450, 250'
    $popForm.Name = 'popForm'
    $popForm.Padding = 5
    $popForm.ShowInTaskbar = $False
    $popForm.StartPosition = 'CenterScreen'
    $popForm.TopMost = $true
    $popForm.Text = $title
    $popForm.KeyPreview = $True
    $popForm.add_Load($popForm_Load)
    #
    # poppic
    #
    $poppic.Location = '20, 20'
    $poppic.Name = 'poppic'
    $poppic.Size = '32,32'
    $poppic.TabIndex = 3
    $poppic.TabStop = $False
    #
    # popLabel
    #
    $popLabel1.AutoSize = $True
    $popLabel1.Font = 'Microsoft Sans Serif, 9pt'
    $popLabel1.Location = '57,20'
    $popLabel1.Name = 'popLabel1'
    $poplabel1.Maximumsize = "350,0"
    $popLabel1.Text = $msg
    $poplabel1.Margin = '0,0,0,0'
    $poplabel1.Anchor = "Top,Left"
    $popLabel1.UseCompatibleTextRendering = $True
    #
    # popLabel2
    #
    $popLabel2.AutoSize = $True
    $popLabel2.Font = 'Microsoft Sans Serif, 9pt, style=Bold'
    $popLabel2.Location = '57,20'
    $popLabel2.Name = 'popLabel2'
    $popLabel2.Maximumsize = "350,0"
    $popLabel2.Text = $msg2
    $popLabel2.Margin = '0,0,0,0'
    $popLabel2.Anchor = "Top,Left"
    $popLabel2.UseCompatibleTextRendering = $True
    #
    # popButton1
    #
    $popButton1.Anchor = "Top,Left"
    $popButton1.Name = 'popButton1'
    $popButton1.Size = '83, 23'
    $popButton1.UseCompatibleTextRendering = $True
    $popButton1.UseVisualStyleBackColor = $True
    #
    # popButton2
    #
    $popButton2.Anchor = "Top,Left"
    $popButton2.Name = 'popButton2'
    $popButton2.Size = '83, 23'
    $popButton2.UseCompatibleTextRendering = $True
    $popButton2.UseVisualStyleBackColor = $True
    #
    # popButton3
    #
    $popButton3.Anchor = "Top,Left"
    $popButton3.Name = 'popButton3'
    $popButton3.Size = '83, 23'
    $popButton3.UseCompatibleTextRendering = $True
    $popButton3.UseVisualStyleBackColor = $True
    #
    # Timer
    #
    $Timer.Interval = 1000
    $Timer.Add_Tick({ Timer_Tick})
    $script:Countdown = 30

    # Resume Form Layout
    $popForm.ResumeLayout()

    #Save the initial state of the form
    $InitialFormWindowState = $popForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $popForm.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $popForm.add_FormClosed($Form_Cleanup_FormClosed)
    #Add shown handler for focus
    $popform.Add_Shown({$popform.Activate()})
    #Show the Form
    $popForm.ShowDialog()

} #End Function

# Function Kill-AppList (FKA)
function Kill-AppList
{
    # Close all the running background apps that we're backing stuff up for
    $applist = ("communicator,lync,lynchtmlconv,onedrive,groove,AeXAgentUIHost,AeXAuditPls,AeXInvSoln,AeXNSAgent,alg,AppleMobileDeviceService,igfxpers,igfxsrvc,igfxtray,OUTLOOK,picpick,googletalk").split(',')
    foreach($app in $applist)
    {
        try{$proc=get-process $app -EA SilentlyContinue}
        catch{$proc= $null}
        if($proc)
        {
            $proc.kill()
        }
    }
}

# Function Reload-WinForm (FRW)
function Reload-WinForm
{
    $script:WinForm.Close()
    $script:WinForm.Dispose()
    $script:TSBP= $null
    $script:USMTPath  = $null
    $script:formexit = $false
    $script:confirmed = $false
    Run-WinForm
}

function Confirm-Cancel
{
    # Make sure that the user really wants to cancel.
    Write-Log "Prompting for Cancel Confirmation."
    $title = "Monsanto User Data Backup - Cancel?"
    $options = "YesNo"
    $style = "Question"
    $message = "You have chosen to Cancel.  Canceling now may result in additional reboots if you choose to run the Backup tool again.`n`nClick Yes to Cancel, or No to Retry." 
    $cancelbox = Trigger-PopUp -title $title -msg $message -options $options -style $style
    switch ($cancelbox)
    {
        "Yes" {
            Write-Log "The Operation was canceled."
            $script:TSBP = $null
            $script:USMTPath = $null
            $script:formexit = $true
            $script:WinForm.Close()
            $script:WinForm.Dispose()
            Final-Cleanup -bad -silent -cancel
        }
        "No" {
            Write-Log "User chose to retry"
            Reload-WinForm
            $script:confirmed = $false
        }
    }
}

# Function Get-USMTStatus (FGU)
function Run-USMT
{
    param($joblist)

    # Define internal function variables
    $today = get-date -format yyyy-MM-dd
    $incount = $joblist.count

    if(!$script:silent)
    {
        # Import Assemblies
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Data")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")

        [System.Windows.Forms.Application]::EnableVisualStyles()
        $formJobProgress = New-Object 'System.Windows.Forms.Form'
        $progressbar1 = New-Object 'System.Windows.Forms.ProgressBar'
        $PhaseLabel = New-Object 'System.Windows.Forms.Label'
        $timerJobTracker = New-Object 'System.Windows.Forms.Timer'
        $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

        $formJobProgress_Load={
            # Start the Job on load.
            if($incount -gt 1)
            {
                for($x=0;$x -lt $incount;$x++)
                {
                    $jobin = $joblist[$x].JobInput
                    $jname = $joblist[$x].JobName
                    $jobsb = $joblist[$x].Jobscript
                    $upsb = $joblist[$x].UpdateScript
                    $comsb = $joblist[$x].CompletedScript

                    Add-JobTracker -Name $jname -JobScript $jobsb -UpdateScript $upsb -CompletedScript $comsb -ArgumentList $jobin
                }
            }
            else {
                $jobin = $joblist.JobInput
                $jname = $joblist.JobName
                $jobsb = $joblist.Jobscript
                $upsb = $joblist.UpdateScript
                $comsb = $joblist.CompletedScript
                Add-JobTracker -Name $jname -JobScript $jobsb -UpdateScript $upsb -CompletedScript $comsb -ArgumentList $jobin
            }
        }
        
        $formMain_FormClosed=[System.Windows.Forms.FormClosedEventHandler]{
            # Stop any pending jobs
            Stop-JobTracker
        }
        
        $timerJobTracker_Tick={
            Update-JobTracker
        }

        $JobTrackerList = New-Object System.Collections.ArrayList
        function Add-JobTracker
        {
            Param(
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [string]$Name, 
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [ScriptBlock]$JobScript,
            $ArgumentList = $null,
            [ScriptBlock]$CompletedScript,
            [ScriptBlock]$UpdateScript)
            
            write-log "Adding Job $Name to Tracker List."
            $jtype = $argumentlist.JobType
            # Start the Job
            switch($jtype)
            {
                "Scanstate" {
                    $sroot = $argumentlist.Scriptroot
                    $usmtp = $argumentlist.USMTPath
                    $lpath = $argumentlist.Logpath
                    $time = $argumentlist.Time
                    $plogpath = $argumentlist.Progpath
                    $username = $argumentlist.Username
                    $largline = $ArgumentList.LastArgline
                    if(($sroot) -and ($usmtp) -and ($lpath) -and ($time) -and ($plogpath) -and ($username) -and ($largline))
                    {
                        write-log "Starting the Scanstate Job"
                        $job = Start-Job -Name $Name -ScriptBlock $jobscript -ArgumentList $sroot,$usmtp,$lpath,$time,$plogpath,$username,$largline
                    }
                    else {
                        write-log "One of the following values are null:`nsroot is $sroot`nusmtp is $usmtp`nlpath is $lpath`ntime is $time`nplogpath is $plogpath`nusername is $username`nlargline is $largline"
                        Final-Cleanup -bad -Username $username
                    }
                }
                "ProgressBar" {
                    $fpath = $argumentlist.Filepath
                    if(!$fpath)
                    {
                        write-log "No Filepath was passed to the progress bar job."
                        Stop-Process scanstate -Force -EA SilentlyContinue
                    }
                    else {
                        write-log "Starting the Progress Bar Job"
                        $job = Start-Job -Name $Name -ScriptBlock $JobScript -ArgumentList $fpath   
                    }
                }

            }
            
            if($job -ne $null)
            {
                # Create a Custom Object to keep track of the Job & Script Blocks
                $psobject = new-object PSObject -Prop @{'Job'=$job;'CompleteScript'=$CompletedScript;'UpdateScript'=$UpdateScript}

                [void]$JobTrackerList.Add($psObject)	
                
                # Start the Timer
                if(-not $timerJobTracker.Enabled)
                {
                    $timerJobTracker.Start()
                }
            }
            elseif($CompletedScript -ne $null)
            {
                # Failed
                write-log "$Name failed, invoking Complete Block."
                Invoke-Command -ScriptBlock $CompletedScript -ArgumentList $null
            }
        }
        function Update-JobTracker
        {
            # Poll the jobs for status updates
            $timerJobTracker.Stop() # Freeze the Timer
            $JobCount = $JobTrackerList.count
            for($index =0; $index -lt $JobCount; $index++)
            {
                $psObject = $JobTrackerList[$index]
                $Jname = $psObject.Job.Name            
                if($psObject -ne $null) 
                {
                    if($psObject.Job -ne $null)
                    {
                        if($psObject.Job.State -ne "Running")
                        {				
                            # Call the Complete Script Block
                            write-log "$Jname is no longer running."
                            if($psObject.CompleteScript -ne $null)
                            {
                                Invoke-Command -ScriptBlock $psObject.CompleteScript -ArgumentList $psObject.Job
                            }
                            
                            $JobTrackerList.RemoveAt($index)
                            Remove-Job -Job $psObject.Job
                            $index-- # Step back so we don't skip a job
                        }
                        elseif($psObject.UpdateScript -ne $null)
                        {
                            # Call the Update Script Block
                            Invoke-Command -ScriptBlock $psObject.UpdateScript -ArgumentList $psObject.Job
                        }
                    }
                }
                else
                {
                    $JobTrackerList.RemoveAt($index)
                    $index-- # Step back so we don't skip a job
                }
            }
            
            if($JobTrackerList.Count -gt 0)
            {
                $timerJobTracker.Start()# Resume the timer	
            }	
        }
        function Stop-JobTracker
        {
            # Stop the timer
            $timerJobTracker.Stop()
            
            # Remove all the jobs
            while($JobTrackerList.Count -gt 0)
            {
                $job = $JobTrackerList[0].Job
                $JobTrackerList.RemoveAt(0)
                Stop-Job $job
                Remove-Job $job
            }
        }
        
        # Customizations

        # Form state definitions and load/unload events
        $Form_StateCorrection_Load=
        {
            # Correct the initial state of the form to prevent the .Net maximized form issue
            $formJobProgress.WindowState = $InitialFormWindowState
        }
        
        $Form_Cleanup_FormClosed=
        {
            # Remove all event handlers from the controls
            try
            {
                $formJobProgress.remove_FormClosed($formMain_FormClosed)
                $formJobProgress.remove_Load($formJobProgress_Load)
                $timerJobTracker.remove_Tick($timerJobTracker_Tick)
                $formJobProgress.remove_Load($Form_StateCorrection_Load)
                $formJobProgress.remove_FormClosed($Form_Cleanup_FormClosed)
            }
            catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
        }

        # Suspend the layout for the form
        $formJobProgress.SuspendLayout()

        # Form customizations
        $formJobProgress.Controls.Add($progressbar1)
        $formJobProgress.Controls.Add($phaselabel)
        $formJobProgress.AutoScaleDimensions = '6, 13'
        $formJobProgress.AutoScaleMode = 'Font'
        $formJobProgress.ClientSize = '284, 81'
        $formJobProgress.FormBorderStyle = 'FixedDialog'
        $formJobProgress.MaximizeBox = $False
        $formJobProgress.MinimizeBox = $False
        $formJobProgress.Name = 'formJobProgress'
        $formJobProgress.StartPosition = 'CenterScreen'
        $formJobProgress.Text = 'Job Progress'
        $formJobProgress.add_FormClosed($formMain_FormClosed)
        $formJobProgress.add_Load($formJobProgress_Load)

        # Phase Label customizations
        $PhaseLabel.Text = "Starting. Please wait ... "
        $PhaseLabel.Size = '260,23'
        $PhaseLabel.Location = '12,12'
        
        # progressbar1 customizations
        $progressbar1.Anchor = 'Top, Left, Right'
        $progressbar1.Location = '12, 46'
        $progressbar1.Name = 'progressbar1'
        $progressbar1.Size = '260, 23'
        $progressbar1.TabIndex = 1
        $progressbar1.Style = 'Continuous'
        
        # timerJobTracker customizations
        $timerJobTracker.add_Tick($timerJobTracker_Tick)
        
        # Resume layout for the form
        $formJobProgress.ResumeLayout()
        
        # Save the initial state of the form
        $InitialFormWindowState = $formJobProgress.WindowState
        # Init the OnLoad event to correct the initial state of the form
        $formJobProgress.add_Load($Form_StateCorrection_Load)
        # Clean up the control events
        $formJobProgress.add_FormClosed($Form_Cleanup_FormClosed)
        # Show the Form
        return $formJobProgress.ShowDialog()
    }
    else {
        write-log "No invoking Progress Bar Form"
        $jobin = $joblist.JobInput
        if($jobin)
        {
            $jname = $joblist.JobName
            $jobsb = $joblist.Jobscript
            $comsb = $joblist.CompletedScript
            $sroot = $jobin.Scriptroot
            $usmtp = $jobin.USMTPath
            $lpath = $jobin.Logpath
            $time = $jobin.Time
            $plogpath = $jobin.Progpath
            $username = $jobin.Username
            $largline = $jobin.LastArgline
            write-log "Our JobName is $jname"
            write-log "Our sroot is $sroot"
            write-log "Our usmtp is $usmtp"
            write-log "Our lpath is $lpath"
            write-log "Our time is $time"
            write-log "Our plogpath is $plogpath"
            write-log "Our username is $username"
            write-log "Our lastargline is $largline"
            if(($sroot) -and ($usmtp) -and ($lpath) -and ($time) -and ($plogpath) -and ($username) -and ($largline))
            {
                write-log "Starting the Scanstate Job"
                $job = Start-Job -Name $jName -ScriptBlock $jobsb -ArgumentList $sroot,$usmtp,$lpath,$time,$plogpath,$username,$largline
            }
            else {
                write-log "One of the following values are null:`nsroot is $sroot`nusmtp is $usmtp`nlpath is $lpath`ntime is $time`nplogpath is $plogpath`nusername is $username`nlargline is $largline"
                Final-Cleanup -bad -silent -Username $username
            }
    
            # Start tracking the job
            $jobdone = $false
            if($job -ne $null)
            {
                # Begin tracking the scanstate job.
                while(!$jobdone)
                {
                    $jobstat = $job.state
                    if($jobstat -ne "Running")
                    {
                        write-log "$jname completed, invoking Complete Block."
                        Invoke-Command -ScriptBlock $comsb -ArgumentList $job
                        $jobdone = $true
                    }
                    else {
                        start-sleep -seconds 5
                    }
                }
            }
            elseif($comsb -ne $null)
            {
                # Failed
                write-log "$jName failed, invoking Complete Block."
                Invoke-Command -ScriptBlock $comsb -ArgumentList $null
            }
    
            # Check to see if we need to update the job
            $JobCount = $JobTrackerList.count
            while($Jobcount -gt 0)
            {
                write-log "Checking status of $jname"
                Update-JobTracker
            }
        }
        else {
            write-log "There was no JobInput object to pull information from."
            Final-Cleanup -bad -silent -cancel
        }
    }
}

# Function Final-Cleanup (FFC)
function Final-Cleanup
{
    param([switch]$bad = $false,[switch]$silent,[switch]$cancel,[string]$username)
    Write-Log "Beginning Final Cleanup"
    # Define our date pattern for log file search later
    $backupdate = Get-Date -f "MM-dd-yyyy"
    if($bad)
    {
        Write-Log "Operation was canceled or had failed." -LogLevel 3
        $lastlogpath = "$approot\LogFiles"
        # Check to see if we're generating a pop up
        if($silent)
        {
            Write-Log "No popup is being generated for this event."
        }
        else
        {
            Write-Log "Generating Popup Options"
            # Bad Pop Up options
            $title = "Monsanto User Data Backup - Incomplete"
            $message = "The User Data Capture Process was interrupted."
        }
    }
    else
    {
        # Define our last log path
        $lastlogPath = "$script:USMTPath\USMT LogFiles"
        # Get our Total Time for email statistic
        $plog = gc $script:progpath
        write-log "Our plog value is $script:progpath"
        $timestamp = $plog[-1].split(',')[2].TrimStart().Trim()
        $totaltime = $timestamp.split(':')
        $timestring = "The backup operation took {0} hours, {1} minutes and {2} seconds." -f $totaltime[0],$totaltime[1],$totaltime[2]
        write-log $timestring
        # Get our total file size transferred
        $fso = New-Object -comobject Scripting.FileSystemObject
        $folder = $fso.GetFolder($script:USMTPath)
        $gsize = ($folder.size)/1GB
        $msize = ($folder.size)/1MB
        if($gsize -lt 1){$totalsize = "$($msize.ToString(".00"))MB"}
        else{$totalsize = "$($gsize.ToString(".00"))GB"}
        $sizestring = "The total compressed size of the captured data was $totalsize"
        write-log $sizestring
        # Good Pop Up options
        $title = "Monsanto User Data Backup - Capture State Complete"
        $message = "The User Data Capture Process is complete."
        $message = "$message`n$timestring`n$sizestring"

        # Bounce CCMExec, because we might have rebooted the machine outside of SCCM while (maybe) running a Task Sequence.
        Write-Log "Bouncing CCMExec"
        try{$ccmsvc = get-service ccmexec}
        catch{$ccmsvc = $null}
        if($ccmsvc) # Test FFC2
        {
            $ccmsvc | Stop-Service
            $ccmsvc | Start-Service
        } # End of Test FFC2
    }
    if(!$silent)
    {
        $options = "OK"
        $style = "Information"
        Trigger-PopUp -title $title -msg $message -options $options -style $style -poptimeout
    }

    # Clean up the registry keys.
    Write-Log "Cleaning up registry keys."
    $usmtpol = "HKLM:\Software\Monsanto\USMT"
    Remove-Item -Path $usmtpol -Recurse -Force

    # Move the logfiles to the USMT Backup Folder
    # Had to move this up so that we could have a path to pass to the backup email task

    if(!(test-path $lastlogPath))
    {
        Write-Log "LastLogPath location does not exist, attempting to create..." -LogLevel 2
        try{$newdir=new-item $lastlogpath -type Directory -Force -ErrorAction SilentlyContinue}
        catch{$newdir = $null}
        if(!$newdir) # Test FFC4A
        {
            Write-Log "Attempting to create the new directory at $lastlogpath failed." -LogLevel 3
            Write-Log "Setting lastlogpath to $logpath"
            $lastlogPath = $logpath
        }
        else
        {
            Write-Log "Successfully created the new directory at $lastlogpath."
        } # End of Test FFC4A
    }
    if(!$cancel)
    {
        # We're not cleaning up after a canceled operation, send email.
        # Send our backup complete email
        $scriptname = "Backup_Complete_Email.ps1"
        $Scriptpath = "$script:Scriptroot\$scriptname"
        $TaskName = "Send Backup Complete Email."
        $TaskDescr = "Send our Backup Complete Email."
        $TaskCommand = "powershell.exe"
        $TaskTrigger = "8"
        if(!$bad)
        {
            # This wasn't a failed backup operation.
            if($script:sharecheck)
            {
                $body = "The process of backing up the User Data for this machine is complete.`n$timestring`n$sizestring`nThe data was backed up to the following location:`n`n$script:USMTPath`n`nThe log files for the data backup transaction are in the following location:`n`n $lastlogpath.`n`nThank you for your patience.  Have a great day!"
            }
            else {
                $body = "The process of backing up your User Data is complete.`n$timestring`n$sizestring`nYour data was backed up to the following location:`n`n$script:USMTPath`n`nThe log files for your data backup transaction are in the following location:`n`n $lastlogpath.`n`nThank you for your patience.  Have a great day!"   
            }
            Write-Log "Sending Backup Complete Email to $script:emailaddress..."
        }
        else
        {
            # This was a failed backup operation.
            if($script:sharecheck)
            {
                $body = "An issue was encountered during the user data backup transaction for this machine. The USMT Backup job started at $script:starttime and failed at $script:endtime with exit code $script:endcode.`n`nThe log files for this data backup transaction are located at $lastlogpath.`n`nThank you for your patience.  Have a great day!"
            }
            else {
                $body = "An issue was encountered during your data backup transaction. The USMT Backup job started at $script:starttime and failed at $script:endtime with exit code $script:endcode.  Please contact your Site Support Technician.`n`nThe log files for your data backup transaction are located at $lastlogpath.`n`nThank you for your patience.  Have a great day!"
            }
            Write-Log "Sending Process Aborted Email to $script:emailaddress......" -LogLevel 2
        }

        # Create and run Send Backup Email Task
        $TaskArgs =  "-WindowStyle Hidden -Executionpolicy Bypass -command `"& '$scriptpath' -em '$script:emailaddress' -body '$body'`""
        $sendemailtask = New-ScheduledTask -TaskName $TaskName -TaskDescr $TaskDescr -TaskTrigger $TaskTrigger -TaskCommand $TaskCommand -TaskArgs $TaskArgs -UserContext $username
        if(@($($sendemailtask.count)) -gt 1)
        {
            for($i=0;$i -lt $($sendemailtask.count);$i++)
            {
                $noteprops = $sendemailtask[$i] | gm -Membertype NoteProperty
                if($noteprops) # Test FFC6A
                {
                    $rnames = $noteprops | Select-Object -expand Name
                    if($rnames -contains "ReturnCode") # Test FFC6A1
                    {
                        $sendemailreturncode=$($sendemailtask[$i].ReturnCode)
                        Write-Log "The ReturnCode of the Send Email Task is $sendemailreturncode"
                    }
                    else
                    {} # End of Test FFC6A1
                }
                else
                {
                } # End of Test FFC6A
            }
            if($sendemailreturncode -eq $null) # Test FFC6B
            {
                Write-Log "The Send Email Task had no ReturnCode" -LogLevel 2
                $sendemailreturncode = "1"
            } # End of Test FFC6B
        }
        else
        {
            write-log "SendEmailTask returns $sendemailtask"
            $noteprops = $sendemailtask | gm -Membertype NoteProperty
            if($noteprops) # Test FFC6C
            {
                $sendemailreturncode = $sendemailtask.ReturnCode
            }
            else
            {
                Write-Log "Task Result gave no ReturnCode" -LogLevel 2
                $sendemailreturncode = "1"
            } # End of Test FFC6C
        }
        if($sendemailreturncode -eq 0)
        {
            Write-Log "Task created successfully."
        }
        else
        {
            Write-Log "The task may not have created successfully with a Return Code of $sendemailreturncode." -LogLevel 2
        }
        Start-Sleep -Seconds 10
        $EmailTask = (Run-SchedTask -name $TaskName)[-1]
        if($EmailTask -eq "0")
        {
            Write-Log "Setting Emailclear to True."
            $emailclear = $true
        }
        else
        {
            Write-Log "Test FFC8 returns $EmailTask for EmailTask, the Task did not run successfully." -LogLevel 3
            Write-Log "Setting Emailclear to False."
            $emailclear = $false
        }
        # Because we had to stow the backup_complete_email log file in the users documents, we'll need to enumerate that and include it.
        $defaultdocspath = "C:\Users\$($script:loguser)\Documents"
        $oneddocspath = "C:\Users\$($script:loguser)\OneDrive - Monsanto\Migrated from My PC\Documents"
        if(!(test-path $defaultdocspath))
        {
            Write-Log "The default desktop path $defaultdocspath does not exist, checking for OneDrive redirected desktop." -LogLevel 2
            if(!(test-path $oneddocspath))
            {
                Write-Log "The OneDrive redirected desktop path $oneddocspath does not exist." -LogLevel 2
            }
            else
            {
                Write-Log "The OneDrive redirected desktop path $oneddocspath exists."
                $emaillogpath = $oneddocspath
            }
        }
        else
        {
            Write-Log "The default mydocuments path $defaultdocspath exists."
            $emaillogpath = $defaultdocspath
        }
        Write-Log "Emaillogpath is set to $emaillogpath"
        # Grab the Backup Email Log first, since it should be static and quick
        $emaillogfiles = gci $emaillogpath | ? {$_.Extension -eq ".log"}
        if(!$emaillogfiles)
        {
            Write-Log "There are no log files in $emaillogpath" -LogLevel 3
        }
        $emaillogmatches = @($emaillogfiles | select -ExpandProperty Name) -like "Backup_Complete_Email-$($backupdate)*"
        if(!$emaillogmatches)
        {
            Write-Log "There are no log files in $emaillogpath that match filename Backup_Complete_Email-$($backupdate)*" -LogLevel 3
        }
        else
        {
            Write-Log "Email log file found."
            $emaillogmatches | % {Move-Item "$emaillogpath\$_" $lastlogpath}
        }
    }
    else{$emailclear = $true}
    #>
    # Kill any running task engine jobs that were left behind
    Write-Log "Cleaning up orphaned task engine processes"
    try{$tasks = get-process taskeng*}
    catch{$tasks = $null}
    if($tasks) # Test FFC9
    {
        $tasks | % {$_.kill()}
    } # End of Test FFC9
    
    # Dispose of the scheduled tasks
    $service = new-object -ComObject("Schedule.Service")
    $service.Connect()
    $rootFolder = $service.GetFolder("\")
    try{$TaskFolder = $service.GetFolder("\USMT")}
    catch{$TaskFolder = $null}
    if($TaskFolder)
    {
        Write-Log "Disposing of the Scheduled Tasks"
        $tasknames = @($TaskFolder.GetTasks($null) | select -ExpandProperty Name)
        if($emailclear) # Test FFC10
        {
            $tasknames | % {$TaskFolder.DeleteTask($_,0)}
            $rootfolder.DeleteFolder($($TaskFolder.Name),0)
        }
        else
        {
            foreach($TName in $tasknames)
            {
                if($TName -ne $TaskName) # Test FFC10A
                {
                    $TaskFolder.DeleteTask($TName,0)
                }
                else
                {
                } # End of Test FFC10A
            }
        } # End of Test FFC10
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($service) > $null
    }
    else
    {
        Write-Log "There were no scheduled tasks to remove."
    }

    if(!$cancel)
    {
        # Start scraping our log files.
        Write-Log "Moving log files to final log path $lastlogpath"
        # Define our log file names
        $LogFileNames = ("USMT-Backup,Reset-USB,Get-LastLogon,scanstate,scanstate-stdout,scanstate-stderror,scanstateprogress").split(',')
        # Move the PSAppLog if this was run from the executable
        $PSAppLog = "$logpath\PS_AppDeployToolkitMain_3.6.9_EN_01_PSAppDeployToolkit_Install.log"
        if(test-path $PSAppLog){Move-Item $PSAppLog $lastlogpath}
        else {write-log "$PSAppLog not found."}
        # Grab the log files that are in our log path
        $applogfiles = gci $logpath | ? {$_.Extension -eq ".log"}
        # Check for log files dropped in the CCM\Logs folder just in case we started from SCCM.
        $ccmlogpath = "C:\Windows\CCM\Logs"
        $ccmlogfiles = gci $ccmlogpath -recurse | ? {$_.Extension -eq ".log"}
        $applogmatches = $LogFileNames | % {($applogfiles | select -ExpandProperty Name) -like "$($_)-$backupdate*"}
        $ccmlogmatches = $LogFileNames | % {($ccmlogfiles | select -ExpandProperty Name) -like "$($_)-$backupdate*"}
        if($applogmatches){$applogmatches | % {Move-Item "$logpath\$_" $lastlogpath}}
        if($ccmlogmatches){$ccmlogmatches | % {Move-Item "$ccmlogpath\$_" $lastlogpath}}
        if(!$script:firstrun)
        {
            $logfile = "$lastlogPath\$script:postbootlogfilename"
            Start-Log $logfile
        }
        else
        {
            $logfile = "$lastlogPath\$script:filename"
            Start-Log $logfile
        }
        write-log "Log Files moved."
    }
    else {
        write-log "Operation canceled, no need to move log files."
    }
    
    if($BackupType -eq "Local")
    {
        # Reblock the USB Drives
        Write-Log "Reblocking USB."
        Reblock-USB
    }
    
    # Cleanup done.
    Write-Log "Cleanup phase complete."
    [environment]::exit(0)
} # End Final Cleanup Function (FFC)

# Function Run-SchedTask (FRS)
function Run-SchedTask
{
    param(
        [string]$name)
        # connect to the local machine. 
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa381833(v=vs.85).aspx
        $service = new-object -ComObject("Schedule.Service")
        $service.Connect()
        $rootFolder = $service.GetFolder("\")
        $TaskFolder = $service.GetFolder("\USMT")
        try{$task = $TaskFolder.GetTask($Name)}
        catch{$task = $null}
        if($task)
        {
            Write-Log "Running $Name"
            $task.Run($null)
            Write-Log "Getting Last RunTime for the task."
            $lastruntime = $task.LastRunTime
            while (!$lastruntime)
            {
                $lastruntime = $task.LastRunTime
                Write-Log "The task has not run yet."
                start-sleep -seconds 5
            }
            # Give it 15 seconds to run
            $StartTime = Get-Date
            $timeout = $false
            $timeouttime = $starttime.addminutes(1)
            $taskgo = $false
            while(!$taskgo)
            {
                $taskstate = $task.State
                switch ($taskstate)
                {
                    "0" {$taskgo = $false;Write-Log "The task has not run yet." -LogLevel 2}
                    "1" {$taskgo = $false;Write-Log "The task is Disabled." -LogLevel 3}
                    "2" {$taskgo = $false;Write-Log "The task is Queued."}
                    "3" {$taskgo = $true;Write-Log "The task is Ready."}
                    "4" {$taskgo = $false;Write-Log "The task is Running."}
                }
                $currenttime = Get-Date
                if($currenttime -gt $timeouttime)
                {
                    $timeout = $true
                    Write-Log "The Task has been running for over a minute.  Canceling." -LogLevel 3
                    $taskgo = $true
                }
                else{start-sleep -seconds 5}
            }
            if($timeout)
            {
                $taskstat = "The task has not run within the timeout value."
                $taskresult = "1"
            }
            else 
            {
                Write-Log "The Task finished, getting results."
                $taskresult = $task.LastTaskResult
                switch ($taskresult)
                {
                    "0" {$taskstat = "The task completed successfully"}
                    "267009" {
                        Write-Log "The task is still running, giving it another 30 seconds."
                        $taskstat = "The task is still running"
                        $newtimeouttime = (get-date).addseconds(30)
                        $newtimeout = $false
                        while(!$newtimeout)
                        {
                            $taskresult = $task.LastTaskResult
                            $newcurrenttime = get-date
                            if($newcurrenttime -lt $newtimeouttime)
                            {
                                if($taskresult -eq 0)
                                {
                                    break
                                }
                                else
                                {
                                    start-sleep -seconds 5
                                }
                            }
                            else
                            {
                                Write-Log "The Task has timed out.  Last Taskresult is $taskresult." -LogLevel 3
                                $newtimeout = $true
                            }
                        }
                    }
                    default {$taskstat = "The task did not complete successfully"}
                }
            }
        }
        else
        {
            $taskstat = "No such task exists"
            $taskresult = "1"
        }
        Write-Log $taskstat
        return $taskresult
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($service) > $null
}

# Function Reblock-USB (FRU)
function Reblock-USB
{
    $rempol = "HKLM:\Software\Policies\Microsoft\Windows\RemovableStorageDevices"
    $polscope = get-childitem $rempol
    $denye = "Deny_Execute"
    $denyr = "Deny_Read"
    $denyw = "Deny_Write"
    foreach($pol in $polscope)
    {
        $polpath = $pol.PSPath
        $deny_execute = Get-RegistryValue -Path $polpath -Value $denye
        if($deny_execute -eq 0){set-itemproperty $polpath -Name $denye -Value 1}
        $deny_read = Get-RegistryValue -Path $polpath -Value $denyr
        if($deny_read -eq 0){set-itemproperty $polpath -Name $denyr -Value 1}
        $deny_write = Get-RegistryValue -Path $polpath -Value $denyw
        if($deny_write -eq 0){set-itemproperty $polpath -Name $denyw -Value 1}
    }

    # Enable Block of USB Write
    $storpol = "HKLM:\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies"
    $protkey = "WriteProtect"
    $exists = Get-RegistryValue -Path $storpol -Value $protkey
    if($exists -eq 0){set-itemproperty $storpol -Name $protkey -Value 1}
    
    # Bounce the storage service and wpdbusenum service
    $storsvc = get-service storsvc
    $wpdsvc = get-service wpdbusenum
    $storsvc | Start-Service
    start-sleep -seconds 5
    $storsvc | Stop-Service
    $wdpsvc | stop-service -EA SilentlyContinue
    $wdpsvc | start-service -EA SilentlyContinue
}

# Function Unblock-USB (FUS)
function Unblock-USB
{
    # Enable External Device Writes
    $rempol = "HKLM:\Software\Policies\Microsoft\Windows\RemovableStorageDevices"
    $polscope = get-childitem $rempol
    $denye = "Deny_Execute"
    $denyr = "Deny_Read"
    $denyw = "Deny_Write"
    foreach($pol in $polscope)
    {
        $polpath = $pol.PSPath
        $deny_execute = Get-RegistryValue -Path $polpath -Value $denye
        if($deny_execute -eq 1){set-itemproperty $polpath -Name $denye -Value 0}
        $deny_read = Get-RegistryValue -Path $polpath -Value $denyr
        if($deny_read -eq 1){set-itemproperty $polpath -Name $denyr -Value 0}
        $deny_write = Get-RegistryValue -Path $polpath -Value $denyw
        if($deny_write -eq 1){set-itemproperty $polpath -Name $denyw -Value 0}
    }

    # Disable Block of USB Write
    $storpol = "HKLM:\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies"
    $protkey = "WriteProtect"
    $exists = Get-RegistryValue -Path $storpol -Value $protkey
    if($exists -eq 1){set-itemproperty $storpol -Name $protkey -Value 0}

    # Bounce the storage service and wpdbusenum service
    $storsvc = get-service storsvc
    $wpdsvc = get-service wpdbusenum
    $storsvc | Start-Service
    start-sleep -seconds 5
    $storsvc | Stop-Service
    $wdpsvc | stop-service -EA SilentlyContinue
    $wdpsvc | start-service -EA SilentlyContinue

    # Set up our scheduled task to unblock GPO
    $scriptname = "Reset-USB.ps1"
    $Scriptpath = "$script:Scriptroot\$scriptname"
    $TaskName = "Unblock USB"
    $TaskDescr = "Unblock USB"
    $TaskTrigger = "0"
    $TaskCommand = "powershell.exe"
    $TaskArgs =  "-Executionpolicy Bypass -NoProfile -File `"$scriptpath`""
    $triggerstring1 = "<QueryList><Query Id='0' Path='Microsoft-Windows-GroupPolicy/Operational'><Select Path='Microsoft-Windows-GroupPolicy/Operational'>*[System[Provider[@Name='Microsoft-Windows-GroupPolicy'] and EventID='8004']]</Select></Query></QueryList>"
    $triggerstring2 = "<QueryList><Query Id='0' Path='Microsoft-Windows-GroupPolicy/Operational'><Select Path='Microsoft-Windows-GroupPolicy/Operational'>*[System[Provider[@Name='Microsoft-Windows-GroupPolicy'] and EventID='8005']]</Select></Query></QueryList>"
    $triggerstring3 = "<QueryList><Query Id='0' Path='Microsoft-Windows-GroupPolicy/Operational'><Select Path='Microsoft-Windows-GroupPolicy/Operational'>*[System[Provider[@Name='Microsoft-Windows-GroupPolicy'] and EventID='8001']]</Select></Query></QueryList>"
    $triggerstring4 = "<QueryList><Query Id='0' Path='Microsoft-Windows-GroupPolicy/Operational'><Select Path='Microsoft-Windows-GroupPolicy/Operational'>*[System[Provider[@Name='Microsoft-Windows-GroupPolicy'] and EventID='8002']]</Select></Query></QueryList>"
    $eventids = ("8004,8005,8001,8002").split(',')
    $NewUnblockTask = New-ScheduledTask -TaskName $TaskName -TaskDescr $TaskDescr -TaskTrigger $TaskTrigger -TaskCommand $TaskCommand -TaskArgs $TaskArgs -EventIds $eventids -EventQueries $triggerstring1,$triggerstring2,$triggerstring3,$triggerstring4
    if(@($($NewUnblockTask.count)) -gt 1) # Test FUS1
    {
        for($i=0;$i -lt $($NewUnblockTask.count);$i++)
        {
            $noteprops = $NewUnBlockTask[$i] | gm -Membertype NoteProperty
            if($noteprops){$Unblockreturncode = $NewUnBlockTask[$i].ReturnCode}
            else{$Unblockreturncode = "An unspecified error occurred at Test FUS1"}
        }
    }
    else
    {
        $noteprops = $NewUnBlockTask | gm -Membertype NoteProperty
        if($noteprops){$Unblockreturncode = $NewUnBlockTask.ReturnCode}
        else{$Unblockreturncode = "An unspecified error occurred at Test FUS1"}
    } # End of Test FUS1
    
    if($Unblockreturncode -eq "0") # Test FUS2
    {
        Write-Log "$taskname scheduled task was completed successfully, running scheduled task."
        $RunUnblockTask = (Run-SchedTask -name $TaskName)[-1]
        if($RunUnblockTask -eq "0") # Test FUS2A
        {
            Write-Log "$taskname ran successfully with result $RunUnBlockTask"
            return $RununBlockTask
            Write-Log "Setting localchkgo to True"
            $script:localchkgo = $true
        }
        else
        {
            Write-Log "Test FUS2A verifies that scheduled Task $taskname was unsuccessful with result $RunUnBlockTask" -LogLevel 2
            Write-Log "Setting RebootNeeded to True."
            $script:USBRebootNeeded = $true
            return "1"
        } # End of Test FUS2A
    }
    else
    {
        Write-Log "Test SB3L2A2B verifies that the $taskname scheduled task creation returned $RunUnblockTask"
        return "1"
    }
}

# Function Drive-Check (FDC)
function Drive-Check {
    Write-Log "Prompting for disk insertion."
    $title = "Monsanto User Data Backup - Insert USB Drive"
    $options = "OK"
    $style = "Information"
    $message = "Temporary Access to removable USB drives has been granted.  Please insert your removable USB Drive now.  Detecting your drive may take up to a minute."
    Trigger-PopUp -title $title -msg $message -options $options -style $style
    $jobsb = {
        start-sleep -Seconds 60
    }
    # Set up the parameters for our loop.
    $timeoutval = $false
    $usbchkval = $false
    $jobname = "TimerJob"
    $jobstarttime = get-date
    $startcheck = $jobstarttime.AddSeconds(30)
    $j = Start-job -Name $jobname -ScriptBlock $jobsb
    while ((!$timeoutval) -and (!$usbchkval)) # Begin Loop FDCL1
    {
        $jobstat = $j.state
        if($jobstat -eq "Completed") # Begin Test FDCL1A
        {
            Write-Log "Job Completed.  Checking for drives one last time."
            $timeoutval = $true
            $script:USBRebootNeeded = $true
            get-job $jobname | remove-job
            $j = $null
            # Check one more time to make sure that the drive isn't available
            $ldletters = @(gwmi win32_logicaldisk | ? {$_.DriveType -eq 2} | select -ExpandProperty DeviceID)
            foreach($letter in $ldletters)
            {
                try{$drive = gwmi win32_volume | ? {$_.DriveLetter -eq $letter}}
                catch{$drive = $null}
                if($drive)
                {
                    Write-Log "Drive $letter is available.  Moving forward."
                    Write-Log "Bouncing the WPDBusEnum Service to allow access to the user"
                    Get-Service WPDBusEnum | stop-service -EA SilentlyContinue
                    Get-Service WPDBusEnum | start-service -EA SilentlyContinue
                    $usbchkval = $true
                    $script:localchkgo = $true
                    $script:USBRebootNeeded = $false
                    stop-job $jobname
                    get-job $jobname | remove-job
                    $j = $null
                    break
                }
            }
        }
        else
        {
                # Timer still going, check to see if the drive has become available.
                $getcurrenttime = get-date
                if($getcurrenttime -lt $startcheck) # Test FDCL1B
                {
                    # Timer Job still running.
                    Write-Log "Timer job still running, waiting for 30 seconds from start time to allow drive access refresh."
                }
                else
                {
                    Write-Log "It's been over 30 seconds since unblock.  Checking for drives."    
                    $ldletters = @(gwmi win32_logicaldisk | ? {$_.DriveType -eq 2} | select -ExpandProperty DeviceID)
                    foreach($letter in $ldletters)
                    {
                        try{$drive = gwmi win32_volume | ? {$_.DriveLetter -eq $letter}}
                        catch{$drive = $null}
                        if($drive)
                        {
                            Write-Log "Drive $letter is available.  Moving forward."
                            Write-Log "Bouncing the WPDBusEnum Service to allow access to the user"
                            Get-Service WPDBusEnum | stop-service -EA SilentlyContinue
                            Get-Service WPDBusEnum | start-service -EA SilentlyContinue
                            $usbchkval = $true
                            $script:localchkgo = $true
                            $script:USBRebootNeeded = $false
                            stop-job $jobname
                            get-job $jobname | remove-job
                            $j = $null
                            break
                        }
                    }
                }
        }
        # Wait 10 seconds and check again.
        start-sleep -seconds 10
    } # End of Loop FDCL1
}

# Function Get-RegistryValue (FGR)
function Get-RegistryValue {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Value
    )

    $valpath = Get-Item -Path $Path
    try{$val = $valpath.GetValue($Value)}
    catch{$val = $null}
    if($val -ne $null)
    {return $val}
    else
    {return $null}
}

# Function Create-RegistryValue (FCR)
function Create-RegistryValue {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Name,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Value
    )
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType String -Force | out-null
}

# Function New-ScheduledTask (FNS)
function New-ScheduledTask {
    param (
        [Parameter()]
	    [string]$TaskName,
        [Parameter()]
        [ValidateScript({Test-Path -Path $_ -PathType 'Leaf' })]
        [string]$FilePath,
        [Parameter()]
        [string]$TaskDescr,
        [Parameter()]
	    [string]$TaskTrigger = "8",
        [Parameter()]
	    $EventIds,
        [Parameter()]
        $EventQueries,
        [Parameter()]
	    [string]$TaskCommand,
        [Parameter()]
        [string]$TaskArgs,
        [Parameter()]
        [string]$UserContext = $null
    )
    write-log "Creating Task $TaskName"
    try{
        # attach the Task Scheduler com object
        $service = new-object -ComObject("Schedule.Service")
        # connect to the local machine. 
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa381833(v=vs.85).aspx
        $service.Connect()
        $rootFolder = $service.GetFolder("\")
        try{$TaskFolder = $service.GetFolder("\USMT")}
        catch{
            $rootFolder.CreateFolder("USMT")
            $TaskFolder = $service.GetFolder("\USMT")
        }
        $TaskDefinition = $service.NewTask(0)
        $TaskDefinition.RegistrationInfo.Description = "$TaskDescr"
        $TaskDefinition.Settings.Enabled = $true
        $TaskDefinition.Settings.AllowDemandStart = $true
        $TaskDefinition.Settings.StartWhenAvailable = $true
        $TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
        $TaskDefinition.Settings.ExecutionTimeLimit = "PT1H"
        $TaskEndTime = [datetime]::Now.AddMinutes(30)
        $triggers = $TaskDefinition.Triggers
        #http://msdn.microsoft.com/en-us/library/windows/desktop/aa383915(v=vs.85).aspx
        # $trigger = $triggers.Create(8) # Creates a "Boot" trigger
        switch ($tasktrigger)
        {
            default {
                        $trigger = $triggers.Create($TaskTrigger) # Creates a custom trigger
                        $Trigger.EndBoundary = $TaskEndTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
                        $trigger.Enabled = $true
                        $taskgo = $true
                    }
            "8" {
                # Using a boot trigger
                $trigger = $triggers.Create($TaskTrigger) # Creates a custom trigger
                $Trigger.EndBoundary = $TaskEndTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
                $trigger.Enabled = $true
                $taskgo = $true
            }
            "9" {
                    # Using a Logon Trigger
                    $trigger = $triggers.Create($TaskTrigger) # Creates a custom trigger
                    $Trigger.EndBoundary = $TaskEndTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
                    $trigger.Enabled = $true
                    $Trigger.UserID = $usercontext
                    $taskgo = $true
                }
            "0" {
                    # Using a combination of event, session state, and startup
                    $TaskDefinition.Settings.RestartCount = 3
                    $TaskDefinition.Settings.RestartInterval = "PT1M"
                    $triggers.Create("8")
                    $triggers.Create("11")
                    if((!$EventIds) -or (!$EventQueries))
                    {
                        $taskgo = $false
                        write-error "You must provide EventIds and EventQueries for this Trigger Type."
                        write-log "You must provide EventIds and EventQueries for this Trigger Type." -LogLevel 3
                        $taskstat = "You must provide EventIds and EventQueries for this Trigger Type."
                        $returncode = 1
                    }
                    else
                    {
                        if($eventids.count -ne $eventqueries.count)
                        {
                            $taskgo = $false
                            write-error "You must provide an EventIds for each EventQuery for this Trigger Type."
                            write-log "You must provide an EventIds for each EventQuery for this Trigger Type." -LogLevel 3
                            $taskstat = "You must provide an EventIds for each EventQuery for this Trigger Type."
                            $returncode = 1
                        }
                        else
                        {
                            $taski = $eventids.count
                            $vars = @{}
                            $taskgo = $true
                            if($taski -gt 1)
                            {
                                for($i=0;$i -lt $taski;$i++)
                                {
                                    $vars["trigger$($i+1)"] = $triggers.Create($TaskTrigger) # Creates a custom trigger
                                    $vars["trigger$($i+1)"].Enabled = $true
                                    $vars["trigger$($i+1)"].Id = $eventids[$i]
                                    $vars["trigger$($i+1)"].Subscription = $eventqueries[$i]
                                    # $vars["trigger$($i+1)"].EndBoundary = $TaskEndTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
                                }
                            }
                            else
                            {
                                $trigger = $triggers.Create($TaskTrigger) # Creates a custom trigger
                                $trigger.enabled = $true
                                $trigger.Id = $eventids
                                $trigger.Subscription = $eventqueries
                                # $trigger.EndBoundary = $TaskEndTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
                            }
                        }
                    }
                }
        }
        # http://msdn.microsoft.com/en-us/library/windows/desktop/aa381841(v=vs.85).aspx
        $Action = $TaskDefinition.Actions.Create(0)
        $action.Path = "$TaskCommand"
        $action.Arguments = "$TaskArgs"
        if($UserContext)
        {
            $un = $usercontext
            $up = $null
            $token = 3
        }
        else {
            write-log "Not using UserContext."
            $TaskDefinition.Principal.RunLevel = 1
            $un = "System"
            $up = $null
            $token = 5
        }
        if($taskgo)
        {
            try{write-log "Successfully created task for $TaskName";$taskreturn=$TaskFolder.RegisterTaskDefinition("$TaskName",$TaskDefinition,6,$un,$up,$token)}
            catch{$taskreturn=$null;$taskstat=$_.Exception.Message;$returncode = $_.Exception.HResult;write-log "Creating task $TaskName encountered error $returncode" -LogLevel 3}
        }
        else
        {
        }
        if($taskreturn)
        {
            $taskstat = "Successfully created scheduled task."
            $returncode = 0
        }
        else
        {
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($service) > $null
        }catch{$taskstat = $_.Exception.Message;$returncode = $_.Exception.HResult;write-log $taskstat}
        $returnobj = new-object PSObject -Prop @{'ResultMessage'=$taskstat;'ReturnCode'=$returncode}
        return $returnobj
}

# Function Validate-FinalPath (FVF)
function Validate-FinalPath
{
    Write-Log "Beginning Validate-FinalPath function."
    $computername = $env:computername
    # Put together our backup path:
    if($script:sharecheck -eq $true)
    {
        # This is a shared machine, back up all users logged in within the last 90 days.
        if(($script:TSBP.endswith(":\"))){$script:USMTPath = "$script:TSBP$computername"}
        else{$script:USMTPath = "$script:TSBP\$computername"}
        $message = "The data for this computer will be backed up to the following backup path:`n`n$script:USMTPath`n`nIs this correct?`n(Click Yes to confirm, No to Retry, Cancel to Exit.)"
    }
    else {
        # Not a shared machine, back up indicated user.
        if(($script:TSBP.endswith(":\"))){$script:USMTPath = "$script:TSBP$($script:TSSUN)_$computername"}
        else{$script:USMTPath = "$script:TSBP\$($script:TSSUN)_$computername"}
        $message = "The data for the selected user will be backed up to the following backup path:`n`n$script:USMTPath`n`nIs this correct?`n(Click Yes to confirm, No to Retry, Cancel to Exit.)"
    }
    # Confirm path with user.
    $title = "Backup Path Confirmation"
    $options = "YesNoCancel"
    $style = "Question"
    $confirmbox = Trigger-PopUp -Title $Title -msg $message -options $options -style $style
    if($confirmbox -eq "Yes")
    {
        # This is our USMT Path
        if((test-path $script:USMTPath) -eq $true)
        {
            # USMT Path already exists, check to make sure that it's okay to overwrite.
            Write-Log "The USMT Backup Path already exists, prompting for action..." -LogLevel 2
            $title = "USMT Backup Path already exists."
            $message = "The USMT Backup Folder already exists.  Do you want to overwrite it?`n`nClick Yes to Overwrite, No to enter a new path, or Cancel to close."
            $options = "YesNoCancel"
            $style = "Warning"
            $owbox = Trigger-PopUp -Title $Title -msg $message -options $options -style $style
            if($owbox -eq "Yes")
            {
                # Try to overwrite the existing backup path
                try{$newdir = new-item $script:USMTPath -type Directory -Force -ErrorAction SilentlyContinue}
                catch{$newdir = $null}
                if($newdir) # Test FVF1A1A
                {
                    # Successfully overwrote the existing folder, set our variables, kill the apps and move on
                    $script:confirmed = $true
                    Write-Log "Successfully created the USMT Backup Folder"
                    Write-Log "The USMT Backup Folder path is $script:USMTPath"
                    Write-Log "Tattooing our BackupPath to the registry...."
                    Create-RegistryValue -Path $RegRebootPath -Name "TSCusBackupPath" -Value $script:TSBP
                    Create-RegistryValue -Path $RegRebootPath -Name "OSDStateStorePath" -Value $script:USMTPath
                    Kill-AppList
                } # Able to overwrite the existing path
                else
                {
                    # Couldn't overwrite the existing folder.
                    Write-Log "Failed to create folder in Backup Path." -LogLevel 3
                    $title = "Failed to create folder in Backup Path."
                    $message = "The Backup Folder could not be created.  Please double check share permissions and re-run the task sequence."
                    $options = "RetryCancel"
                    $style = "Error"
                    $failbox = Trigger-PopUp -Title $Title -msg $message -options $options -style $style
                    if($failbox -eq "Retry") # Test FVF1A1A1
                    {
                        Write-Log "User chose to retry."
                        Reload-WinForm
                        $script:confirmed = $false
                    }
                    else
                    {
                        Write-Log "The user opted to cancel.  Exiting."
                        $script:formexit = $true
                        $script:confirmed = $false
                    } 
                } 
            } 
            if($owbox -eq "No") 
            {
                Write-Log "User chose to retry."
                Reload-WinForm
                $script:confirmed = $false
            } 
            if($owbox -eq "Cancel") 
            {
                Write-Log "The user opted to cancel.  Exiting."
                $script:formexit = $true
                $script:confirmed = $false
            }
        } 
        else
        {
            # The path doesn't exist, try to create it.
            try{$newdir = new-item $script:USMTPath -type Directory -Force -ErrorAction SilentlyContinue}
            catch{$newdir = $null}
            if($newdir) # Test FVF1A2
            {
                # Path could be created, set our variables, kill the apps and move on
                $script:confirmed = $true
                Create-RegistryValue -Path $RegRebootPath -Name "OSDStateStorePath" -Value $script:USMTPath
                Write-Log "The USMT Backup Folder path is $script:USMTPath"
                Write-Log "Tattooing our BackupPath to the registry...."
                Create-RegistryValue -Path $RegRebootPath -Name "TSCusBackupPath" -Value $script:TSBP
                Create-RegistryValue -Path $RegRebootPath -Name "OSDStateStorePath" -Value $script:USMTPath
                Kill-AppList
            } # Path could be created
            else
            {
                # Path couldn't be created  
                Write-Log "Failed to create folder in Backup Path." -LogLevel 3
                $title = "Failed to create folder in Backup Path."
                $message = "The Backup Folder could not be created.  Please double check share permissions and re-run the task sequence."
                $options = "RetryCancel"
                $style = "Error"
                $failbox = Trigger-PopUp -Title $Title -msg $message -options $options -style $style
                if($failbox -eq "Retry")
                {
                    Write-Log "User chose to retry."
                    Reload-WinForm
                    $script:confirmed = $false
                }
                else
                {
                    Write-Log "The user opted to cancel."
                    $script:formexit = $true
                    $script:confirmed = $false
                } 
            } 
        }
    } # User confirmed path at prompt
    if($confirmbox -eq "No")
    {
        Write-Log "User chose to retry."
        Reload-WinForm
        $script:confirmed = $false
    } # User chose to retry
    if($confirmbox -eq "Cancel") # Test FVF1
    {
        Write-Log "User chose to cancel."
        $script:formexit = $true
        $script:confirmed = $false
    }
}

# Function Run-WinForm (FWF)
function Run-WinForm
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Windows.Forms.Application]::EnableVisualStyles()
    
    # Set our initial validating conditions for formexit and confirm
    $script:formexit = $false
    $script:confirmed = $false

    # Create our Nested Functions
    # Validation Functions
    # Check to see if the text field is null or has changed.
    function Validate-Tag
    {
        Param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNull()]
            [string]$textfield,
            [string]$tagfield
            )
        if($textfield -eq $tagfield)
        {
            # Text fields are the same
            return $true
        }
        else
        {
            # Text fields are different
            return $false
        }
    }

    function Browse-ForFolder
    {
        $o = new-object -comobject Shell.Application
        $folder = $o.BrowseForFolder(0,"Select location to store user backup",4213,17)
        $fstest = $folder.self.IsFileSystem
        if($fstest -eq $true){
            $selectedfolder = $folder.self.path
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($o) > $null
        }
        else{
            $selectedfolder = ''
        }
        return $selectedfolder
    }

    # Create our objects:
    $script:WinForm = New-Object 'System.Windows.Forms.Form'
    $script:OKButton = New-Object 'System.Windows.Forms.Button'
    $script:CancelButton = New-Object 'System.Windows.Forms.Button'
    $script:BrowseButton = New-Object 'System.Windows.Forms.Button'
    $script:objpathLabel = New-Object 'System.Windows.Forms.Label'
    $script:objuserLabel = New-Object 'System.Windows.Forms.Label'
    $script:objpathTextBox = New-Object 'System.Windows.Forms.TextBox'
    $script:objuserTextBox = New-Object 'System.Windows.Forms.TextBox'
    $script:PathErrorProvider = New-Object 'System.Windows.Forms.ErrorProvider'
    $script:UserErrorProvider = New-Object 'System.Windows.Forms.ErrorProvider'
    $script:InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

    # Set up our initial Form Load Conditions
    $WinForm_Load={
        $wformh = $winform.Size | select -ExpandProperty Height
        $wformw = $winform.Size | select -ExpandProperty Width
        $wclienth = $winform.ClientSize | select -ExpandProperty Height
        $wclientw = $winform.ClientSize | select -ExpandProperty Width
        $cancelbuttonw = $cancelButton.Size | select -ExpandProperty Width
        $browsebuttonw = $browseButton.Size | select -ExpandProperty Width
        $cancelbuttonh = $cancelButton.Size | select -ExpandProperty Height
        $cancelbuttonwpad = $cancelbuttonw + 10
        $cancelbuttonhpad = $cancelbuttonh + 10
        $usertextwpad = $($objuserTextbox.width) + 20

        if($script:sharecheck -eq $false)
        {
            write-log "The user field should be visible on the form."
            $objUserLabel.visible = $True
            $objUserTextBox.visible = $true
            $userlabeltop = $($objpathTextBox.bottom) + 5
            $objUserLabel.Top = $userlabeltop
            $objUserLabel.Left = $objpathTextBox.Left
            $objUserTextBox.Top = $objUserLabel.Bottom
            $objUserTextBox.Left = $objUserLabel.Left
            $usertextbottom = $($objuserTextbox.bottom)
            $textbottom = $usertextbottom + $cancelbuttonhpad + 10
            $newformmaxh = $textbottom + $cancelbuttonhpad + 5
            $newcancelbuttonx = $usertextwpad - $cancelbuttonw
        }
        else {
            $objUserLabel.visible = $false
            $objUserTextBox.visible = $false
            $pathtextbottom = $($objPathTextBox.bottom)
            $textbottom = $pathtextbottom + $cancelbuttonhpad + 10
            $newformmaxh = $winform.Maximumsize | select -ExpandProperty Height
            $newcancelbuttonx = $browsebutton.Left - ($cancelbuttonw - $browsebuttonw)
        }

        # Declare new button y value
        $WinForm.Maximumsize = "$wformw,$newformmaxh"
        $WinForm.Size = "$wformw,$newformmaxh"
        $WinForm.ClientSize = "$wclientw,$textbottom"
        $newbuttony = $textbottom - $cancelbuttonhpad
        $cancelbutton.Location = "$newcancelbuttonx,$newbuttony"
        $okbutton.location = "20,$newbuttony"
    }

    # Set up Validation when the form closes with OK.
    $WinForm_FormClosing=[System.Windows.Forms.FormClosingEventHandler]{
        #Event Argument: $_ = [System.Windows.Forms.FormClosingEventArgs]
            #Validate only on OK Button
            if($WinForm.DialogResult -eq "OK")
            {
                #Validate the Child Control and Cancel if any fail
                $_.Cancel = -not $WinForm.ValidateChildren()
            }
            else {
                write-log "The Form was canceled before being completed."
            }
    }

    $Form_StateCorrection_Load=
    {
        $WinForm.WindowState = $script:InitialFormWindowState
    }

    $Form_Cleanup_FormClosed=
    {
        $WinForm.remove_Load($WinForm_Load)
        $WinForm.remove_Load($Form_StateCorrection_Load)
        $WinForm.remove_FormClosed($Form_Cleanup_FormClosed)
    }

    # Set our Error Handler options
    $PathErrorProvider.BlinkStyle = "NeverBlink"
    $PathErrorProvider.ContainerControl = $WinForm
    $UserErrorProvider.BlinkStyle = "NeverBlink"
    $UserErrorProvider.ContainerControl = $WinForm

    # Define our validating conditions
    # Path validating
    $objpathTextBox_Validating = [System.ComponentModel.CancelEventHandler]{
        $_.Cancel = $true
        $title = "Backup Path Invalid"
        $options = "OK"
        $style = "Error"
        try{
            $_.Cancel = Validate-Tag $objpathTextBox.Text $objpathTextBox.Tag
            if($_.Cancel)
            {
                $msg = "Please enter a valid Backup Path."
                Trigger-PopUp -title $title -msg $msg -options $options -style $style
                $PathErrorProvider.SetError($this,$msg)
                $PathErrorProvider.SetIconAlignment($this, [System.Windows.Forms.ErrorIconAlignment]::MiddleLeft)
            }
            else {
                $objpathTextBox.ForeColor = 'WindowText'
            }
        }
        catch [System.Management.Automation.ParameterBindingException]{
            $msg = "The Backup Path Field cannot be blank."
            Trigger-PopUp -title $title -msg $msg -options $options -style $style
            $PathErrorProvider.SetError($this,$msg)
            $PathErrorProvider.SetIconAlignment($this, [System.Windows.Forms.ErrorIconAlignment]::MiddleLeft)
        }
    }

    # Path validated
    $objpathTextBox_Validated ={
        # Pass the calling control and clear error message
        $PathErrorProvider.SetError($this, "")
    }

    $Winform.SuspendLayout()
    # Add the controls to our form
    $WinForm.Controls.Add($OKButton)
    $WinForm.Controls.Add($CancelButton)
    $WinForm.Controls.Add($BrowseButton)
    $WinForm.Controls.Add($objpathTextBox)
    $WinForm.Controls.Add($objpathLabel)

    # Customize the form
    $WinForm.MaximumSize = '320,150'
    $WinForm.Size = '320,150'
    $WinForm.Text = "USMT Backup Path"
    $WinForm.StartPosition = "CenterScreen"
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $WinForm.Icon = $Icon
    $WinForm.Topmost = $True
    $WinForm.AcceptButton = $OKButton

    # Allow for Enter and Escape Key action on Form.
    $WinForm.KeyPreview = $True
    $WinForm.Add_KeyDown(
        {
            # Use the OK Button Click Event for the OK button.
            if ($_.KeyCode -eq "Enter"){$OKButton.PerformClick()}
            # Use the Cancel Button Click Event for the Escape button.
            if ($_.KeyCode -eq "Escape"){Confirm-Cancel}
    })
    $Winform.Add_Load($WinForm_Load)

    # Customize the OK Button
    $OKButton.Anchor = "Top,Left"
    $OKButton.Size = '75,23'
    $OKButton.Text = "&OK"
    $OKButton.TabIndex = 3
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Define our Button Click event
    $OKButton.Add_Click(
    {
        $script:TSBP = $objpathTextBox.Text
        write-log "The TSBackup Path is $script:TSBP"
        if($script:sharecheck -eq $false)
        {
            $callinguser = $username
            $backedupuser = $objuserTextBox.Text
            if($callinguser -ne $backedupuser)
            {
                # The machine is being backed up for the owner by another user, likely a tech.  Include them in the email.
                $script:TSSUN = $backedupuser.split('\')[1]
                $bupem = "$($TSSUN)@monsanto.com"
                $opun = $callinguser.split('\')[1]
                $opem = "$($opun)@monsanto.com"
                $script:emailaddress = "$bupem;$opem"
                write-log "The user running the tool is $opun, the user being backed up is $script:TSSUN"
                $script:backedupuser = $backedupuser
                $script:loguser = $opun
            }
            else {
                # The user is backing up their own data.
                write-log "The Backed Up User is the current operator."
                $script:TSSUN = $username.split('\')[1]
                $bupem = "$($TSSUN)@monsanto.com"
                $script:emailaddress = "$bupem"
                $script:backedupuser = $username
                $script:loguser = $script:TSSUN
            }
        }
        else {
            # This is a shared machine backup, send the email to the username running the process
            $script:emailaddress = "$($username.split('\')[1])@monsanto.com"
        }
        $WinForm.Close()
        $WinForm.Dispose()
    })

        # Customize the Browse Button
        $BrowseButton.Location = '207,25'
        $BrowseButton.Size = '60,21'
        $BrowseButton.Text = "&Browse"
        $BrowseButton.CausesValidation = $false
        $BrowseButton.TabIndex = 1
    
        # Define our Button Click Event
        $BrowseButton.Add_Click({
            $Browsing = $true
            $Winform.Visible =  $false
            $BrowseInput = Browse-ForFolder
            if(![string]::IsNullOrEmpty($BrowseInput))
            {
                $objpathTextBox.Text = $BrowseInput
                $Winform.Visible = $true
            }
            else
            {
                $Winform.Visible = $true
            }
        })

   # Customize the Cancel Button
   $CancelButton.Size = '75,23'
   $CancelButton.Text = "&Cancel"
   $CancelButton.CausesValidation = $false
   $CancelButton.TabIndex = 4
   $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # Define our Button Click Event
    $CancelButton.Add_Click({
            write-log "The Operation was canceled."
            $script:TSBP = $null
            $script:USMTPath = $null
            $script:formexit = $true
            $WinForm.Close()
            $WinForm.Dispose()
    })

    # Customize our Path Text Box and Label
    # Path Text Box
    $objpathTextBox.Location = '20,25'
    $objpathTextBox.Size = '182,23'
    $objpathTextBox.Text = "\\<server>\<share>"
    $objpathTextBox.Tag = "\\<server>\<share>"
    $objpathTextBox.TabIndex = 0
    $objpathTextBox.Add_Validating($objpathTextBox_Validating)
    $objpathTextBox.Add_Validated($objpathTextBox_Validated)

    # Path Label
    $objpathLabel.TextAlign ="BottomLeft"
    $objpathLabel.Location = '20,5'
    $objpathLabel.Size = '280,20'
    $objpathLabel.Text = "Please enter the USMT Backup Path:"

    if($script:sharecheck -eq $false)
    {
        # Customize our User Text Box and Label
        # User Text Box
        # $objuserTextBox.Location = '20,90'
        $objuserTextBox.Size = '260,40'
        $objuserTextBox.Text = "<Domain>\<UserName>"
        $objuserTextBox.Tag = "<Domain>\<UserName>"
        $objuserTextBox.TabIndex = 2
        $objuserTextBox.Add_Validating($objUserTextBox_Validating)
        $objuserTextBox.Add_Validated($objUserTextBox_Validated)

        # User Label
        $objuserLabel.TextAlign ="BottomLeft"
        # $objuserLabel.Location = '20,55'
        $objuserLabel.Size = '280,20'
        $objuserLabel.Text = "Please enter the User ID of the Primary User:"

        # Set up User validating
        $objUserTextBox_Validating = [System.ComponentModel.CancelEventHandler]{
            $_.Cancel = $true
            $title = "UserName Invalid"
            $options = "OK"
            $style = "Error"
            try{
                $msg = "Please enter a valid UserName in Domain\Username format."
                $_.Cancel = Validate-Tag $objuserTextBox.Text $objuserTextBox.Tag
                if($_.Cancel)
                {
                    Trigger-PopUp -title $title -msg $msg -options $options -style $style
                    $UserErrorProvider.SetError($this,$msg)
                    $UserErrorProvider.SetIconAlignment($this, [System.Windows.Forms.ErrorIconAlignment]::MiddleLeft)
                }
                else {
                    if($objuserTextBox.Text -notlike "*\*")
                    {
                        $_.Cancel = $true
                        Trigger-PopUp -title $title -msg $msg -options $options -style $style
                        $UserErrorProvider.SetError($this,$msg)
                        $UserErrorProvider.SetIconAlignment($this, [System.Windows.Forms.ErrorIconAlignment]::MiddleLeft)
                    }
                    else {
                        $objuserTextBox.ForeColor = 'WindowText'
                    }
                }
            }
            catch [System.Management.Automation.ParameterBindingException]{
                $msg = "The UserName Field cannot be blank."
                Trigger-PopUp -title $title -msg $msg -options $options -style $style
                $UserErrorProvider.SetError($this,$msg)
                $UserErrorProvider.SetIconAlignment($this, [System.Windows.Forms.ErrorIconAlignment]::MiddleLeft)
            }
        }

        # User validated
        $objUserTextBox_Validated ={
            # Pass the calling control and clear error message
            $UserErrorProvider.SetError($this, "")
        }

        # Include our User Controls
        $WinForm.Controls.Add($objuserTextBox)
        $WinForm.Controls.Add($objuserLabel)
    }

    $Winform.ResumeLayout()

    #Save the initial state of the form
    $InitialFormWindowState = $WinForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $WinForm.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $WinForm.add_FormClosed($Form_Cleanup_FormClosed)
    #Add shown handler for focus
    $WinForm.Add_Shown({$WinForm.Activate()})
    #Show the Form
    $WinForm.ShowDialog()
}

Function Get-LoggedOnUser {
    <#
    .SYNOPSIS
        Get session details for all local and RDP logged on users.
    .DESCRIPTION
        Get session details for all local and RDP logged on users using Win32 APIs. Get the following session details:
         NTAccount, SID, UserName, DomainName, SessionId, SessionName, ConnectState, IsCurrentSession, IsConsoleSession, IsUserSession, IsActiveUserSession
         IsRdpSession, IsLocalAdmin, LogonTime, IdleTime, DisconnectTime, ClientName, ClientProtocolType, ClientDirectory, ClientBuildNumber
    .EXAMPLE
        Get-LoggedOnUser
    .NOTES
        Description of ConnectState property:
        Value		 Description
        -----		 -----------
        Active		 A user is logged on to the session.
        ConnectQuery The session is in the process of connecting to a client.
        Connected	 A client is connected to the session.
        Disconnected The session is active, but the client has disconnected from it.
        Down		 The session is down due to an error.
        Idle		 The session is waiting for a client to connect.
        Initializing The session is initializing.
        Listening 	 The session is listening for connections.
        Reset		 The session is being reset.
        Shadowing	 This session is shadowing another session.
        
        Description of IsActiveUserSession property:
        If a console user exists, then that will be the active user session.
        If no console user exists but users are logged in, such as on terminal servers, then the first logged-in non-console user that is either 'Active' or 'Connected' is the active user.
        
        Description of IsRdpSession property:
        Gets a value indicating whether the user is associated with an RDP client session.
    .LINK
        http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    Param (
    )
    Try {
        Write-Log 'Get session information for all logged on users.'
        Write-Output -InputObject ([PSADT.QueryUser]::GetUserSessionInfo("$env:ComputerName"))
    }
    Catch {
        Write-Log "Failed to get session information for all logged on users." -LogLevel 3
    }
}

function Compile-Scriptblocks
{
    # Define our username variable
    $username = $script:loguser
    $backedupuser = $script:backedupuser
    # Create our joblist array here
    $joblist = new-object System.Collections.ArrayList
    # Define some function-level variables here
    $jobstarttime = get-date
    $time = $jobstarttime.ToString("MM-dd-yyyy-hh-mm")
    $scriptroot = $script:scriptroot
    $usmtpath = $script:usmtpath
    $logpath = $script:logpath
    $proglog = "scanstateprogress-$time.log"
    $script:progpath = "$script:logpath\$proglog"
    # Check to see if this is a shared machine
    if($script:sharecheck)
    {
        $lastargline = "/vsc /uel:90 /ue:*\*app-* /ue:*\*svc-* /ue:%COMPUTERNAME%\* /ue:*\*admin* /encrypt:AES_256 /key:M0n$@nt0!2o17"
    }
    else {
        $lastargline = "/ue:*\* /ui:$backedupuser /encrypt:AES_256 /key:M0n$@nt0!2o17"
    }
    # Set up our scriptblocks
    Write-Log "Compiling Scriptblocks for Scanstate Process"
    # Scanstate jobsb can be static across all run types, defining up here.
    $ssjobsb = {
        
        param($sroot,$usmtp,$lpath,$time,$plogpath,$username,$largline)
        $now = $time
        $scriptroot = $sroot
        $approot = split-path -Parent -Path $scriptroot
        $usmtpath = $usmtp
        $usmtroot = "$approot\amd64"
        $logpath = $lpath
        $scanstate = "$usmtroot\scanstate.exe"
        $statelog = "scanstate-$now.log"
        $progpath = $plogpath
        $lastargline = $largline
        $stdoutlogname = "scanstate-stdout-$now.log"
        $stderrlogname = "scanstate-stderror-$now.log"

        $arglist = @(
            "/o /localonly /efs:copyraw /v:5 /c",
            "$USMTPath",
            "/l:$logpath\$statelog /progress:$progpath",
            "/i:$usmtroot\MonCustom.xml /i:$usmtroot\migdocs.xml /i:$usmtroot\migapp.xml",
            "$lastargline"
        )
        $usmtjob = Start-Process -FilePath $scanstate -ArgumentList $arglist -NoNewWindow -RedirectStandardOutput "$logpath\$stdoutlogname" -RedirectStandardError "$logpath\$stderrlogname" -Wait -Passthru
        $starttime = $usmtjob.StartTime
        $exittime = $usmtjob.ExitTime
        $exitcode = $usmtjob.ExitCode
        $exitobj = new-object PSObject -Prop @{'ExitCode'=$exitcode;'StartTime'=$starttime;'ExitTime'=$exittime}
        $exitobj
    }
    if($script:silent)
    {
        write-log "No Progress Bar Job invoked for Silent run Completed Scriptblock."
        # Silent run completed scriptblock won't have any progress bar stuff
        $sscompletedsb = {
            param([System.Management.Automation.Job]$job)
            $jobname = $job.Name
            $jobstate = $job.state
            $result = receive-job $jobname -keep
            $starttime = $result.StartTime
            $endtime = $result.ExitTime
            $endcode = $result.ExitCode
            if(($endcode -eq 0) -or ($endcode -eq 3))
            {
                write-log "The USMT Backup job started at $starttime and completed successfully at $endtime with exit code $endcode"
                Final-Cleanup -Silent -Username $username
            }
            else {
                write-log "The USMT Backup job started at $starttime and failed at $endtime with exit code $endcode"
                Final-Cleanup -Silent -bad -Username $username
            }
        }
    }
    else {
        write-log "Progress Bar Job invoked for Interactive run Completed Scriptblock."
        # Interactive run completed scriptblock will need to update and address progress bar form
        $sscompletedsb = {
            param([System.Management.Automation.Job]$job)
            if($job)
            {
                $jobname = $job.Name
                $jobstate = $job.state
                $result = receive-job $jobname -keep
                $starttime = $result.StartTime
                $endtime = $result.ExitTime
                $endcode = $result.ExitCode
                if(($endcode -eq 0) -or ($endcode -eq 3))
                {
                    write-log "The USMT Backup job started at $starttime and completed at $endtime with exit code $endcode"
                    $progressbar1.Value = 100
                    $labeltext = "User Data Backup Process Complete"
                    $PhaseLabel.Text = $labeltext
                    $formJobProgress.Close()
                    $formJobProgress.Dispose()
                    Final-Cleanup -Username $username
                }
                else {
                    write-log "The USMT Backup job started at $starttime and failed at $endtime with exit code $endcode"
                    $progressbar1.Value = 100
                    $labeltext = "User Data Backup Process Failed"
                    $PhaseLabel.Text = $labeltext
                    $formJobProgress.Close()
                    $formJobProgress.Dispose()
                    Final-Cleanup -bad -Username $username
                }
            }
            else {
                write-log "The USMT Backup job did not start."
                Final-Cleanup -bad -Username $username
            }
        }
    }

    # Add our scriptblock and arguments to an object.
    # Create a debug array
    $jobinobj = new-object PSObject -Prop @{'Scriptroot'=$scriptroot;'USMTPath'=$usmtpath;'Logpath'=$logpath;'Time'=$time;'Progpath'=$progpath;'Username'=$username;'LastArgline'=$lastargline;'JobType'="Scanstate"}
    # Define our jobname
    $jobname = "USMTBackup"
    # Add our input object to a new object with our other scriptblocks.
    $jobobj = new-object PSObject -Prop @{'Jobname'=$jobname;'Jobscript'=$ssjobsb;'JobInput'=$jobinobj;'CompletedScript'=$sscompletedsb}
    # Add our objects to the joblist
    [void]$joblist.add($jobobj)

    # Create objects for Progress Bar Form if interactive
    if(!$script:silent)
    {
        # Create the scriptblocks for the Progress Bar Form
        Write-Log "Compiling Scriptblocks for Progress Bar Form"
        # Create the scriptblock for our Progress Bar Form
        $pbjobsb = {
            # This is where we put our scriptblock.
            param($filepath)
            $Finishstring = "Successful run"
            $phasestring = "PHASE"
            $percentstring = "totalPercentageCompleted"
            $phasearr = @()
            $percentarr = @()
            $lastcount = 0
            do
            {
                $match = $false
                $progin = @(get-content $filepath)
                if($progin.count -gt 0)
                {
                    # The file has content.
                    if(($progin.count) -gt $lastcount)
                    {
                        # The content has changed
                        $startloop = $lastcount
                        $lastcount = $progin.count
                        $currentchunk = $progin[$startloop..$lastcount]
                        foreach($line in $currentchunk)
                        {
                            try{$phasetrigger = $line | ? {$_ -clike "*$($phasestring)*"}}
                            catch{$phasetrigger = $null}
                            if($phasetrigger)
                            {
                                $phase = $phasetrigger.split(',')[-1].TrimStart().Trim()
                                if(!($phasearr -match $phase))
                                {
                                    $phasearr+=$phase
                                    $currentstatus = $phasearr[-1]
                                }
                            }
                            try{$percenttrigger = $line | ? {$_ -clike "*$($percentstring)*"}}
                            catch{$percenttrigger = $null}
                            if($percenttrigger)
                            {
                                $percent = [int]($percenttrigger.split(',')[-1].TrimStart().Trim())
                                if(!($percentarr -match $percent))
                                {
                                    $percentarr+=$percent
                                    $currentpct = $percentarr[-1]
                                }
                            }
                            try{$matchtrigger = $currentchunk | ? {$_ -clike "*$($Finishstring)*"}}
                            catch{$matchtrigger = $null}
                            if($matchtrigger)
                            {
                                $match = $true
                            }
                            if(!$currentpct){$currentpct = 0}
                            $resultval = new-object PSobject -Prop @{'Status'=$currentstatus;'Percent'=$currentpct}
                            $resultval
                        }
                    }
                    else {}
                }
                else
                {} # The file has no content.
            }
            while($match -ne $true)
        }
        $jobinobj = new-object PSObject -Prop @{'Filepath'=$progpath;'JobType'="ProgressBar"}
        $pbcompletedsb = {
            Param($Job)
            $progressbar1.Value = 100
            $labeltext = "User Data Backup Process Complete"
            $PhaseLabel.Text = $labeltext
        }
        $pbupdatesb = {
            Param($Job)
            $results = Receive-Job -Job $Job -keep
            $uniqueresults = @($results | sort Percent,Status -unique | select Percent,Status)
            $currentpct = $uniqueresults[-1].Percent
            $currentstatus = $uniqueresults[-1].Status
            if($currentstatus)
            {
                $labeltext = "Executing $currentstatus Phase.  Please wait..."
            }
            else {
                $labeltext = "Starting. Please wait ... "
            }
            $PhaseLabel.Text = $labeltext
            $progressbar1.Value = $currentpct
        }
        # Add our scriptblocks to an object
        $jobobj = new-object PSObject -Prop @{'Jobname'="Progress Bar Updater";'Jobscript'=$pbjobsb;'JobInput'=$jobinobj;'UpdateScript'=$pbupdatesb;'CompletedScript'=$pbcompletedsb}
        # Add our object to the joblist.
        [void]$joblist.Add($jobobj)
        write-log "Invoking Run-USMT"
        Run-USMT -joblist $joblist
    }
    else {
        write-log "Invoking Run-USMT"
        Run-USMT -joblist $jobobj
    }
}

function Environment-Check
{
    # Evaluate our running user name
    $username = $script:username
    $unlist = $username.split(',')
    # ProcUser is the process context that the script is running under.
    $procuser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    # Pass our username
    if(-not [string]::IsNullOrEmpty($username))
    {
        if($unlist.count -gt 1)
        {
            # RunningUser is currently logged onto the machine.
            $runninguser = $unlist[1]
            $usermsg = "$usermsg`nRunninguser is set to $runninguser."
            # Username is the value passed by the running user
            $username = $unlist[0]
            $usermsg = "$usermsg`nUsername is set to $username"
        }
        else {
            $usermsg = "$usermsg`nUsername value is set to $username"
        }
        if($procuser -ne $username)
        {
            if(-not [string]::IsNullOrEmpty($runninguser))
            {
                $usermsg = "$usermsg`nScript run initiated for $username by logged in user $runninguser, executed through $procuser"
            }
            else {
                $usermsg = "$usermsg`nScript run initiated for $username, executed through $procuser"
            }
        }
        else {
            $usermsg = "Run in context of $username"
        }
    }
    else {
        $username = $procuser
        $usermsg = "Run without a username value, resolved username context to $username"
    }
    # Pass our Computername
    $computername = $env:computername
    # Define our parent directory.
    $scriptroot = Get-ScriptDirectory
    $approot = split-path -Parent -Path $scriptroot
    $usmtroot = "$approot\amd64"

    # Define our root registry key
    $MonRegPath = "HKLM:\Software\Monsanto"
    $RegRebootPath = "$MonRegPath\USMT"

    # Set up variables for our registry key values
    $BackupTypeKey = "USMTBackupType"
    $SharedKey = "SharedMachine"
    $UserKey = "LoggedOnUser"

    $testrootpath = test-path $RegRebootPath
    if(!$testrootpath)
    {
        # This is our first run
        New-Item -Path $MonRegPath -Name USMT -Force | out-null
        $firstrun = $true
        $runcheck = "This is our first run."
    }
    else {
        # This is not our first run
        $firstrun = $false
        $runcheck = "This is not our first run."
    }

    # Check to see if we're running silent
    if($script:Silent)
    {
        $silentcheck = "Silent run of the script has been initiated."
    }
    else {
        $silentcheck = "Interactive run of the script has been initiated."
    }

    # Get the current time
    $now = Get-Date -Format "MM-dd-yyyy-hh-mm"

    # Check to see if we're running in a Task Sequence
    try{$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment}
    catch{$tsenv = $null}
    if($tsenv)
    {
        # We're running from a Task Sequence
        # Set our initial logpath to the default logpath
        $logPath = $tsenv.Value("_SMSTSLogPath")
        # Set our logfile name properties
        # Our initial filename is defined in our Get-Scriptdirectory function
        $scriptname = "$($script:filename).ps1"
        $filename = "$($script:filename)-$now.log"
        $logFile = "$logPath\$filename"
        $envmsg = "Running from Task Sequence."
        # Check for our Username value
        try{
            # Grab the TS Variable value 
            $username = $tsenv.value("TSCusCurrentUser")
        }
        catch{
            $username = $null
        }
        
        #Hide the progress dialog
        $TSProgressUI = new-object -comobject Microsoft.SMS.TSProgressUI
        $TSProgressUI.CloseProgressDialog()
    }
    else 
    {
        # We're running from the .exe or manually.
        # Create a logpath one level above the scriptroot, and define it as the logpath
        $logPath = "$approot\LogFiles"
        if(!(test-path $logPath)){new-item $logpath -type Directory -Force -ErrorAction SilentlyContinue}
        # Set our logfile name properties
        # Our initial filename is defined in our Get-Scriptdirectory function
        $script:scriptname = "$($script:filename).ps1"
        $script:filename = "$($script:filename)-$now.log"
        $logFile = "$logPath\$filename"
        $envmsg = "Not running from Task Sequence"
    }

    # Start Logging
    Start-Log -FilePath $logfile
    Write-Log $runcheck
    Write-Log $SilentCheck
    write-log $envmsg
    Write-Log "LogPath is set to $logpath"
    write-log $usermsg

    # Now that we've got our environment sorted out, set our script level variables here.
    $script:scriptroot = $scriptroot
    $script:approot = $approot
    $script:usmtroot = $usmtroot
    $script:logPath = $logPath
    $script:logfile = $logfile
    $script:RegRebootPath = $RegRebootPath
    $script:username = $username
    $script:now = $now

    if($firstrun)
    {
        # Perform first run stuff
        if(!$script:silent)
        {
            if(!$script:shared)
            {
                # Not run with the shared switch, have to check.
                write-log "Not run with the shared switch, have to check."
                $title = "Monsanto User Data Backup - Shared Computer Usage"
                $message = "Is this machine shared by multiple users on a regular basis?  If so, data from all users who have logged in within the last 90 days will be captured."
                $options = "YesNo"
                $style = "Question"
                $sharecheck = Trigger-PopUp -title $title -msg $message -options $options -style $style
                switch($sharecheck)
                {
                    "Yes" {
                            write-log "This is a shared machine"
                            $script:sharecheck = $true
                    }
                    "No" {
                        write-log "This is not a shared machine"
                            $script:sharecheck = $false
                    }
                }
            }
            else {
                write-log "Run with the shared switch, no need to check."
                $script:sharecheck = $true
            }
            
            # Determine our backup type
            $title = "Monsanto User Data Backup - USB Storage"
            $message = "Will the backup be stored on a USB or other removable drive, directly connected to this PC?  If you are backing up to a network location (i.e. \\server\share\backup), choose No."
            $options = "YesNo"
            $style = "Question"
            $usbcheck = Trigger-PopUp -title $title -msg $message -options $options -style $style
            switch ($usbcheck)
            {
                "Yes" {
                    $BackupType = "Local"
                }
                "No" {
                    $BackupType = "Network"
                }
            }
            Write-Log "BackupType $BackupType has been selected, writing it to the Registry."
            #Ask about CheckDisk
            $title = "Run CheckDisk?"
            $message = "Do you want to run CheckDisk to check the disk for errors?`n`nIf you click Yes, the system will restart and initiate a ChkDsk on the local C: drive."
            $options = "YesNo"
            $style = "Question"
            $chkdiskcheck = Trigger-PopUp -title $title -msg $message -options $options -style $style
            switch ($chkdiskcheck)
            {
                "Yes" {
                        Write-Log "User selected Yes to CheckDisk."
                        # Create the RebootNeeded Registry Key and set to true.
                        $ChkDskRebootNeeded = $true
                        $cdrive = Get-WMIObject -class Win32_LogicalDisk -Filter 'DeviceID="C:"'
                        $cdrive.chkdsk($false,$false,$false,$true,$false,$true)
                        $chkdskgo = $false
                }
                "No" {
                        # This makes life easier.
                        Write-Log "User selected No to CheckDisk."
                        # Create the RebootNeeded Registry Key and set to false.
                        $ChkDskRebootNeeded = $false
                        $chkdskgo = $true
                }
            }
        }
        else {
            # Silent run of the script.  This should always be firstrun level stuff, because a silent run of the script shouldn't involve a reboot.
            # Define our email address
            $emlist = $script:emailaddress.split('')
            if($emlist.count -gt 1){$emailaddress = $emlist -join ';'}
            # Set Chkdskgo to True
            write-log "Setting Chkdskgo to True."
            $chkdskgo = $true
            #  Set our BackupType
            write-log "Backup type has to be set to network for Silent runs."
            $BackupType = "Network"
            # Define our USMTPath
            $pathstringtest = [string]::IsNullOrEmpty($script:USMTPath)
            if(!$pathstringtest)
            {
                $USMTPath = $script:USMTPath
                # Ready to assemble our final path
                if(!$script:shared)
                {
                    write-log "Silent run of the script, no shared switch, setting sharedcheck to false."
                    $script:sharecheck = $false
                    $callinguser = $procuser
                    $backedupuser = $username
                    if($callinguser -ne $backedupuser)
                    {
                        $TSSUN = $backedupuser.split('\')[1]
                        $bupem = "$($TSSUN)@monsanto.com"
                        # Check to see if we're running the script as System
                        if($callinguser -ne "NT AUTHORITY\SYSTEM")
                        {
                            write-log "CallingUser is $callinguser, Backedup User is $backedupuser"
                            # The machine is being backed up manually for the owner by another user, likely a tech.  Include them in the email.
                            $opem = "$($callinguser.split('\')[1])@monsanto.com"
                            $emadds =$opem
                            $script:loguser = $opun
                        }
                        else {
                            # The script is being run through the executable, or by SCCM, don't try to send an email to SYSTEM
                            if(-not [string]::IsNullOrEmpty($runninguser))
                            {
                                $opem = "$($runninguser.split('\')[1])@monsanto.com"
                                $emadds = "$bupem;$opem"
                                $script:loguser = $($runninguser.split('\')[1])
                            }
                            else {
                                write-log "CallingUser is $callinguser, only sending email to $backedupuser"
                                $emadds = "$bupem"
                                $script:loguser = $TSSUN
                            }
                        }
                        $script:backedupuser = $backedupuser
                    }
                    else {
                        # The user is backing up their own data.
                        write-log "The Backed Up User is the current operator."
                        $TSSUN = $username.split('\')[1]
                        $bupem = "$($TSSUN)@monsanto.com"
                        $emadds = $bupem
                        $script:backedupuser = $username
                        $script:loguser = $TSSUN
                    }
                    # Check the USMTPath and glue it together 
                    if(($USMTPath.endswith(":\"))){$USMTPath = "$USMTPath$($TSSUN)_$computername"}
                    else{$USMTPath = "$USMTPath\$($TSSUN)_$computername"}
                }
                else {
                    write-log "Silent run of the script using shared switch, setting sharedcheck to true."
                    $TSSUN = $username.split('\')[1]
                    $script:sharecheck = $true
                    $emadds = "$TSSUN@monsanto.com"
                    # Check the USMTPath and glue it together 
                    if(($USMTPath.endswith(":\"))){$USMTPath = "$USMTPath$computername"}
                    else{$USMTPath = "$USMTPath\$computername"}
                    $script:loguser = $TSSUN
                }
                # Determine whether or not we need to add email addresses
                if(!$emailaddress)
                {
                    # Run without an explicit email address
                    write-log "The script was run without an explicit email address"
                    $script:emailaddress = $emadds
                }
                else{
                    write-log "The script was run with an explicit email address of $emailaddress, adding it to the list"
                    $script:emailaddress = "$emailaddress;$emadds"
                }
                write-log "LogUser value set to $script:loguser"
                write-log "USMTPath is $USMTPath"
                write-log "Sending email to $script:emailaddress"
                # Path assembled, time to see if it needs to be created
                if(!(test-path $usmtpath) -eq $true)
                {
                    Write-Log "The USMTPath does not exist, attempting to create it now..." -LogLevel 2
                    try{$newdir = new-item $USMTPath -type Directory -Force -ErrorAction SilentlyContinue}
                    catch{$newdir = $null}
                    if($newdir)
                    {
                        # Path was created successfully
                        Write-Log "All required parameters for Silent script run are valid, USMTPath was created, compiling script blocks."
                    } 
                    else
                    {
                        # Path couldn't be created, abort.
                        $errmsg = "The USMTPath could not be located or created.  Please check the path and try again."
                        write-log $errmsg -LogLevel 3
                        $returncode = 1
                        exit $returncode
                    }
                }
                else {
                    Write-Log "All required parameters for Silent script run are valid, USMTPath exists, compiling script blocks."
                }
            }
            else {
                $errmsg = "Silent run of the script has been initiated, but no USMTPath was passed."
                write-log $errmsg -LogLevel 3
                $returncode = 1
                Final-Cleanup -bad -silent -cancel
            }
            # Set our script level USMTPath
            $script:USMTPath = $USMTPath
        }

        # Check to see if we need to Unblock/Reboot for USB
        do
        {
            switch($BackupType)
            {
                "Local" {
                        # Set localchkgo val to false before evaluating.
                        $script:localchkgo = $false
                        # Check to see if we have any USB Drives plugged in right now.
                        Write-Log "Checking for any plugged in drives."
                        $usbdrives = @(gwmi win32_diskdrive | ? {(($_.MediaType -eq "Removable Media") -and ($_.InterfaceType -eq "USB"))})
                        if($usbdrives.count -gt 0) # Test SB3L2A1
                        {
                            # We do.  Prompt to remove before beginning Unblock.
                            Write-Log "Prompting for removal of USB Drive."
                            $title = "Monsanto User Data Backup - Remove USB Drive"
                            $options = "OK"
                            $style = "Warning"
                            $message = "A removable USB Drive has been detected.  If this is the drive that you intend to back up to, please remove it before clicking OK."
                            $message2 = "If you do not remove the USB drive before clicking OK a reboot will be triggered in order to make the drive available."
                            Trigger-PopUp -title $title -msg $message -msg2 $message2 -options $options -style $style
                            $usbdrives = @(gwmi win32_diskdrive | ? {(($_.MediaType -eq "Removable Media") -and ($_.InterfaceType -eq "USB"))})
                            if($usbdrives.count -gt 0) # Test SB3L2A1A
                            {
                                # They haven't removed the USB Drive, so we'll need to reboot.
                                Write-Log "There are still removeable drives plugged in despite the warning." -LogLevel 3
                                Write-Log "Prompting the user again."
                                $title = "Monsanto User Data Backup - Remove USB Drive"
                                $options = "RebootRetryCancel"
                                $style = "Warning"
                                $message = "A removable USB Drive is still connected to the machine.  If this is the drive that you intend to back up to, please remove it and click Retry."
                                $message2 = "If the USB Drive cannot be removed at the present time, please click Reboot.  If you wish to Cancel and start over later, click Cancel."
                                $drivepop2 = Trigger-PopUp -title $title -msg $message -msg2 $message2 -options $options -style $style
                                switch ($drivepop2)
                                {
                                    "OK" {
                                        Write-Log "The user chose to Reboot."
                                        $script:USBRebootNeeded = $true
                                        Unblock-USB
                                    }
                                    "Retry" {
                                        Write-Log "The user chose to Retry."
                                    }
                                    "Cancel" {
                                        Write-Log "The user chose to Cancel."
                                        Confirm-Cancel
                                    }
                                }
                            }
                            else
                            {
                                Write-Log "The USB Drive has been removed.  Unblocking and prompting for insertion of USB Drive."
                                $script:USBRebootNeeded = $false
                                $unblockgo = $true
                            }
                        }
                        else
                        {
                            Write-Log "No Removable USB Drives are detected"
                            $script:USBRebootNeeded = $false
                            $unblockgo = $true
                        }
                        if($unblockgo)
                        {
                            Write-Log "Proceeding to unblock."
                            # Run the unblock function before rebooting the machine
                            $Unblock = Unblock-USB
                            if($unblock -eq "0")
                            {
                                Write-Log "Checking for USB Drive readiness."
                                $script:USBRebootNeeded = $false
                                Drive-Check
                            }
                            else
                            {
                                Write-Log "USB Drive check failed.  We will most likely need to reboot." -LogLevel 2
                                $title = "Monsanto User Data Backup - Reboot Required"
                                $options = "OK"
                                $style = "Warning"
                                $message = "There was an issue with detecting your USB drive.  To enable USB storage access, the PC will be rebooted when you click OK."
                                Trigger-PopUp -title $title -msg $message -options $options -style $style
                                $script:USBRebootNeeded = $true
                            }
                        }
                }
                "Network" {
                    $script:localchkgo = $true
                    $script:USBRebootNeeded = $false
                }
            }
        }
        until($script:USBRebootNeeded -ne $null)

        # Tattoo our values to the registry, in case we need to reboot.
        Create-RegistryValue -Path $RegRebootPath -Name $SharedKey -Value $script:sharecheck
        Create-RegistryValue -Path $RegRebootPath -Name $BackupTypeKey -Value $BackupType
        Create-RegistryValue -Path $RegRebootPath -Name $UserKey -Value $username
    }
    else {
        # Not our first run, rename logfile and prep for post-reboot script run
        $postbootlogfilename = "$($filename.substring(0,$($filename.length-4)))_PostReboot.log"
        $newpath = "$logpath\$postbootlogfilename"
        Rename-Item $logfile $postbootlogfilename
        $logfile = $newpath
        Start-Log -FilePath $logfile
        $chkdskgo = $true
        $script:localchkgo = $true
        $BackupType = Get-RegistryValue -Path $RegRebootPath -Value $BackupTypeKey
        write-log "Our backup type is $BackupType"
        $script:sharecheck = Get-RegistryValue -Path $RegRebootPath -Value $SharedKey
        write-log "Our sharecheck value is set to $script:sharecheck"
        $script:username = Get-RegistryValue -Path $RegRebootPath -Value $UserKey
        write-log "Our username is $script:username"
        # Check to see if we need to run Drive-Check
        if($BackupType -eq "Local")
        {
            # We've rebooted for USB or for Checkdisk, either way we need to detect the drive.
            write-log "Rebooted for USB, running Drive-Check"
            Drive-Check
        }
        if($script:cancel)
        {
            # Debug switch, reset everything
            write-log "Run with the cancel switch, reset the environment."
            Final-Cleanup -bad -silent -cancel
        }
    }

    # Check to see if we're good to start the process
    if(($chkdskgo -eq $true) -and ($script:localchkgo -eq $true))
    {
        # Chkdsk and USB checks are both good, ready to move forward.
        Write-Log "Chkdskgo is $chkdskgo and localchkgo is $script:localchkgo"
        if(!$script:silent)
        {
            while($script:TSBP -eq $null)
            {
                # TSBP is null, which means that the user either hasn't run the form yet, they've hit cancel, or something has gone wrong.
                if(!$script:formexit)
                {
                    # The user didn't hit cancel, so either the user hasn't run the form yet, or something has gone wrong.
                    # Get our path
                    Run-WinForm
                }
                else
                {
                    Write-Log "The user chose to cancel."
                    Confirm-Cancel
                }
            }
        }
        else {
            # Set confirmed to True and Formexit to False.
            $script:confirmed = $true
            $script:formexit = $false
        }
        
        # Beginning USMTPath confirmation loop
        while($script:confirmed -eq $false)
        {
            Validate-FinalPath
            if($script:formexit -eq $true)
            {
                Write-Log "The user chose to cancel."
                Confirm-Cancel
            }
            # Haven't confirmed our USMTPath yet.
        }
        if(($script:confirmed -eq $true) -and ($script:formexit -eq $false))
        {
            Write-Log "Confirmed is true and formexit is false."
            Compile-Scriptblocks
        }
        else {
            write-Log "Confirmed is $script:confirmed and Formexit is $script:Formexit"
        }
    }
    else {
        Write-Log "Chkdskgo is $chkdskgo and localchkgo is $script:localchkgo"
        # We probably need to reboot.
        if(($script:USBRebootNeeded -eq $true) -or ($ChkDskRebootNeeded -eq $true))
        {
            # Create a scheduled task to rerun our script at next logon.
            Write-Log "Setting up task to rerun our script on next login."
            $Scriptpath = "$Scriptroot\$scriptname"
            Write-Log "Our scriptpath is $scriptpath"
            $Name = "Rerun USMT-Backup.ps1"
            $Descr = "Rerun the script after reboot to unblock USB."
            $TaskCommand = "$scriptroot\ServiceUI_x64.exe"
            Write-Log "Our taskcommand is $TaskCommand"
            $TaskArgs = "`"c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe`" -NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file $scriptpath"
            Write-Log "Our TaskArgs are $TaskArgs"
            Write-Log "Creating new scheduled task."
            New-ScheduledTask -TaskName $Name -TaskDescr $Descr -TaskTrigger "9" -TaskCommand $TaskCommand -TaskArgs $TaskArgs

            Write-Log "USBRebootNeeded returns $script:USBRebootNeeded and ChkDskRebootNeeded returns $ChkDskRebootNeeded"
            Write-Log "Prompting the User to confirm reboot"
            $title = "Monsanto User Data Backup - Reboot Pending"
            $message = "This machine will need to be rebooted before the data backup process can continue."
            $message2 = "Please click Reboot in order to reboot, or Cancel in order to Cancel."
            $options = "RebootCancel"
            $style = "Question"
            $rebootcheck = Trigger-PopUp -title $title -msg $message -msg2 $message2 -options $options -style $style
            switch ($rebootcheck)
            {
                "OK" {
                    Write-Log "The user clicked the Reboot button.  Rebooting the machine." -LogLevel 2
                    $prebootlogfilename = "$($filename.substring(0,$($filename.length-4)))_PreReboot.log"
                    Write-Log "Renaming Pre-Boot Logfile"
                    Rename-Item $logfile $prebootlogfilename
                    Restart-Computer -Force
                }
                "Cancel" {
                    Write-Log "The user chose to cancel."
                    Confirm-Cancel
                }
            }
        }
        else
        {
            Write-Log "USBRebootNeeded returns $script:USBRebootNeeded and ChkDskRebootNeeded returns $ChkDskRebootNeeded"
        }
    }
}

# Begin script here - Script Body (SB)
Environment-Check