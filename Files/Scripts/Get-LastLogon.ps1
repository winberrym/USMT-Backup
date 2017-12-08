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

Function Get-LastLogon
{
<#

.SYNOPSIS
	This function will list the last user logged on or logged in.

.DESCRIPTION
	This function will list the last user logged on or logged in.  It will detect if the user is currently logged on
	via WMI or the Registry, depending on what version of Windows is running on the target.  There is some "guess" work
	to determine what Domain the user truly belongs to if run against Vista NON SP1 and below, since the function
	is using the profile name initially to detect the user name.  It then compares the profile name and the Security
	Entries (ACE-SDDL) to see if they are equal to determine Domain and if the profile is loaded via the Registry.

.PARAMETER ComputerName
	A single Computer or an array of computer names.  The default is localhost ($env:COMPUTERNAME).

.PARAMETER FilterSID
	Filters a single SID from the results.  For use if there is a service account commonly used.
	
.PARAMETER WQLFilter
	Default WQLFilter defined for the Win32_UserProfile query, it is best to leave this alone, unless you know what
	you are doing.
	Default Value = "NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	
.EXAMPLE
	$Servers = Get-Content "C:\ServerList.txt"
	Get-LastLogon -ComputerName $Servers

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILIHTE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True
	
.EXAMPLE
	Get-LastLogon -ComputerName svr01, svr02 -FilterSID S-1-5-21-012345678-0123456789-012345678-012345

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILIHTE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True

.LINK
	http://msdn.microsoft.com/en-us/library/windows/desktop/ee886409(v=vs.85).aspx
	http://msdn.microsoft.com/en-us/library/system.security.principal.securityidentifier.aspx

.NOTES
	Author:	 Brian C. Wilhite
	Email:	 bwilhite1@carolina.rr.com
	Date: 	 "09/20/2012"
	Updates: Added FilterSID Parameter
	         Cleaned Up Code, defined fewer variables when creating PSObjects
	ToDo:    Clean up the UserSID Translation, to continue even if the SID is local
#>

[CmdletBinding()]
param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	[Alias("CN","Computer")]
	[String[]]$ComputerName="$env:COMPUTERNAME",
	[String]$FilterSID,
	[String]$WQLFilter="NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	)

Begin
	{
		#Adjusting ErrorActionPreference to stop on all errors
		$TempErrAct = $ErrorActionPreference
		$ErrorActionPreference = "Stop"
		#Exclude Local System, Local Service & Network Service
	}#End Begin Script Block

Process
	{
		Foreach ($Computer in $ComputerName)
			{
				$Computer = $Computer.ToUpper().Trim()
				Try
					{
						#Querying Windows version to determine how to proceed.
						$Win32OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer
						$Build = $Win32OS.BuildNumber
						
						#Win32_UserProfile exist on Windows Vista and above
						If ($Build -ge 6001)
							{
								If ($FilterSID)
									{
										$WQLFilter = $WQLFilter + " AND NOT SID = `'$FilterSID`'"
									}#End If ($FilterSID)
								$Win32User = Get-WmiObject -Class Win32_UserProfile -Filter $WQLFilter -ComputerName $Computer
								$LastUser = $Win32User | Sort-Object -Property LastUseTime -Descending | Select-Object -First 1
								$Loaded = $LastUser.Loaded
								$Script:Time = ([WMI]'').ConvertToDateTime($LastUser.LastUseTime)
								
								#Convert SID to Account for friendly display
								$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($LastUser.SID)
								$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
							}#End If ($Build -ge 6001)
							
						If ($Build -le 6000)
							{
								If ($Build -eq 2195)
									{
										$SysDrv = $Win32OS.SystemDirectory.ToCharArray()[0] + ":"
									}#End If ($Build -eq 2195)
								Else
									{
										$SysDrv = $Win32OS.SystemDrive
									}#End Else
								$SysDrv = $SysDrv.Replace(":","$")
								$Script:ProfLoc = "\\$Computer\$SysDrv\Documents and Settings"
								$Profiles = Get-ChildItem -Path $Script:ProfLoc
								$Script:NTUserDatLog = $Profiles | ForEach-Object -Process {$_.GetFiles("ntuser.dat.LOG")}
								
								#Function to grab last profile data, used for allowing -FilterSID to function properly.
								function GetLastProfData ($InstanceNumber)
									{
										$Script:LastProf = ($Script:NTUserDatLog | Sort-Object -Property LastWriteTime -Descending)[$InstanceNumber]							
										$Script:UserName = $Script:LastProf.DirectoryName.Replace("$Script:ProfLoc","").Trim("\").ToUpper()
										$Script:Time = $Script:LastProf.LastAccessTime
										
										#Getting the SID of the user from the file ACE to compare
										$Script:Sddl = $Script:LastProf.GetAccessControl().Sddl
										$Script:Sddl = $Script:Sddl.split("(") | Select-String -Pattern "[0-9]\)$" | Select-Object -First 1
										#Formatting SID, assuming the 6th entry will be the users SID.
										$Script:Sddl = $Script:Sddl.ToString().Split(";")[5].Trim(")")
										
										#Convert Account to SID to detect if profile is loaded via the remote registry
										$Script:TranSID = New-Object System.Security.Principal.NTAccount($Script:UserName)
										$Script:UserSID = $Script:TranSID.Translate([System.Security.Principal.SecurityIdentifier])
									}#End function GetLastProfData
								GetLastProfData -InstanceNumber 0
								
								#If the FilterSID equals the UserSID, rerun GetLastProfData and select the next instance
								If ($Script:UserSID -eq $FilterSID)
									{
										GetLastProfData -InstanceNumber 1
									}#End If ($Script:UserSID -eq $FilterSID)
								
								#If the detected SID via Sddl matches the UserSID, then connect to the registry to detect currently loggedon.
								If ($Script:Sddl -eq $Script:UserSID)
									{
										$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"Users",$Computer)
										$Loaded = $Reg.GetSubKeyNames() -contains $Script:UserSID.Value
										#Convert SID to Account for friendly display
										$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($Script:UserSID)
										$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
									}#End If ($Script:Sddl -eq $Script:UserSID)
								Else
									{
										$User = $Script:UserName
										$Loaded = "Unknown"
									}#End Else

							}#End If ($Build -le 6000)
						
						#Creating Custom PSObject For Output
						New-Object -TypeName PSObject -Property @{
							Computer=$Computer
							User=$User
							SID=$Script:UserSID
							Time=$Script:Time
							CurrentlyLoggedOn=$Loaded
							} | Select-Object Computer, User, SID, Time, CurrentlyLoggedOn
							
					}#End Try
					
				Catch
					{
						If ($_.Exception.Message -Like "*Some or all identity references could not be translated*")
							{
								Write-Warning "Unable to Translate $Script:UserSID, try filtering the SID `nby using the -FilterSID parameter."	
								Write-Warning "It may be that $Script:UserSID is local to $Computer, Unable to translate remote SID"
							}
						Else
							{
								Write-Warning $_
							}
					}#End Catch
					
			}#End Foreach ($Computer in $ComputerName)
			
	}#End Process
	
End
	{
		#Resetting ErrorActionPref
		$ErrorActionPreference = $TempErrAct
	}#End End

}# End Function Get-LastLogon

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

#region Function Get-LoggedOnUser
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
		Write-Log "Get session information for all logged on users."
		Write-Output -InputObject ([PSADT.QueryUser]::GetUserSessionInfo("$env:ComputerName"))
	}
	Catch {
		$errmsg = $_.Exception.Message
		Write-Log "Failed to get session information for all logged on users. `n$errmsg" -LogLevel 3
	}
}

function Environment-Check
{
    # Define our parent directory.
    $scriptroot = Get-ScriptDirectory
    $approot = split-path -Parent -Path $scriptroot

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
		# Our initial filename is defined in our Get-Scriptdirectory function
		$scriptname = "$($filename).ps1"
		$filename = "$($filename)-$now.log"
		$logFile = "$logPath\$filename"
		$envmsg = "Running from Task Sequence."
	}
	# Start Logging
	Start-Log -FilePath $logfile
	Write-Log $envmsg
	Write-Log "LogPath is set to $logpath"
	
	# See if there are any users currently logged on
	[psobject[]]$LoggedOnUserSessions = Get-LoggedOnUser
	[string[]]$usersLoggedOn = $LoggedOnUserSessions | ForEach-Object { $_.NTAccount }

	If ($usersLoggedOn) {
		#  Get account and session details for the logged on user session that the current process is running under. Note that the account used to execute the current process may be different than the account that is logged into the session (i.e. you can use "RunAs" to launch with different credentials when logged into an account).
		[psobject]$CurrentLoggedOnUserSession = $LoggedOnUserSessions | Where-Object { $_.IsCurrentSession }
		
		#  Get account and session details for the account running as the console user (user with control of the physical monitor, keyboard, and mouse)
		[psobject]$CurrentConsoleUserSession = $LoggedOnUserSessions | Where-Object { $_.IsConsoleSession }
		
		## Determine the account that will be used to execute commands in the user session when toolkit is running under the SYSTEM account
		#  If a console user exists, then that will be the active user session.
		#  If no console user exists but users are logged in, such as on terminal servers, then the first logged-in non-console user that is either 'Active' or 'Connected' is the active user.
		[psobject]$RunAsActiveUser = $LoggedOnUserSessions | Where-Object { $_.IsActiveUserSession }
	}
	
	If ($usersLoggedOn) {
		Write-Log "The following users are logged on to the system: [$($usersLoggedOn -join ', ')]."
		#  Check if the current process is running in the context of one of the logged in users
		If ($CurrentLoggedOnUserSession) {
			write-Log "Current process is running with user account [$ProcessNTAccount] under logged in user session for [$($CurrentLoggedOnUserSession.NTAccount)]."
			$userval = $CurrentLoggedOnUserSession.NTAccount
		}
		Else {
			write-Log "Current process is running under a system account [$ProcessNTAccount]."
		}
		
		#  Display account and session details for the account running as the console user (user with control of the physical monitor, keyboard, and mouse)
		If ($CurrentConsoleUserSession) {
			write-Log "The following user is the console user [$($CurrentConsoleUserSession.NTAccount)] (user with control of physical monitor, keyboard, and mouse)."
		}
		Else {
			write-Log 'There is no console user logged in (user with control of physical monitor, keyboard, and mouse).'
		}
		
		#  Display the account that will be used to execute commands in the user session when toolkit is running under the SYSTEM account
		If ($RunAsActiveUser) {
			write-Log "The active logged on user is [$($RunAsActiveUser.NTAccount)]."
		}
	}
	Else {
		write-Log 'No users are logged on to the system.'
		# Get our Last Logged on user value
		try{$userval = (Get-LastLogon).User.Value}
		catch{$userval = $null}
	}

	# Check for our userval value
	if($userval)
	{
		write-log "Recording Username $userval to the TS Variable..."
		$tsenv.Value('TSCusCurrentUser') = $userval
	}
	else {
		write-log "No Username value was returned." -LogLevel 3
	}
}

Environment-Check