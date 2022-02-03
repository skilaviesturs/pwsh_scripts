<#Get-ExpoUpdateServer.ps1
.SYNOPSIS
This script checks computer on Windows updates, schedule Windows update TaskJob

.DESCRIPTION
This script checks computer on Windows updates, schedule Windows update TaskJob

.EXAMPLE
On local computer:
PS> Get-ExpoUpdateServer.ps1 [-install<SwitchParameter>][-version<SwitchParameter>][<CommonParameters>]

On remote computer:
PS> Invoke-Command -ComputerName [Name] -FilePath ".\Get-ExpoUpdateServer.ps1" -Verbose -ArgumentList [Install,InstallAutoReboot]

.NOTES
	Author:	Viesturs Skila
	Version: 1.1.4
#>
[CmdletBinding()] 
param (
	[Parameter(Position = 0,Mandatory=$true)]
	[switch]$Install,
	[Parameter(Position = 1,Mandatory=$true)]
	[switch]$InstallAutoReboot,
    [switch]$version
)

begin {
#    Clear-Host
	Import-Module PSWindowsupdate
	$CurVersion = "1.1.4"
	if ( $version ) {
		Write-Host "`n$CurVersion`n"
		Exit
	}
    $ScriptWatch = [System.Diagnostics.Stopwatch]::startNew()
	$Global:LogFile = "C:\ExpoWindowsUpdate$(Get-Date -Format "yyyyMMdd").log"
	$Computername = $env:COMPUTERNAME
	$Script:ThereIsUpdates = 0
	$PendingReboot = $False
	$Script:forReport = @()
	if ($null -ne $using) {
        # $using is only available if this is being called with a remote session
        $VerbosePreference = $using:VerbosePreference
    }

	$ObjectReturn = New-Object -TypeName psobject -Property @{
		Computer		= $Computername ;
		ReturnLogs		= $null ;
		PendingReboot	= $null ;
		Updates			= $null ;
		AutoReboot		= $false ;
		ScheduledTask	= $null	;
		ErrorMsg		= $null	;
	}

	#Looking for old log files; delete them if found
	if ( Test-Path -Path "C:\*" -Include "ExpoWindowsUpdate*" -PathType Leaf ) {
		Remove-Item -Path "C:\*" -Include "ExpoWindowsUpdate*"
	}

	Function Write-msg { 
		[CmdletBinding(DefaultParameterSetName="default")]
		[alias("wrlog")]
		Param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$text,
			[switch]$log,
			[switch]$bug
		) #param
		try {
			#write-debug "[wrlog] Log path: $log"
				if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO'}
				$timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
				if ( $log -and $bug ) {
					Write-Warning $flag"`t"$text	
					Write-Output "$timeStamp | $flag | $text" | Out-File $Global:LogFile -Append -ErrorAction Stop
					$Script:forReport += @("$timeStamp | $flag | $text")
					} elseif ( $log ) {
						Write-Verbose $flag"`t"$text
						Write-Output "$timeStamp | $flag | $text" | Out-File $Global:LogFile -Append -ErrorAction Stop
						$Script:forReport += @("$timeStamp | $flag | $text")
					} else {
						Write-Verbose $flag"`t"$text
						$Script:forReport += @("$timeStamp | $flag | $text")
				} #else
		}
		catch {
			Write-Warning "[Write-msg] $($_.Exception.Message)"
			return
		}#endOftry
	}#endOffunction
	Function Stop-Watch {
        [CmdletBinding()] 
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [object]$Timer,
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        $Timer.Stop()
        if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)"} else { $bMin = "$($Timer.Elapsed.Minutes)" }
        if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)"} else { $bSec = "$($Timer.Elapsed.Seconds)" }
        if ($Name -notlike 'JOBers') {
			Write-msg -log -text "[$Name] finished in $(
				if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
				elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
				else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
				)"
		}
		else {
			Write-Host " done in $(
				if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
				elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
				else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
				)" -ForegroundColor Yellow -BackgroundColor Black
		}
    }#endOffunction
	
    Write-msg -log -text "[-----] Script started in [$(if ($install) { if ( $InstallAutoReboot ) {"Install-AutoReboot"} else {"Install-IgnoreReboot"} } 
		else {"Check"})] mode"
	
    function Get-PSVersion {
		[OutputType('bool')]
	    $version=$PSVersionTable.PSVersion.Major
	    if ($version -ge 5){return $True;} else {return $False;}
    }
	
    function Get-PendingUpdates {
		try {
			$UpdateResults = @()
			$updatesession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computername))
			$UpdateSearcher = $updatesession.CreateUpdateSearcher()
			$searchresult = $updatesearcher.Search("IsInstalled=0") # 0 = NotInstalled | 1 = Installed
			$Script:ThereIsUpdates = $searchresult.Updates.Count
			if ( $Script:ThereIsUpdates -gt 0 ) {
				#Updates are waiting to be installed
				Write-msg -log -text "$Computername | Summary | [$Script:ThereIsUpdates] $(if ( $Script:ThereIsUpdates -gt 1 ) {"are"} else {"is"} ) waiting to be installed"   
				foreach ( $entry in $searchresult.Updates ) {
					Write-msg -log -text "$Computername | Update | KB$($entry.KBArticleIDs) Title:$($entry.Title)"
					$UpdateResults += @("$Computername | KB$($entry.KBArticleIDs) | Title:$($entry.Title)")
				}#foreach
			} else {
				Write-msg -log -text "$Computername | Update | False"
				$UpdateResults = "$Computername | Update | False"
			}#if-condition
		} catch {
			Write-msg -log -bug -text "[listPendingUpdates] $($_.Exception.Message)"
		}#endOftry
		return $UpdateResults
    }#endOffunction

	Function Test-RegistryKey {
        [OutputType('bool')]
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Key
        )
    
		if ( Get-Item -Path $Key -ErrorAction Ignore ) { $true }
    }#endOffunction

	Function Test-RegistryValue {
        [OutputType('bool')]
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Key,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Value
        )

		if ( Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore ) { $true }
    }#endOffunction

	Function Test-RegistryValueNotNull {
        [OutputType('bool')]
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Key,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Value
        )
	
		if ( ($regVal = Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore  ) -and $regVal.($Value)) { $true }
    }#endOffunction

	Function Get-RebootRequired {
		[OutputType('bool')]
		$tests = @(
			{ Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' }
			{ Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress' }
			{ Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' }
			{ Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending' }
			{ Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting' }
		 <#       { Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations' }
			{ Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations2' }
			{ 
				# Added test to check first if key exists, using "ErrorAction ignore" will incorrectly return $true
				'HKLM:\SOFTWARE\Microsoft\Updates' | Where-Object { test-path $_ -PathType Container } | ForEach-Object {            
					(Get-ItemProperty -Path $_ -Name 'UpdateExeVolatile' | Select-Object -ExpandProperty UpdateExeVolatile) -ne 0 
				}
			}
			{ Test-RegistryValue -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Value 'DVDRebootSignal' }
	
			{ Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttemps' }
			{ Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'JoinDomain' }
			{ Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'AvoidSpnSet' }
			{
				# Added test to check first if keys exists, if not each group will return $Null
				# May need to evaluate what it means if one or both of these keys do not exist
				( 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' | Where-Object { test-path $_ } | ForEach-Object { (Get-ItemProperty -Path $_ ).ComputerName } ) -ne 
				( 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' | Where-Object { Test-Path $_ } | ForEach-Object { (Get-ItemProperty -Path $_ ).ComputerName } )
			}
		
			{
				# Added test to check first if key exists
				'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending' | Where-Object { 
					(Test-Path $_) -and (Get-ChildItem -Path $_) } | ForEach-Object { $true }
			}
		# #>
		)
		foreach ($test in $tests) {
			#Write-msg -text "Running scriptblock: [$($test.ToString())]"
			if (& $test) {
				$true
				break
			}
		}#endForEach
	}#endOffunction

}#endOfbegin

process {
	if ( Get-PSVersion ){
		$PendingReboot = Get-RebootRequired
		$ObjectReturn.PendingReboot = "$Computername | Reboot | $(if ( $PendingReboot ) {"[Pending] reboot"} else {"False"})"
		$ObjectReturn.Updates = Get-PendingUpdates
            if ( $Install ) {
				if ( Test-Path -Path "C:\*" -Include "*SheduledWUjob*" -PathType Leaf ) {
					Remove-Item -Path "C:\*" -Include "*SheduledWUjob*" -ErrorAction Ignore
				}
                if ( $Script:ThereIsUpdates -gt 0 ) {
					
					if ( $InstallAutoReboot ) {
						$Script = { Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -AutoReboot -Verbose | Out-File "C:\ExpoSheduledWUjob.log" -Force }
						Invoke-WUjob -Script $Script -RunNow -Confirm:$false
						Write-msg -log -text "$Computername | ScheduledTask | Job has been scheduled - auto-restart [$InstallAutoReboot]"
						$ObjectReturn.ScheduledTask = "$Computername | ScheduledTask | Job has been scheduled - auto-restart [$InstallAutoReboot]"
						$ObjectReturn.AutoReboot = $True
					} 
					else {
						$Script = { Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -IgnoreReboot -Verbose | Out-File "C:\ExpoSheduledWUjob.log" -Force }
						Invoke-WUjob -Script $Script -RunNow -Confirm:$false
						$ObjectReturn.ScheduledTask = "$Computername | ScheduledTask | Job has been scheduled - auto-restart [$InstallAutoReboot]"
						Write-msg -log -text "$Computername | ScheduledTask | Job has been scheduled - auto-restart [$InstallAutoReboot]"
					}#endelse
				}#endif
				else {
					Write-msg -log -text "$Computername | ScheduledTask | no updates. end."
					$ObjectReturn.ScheduledTask = "$Computername | ScheduledTask | no updates. end."
				}
			} elseif ( $PendingReboot -and -not ( $Install ) ) {
				Write-msg -log -text "$Computername | Reboot | $(if ( $PendingReboot ) {"True"} else {"False"})"
            }#endif
	} else {
	Write-msg -log -bug -text "Script's prerequisite is running PowerShell version 5.0 on computer [$Computername]."
	$ObjectReturn.ErrorMsg = "Script's prerequisite is running PowerShell version 5.0 on computer [$Computername]."
	} #endif

}#endOfbegin

end {
    Stop-Watch -Timer $scriptWatch -Name Script
	$ObjectReturn.ReturnLogs = $Script:forReport
	return $ObjectReturn
}#endOfend
