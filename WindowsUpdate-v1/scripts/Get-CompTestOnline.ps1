<#
.SYNOPSIS
Skenējam datorvārdus, ip adreses ar ICMP un atgriežam sarakstu ar datoru vārdiem un IP adresēm, kas atbildēja.

.DESCRIPTION
Skripts skenē norādītos DNS vai IP adreses. Skripts atbalsta parametru padošanu pipeline

.PARAMETER ComputerNames
Norādam tīkla segmentu bez pēdējā punkta, piemēram, "192.168.0"

.EXAMPLE
Get-CompTestOnline.ps1 "192.168.0.2"
Norādam bez parametra 

.EXAMPLE
Get-CompTestOnline.ps1 -Network "192.168.0.4", "computer.ltb.lan"
Norādam ar parametru

.EXAMPLE
Get-Content .\expo-segments.txt | Get-CompTestOnline.ps1
Padodam parametru no pipeline

.NOTES
Author:	Viesturs Skila
Version: 1.1.2
#>
[CmdletBinding(DefaultParameterSetName = 'Name')]
param(
	[Parameter(Position = 0,
		ParameterSetName = 'Name',
		Mandatory=$true,
		ValueFromPipeline)]
	[ValidateNotNullOrEmpty()]
	[string[]]$Name,
	
	[Parameter(Position = 0,
		ParameterSetName = 'inPath',
		Mandatory=$true,
		ValueFromPipeline)]
	[ValidateScript( {
		if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
			Write-Host "File does not exist"
			throw
		}#endif
		return $True
	} ) ]
	[System.IO.FileInfo]$inPath,

	[Parameter(Mandatory = $False, ParameterSetName = 'Help')]
	[switch]$Help = $False
)

BEGIN {
	#Skripta tehniskie mainīgie
	$CurVersion		= "1.1.2"
	#Skritpa konfigurācijas datnes
	$__ScriptName	= $MyInvocation.MyCommand
	$__ScriptPath	= Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
	#Žurnalēšanai
	#$LogFileDir		= "log"
	#$LogFile		= "$LogFileDir\RemoteJob_$(Get-Date -Format "yyyyMMdd")"
	$jsonFile		= "$__ScriptPath\data\PingOnlineComp.dat"
	$LogObject =@()
	$Output = @()
	
	if ($Help) {
		Write-Host "`nVersion:[$CurVersion]`n"
		#$text = (Get-Command "$__ScriptPath\$__ScriptName" ).ParameterSets | Select-Object -Property @{n = 'Parameters'; e = { $_.ToString() } }
		$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
		$text | ForEach-Object { Write-Host $($_) }
		Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
		Exit
	}#endif

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
		Write-Host "`r[TestOnline] done in $(
			if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
			elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
			else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
			)" -ForegroundColor Yellow -BackgroundColor Black
	}#endOffunction
	function Test-OnlineFast {
		param(
			[Parameter(Mandatory,ValueFromPipeline)]
			[string[]]$ComputersName,
			$TimeoutMillisec = 2000
		)

		BEGIN{

			[Collections.ArrayList]$bucket = @()
	
			$StatusCode_ReturnValue = @{
				0='Success'
				11001='Buffer Too Small'
				11002='Destination Net Unreachable'
				11003='Destination Host Unreachable'
				11004='Destination Protocol Unreachable'
				11005='Destination Port Unreachable'
				11006='No Resources'
				11007='Bad Option'
				11008='Hardware Error'
				11009='Packet Too Big'
				11010='Request Timed Out'
				11011='Bad Request'
				11012='Bad Route'
				11013='TimeToLive Expired Transit'
				11014='TimeToLive Expired Reassembly'
				11015='Parameter Problem'
				11016='Source Quench'
				11017='Option Too Big'
				11018='Bad Destination'
				11032='Negotiating IPSEC'
				11050='General Failure'
			}

			$statusFriendlyText = @{
				Name = 'Status'
				Expression = { 
					if ( $null -eq $_.StatusCode ) {
						"Unknown"
					}#endif
					else {
						$StatusCode_ReturnValue[([int]$_.StatusCode)]
					}#endelse
				}#endExpr
			}#endparam

			$IsOnline = @{
				Name = 'Online'
				Expression = { $_.StatusCode -eq 0 }
			}#endparam

			$DNSName = @{
				Name = 'DNSName'
				Expression = { if ($_.StatusCode -eq 0) { 
						if ($_.Address -like '*.*.*.*') 
						{ [Net.DNS]::GetHostByAddress($_.Address).HostName  } 
						else  
						{ [Net.DNS]::GetHostByName($_.Address).HostName  } 
					}#endif
				}#endExpr
			}#endparam
		}#EndOfBegin
	
		PROCESS {

			$ComputersName | ForEach-Object {
				$null = $bucket.Add($_)
			}#endforeach
		}#EndOfProcess

		END {

			$query = $bucket -join "' or Address='"
			
			Get-CimInstance -ClassName Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" |
			Select-Object -Property $DNSName, Address, $IsOnline, $statusFriendlyText, StatusCode
		}#EndOfEnd
	}#endOffunction

	if ( $PSCmdlet.ParameterSetName -eq 'Name' ) {
		$Computers = $Name | Get-Unique
	}#endif
	else {
		$Computers = Get-Content -Path $InPath | Where-Object { $_ -ne "" } `
		| Where-Object { -not $_.StartsWith('#') }  | Sort-Object | Get-Unique
	}#endelse

	if ( (Test-Path -Path $jsonFile -Type Leaf) ) {
		$Script:DataArchive = Get-Content -Path $jsonFile -Raw | ConvertFrom-Json
	}#endif
	else {
		$Script:DataArchive = @()
	}#endelse
	Write-Host "[TestOnline] got for testing [$($Computers.Count)] $(if($Computers.Count -eq 1){"computer"}else{"computers"})"
	Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
	$jobWatch = [System.Diagnostics.Stopwatch]::startNew()
}#EndOfBegin

PROCESS {
	
	$result = @()
	Write-Host -NoNewline "." -ForegroundColor Yellow -BackgroundColor Black
	$result = $Computers | Test-OnlineFast
	$resultAll += @($result)

}#EndOfProcess

END {
	$resultAll | Where-Object -Property Status -eq "Success" | Where-Object -Property DNSName -notlike '' | ForEach-Object {
		$__comp = [string]$_.Address
		try {
			$probeComp = Get-CimInstance -ClassName win32_networkadapterconfiguration -ComputerName $_.DNSName -ErrorAction Stop `
			| Where-Object -Property IPEnabled -eq 'true' `
			| Where-Object -Property DefaultIPGateway -ne $null `
			| Select-Object IPAddress, MacAddress
			$probeComp | Add-Member -MemberType NoteProperty -Name WinRMservice -Value $True
			$Output += @( New-Object -TypeName psobject -Property @{
					PipedName	= [string]$_.Address.ToLower();
					DNSName		= [string]$_.DNSName.ToLower();
					AddDate		= [System.DateTime](Get-Date);
					MacAddress	= [string]$probeComp.MacAddress;
					IPAddress	= [string]$probeComp.IPAddress[0];
					WinRMservice = $true;
				}#endobject
				)
			}#endtry
			catch {
				$LogObject += @("[TestOnline] [ERROR] WinRM service of [$__comp][$([System.Net.Dns]::GetHostAddresses($__comp))] is not accessible.")
		}#endcatch
	}#endforeach

	$resultAll | Where-Object -Property StatusCode -ne 0 | ForEach-Object {
		if ( $_.Status -eq 'Unknown' ) {
			$LogObject += @("[TestOnline] [$($_.Address)][not found such host]`t- $($_.Status)")
		}#endif
		else {
			$LogObject += @("[TestOnline] [$($_.Address)][$([System.Net.Dns]::GetHostAddresses($($_.Address)))]`t- $($_.Status)")
		}#endelse
	}#endforeach
	
	$result = Stop-Watch -Timer $jobWatch -Name JOBers
	$LogObject | ForEach-Object { Write-Warning "$_" }

	if ( $Output.count -gt 0) {
		$Output | ConvertTo-Json | Out-File $jsonFile -Force
		return $Output
	}#endif
	else {
		return $false
	}
}#EndOfEnd
