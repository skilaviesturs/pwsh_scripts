<#
.SYNOPSIS
Skenējam tīkla segmentu/us ar ICMP un atgriežam sarakstu ar datoru vārdiem un IP adresēm, kas atbildēja.

.DESCRIPTION
Skripts skenē norādītos IP adrešu segmentus. Skripts atbalsta parametru padošanu pipeline

.PARAMETER Network
Norādam tīkla segmentu bez pēdējā punkta, piemēram, "192.168.0"

.INPUTS
Padodam IP adrešu segmentu, piemēram, "192.168.0", vai sarakstu

.OUTPUTS
Atgriež objektu ar DNSName, Address, Online

.EXAMPLE
ping-segment.ps1 "192.168.0"
Norādam bez parametra IP adrešu segmentu vai sarakstu

.EXAMPLE
ping-segment.ps1 -Network "192.168.0", "192.168.1"
Norādam ar parametru IP adrešu segmentu vai sarakstu

.EXAMPLE
Get-Content .\expo-segments.txt | .\ping-segment.ps1
Ielādējam IP adrešu segmentu sarakstu no datnes 

.NOTES
Author:	Viesturs Skila
Version: 1.0.1
#>
[CmdletBinding()] 
param(
	[Parameter(Mandatory,ValueFromPipeline)]
	[string[]]$Network
)

BEGIN {
	$Script:output = @()
	$Octs = (1..254)
	$jsonFile = ".\ping-segment.dat"

	if ( (Test-Path -Path $jsonFile -Type Leaf) ) {
		$Script:DataArchive = Get-Content -Path $jsonFile -Raw | ConvertFrom-Json
	}#endif
	else {
		$Script:DataArchive = @()
	}#endelse

	function Test-OnlineFast {
		param(
			[Parameter(Mandatory,ValueFromPipeline)]
			[string[]]$ComputerName,
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
						"N/A"
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

			$ComputerName | ForEach-Object {
				$null = $bucket.Add($_)
			}#endforeach
		}#EndOfProcess

		END {

			$query = $bucket -join "' or Address='"
			
			Get-CimInstance -ClassName Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" |
			Select-Object -Property $DNSName, Address, $IsOnline, $statusFriendlyText
		}#EndOfEnd
	}#endOffunction
	
	Function Set-DataArchive {
		param(
			[Parameter(Mandatory,ValueFromPipeline)]
			[Object[]]$Object
			)

		BEGIN {
			
			if ( $Object.count -eq 0 ) {
				Write-Host "Object is $null."
				return
			}#endif
			
		}#endOfBegin
		
		PROCESS {
			
			if ( $Script:DataArchive.count -eq 0) {
                    $Script:DataArchive += @{
                        Address		= $object.Address;
                        DNSName		= $object.DNSName;
                        AddDate		= Get-Date;
                    }
                }#endif
                elseif ( -NOT $Script:DataArchive.Path.Contains($_.FullName) ) {
                    $Script:DataArchive += @{
                        Address		= $object.Address;
                        DNSName		= $object.DNSName;
                        AddDate		= Get-Date;
                    }
                }#endelseif
				
		}#endOfProcess
		
		END{}
		
	}#endOfFunction
	
	
}#EndOfBegin

PROCESS {
	
	$result = @()
	$Computers = @()
	Write-Host -NoNewline "Scanning network: [$Network.0/24] ..."
	foreach ( $Oct in $Octs ) {
		$Computers += "$Network.$Oct"
	}#endforeach
	Write-Host -NoNewline "..."
	$result = $Computers | Test-OnlineFast | Where-Object -Property Status -eq "Success" | Where-Object -Property DNSName -notlike ''
	$Script:output += $result
	Write-Host -NoNewline "...`tfound [$(if ($result.count -le 9) {"  $($result.count)"} 
		elseif ($result.count -le 99) {" $($result.count)"} 
		else {"$($result.count)"} )]of[$($Octs.count)] objects.`n"

}#EndOfProcess

END {
	$Script:output | Sort-Object -Property DNSName
	Write-Host "Found [$($Script:output.count)] objects."
	
	$Script:output | Set-DataArchive $_
	
	if ( $Script:DataArchive.count -gt 0) {
		$Script:DataArchive | ConvertTo-Json | Out-File $jsonFile -Force
	}#endif
}#EndOfEnd
