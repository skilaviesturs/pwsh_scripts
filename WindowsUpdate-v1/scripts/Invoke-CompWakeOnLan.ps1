<#
.SYNOPSIS
Skripts datu bāzē atrod nepieciešamos parametrus datora attālinātai sāknēšanai

.PARAMETER ComputerName
Datora vārds

.PARAMETER DataArchiveFile
Datu bāzes faila atrašanās vieta

.NOTES
	Author:	Viesturs Skila
	Version: 1.1.1
#>
[CmdletBinding(DefaultParameterSetName = 'Name')]
param (
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline, ParameterSetName = 'Name',
        HelpMessage = "Name of computer"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName,

    [Parameter(Position = 1, Mandatory = $true, ParameterSetName = 'Name',
        HelpMessage = "Path to archive file."
    )]
    [System.IO.FileInfo]$DataArchiveFile,

    [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Name',
        HelpMessage = "Path to archive file."
    )]
    [System.IO.FileInfo]$CompTestOnlineFile,

    [Parameter(Position = 0, Mandatory = $true,
        ParameterSetName = 'Help'
    )]
    [switch]$Help
)
BEGIN {
    <# ---------------------------------------------------------------------------------------------------------
	Skripta konfigurācijas datnes
	--------------------------------------------------------------------------------------------------------- #>
    $CurVersion = "1.1.1"
    $scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    #Žurnalēšanai
    #$LogFileDir		= "log"
    #$LogFile		= "$LogFileDir\RemoteJob_$(Get-Date -Format "yyyyMMdd")"

    $LogObject = @()

    if ($Help) {
        Write-Host "`nVersion:[$CurVersion]`n"
        $text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
        $text | ForEach-Object { Write-Host $($_) }
        Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
        Exit
    }#endif
    Function Stop-Watch {
        [CmdletBinding()] 
        param 
        (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [object]$Timer,
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        $LogObject = @()
        $Timer.Stop()
        if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)" } else { $bMin = "$($Timer.Elapsed.Minutes)" }
        if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)" } else { $bSec = "$($Timer.Elapsed.Seconds)" }
        $LogObject += @( "[$Name] finished in $(
            if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
            elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
            else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
            )"
        )
        return $LogObject
    }#endOffunction
    function Invoke-WakeOnLan {
        param
        (
            # one or more MACAddresses
            [Parameter(Mandatory = $true, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            # mac address must be a following this regex pattern:
            [ValidatePattern('^([0-9A-F]{2}[:-]){5}([0-9A-F]{2})$')]
            [string[]]
            $MacAddress 
        )
 
        begin {
            $LogObject = @()
            # instantiate a UDP client:
            $UDPclient = [System.Net.Sockets.UdpClient]::new()
        }
        process {
            foreach ($_ in $MacAddress) {
                try {
                    $currentMacAddress = $_
        
                    # get byte array from mac address:
                    $mac = $currentMacAddress -split '[:-]' |
                    # convert the hex number into byte:
                    ForEach-Object {
                        [System.Convert]::ToByte($_, 16)
                    }
 
                    #region compose the "magic packet"
                    # create a byte array with 102 bytes initialized to 255 each:
                    $packet = [byte[]](, 0xFF * 102)
        
                    # leave the first 6 bytes untouched, and
                    # repeat the target mac address bytes in bytes 7 through 102:
                    6..101 | Foreach-Object { 
                        # $_ is indexing in the byte array,
                        # $_ % 6 produces repeating indices between 0 and 5
                        # (modulo operator)
                        $packet[$_] = $mac[($_ % 6)]
                    }
                    #endregion

                    # connect to port 400 on broadcast address:
                    $UDPclient.Connect(([System.Net.IPAddress]::Broadcast), 4000)
        
                    # send the magic packet to the broadcast address:
                    $null = $UDPclient.Send($packet, $packet.Length)
                    $LogObject += @("[Waker] [SUCCESS] sent magic packet to $currentMacAddress from [$($env:computername)]...")
                }#endtry
                catch {
                    $LogObject += @("[Waker] [ERROR] unable to send ${mac}: $_" )
                }#endcatch
            }#endforeach
        }
        end {
            # release the UDF client and free its memory:
            $UDPclient.Close()
            $UDPclient.Dispose()
            return $LogObject 
        }#endOfend
    }#endOffunction

    if ( Test-Path $DataArchiveFile -PathType Leaf ) {
        try {
            $DataArchive = @(Import-Clixml -Path $DataArchiveFile -ErrorAction Stop)
        }#endtry
        catch {
            $DataArchive = @()
        }#endcatch
    }#endif
    else {
        $DataArchive = @()
    }#endif
    #$LogObject += @("[Waker] [INFO] got:[$($ComputerName.Count)]")
    #$LogObject += @("[Waker] [INFO] got DataArchive:[$($DataArchive.Count)]")
}#endOfBEGIN

PROCESS {
    [string]$HostDNSName = $null
    [string]$TargetMacAddress = $null
    try {
        [string]$IPAddress = [System.Net.Dns]::GetHostAddresses($ComputerName)
        [string]$Mask = $IPAddress -match '^([0-9]{1,3}[.]){2}([0-9]{1,3})'
    }#endtry
    catch {
        $LogObject += @("[Waker] [ERROR] Computer [$ComputerName] is not registered in DNS. Exit.")
    }
    #Atrodam IP adreses segmentu xxx.xxx.xxx
    if ( $Mask ) {
        [string]$Pattern = $matches[0]
        #$LogObject += @("[Waker] [INFO] [$ComputerName] belongs to net segment [$Pattern]")
        if ( $DataArchive.Count -gt 0 ) {
            #Atrodam arhīvā mērķa datora ierakstu, lai noteiktu mac adresi
            foreach ( $rec in $DataArchive ) {
                if ( $rec.DNSName -like $ComputerName -or $rec.PipedName -like $ComputerName ) {
                    $TargetMacAddress = $rec.MacAddress
                    $LogObject += @("[Waker] [INFO] [$ComputerName] has mac address [$TargetMacAddress]")
                    break
                }#endif
            }#endforeach
            #Atrodam arhīvā datoru, kas atrodas tajā pašā segmentā, lai no tā varētu pamodināt guļošo
            foreach ( $rec in $DataArchive ) {
                if ( $rec.IPAddress -match "$Pattern`*" -and ( $rec.DNSName -notlike $ComputerName -or $rec.PipedName -notlike $ComputerName ) ) {
                    #Pārbaudam vai atrastais remote dators ir online
                    #$LogObject += @("[Waker] [INFO]  [Invoke-Expression] `& `"$CompTestOnlineFile`" `-Name $($rec.DNSName) ")
                    $OnlineRemoteComps = Invoke-Expression "& `"$CompTestOnlineFile`" `-Name $($rec.DNSName) "
                    #Pārbaudam vai atrastā remote datora WinRM serviss darbojas
                    if ( $OnlineRemoteComps.WinRMservice ) {
                        $HostDNSName = $rec.DNSName
                        $LogObject += @("[Waker] [INFO] found online neighbor [$HostDNSName]:[$($rec.IPAddress)] on the same net.")
                        break
                    }#endif
                }#endif
            }#endforeach
            if ( [string]::IsNullOrWhitespace($HostDNSName) ) {
                $LogObject += @("[Waker] [ERROR] there is no entry for the computer on the same net [$Pattern] in the database. Exit.")
            }#endif
            elseif ( [string]::IsNullOrWhitespace($TargetMacAddress) ) {
                $LogObject += @("[Waker] [ERROR] there is no entry for the computer [$ComputerName] mac address in the database. Exit.")
            }#endelseif
            else {
                $LogObject += @("[Waker] [INFO] going to WakeOnLan [$TargetMacAddress] from remote host [$HostDNSName].")
                $result = Invoke-Command -Computername $HostDNSName -ScriptBlock ${Function:Invoke-WakeOnLan} -ArgumentList $TargetMacAddress
                $result | ForEach-Object { $LogObject += @($_) }
                <#Veicam ping
                #Invoke-Expression "& ping.exe -a -4 -n 50 $ComputerName "
                #Pārbaudam vai dators ir on-line
                #Write-Verbose "[Waker]:[Invoke-Expression] `& `"$CompTestOnlineFile`" `-Name $ComputerName "
                $WokenComp = Invoke-Expression "& `"$CompTestOnlineFile`" `-Name $ComputerName "
                $WokenComp | Format-Table * -AutoSize | Out-String -Stream | Where-Object { $_ -ne "" } | ForEach-Object {  Write-Host "$_" }
                #>
            }#endelse
        }#endif
    }#endif

}#endOfPROCESS

END {
    $result = Stop-Watch -Timer $scriptWatch -Name Script
    $result | ForEach-Object { $LogObject += @($_) }
    $Output = @()
    $i = 0
    $LogObject | ForEach-Object {
        $Output += @( New-Object -TypeName psobject -Property @{
                id       = $i;
                Computer = [string]$ComputerName;
                Message  = [string]$_;
            }#endobject
        )
        $i++
    }#endforeach

    return $Output
}#endOfEND