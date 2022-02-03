<#Get-ExpoCompAsset.ps1
.SYNOPSIS
This script checks computer's hardware, operational systems paremeters

.DESCRIPTION
This script checks computer's hardware, operational systems paremeters

.EXAMPLE
On local computer:
Get-ExpoAsset [-ServicesList<String>][-version<SwitchParameter>][<CommonParameters>]

On remote computer:
PS> Invoke-Command -ComputerName [Name] -FilePath ".\Get-ExpoAsset.ps1" -Verbose -ArgumentList [-ServicesList<String>]

.NOTES
	Author:	Viesturs Skila
	Version: 1.0.5
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [switch]$Name,
    [switch]$version
    )

begin {
    #Clear-Host
    $CurVersion = "1.0.5"
	if ( $version ) {
		Write-Host "`n$CurVersion`n"
		Exit
	}#endif
    $StopWatch = [System.Diagnostics.Stopwatch]::startNew()
    $Global:LogFile = "C:\ExpoAssets$(Get-Date -Format "yyyyMMdd").log"
	$Computername = $env:COMPUTERNAME
    $isServiceList = ( -not [string]::IsNullOrEmpty($ServicesToCheck) )

    #Looking for old log files; delete them if found
	if ( Test-Path -Path "C:\*" -Include "ExpoAssets*" -PathType Leaf ) {
		Remove-Item -Path "C:\*" -Include "ExpoAssets*"
	}#endif
  
    <# ---------------------
        Define template CustomObject's
    #>
    
    $Object = [PSCustomObject][ordered]@{
        'aComputerName' = $Computername
        'bOS' = $null
        'bWinVersion' = $null
        'bVersion' = $null
        'bProductType' = $null
        'cRAM' = $null
        'dCPU' = $null
        'dCPUCores' = $null
        'dCPULogical' = $null
        'eBiosManufacturer' = $null
        'eBiosSeralNumber' = $null
        'fManufacturer' = $null
        'fSystemFamily' = $null
        'fModel' = $null

        'gCsIPv4' = $null
        'gDNS1' = $null
        'gDNS2' = $null

        'hDiskType' = $null
        'hDiskHealth' = $null
        'hDiskModel' = $null
    }
    #>
    $Services = [PSCustomObject]@{
        'aComputerName' = $Computername
    }
    $Disks = [PSCustomObject]@{
        'aComputerName' = $Computername
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
					Write-Output "$timeStamp`t$flag`t$text" | Out-File $Global:LogFile -Append
					} elseif ( $log ) {
						Write-Verbose $flag"`t"$text
						Write-Output "$timeStamp`t$flag`t$text" | Out-File $Global:LogFile -Append
					} else {
						Write-Verbose $flag"`t"$text
				} #else
			
		}
		catch {
			Write-Warning "[Write-msg] $($_.Exception.Message)"
			Exit 1
		}
	} #endOffunction
    
} #endOfbegin

PROCESS {

    Write-msg -log -text "[-----] Script started"
    $WinCurrentVersion = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

    #ProductType
    $varProductType = (Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property ProductType).ProductType
    switch ($varProductType) {
        1 { $Object.bProductType = 'Workstation' }
        2 { $Object.bProductType = 'DomainController' }
        3 { $Object.bProductType = 'Server' }
        default { $Object.bProductType = 'unknown' }
    }    

    if ( $WinCurrentVersion.CurrentMajorVersionNumber -ge 6.3 -or $WinCurrentVersion.CurrentVersion -ge 6.3 ) {
        $info = Get-ComputerInfo
        $Object.bOS = $info.WindowsProductName
        if ( ($info.CsProcessors.Name).count -eq 1 ) {
            $Object.dCPU = $info.CsProcessors.Name
        } else {
            $Object.dCPU = $info.CsProcessors.Name[0]
        }
        $Object.dCPUCores = ($info.CsProcessors.NumberOfCores | Measure-Object -Sum).Sum
        $Object.dCPULogical = ($info.CsProcessors.NumberOfLogicalProcessors | Measure-Object -Sum).Sum
        $Object.eBiosManufacturer = $info.BiosManufacturer
        $object.eBiosSeralNumber = $info.BiosSeralNumber
        $Object.fManufacturer = $info.CsManufacturer
        $Object.fModel = $info.CsModel
        
        if ( $Object.bProductType -eq 'Workstation' ) {
            $Object.bVersion = "$($info.OsVersion).$($WinCurrentVersion.ubr)"
            $Object.fSystemFamily = $info.CsSystemFamily
            $disk = Get-PhysicalDisk
            $pcDeviceId = ( Get-PhysicalDisk -DeviceNumber 0 | Get-StorageReliabilityCounter ).DeviceID
            
            $Object.hDiskType = $disk.MediaType
            $Object.hDiskHealth = ( Get-PhysicalDisk -DeviceNumber $pcDeviceId | Get-StorageReliabilityCounter).Wear
            $Object.hDiskModel = $disk.Model
        }
    }
    
    #get BIOS
    #$bios = Get-CimInstance Win32_BIOS | Select-Object *

    #get installed programs
    #$programms = Get-CimInstance win32_product | Select-Object Name,InstallDate

    $Object.cRAM = ( Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum /1GB
    $network = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'"
    $Object.bWinVersion = (Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property Version).Version
    
    $Object.gCsIPv4 = $network.IPAddress[0]
    $Object.gDNS1 = $network.DNSServerSearchOrder[0]
    $Object.gDNS2 = $network.DNSServerSearchOrder[1]
    
        
	# Looking for servers's disk drives and each drive size and free space
    $devices = Get-CimInstance -ClassName win32_LogicalDisk -Filter "DriveType = '3'" -Property DeviceID, Size, FreeSpace
    foreach ($device in $devices) {
        $driveSizeID = "hDisk $($device.DeviceID) size"
        $freeSpaceID = "hDisk $($device.DeviceID) size free"
        $dSize = [Math]::Round( ( $device.Size / 1GB ),2)
        $dFreeSpace = [Math]::Round( ( $device.FreeSpace / 1GB ),2 )
    
        if ( Get-Member -InputObject $Disks -Name $driveSizeID ) { $Disks.$driveSizeID = $dSize }
            else { $Disks | Add-Member -MemberType NoteProperty -Name $driveSizeID -Value $dSize }
        if ( Get-Member -InputObject $Disks -Name $freeSpaceID) { $Disks.$freeSpaceID = $dFreeSpace }
            else { $Disks | Add-Member -MemberType NoteProperty -Name $freeSpaceID -Value $dFreeSpace }
    } #endforeach

    if ( $isServiceList ) {
        foreach ( $service in $ServicesToCheck ) {
            if ( Get-CimInstance -ClassName Win32_Service -filter "Name like '%$service%'" ) { 
                $Services."$service" = 'True'
            } else {
                $Services."$service" = 'False'
            }#endif
        }#endforeach
    }#endif

    #   $tempToExcel += Join-Object -Left $SrvInfo -Right $Services -LeftJoinProperty ComputerName -RightJoinProperty ComputerName -Type OnlyIfInBoth
    
    # Add $Disks info to merged table
    $OutputObject += Join-Object -Left $Object -Right $Disks -LeftJoinProperty aComputerName -RightJoinProperty aComputerName -Type OnlyIfInBoth

    if ( $Name ) {
        $OutputObject | Format-List * | Out-String -Stream | 
            Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" }
    }#endif
    else {
        return $OutputObject
    }#endelse

}#endOfPROCESS

END {
    $stopwatch.Stop()
    Write-msg -log -text "[-----] Script finished in $([math]::Round(($stopwatch.Elapsed).TotalSeconds,3)) seconds."
}#endOfEND
