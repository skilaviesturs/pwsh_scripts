<#
.SYNOPSIS
Skripts atvieglo administratora windows jaunināšanas procesu, tehnisko parametru apkopošanu, attālināto darbstaciju startēšanu, pārsāknēšanu un izslēgšanu.

.DESCRIPTION
Skripts nodrošina:
[*] Windows jauninājumu pārbaudi, uzstādīšanu un datortehnikas pārsāknēšanu, ja to pieprasa jauninājums,
[*] pārbaudi datortehnikas nepieciešamībai uz pārsāknēšanu,
[*] attālināto datortehnikas sāknēšanu, restartu un izslēgšanu,
[*] datortehnikas tehnikos parametru un uzstādītās programmatūras pārskata izveidi.

.PARAMETER Name
Obligāts lauks.
Norādam datora vārdu.

.PARAMETER InPath
Obligāts lauks.
Norādam datoru saraksta datnes atrašanās vietu.

.PARAMETER Check
Kopā ar [-Name] vai [-InPath].
Veic konkrētā vai sarakstā norādīto datoru pārbaudi uz windows jauninājumiem. Rezultātu izvada ekrānā

.PARAMETER Update
Kopā ar [-Name] vai [-InPath].
Windows jauninājumu uzstādīšana.

.PARAMETER AutoReboot
Tikai kopā ar [-Update].
Automātisks datortehnikas restarts, ja jauninājums to pieprasa.

.PARAMETER WakeOnLan
Tikai kopā ar [-Name].
Veic norādītās datortehnikas pamodināšanu ar Magic paketes palīdzību.

.PARAMETER Stop
Tikai kopā ar [-Name].
Attālināti apstādina (shutdown) norādīto datoru.

.PARAMETER Reboot
Tikai kopā ar [-Name].
Attālināti restartē norādīto datoru un gaida, kamēr dators būs gatavs Powershell komandu izpildei

.PARAMETER NoWait
Tikai kopā ar [-Reboot].
Negaida, kamēr dators veiks pārsāknēšanas procedūru.

.PARAMETER Trace
Tikai kopā ar [-Name].
Tiešaistē seko līdzi Windows update žurnalēšanas datnes satura izmaiņām.

.PARAMETER EventLog
Tikai kopā ar [-Name] vai [-InPath]
Veic sarakstā norādīto serveru pārbaudi uz windows jauninājumu notikumiem datoru sistēmas notikumu žurnālā.

.PARAMETER Days
Tikai kopā ar [-EventLog]
Norādam par kādu periodu pagātnē tiek skatīti notikumi. Pēc noklusējuma - 30 dienas.

.PARAMETER OutPath
Tikai kopā ar [-InPath] un [-EventLog]
Norāda datnes vārdu, kurā tiks ierakstīts skripta rezultāts. Ja parametrs nav norādīts, rezultāts tiek izvadīts uz ekrāna.

.PARAMETER Asset
Tikai kopā ar [-Name].
Sagatavo datora tehnisko parametru un uzstādītās programmatūras un to versiju pārskatu.

.PARAMETER Include
Tikai kopā ar [-Asset]
Atlasa programmatūru pēc norādītā paterna. Atbalsta wildcard parametrus - *,?,[a-z] un [abc].
Vairāk informācijas šeit: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-5.1

.PARAMETER Exclude
Tikai kopā ar [-Asset]
Atlasa programmatūru, izņemot norādītajam paternam atbilstošo. Atbalsta wildcard parametrus - *,?,[a-z] un [abc].
Vairāk informācijas šeit: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-5.1

.PARAMETER Hardware
Tikai kopā ar [-Asset]
Pārskatā iekļauj datortehnikas tehniskos parametrus.

.PARAMETER NoSoftware
Tikai kopā ar [-Asset]
Pārskatā neiekļauj uzstādīto programmatūru.

.PARAMETER Install
Tikai kopā ar [-Name] vai [-InPath]
Norādam programmatūras pakotnes atrašanās vietu un tā tiek uzstādīta uz uz datortehnikas. Uzstādīšanas žurnalēšanas datne atrodama C:/temp mapē.

.PARAMETER Uninstall
Tikai kopā ar [-Name] vai [-InPath]
Norādam programmatūras unikālo identifikatoru (Identifying Number) un norādītā programmatūra tiek novēkta no datora.
Identifying Number atrodams [-Asset] programmatūras pārskatā.

.PARAMETER ScriptUpdate
Pārbauda vai nav jaunākas skripta versijas datne norādītajā skripta etalona mapē.
Ja atrod - kopē uz darba direktoriju un beidz darbu.

.PARAMETER Help
Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.

.EXAMPLE
WindowUpdate.ps1 -Name EX00001
Pārbauda datoram EX00001 pieejamos windows jauninājumus

.EXAMPLE
WindowUpdate.ps1 EX00001 -Asset
Sagatavo un parāda ekrānā datora EX00001 tehniskos parametrus

.EXAMPLE
WindowUpdate.ps1 -InPath .\computers.txt -Asset
Sagatavo tehnisko parametru atskaiti Excel formātā visām .\computers.txt norādītajām vienībām

.EXAMPLE
WindowUpdate.ps1 -InPath .\computers.txt
Pārbauda .\computers.txt datnē norādītajai datortehnikai pieejamos windows jauninājumus

.EXAMPLE
WindowUpdate.ps1 -InPath .\servers.txt -IsPendingReboot
Pārbauda vai .\servers.txt datnē norādītajai datortehnikai ir nepieciešams restarts.
Skripta log failā ir norāde uz CSV datnes atrašanās vietu.

.EXAMPLE
WindowUpdate.ps1 -Name EX00001 -Update
Izveido darbstacijā ScheduledTask Windows update uzdevumu, kas izpildās nekavējoši un uzsāk datortehnikai pieejamo jauninājumu lejupielādi, uzstādīšanu.
Tiek ignorēts jauninājuma pieprasījums pēc datoretehnikas pārsāknēšanas. Nepieciešams patstāvīgi pārliecināties, ka jauninājums uzstādījies pilnā apjomā un pārsāknēt darbstaciju.

.EXAMPLE
WindowUpdate.ps1 -InPath .\servers.txt -Update -AutoReboot
Sarakstā norādītajiem serveriem tiek izveidots ScheduledTask Windows update uzdevums, kas izpildās nekavējoši un uzsāk datortehnikai pieejamo jauninājumu lejupielādi, uzstādīšanu.
Ja jauninājuma pilnīgai uzstādīšanai ir nepeiciešams restarts - serveri tiek restartēti pēc jauninājuma pieprasījuma.

.EXAMPLE
WindowUpdate.ps1 -Name reja -RemoteReboot
Attālināti tiek pārsāknēta dators reja. Pārsāknēšana nav atceļama un tiek izpildīta nekavējoši.

.EXAMPLE
WindowUpdate.ps1 -EventLog EX00001
Pārbauda datora EX00001 notikumu žurnālu. Rāda tikai kopsavilkumu.

.EXAMPLE
WindowUpdate.ps1 -EventLog .\computers.txt -Details
Sagatavo .\computers.txt datnē norādītajiem datoru notikumu žurnāla ierakstus. Parāda detalizētu atskaiti.
Norādītai datnei ir jāeksistē, pretējā gadījumā tiek izvadīta kļūda.

.EXAMPLE
WindowUpdate.ps1 -EventLog EX00001 -Days 7
Pārbauda datora EX00001 notikumu žurnālu par notikumiem pēdējās 7 dienās

.EXAMPLE
WindowUpdate.ps1 -Install "C:\install\notepad\npp.8.1.9.3.Installer.x64.exe" EX00001
Uzstāda norādīto programmatūras pakotni uz datora.

.EXAMPLE
WindowUpdate.ps1 -Uninstall '{F914A43C-9614-4100-B94E-BE2D5EC2E5E2}' EX00001
Noņem norādīto programmatūras pakotni no datora. Programmatūras unikālais identifikators atrodams Asste pārskatā.

.EXAMPLE
WindowUpdate.ps1 -ScriptUpdate
Skripta piespiedu pārbaude uz skripta jauninājumu pieejamību skripta etalona mapē, kas tiek norādīta skripta mainīgajā $UpdateDir.

.NOTES
	Author:	Viesturs Skila
	Version: 2.5.4
#>
[CmdletBinding(DefaultParameterSetName = 'InPathCheck')]
param (
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathCheck',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathUpdate',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathEventLog',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathAsset',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPath4Install',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPath4Uninstall',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'WakeOnLanInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'RebootInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'StopInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$InPath,

	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'Name4Uninstall',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'Name4Install',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameAsset',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameEventLog',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameTrace',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'RebootName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'StopName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'WakeOnLanName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameUpdate',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameCheck',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[string]$Name,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathCheck')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameCheck')]
	[switch]$Check = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathUpdate')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameUpdate')]
	[switch]$Update = $False,
	
	[Parameter(Position = 2, Mandatory = $False, ParameterSetName = 'InPathUpdate')]
	[Parameter(Position = 2, Mandatory = $False, ParameterSetName = 'NameUpdate')]
	[switch]$AutoReboot = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'WakeOnLanInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'WakeOnLanName')]
	[switch]$WakeOnLan = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'RebootInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$Reboot = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$NoWait = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'StopInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'StopName')]
	[switch]$Stop = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'StopInPath')]
	[Parameter(Mandatory = $False, ParameterSetName = 'RebootInPath')]
	[Parameter(Mandatory = $False, ParameterSetName = 'StopName')]
	[Parameter(Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$Force = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameTrace')]
	[switch]$Trace = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameEventLog')]
	[switch]$EventLog = $False,
	
	[Parameter(Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameEventLog')]
	[int]$Days,
	
	[Parameter(Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[switch]$OutPath = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$Asset = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[SupportsWildcards()]
	[string]$Include = '*',

	[Parameter(Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[SupportsWildcards()]
	[string]$Exclude = '',
	
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$Hardware = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$NoSoftware = $False,

	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'InPath4Install',
		HelpMessage = "Path of exe or msi file.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".msi|.exe") {
				Write-Host "The file specified in the path argument must be msi or exe file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'Name4Install',
		HelpMessage = "Path of exe or msi file.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".msi|.exe") {
				Write-Host "The file specified in the path argument must be msi or exe file"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$Install,

	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'InPath4Uninstall')]
	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'Name4Uninstall')]
	[string]$Uninstall,

	#Helper slēdži
	[Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'ScriptUpdate')]
	[switch]$ScriptUpdate = $False,

	[Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
	[switch]$Help = $False
)
BEGIN {
	$UpdateDir = "\\beluga\install\Scripts\ExpoRemoteJobs"
	#$PsBoundParameters
	<# ---------------------------------------------------------------------------------------------------------
		Zemāk veicam izmaiņas, ja patiešām saprotam, ko darām.
	--------------------------------------------------------------------------------------------------------- #>
	#Skripta tehniskie mainīgie
	$CurVersion = "2.5.4"
	$scriptWatch	= [System.Diagnostics.Stopwatch]::startNew()
	#Skritpa konfigurācijas datnes
	$__ScriptName	= $MyInvocation.MyCommand
	$__ScriptPath	= Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
	#Atskaitēm un datu uzkrāšanai
	$ReportPath = "$__ScriptPath\result"
	$DataDir = "$__ScriptPath\lib\data"
	$BackupDir = "$__ScriptPath\lib\data\backup"
	#Helper scriptu bibliotēkas
	$CompUpdateFileName = "scripts\Set-CompUpdate.ps1"
	$CompProgramFileName = "scripts\Set-CompProgram.ps1"
	$CompAssetFileName = "scripts\Get-CompAsset.ps1"
	$CompSoftwareFileName	= "scripts\Get-CompSoftware.ps1"
	$CompEventsFileName = "scripts\Get-CompEvents.ps1"
	$CompTestOnlineFileName	= "scripts\Get-CompTestOnline.ps1"
	$CompWakeOnLanFileName	= "scripts\Invoke-CompWakeOnLan.ps1"
	$WinUpdFile = "$__ScriptPath\$CompUpdateFileName"
	$CompProgramFile = "$__ScriptPath\$CompProgramFileName"
	$CompAssetFile	= "$__ScriptPath\$CompAssetFileName"
	$CompSoftwareFile = "$__ScriptPath\$CompSoftwareFileName"
	$CompEventsFile = "$__ScriptPath\$CompEventsFileName"
	$CompTestOnlineFile	= "$__ScriptPath\$CompTestOnlineFileName"
	$CompWakeOnLanFile	= "$__ScriptPath\$CompWakeOnLanFileName"
	#Žurnalēšanai
	$LogFileDir = "$__ScriptPath\log"
	$LogFile = "$LogFileDir\RemoteJob_$(Get-Date -Format "yyyyMMdd")"
	#$XLStoFile		= "$ReportPath\Report-$(Get-Date -Format "yyyyMMddHHmmss").xls"
	$OutputToFile	= "$ReportPath\Report-$(Get-Date -Format "yyyyMMdd").txt"
	$TraceFile = "C:\ExpoSheduledWUjob.log"
	$DataArchiveFile = "$DataDir\DataArchive.dat"
	$VerifyCompsResultFile = "$DataDir\VerifyCompsResult.dat"
	#$ComputerName = ( -not [string]::IsNullOrEmpty($Name) )
	$ScriptRandomID	= Get-Random -Minimum 100000 -Maximum 999999
	$ScriptUser = Invoke-Command -ScriptBlock { whoami }
	$RemoteComputers = @()
	$MaxJobsThreads	= 40

	Get-PSSession | Remove-PSSession
	if ( -not ( Test-Path -Path $LogFileDir ) ) { $null = New-Item -ItemType "Directory" -Path $LogFileDir }
	if ( -not ( Test-Path -Path $ReportPath ) ) { $null = New-Item -ItemType "Directory" -Path $ReportPath }
	if ( -not ( Test-Path -Path $DataDir ) ) { $null = New-Item -ItemType "Directory" -Path $DataDir }
	if ( -not ( Test-Path -Path $BackupDir ) ) { $null = New-Item -ItemType "Directory" -Path $BackupDir }

	if ($Help) {
		Write-Host "`nVersion:[$CurVersion]`n"
		#$text = (Get-Command "$__ScriptPath\$__ScriptName" ).ParameterSets | Select-Object -Property @{n = 'Parameters'; e = { $_.ToString() } }
		$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
		$text | ForEach-Object { Write-Host $($_) }
		Write-Host "For more info write `'Get-Help `.`\$__ScriptName -Examples`'"
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
		if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)" } else { $bMin = "$($Timer.Elapsed.Minutes)" }
		if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)" } else { $bSec = "$($Timer.Elapsed.Seconds)" }
		if ($Name -notlike 'JOBers') {
			Write-msg -log -text "[$Name] finished in $(
				if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
				elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
				else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
			)"
			Write-Host "[$Name] finished in $(
			if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
			elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
			else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
			)"
		}
		else {
			Write-Host "`rJobs done in $(
				if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
				elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
				else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
				)" -ForegroundColor Yellow -BackgroundColor Black
		}
	}#endOffunction

	Function Write-msg { 
		[Alias("wrlog")]
		Param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$text,
			[switch]$log,
			[switch]$bug
		) #param

		try {
			#write-debug "[wrlog] Log path: $log"
			if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO' }
			$timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
			if ( $log -and $bug ) {
				Write-Warning "[$flag] $text"	
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" | Out-File "$LogFile.log" -Append -ErrorAction Stop
			}#endif
			elseif ( $log ) {
				Write-Verbose "[$flag] $text"
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" | Out-File "$LogFile.log" -Append -ErrorAction Stop
			}#endif
			else {
				Write-Verbose "$flag [$ScriptRandomID] $text"
			} #else
		}#endtry
		catch {
			Write-Warning "[Write-msg] $($_.Exception.Message)"
			return
		}#endOftry
	}#endOffunction

	Function Write-ErrorMsg {
		[CmdletBinding()]
		Param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[object]$InputObject,
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$Name
		)#param
		Write-msg -log -bug -text "[$name] Error: $($InputObject.Exception.Message)"
		$string_err = $InputObject | Out-String
		Write-msg -log -bug -text "$string_err"
	}#endOffunction

	Function Get-ScriptFileUpdate {
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[String]$FileName,
			[Parameter(Position = 1, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[String]$ScriptPath,
			[Parameter(Position = 2, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[String]$UpdatePath
		)
		try { $NewFile = Get-ChildItem "$UpdatePath\$FileName" -ErrorAction Stop } catch {}
		try { $ScriptFile = Get-ChildItem "$ScriptPath\$FileName"  -ErrorAction Stop } catch {}
		if ( $NewFile.count -gt 0 ) {
			if ( $NewFile.LastWriteTime -gt $ScriptFile.LastWriteTime ) {
				Write-msg -log -text "[ScriptUpdate] Found update for script [$FileName]"
				Write-msg -log -text "[ScriptUpdate] Old version [$($ScriptFile.LastWriteTime)], [$($ScriptFile.FullName)]"
				Write-msg -log -text "[ScriptUpdate] New version [$($NewFile.LastWriteTime)], [$($NewFile.FullName)]"
				try {
					Copy-Item -Path $NewFile.FullName -Destination "$(Split-Path -Path "$ScriptPath\$FileName")" -Force -ErrorAction Stop
					Write-msg -log -text "[ScriptUpdate] New version deployed."
				}#endtry
				catch {
					#Write-msg -log -bug -text "[ScriptUpdate] [$FileName] $($_.Exception.Message)"
					Write-ErrorMsg -Name 'ScriptUpdate' -InputObject $_
				}#endcatch
			}#endif
		}#endif

	}#endOffunction

	Function Get-NormaliseDiskLabelsForExcel {
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[System.Object]$Computers
		)
		$TemplateDisks = [PSCustomObject][ordered]@{}
		#Get uniqe disk label's collection from computers and write in array
		foreach ($computer in $computers) {
			foreach ( $Item in $computer.PsObject.Properties ) {
				if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize' ) {
					if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
						$TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
					}#endif
				}#endif
				if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize\sFree' ) {
					if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
						$TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
					}#endif
				}#endif
			}#endforeach
		}#endforeach
		#Add to each computer's properties missing disk label
		foreach ( $computer in $computers ) {
			foreach ($label in $TemplateDisks.PsObject.Properties) {
				if ( -not ( Get-Member -InputObject $computer -Name $label.Name ) ) {
					Add-Member -InputObject $computer -NotePropertyName $label.Name -NotePropertyValue 'none' -Force
				}#endif
			}#endforeach
		}#endforeach
		return $computers
	}#endOffunction

	Function Set-Jobs {
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[string[]]$Computers,
			[Parameter(Position = 1, Mandatory = $true)]
			[string]$ScriptBlockName,
			[Parameter(Position = 2, Mandatory = $false)]
			[string]$Argument1
		)
		<# ---------------------------------------------------------------------------------------------------------
		#	Definējam komandu blokus:
		#	SBVerifyComps : pārbaudam datora gatavību darbam ar skriptu
		--------------------------------------------------------------------------------------------------------- #>
		$SBVerifyComps = {
			try {
				$Computer = $args[0]
				#Write-host "[Comp,block1] $Computer"
				If ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop ) {
					$isPingTest = $True
					$RemSess = New-PSSession -ComputerName $Computer -ErrorAction Stop
					if ( ( Invoke-Command -Session $RemSess -ScriptBlock { $PSVersionTable.PSVersion.Major } ) -ge 5 ) { $isPSversion = $True }
			
					#check module 'PSWindowsUpdate' is installed, if not, copy from script's root directory to remote computer
					if (-not ( Invoke-Command -Session $RemSess -ScriptBlock { Get-Module -ListAvailable -Name 'PSWindowsUpdate' } )) {
						Copy-Item "scripts\modules\PSWindowsUpdate\" -Destination "C:\Program Files\WindowsPowerShell\Modules\" -ToSession $RemSess -Recurse -ErrorAction Stop
						$msgPSWUModuleInf = "PSWindowsUpdate module installed on [$computer]."
						$isPSWUModule = $True
					}#endif
					else {
						$isPSWUModule = $True
					}#endelse
					if (-not ( Invoke-Command -Session $RemSess -ScriptBlock { Get-Module -ListAvailable -Name 'Join-Object' } )) {
						Copy-Item "scripts\modules\Join-Object\" -Destination "C:\Program Files\WindowsPowerShell\Modules\" -ToSession $RemSess -Recurse -ErrorAction Stop
						$msgJoinObjectInf = "Join-Object module installed on [$computer]."
						$isJoinObject = $True
					}#endif
					else {
						$isJoinObject = $True
					}#endelse
					#cheking computer is set ExecutionPolicy = RemoteSigned and LanguageMode = FullLanguage
					$psLanguage = Invoke-Command -Session $RemSess -ScriptBlock { $ExecutionContext.SessionState.LanguageMode }
					if ( $psLanguage.value -notlike 'FullLanguage' ) {
						$msgLanguageBug = "Computer [$Computer] is not ready for PSRemote: LanguageMode is set [$($psLanguage.value)]"
					}#endif
					else {
						$isLanguage = $True
					}#endelse
					$policy = Invoke-Command -Session $RemSess -ScriptBlock { Get-ExecutionPolicy }
					if ( $policy.value -notlike 'Unrestricted' -and $policy.value -notlike 'RemoteSigned' ) {
						$msgPolicyBug = "Computer [$Computer] is not ready for PSRemote: policy is set [$($policy.value)]"
					}#endif
					else {
						$isPolicy = $True
					}#endelse
					if ( $RemSess.count -gt 0 ) {
						Remove-PSSession -Session $RemSess
					}#endif
				}#endif
				else {
					$msgCatchErr = "Computer [$Computer] is not accessible."
				}#endelse
				$ObjectReturn = New-Object -TypeName psobject -Property @{
					Computer         = $Computer ;
					isPingTest       = if ( $isPingTest ) { $isPingTest } else { $False } ;
					isPSversion      = if ( $isPSversion ) { $isPSversion } else { $False } ;
					isPSWUModule     = if ( $isPSWUModule ) { $isPSWUModule } else { $False } ;
					msgPSWUModuleInf	= if ( $msgPSWUModuleInf ) { $isPSWUModule } else { $null } ;
					isJoinObject     = if ( $isJoinObject ) { $isJoinObject } else { $False } ;
					msgJoinObjectInf	= if ( $msgJoinObjectInf ) { $msgJoinObjectInf } else { $null } ;
					isLanguage       = if ( $isLanguage ) { $isLanguage } else { $False } ;
					msgLanguageBug   = if ( $msgLanguageBug ) { $msgLanguageBug } else { $null } ;
					isPolicy         = if ( $isPolicy ) { $isPolicy } else { $False } ;
					msgPolicyBug     = if ( $msgPolicyBug ) { $msgPolicyBug } else { $null } ;
					msgCatchErr      = if ( $msgCatchErr ) { $msgCatchErr } else { $null } ;
				}#endobject
				return $ObjectReturn
			}#endtry
			catch {
				$msgCatchErr = "$_"
				$ObjectReturn = New-Object -TypeName psobject -Property @{
					Computer         = $Computer ;
					isPingTest       = if ( $isPingTest ) { $isPingTest } else { $False } ;
					isPSversion      = if ( $isPSversion ) { $isPSversion } else { $False } ;
					isPSWUModule     = if ( $isPSWUModule ) { $isPSWUModule } else { $False } ;
					msgPSWUModuleInf	= if ( $msgPSWUModuleInf ) { $isPSWUModule } else { $null } ;
					isJoinObject     = if ( $isJoinObject ) { $isJoinObject } else { $False } ;
					msgJoinObjectInf	= if ( $msgJoinObjectInf ) { $msgJoinObjectInf } else { $null } ;
					isLanguage       = if ( $isLanguage ) { $isLanguage } else { $False } ;
					msgLanguageBug   = if ( $msgLanguageBug ) { $msgLanguageBug } else { $null } ;
					isPolicy         = if ( $isPolicy ) { $isPolicy } else { $False } ;
					msgPolicyBug     = if ( $msgPolicyBug ) { $msgPolicyBug } else { $null } ;
					msgCatchErr      = if ( $msgCatchErr ) { $msgCatchErr } else { $null } ;
				}#endobject
				if ( $RemSess.count -gt 0 ) {
					Remove-PSSession -Session $RemSess
				}#endif
				return $ObjectReturn
			}#endcatch
		}#endblock

		<# ---------------------------------------------------------------------------------------------------------
		# SBWindowsUpdate: izsaucam Windows update uz attālinātās darbstacijas
		---------------------------------------------------------------------------------------------------------#>
		$SBWindowsUpdate = {
			$Computer = $args[0]
			$WinUpdFile = $args[1]
			$Update = $args[2]
			$AutoReboot = $args[3]
			$OutputResults = Invoke-Command -ComputerName $Computer -FilePath $WinUpdFile -ArgumentList ($Update, $AutoReboot)
			return $OutputResults
		}#endblock

		<# ---------------------------------------------------------------------------------------------------------
		# SBInstall: izsaucam programmas uzstādīšanas procesu
		# Set-CompProgram.ps1 [-ComputerName] <string> [-InstallPath <FileInfo>] [<CommonParameters>]
		---------------------------------------------------------------------------------------------------------#>
		$SBInstall = {
			$Computer = $args[0]
			$CompProgramFile = $args[1]
			$Install = $args[2]
			$OutputResults = Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-InstallPath $Install "
			return $OutputResults
		}#endblock

		<# ---------------------------------------------------------------------------------------------------------
		# SBInstall: izsaucam programmas noņemšanas procesu
		# Set-CompProgram.ps1 [-ComputerName] <string> [-CryptedIdNumber <string>] [<CommonParameters>]
		---------------------------------------------------------------------------------------------------------#>
		$SBUninstall = {
			$Computer = $args[0]
			$CompProgramFile = $args[1]
			$EncryptedParameter = $args[2]
			$OutputResults = Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-CryptedIdNumber $EncryptedParameter "
			return $OutputResults
		}#endblock

		<# ---------------------------------------------------------------------------------------------------------
		# SBWakeOnLan: izsaucam programmas noņemšanas procesu
		# Invoke-CompWakeOnLan.ps1 [-ComputerName] <string[]> [-DataArchiveFile] <FileInfo> [-CompTestOnline] <FileInfo> [<CommonParameters>]
		---------------------------------------------------------------------------------------------------------#>
		$SBWakeOnLan = {
			$Computer = $args[0]
			$CompWakeOnLanFile = $args[1]
			$DataArchiveFile = $args[2]
			$CompTestOnlineFile = $args[3]
			$OutputResults = Invoke-Expression "& `"$CompWakeOnLanFile`" `-ComputerName $Computer `-DataArchiveFile `"$DataArchiveFile`" `-CompTestOnline `"$CompTestOnlineFile`" "
			return $OutputResults
		}#endblock
		<# ---------------------------------------------------------------------------------------------------------
			[JOBers] kods
		--------------------------------------------------------------------------------------------------------- #>
		$jobWatch = [System.Diagnostics.Stopwatch]::startNew()
		$Output = @()
		Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
		ForEach ( $Computer in $Computers ) {
			While ($(Get-Job -state running).count -ge $MaxJobsThreads) {
				Start-Sleep -Milliseconds 10
			}#endWhile
			if ( $ScriptBlockName -eq 'SBVerifyComps' ) { 
				$null = Start-Job -Name "$($Computer)" -Scriptblock $SBVerifyComps -ArgumentList $Computer 
			}#endif
			if ( $ScriptBlockName -eq 'SBWindowsUpdate' ) { 
				$null = Start-Job -Name "$($Computer)" -Scriptblock $SBWindowsUpdate -ArgumentList $Computer, $WinUpdFile, $Update, $AutoReboot 
			}#endif
			if ( $ScriptBlockName -eq 'SBInstall' ) { 
				Write-Verbose "[StartJob] Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile, $Install"
				$null = Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile, $Install
			}#endif
			if ( $ScriptBlockName -eq 'SBUninstall' ) { 
				Write-Verbose "[StartJob] Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1"
				$null = Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1
			}#endif
			if ( $ScriptBlockName -eq 'SBWakeOnLan' ) { 
				Write-Verbose "[StartJob] Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile"
				$null = Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile
			}#endif
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
		}#endForEach
		While (Get-Job -State "Running") {
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep 10
		}
		#Get information from each job.
		foreach ( $job in Get-Job ) {
			$result = @()
			$result = Receive-Job -Id ($job.Id)
			if ( $result -or $result.count -gt 0 ) {
				$Output += $result
			}#endif
		}#endforeach
		Stop-Watch -Timer $jobWatch -Name JOBers
		Get-Job | Remove-Job

		if ( $Output.Count -gt 0 ) {
			Return $Output
		}#endif
		else {
			Return $False
		}#endelse
	}#endOfFunction

	Function Get-VerifyComputers {
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[string[]]$ComputerNames
		)
		Write-Host "[VerifyComp] got for testing [$($ComputerNames.Count)] $(if($ComputerNames.Count -eq 1){"computer"}else{"computers"})"
		Write-Verbose "[Function:VerifyComputers]:got:[$($ComputerNames)]"
		$JobResults = Set-Jobs -Computers $ComputerNames -ScriptBlockName 'SBVerifyComps'
		#$JobResults
		$JobResults = $JobResults | Where-Object -Property Computer -ne $null
		if ( $JobResults -ne $False -or $JobResults -notlike 'False' -or $JobResults.Count -gt 0 ) {
			#Analizējam no job saņemtos rezultātus
			$DelComps = @()
			#Analizējam JObbu atgrieztos rezultātus, ievietojam DelComps
			$JobResults | Foreach-Object {
				if ( -not $_.CatchErr ) {
					if ($_.isPingTest -eq $False ) {
						$DelComps += $_.Computer
						Write-msg -log -text "[JobResults] computer [$($_.Computer)] is not accessible"
					}#endif
					else {
						if ($_.isPSversion -eq $False ) { 
							$DelComps += $_.Computer
							Write-msg -log -text "[JobResults] [$($_.Computer)] Powershell version is less than 5.0"
							Write-Verbose "[JobResults] [$($_.Computer)] $($_.msgCatchErr)"
						}#endif
						elseif ($_.isPSWUModule -eq $False ) { 
							$DelComps += $_.Computer
							Write-msg -log -text "[JobResults] $($_.msgPSWUModuleInf)"
						}#endif
						elseif ($_.isJoinObject -eq $False ) {
							$DelComps += $_.Computer 
							Write-msg -log -text "[JobResults] $($_.msgJoinObjectInf)"
						}#endif
						elseif ($_.isLanguage -eq $False ) {
							$DelComps += $_.Computer 
							Write-msg -log -bug -text "[JobResults] $($_.msgLanguageBug)"
						}#endif
						elseif ($_.isPolicy -eq $False ) {
							$DelComps += $_.Computer 
							Write-msg -log -bug -text "[JobResults] $($_.msgPolicyBug)"
						}#endif
					}#endelse
				}#endif
				else {
					Write-Host "[JobResults] $($_.CatchErr)"
					Write-msg -log -bug -text "[JobResults] $($_.CatchErr)" 
				}#endelse
			}#endforeach
			#Atbrīvojamies no dublikātiem, ja tādu ir
			$DelComps = $DelComps | Get-Unique
			#Parsējam imput masīvu un papildinām DelComps ar dzēšamajiem datoriem, kas nav atbildējuši uz ping
			$ComputerNames | ForEach-Object {
				if ( $JobResults.Computer.Contains($_) -eq $False ) {
					$DelComps += $_
				}#endif
			}#endforeach
			#Aizvācam no input masīva visus datorus, kas nav izturējuši pārbaudi
			$DelComps | ForEach-Object { 
				if ( $ComputerNames.Contains($_) ) {
					$ComputerNames = $ComputerNames -ne $_
					Write-msg -log -text "[VerifyComputers] Computer [$_] is not ready PSRemote. Removed."
				}#endif
			}#endforeach
		}#endif
		else {
			Write-msg -log -bug -text "[JobResults] Oopps!! Job returned nothing." 
		}#endelse
		if ( $JobResults.GetType().BaseType.name -eq 'Object' ) {
			$tmpJobResults = $JobResults.psobject.copy()
			$JobResults = @()
			$JobResults += @($tmpJobResults)
		}#endif
		if ( $JobResults.count -gt 0) {
			Write-Verbose "[Function:VerifyComputers]:return:[$($ComputerNames)]"
			$JobResults | Export-Clixml -Path $VerifyCompsResultFile -Depth 10 -Force
			return $ComputerNames
		}#endif
		else {
			return $false
		}
	}#endOffunction

	<# ---------------------------------------------------------------------------------------------------------
	Funkciju definēšanu beidzām
	IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
	--------------------------------------------------------------------------------------------------------- #>
	Write-msg -log -text "[-----] Script started in [$(if ($Trace) {"Trace"}
		elseif ($ScriptUpdate) {"ScriptUpdate"}
		elseif ($Asset) {"Asset"}
		elseif ($Update) {"Update"} 
		elseif ($Reboot) {"Reboot"}
		elseif ($Stop) {"Stop"} 
		elseif ( $PSCmdlet.ParameterSetName -eq "NameEventLog" ) {"Name-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -eq "InPathEventLog" ) {"InPath-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLan*" ) {"WakeOnLan"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Install" ) {"Install"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Uninstall" ) {"Uninstall"}
		else {"Check"})] mode. Used value [$(if ($Name){"Name:[$Name]"}
			elseif ($InPath) {"InPath:[$InPath]"} 
			else {"none"})]"
	Write-Host "`n[-----] Script started in [$(if ($Trace) {"Trace"}
		elseif ($ScriptUpdate) {"ScriptUpdate"}
		elseif ($Asset) {"Asset"}
		elseif ($Update) {"Update"} 
		elseif ($Reboot) {"Reboot"}
		elseif ($Stop) {"Stop"}
		elseif ( $PSCmdlet.ParameterSetName -eq "NameEventLog" ) {"Name-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -eq "InPathEventLog" ) {"InPath-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLan*" ) {"WakeOnLan"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Install" ) {"Install"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Uninstall" ) {"Uninstall"}
		else {"Check"})] mode. Used value [$(if ($Name){"Name:[$Name]"}
			elseif ($InPath) {"InPath:[$InPath]"} 
			else {"none"})]"
			
	#Scripts pārbauda vai nav jaunākas versijas repozitorijā 
	if ( Test-Path -Path $UpdateDir ) {
		Get-ScriptFileUpdate $__ScriptName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompUpdateFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompProgramFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompAssetFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompSoftwareFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompEventsFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompTestOnlineFileName $__ScriptPath $UpdateDir
		Get-ScriptFileUpdate $CompWakeOnLanFileName $__ScriptPath $UpdateDir
		if ($ScriptUpdate) {
			Stop-Watch -Timer $scriptWatch -Name Script
			exit 
		}#endif
	}#endif
	else {
		if ($ScriptUpdate) {
			Write-msg -log -text "[Update] directory [$($UpdateDir)] not available."
			Stop-Watch -Timer $scriptWatch -Name Script
			exit 
		}#endif
	}#endif
	
	#Check for script helper files; if not aviable - exit
	if ( -not ( Test-Path -Path $WinUpdFile -PathType Leaf )) {
		Write-msg -log -bug -text "File [$WinUpdFile] not found. Exit."
		Stop-Watch -Timer $scriptWatch -Name Script
		Exit
	}#endif
	if ( -not ( Test-Path -Path $CompTestOnlineFile -PathType Leaf )) {
		Write-msg -log -bug -text "File [$CompTestOnlineFile] not found. Exit."
		Stop-Watch -Timer $scriptWatch -Name Script
		Exit
	}#endif
	
	<# ---------------------------------------------------------------------------------------------------------
	pārbaudām katras datortehnikas gatavību strādāt PSRemote režīmā, vai ir nepieciešamās bilbiotēkas
	To darām trīs soļos: 
	[1] ielādējam sarakstu un pingojam,
	[2] ielādējam arhīvu un pārbaudam vai TTL nav beidzies,
	[3] ja TTL beidzies, veicam datortehnikas pilno pārbaudi uz PSRemote
	--------------------------------------------------------------------------------------------------------- #>
	# ielādējam inut parametrus
	if ( $PSCmdlet.ParameterSetName -like "Name*" -or 
		$PSCmdlet.ParameterSetName -eq "RebootName" -or 
		$PSCmdlet.ParameterSetName -eq "StopName" -or 
		$PSCmdlet.ParameterSetName -like "Asset*" ) {
		$InComputers = @($Name.ToLower())
		$ArgumentListOnline = "`-Name $InComputers"
	}#endif
	elseif ( $PSCmdlet.ParameterSetName -like "InPath*" -or
		$PSCmdlet.ParameterSetName -eq "RebootInPath" -or 
		$PSCmdlet.ParameterSetName -eq "StopInPath" ) {
		$InComputers = @(Get-Content $InPath | Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  | ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique )
		$ArgumentListOnline = "`-Inpath $InPath"
	}#endelseif
	elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLanName" ) {
		$RemoteComputers = @($Name.ToLower())
	}
	elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLanInPath" ) {
		$RemoteComputers = @(Get-Content $InPath | Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  | ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique )
	}
	
	if ( $PSCmdlet.ParameterSetName -like "Name*" -or 
		$PSCmdlet.ParameterSetName -like "InPath*" -or
		$PSCmdlet.ParameterSetName -like "Reboot*" -or 
		$PSCmdlet.ParameterSetName -like "Stop*" -or 
		$PSCmdlet.ParameterSetName -like "Asset*" ) {
		#ielādējam datu arhīvu
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
		
		Write-Verbose "1.0:[Input]:got:[$($InComputers.count)]"
		<#
		Write-Verbose "1.1:[DataArchive]-=> []---------------------------------------------------------"
		$DataArchive | Sort-Object -Property AddDate -Descending `
		| Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize  `
		| Out-String -Stream | Where-Object { $_ -ne "" } `
		| ForEach-Object { Write-Verbose "$_" }
		Write-Verbose "--------------------------------------------------------------------------------"
		#>

		$OfflineComputers = @()
		Write-Verbose "`& `"$CompTestOnlineFile`" $ArgumentListOnline "
		$OnlineComps = Invoke-Expression "& `"$CompTestOnlineFile`" $ArgumentListOnline "

		Write-Verbose "OnlineComps:[$(if ( $OnlineComps ) {"Atgriezts masīvs"}  else {"Atgriezts tukšs masīvs"})]"
		if ( $OnlineComps ) {
			Write-Verbose "1.2:[OnlineComps]---------------------------------------------------------------"
			$OnlineComps | Sort-Object -Property AddDate -Descending `
			| Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize  `
			| Out-String -Stream | Where-Object { $_ -ne "" } `
			| ForEach-Object { Write-Verbose "$_" }
			Write-Verbose "--------------------------------------------------------------------------------"
			Write-Verbose "1.3:OnlineComps:[$(if ( $OnlineComps.GetType().BaseType.name -eq 'Array' -and $OnlineComps.count -gt 0 ) {"Array"} `
				elseif ( $OnlineComps.GetType().BaseType.name -eq 'Object' ) {"Object"} else {"Other"})]; DataArchive:[$(`
				if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -gt 0 ) {"Array"} `
				elseif ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -eq 0 ) {"Empty Array"} `
				elseif ( $DataArchive.GetType().BaseType.name -eq 'Object' ) {"Object"} else {"Other"})]`
				"
		}#endif
		
		#Ja OnlineComps atgriež vienu objektu un tas nav tukšs, tad to pārveidojam to par objektu masīvu
		if ( $OnlineComps.GetType().BaseType.name -eq 'Object' ) {
			$tmpOnlineComps = $OnlineComps.psobject.copy()
			$OnlineComps = @()
			$OnlineComps += @($tmpOnlineComps)
		}#endif

		if ( $OnlineComps -eq $false) {
			Write-Verbose "[OnlineComps] returned empty object"
			$RemoteComputers = @()
		}#endif
		#ja atgriezts objektu masīvs
		elseif ( $OnlineComps.GetType().BaseType.name -eq 'Array' -and $OnlineComps.count -gt 0 ) {
			#Atlasām offline datorus
			$InComputers | ForEach-Object {
				if ( $OnlineComps.PipedName.Contains($_) -eq $false) {
					Write-Verbose "1:[OnlineComps]:[$($_)] -=> [OfflineComputers]] "
					$OfflineComputers += @($_)
				}#endif
			}#endforeach
			[string]$vstring = $null
			$OnlineComps.DNSName | ForEach-Object { $vstring += "[$_], " }
			Write-Verbose "2.0:[OnlineComps]: $vstring"
			Write-Verbose "OfflineComputers:[$OfflineComputers]"
			
			$_OnlineComps = @()
			<# ---------------------------------------------------------------------------------------------------------
				pārbaudam ARHĪVĀ vai datoram  nav beidzies TTL derīgums: 12 stundas)
			--------------------------------------------------------------------------------------------------------- #>
			#izlaižam arhīva soli, ja [1] arhīvs nesatur vērtības vai [2] Online ir tikai viens ieraksts
			if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -gt 0 -and $OnlineComps.count -gt 1 ) {

				foreach ( $comp in $OnlineComps) {
					:OutOfNestedForEach_LABEL
					foreach ( $record in $DataArchive) {
						if ( $comp.DNSName -eq $record.DNSName ) {
							#ja dators pēdējo reizi verificēts pirms 12h, tad pārbaudām
							if ($record.AddDate -lt (Get-Date).AddHours(-12) ) {
								$_OnlineComps += @($record.DNSName)
								Write-Verbose "2.1:[Archive]:[$($record.DNSName)]-=> [_OnlineComps]"
								break :OutOfNestedForEach_LABEL
							}#endif
							else {
								$RemoteComputers += @($record.DNSName)
								Write-Verbose "2.2:[Archive]:[$($record.DNSName)]-=> [RemoteComputers]"
								break :OutOfNestedForEach_LABEL
							}#endelse
						}#endif
					}#endforeach
				}#endforeach

				#pārbaudam uz online ierakstu esamību, kas netika atrasti arhīvā - ja ir, tad liekam _Online
				$OnlineComps.DNSName | ForEach-Object {
					if ( $_OnlineComps.Contains($_) -eq $false -and $RemoteComputers.Contains($_) -eq $false ) {
						Write-Verbose "2.3:[Online]:[$($record.DNSName)]-=> [RemoteComputers]"
						$_OnlineComps += @($_)
					}#endif
				}
			}#endif
			else {
				$_OnlineComps = $OnlineComps | ForEach-Object { @($_.DNSName) }
			}#endelse
			Write-Verbose "2.4:_OnlineComps:[$_OnlineComps]; RemoteComputers[$RemoteComputers]"

			<# ---------------------------------------------------------------------------------------------------------
				ja atrasti ieraksti, kam TTL beidzies, veicam datortehnikas veicam padziļināto atbilstības pārbaudi
			--------------------------------------------------------------------------------------------------------- #>
			if ( $_OnlineComps.count -gt 0 ) {

				$VerifiedComps = Get-VerifyComputers -ComputerNames $_OnlineComps
				
				Write-Verbose "3.0: Got from VerifiedComps:[$VerifiedComps]"

				#Ja pārbaudi neiztur neviens dators
				if ( $VerifiedComps -eq $false ) {
					$OfflineComputers += $_OnlineComps | ForEach-Object { @($_) }
				}#endif
				elseif ( $VerifiedComps.count -gt 0 ) {
					foreach ( $comp in $OnlineComps) { 
						Write-Verbose "3.1:[OnlineComps] [$($comp.DNSName)]:[$($comp.PipedName)]:[$($comp.AddDate)]:[$($comp.MacAddress)]"
						#Ja izturēja pārbaudi - liekam remote un papildinām arhīvu, ja ne - offline
						if ( $VerifiedComps.Contains($comp.DNSName) -eq $true ) {
							Write-Verbose "3.2:[OnlineComps]:[$($comp.DNSName)]-=> [RemoteComputers]"
							$RemoteComputers += @($comp.DNSName)
							#salāgojam online un arhīva ierakstus
							if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -eq 0 ) {
								Write-Verbose "3.3:[OnlineComps]:[$($comp.DNSName)]-=> [DataArchive]"
								$DataArchive = @(
									New-Object PSObject -Property @{
										PipedName    = $comp.PipedName.ToLower();
										DNSName      = $comp.DNSName.ToLower();
										AddDate      = [System.DateTime](Get-Date);
										MacAddress   = $comp.MacAddress;
										IPAddress    = $comp.IPAddress;
										WinRMservice = $comp.WinRMservice;
									}#endobject
								)
							}#endif
							elseif ( $DataArchive.DNSName.Contains($comp.DNSName) -eq $false ) {
								Write-Verbose "3.4:[OnlineComps]:[$($comp.DNSName)]-=> [DataArchive]"
								$DataArchive += @(
									New-Object PSObject -Property @{
										PipedName    = $comp.PipedName.ToLower();
										DNSName      = $comp.DNSName.ToLower();
										AddDate      = [System.DateTime](Get-Date);
										MacAddress   = $comp.MacAddress;
										IPAddress    = $comp.IPAddress;
										WinRMservice = $comp.WinRMservice;
									}#endobject
								)
							}#endelseif
							foreach ( $row in $DataArchive ) {
								if ( $comp.DNSName -eq $row.DNSName ) {
									$row.AddDate = [System.DateTime](Get-Date)
								}
							}#endforeach
						}#endif
						else {
							if ( $RemoteComputers.Contains($comp.DNSName) -eq $false ) {
								Write-Verbose "3.5:[OnlineComps]:[$($_)]-=> [OfflineComputers] "
								$OfflineComputers += @($comp.DNSName)
							}#endif
						}#endelse
					}#endforeach
				}#endelseif
				else {
					$OfflineComputers += $_OnlineComps | ForEach-Object { @($_) }
				}#endelse
			}#endif
			else {
				Write-Verbose "3.0:[VerifiedComps]: -= SKIPPED =-"
			}#endif
		}#endif
		else {
			Write-msg -log -bug -text "[OnlineComps] returned Object type [Other] or [Empty]"
			$RemoteComputers = @()
		}#endelse
		<#
		Write-Verbose "4.0:[]-=> [DataArchive] --------------------------------------------------------"
		$DataArchive | Sort-Object -Property AddDate -Descending `
		| Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize `
		| Out-String -Stream | Where-Object { $_ -ne "" } `
		| ForEach-Object { Write-Verbose "$_" }
		Write-Verbose "--------------------------------------------------------------------------------"
		#>
		#ierakstām datus arhīvā
		Copy-Item -Path $DataArchiveFile -Destination "$BackupDir\DataArchive-$(Get-Date -Format "yyyyMMddHHmm").bck"
		$DataArchive | Export-Clixml -Path $DataArchiveFile -Depth 10 -Force

		Write-Verbose "[Online]:[$($RemoteComputers.count)],[Offline]:[$($OfflineComputers.Count)]"
	}#endif
}#endOfbegin

<# ---------------------------------------------------------------------------------------------------------
	ŠEIT SĀKAS PAMATA DARBS
--------------------------------------------------------------------------------------------------------- #>
PROCESS {
	<# ---------------------------------------------------------------------------------------------------------
		Reboot vai Stop - prasām apliecinājumus un darām darbu
	--------------------------------------------------------------------------------------------------------- #>
	if ( ( $RemoteComputers.count -gt 0 ) -and `
		( $PSCmdlet.ParameterSetName -like "Reboot*" -or `
				$PSCmdlet.ParameterSetName -like "Stop*" ) ) {
		
		Write-Host "Please be sure in what you are going to do!!!`n---------------------------------------------" -ForegroundColor Yellow
		foreach ( $computer in $RemoteComputers ) {
			Write-msg -log -text "[Main] Confirm $(if($Reboot) {"reboot"} else {"shutdown"}) of [$computer]"
			$answer = Read-Host -Prompt "Please confirm remote $(if($Reboot) {"reboot"} else {"shutdown"}) of [ $computer ]:`t`tType [Yes/No]"
			if ( $answer -like 'Yes' ) {
				Write-msg -log -text "[Main] [$Computer] going to be $(if($Reboot) {"rebooted"} else {"shutdowned"})."
				if ($Reboot) {
					try {
						if ($NoWait -or $PSCmdlet.ParameterSetName -eq "RebootInPath") {
							$parameters = @{
								ComputerName = $Computer
							}#endsplat
						}#endif
						else {
							$parameters = @{
								ComputerName = $Computer
								Wait         = $True
								For          = 'Powershell'
								Timeout      = 300
								Delay        = 2 
							}#endsplat
						}#endelse
						if ($Force) {
							$parameters.Add( 'Force', $Force ) 
						}#endif
						$parameters.Add( 'ErrorAction', 'Stop' )

						Restart-Computer @parameters
						Write-msg -log -text "[Reboot] [$Computer] successfully."
						Write-Host "[Reboot] [$Computer] successfully."
					}#endtry
					catch {
						Write-ErrorMsg -Name 'Reboot' -InputObject $_
					}#endcatch
				}#endif
				if ($Stop) {
					try {
						$parameters = @{
							ComputerName = $Computer
							ErrorAction  = 'Stop'
						}#endsplat
						if ($Force) { 
							$parameters.Add( 'Force', $Force ) 
						}#endif

						Stop-Computer @parameters
						Write-msg -log -text "[Stop] [$Computer] successfully."
						Write-Host "[Stop] [$Computer] successfully."
					}#endtry
					catch {
						Write-ErrorMsg -Name 'Stop' -InputObject $_
					}#endcatch
				}#endif
			}#endif
			else {
				Write-msg -log -text "[$(if($Reboot) {"Reboot"} else {"Stop"})] [$Computer] $(if($Reboot) {"reboot"} else {"shutdown"}) canceled."
				Write-Host "[$(if($Reboot) {"Reboot"} else {"Stop"})] [$Computer] $(if($Reboot) {"reboot"} else {"shutdown"}) canceled."
			}#endelse
		}#endForEach
	}#endif
	<# ---------------------------------------------------------------------------------------------------------
		izpildam Asset skriptu			- $CompAssetFileName	= "lib\Get-CompAsset.ps1"
		izpildam Get-Software skriptu	- $CompSoftwareFileName = "lib\Get-CompSoftware.ps1"
	--------------------------------------------------------------------------------------------------------- #>
	if ( $RemoteComputers.count -gt 0 -and $Asset ) {
		try {
			$CompSession = New-PSSession -ComputerName $RemoteComputers -ErrorAction Stop
			Write-msg -log -text "[Asset]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
			Write-Host "[Asset]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
			
			$__throwMessage = $null
			if (-NOT ( Test-Path -Path $CompAssetFile -PathType Leaf )) {
				throw $__throwMessage = "[Asset] script file [$CompAssetFile] not found. Fatal error."
			}#endif
			elseif (-NOT ( Test-Path -Path $CompSoftwareFile -PathType Leaf )) {
				throw $__throwMessage = "[Asset] script file [$CompSoftwareFile] not found. Fatal error."
			}

			if ( $PSCmdlet.ParameterSetName -like "NameAsset" ) {
				if ( $Hardware ) {
					Write-Host "`n[Computer]====================================================================================================" -ForegroundColor Yellow
					$result = Invoke-Command -Session $CompSession -FilePath $CompAssetFile -ArgumentList ($true)
					$result | Format-List | Out-String -Stream | Where-Object { $_ -ne "" } `
					| ForEach-Object { Write-Verbose "$_" }
				}#endif
				if ( -NOT $NoSoftware ) {
					Write-Host "`n[Software]====================================================================================================" -ForegroundColor Yellow
					$result = Invoke-Command -Session $CompSession -FilePath $CompSoftwareFile -ArgumentList ($Include, $Exclude) `
					| Sort-Object -Property DisplayName `
					| Select-Object PSComputerName, @{name = 'Name'; expression = { $_.DisplayName } }, @{name = 'Version'; expression = { $_.DisplayVersion } }, Scope, IdentifyingNumber, @{name = 'Arch'; expression = { $_.Architecture } } -Unique
					if ($result) {
						$result	| Format-Table Name, Version, IdentifyingNumber, Scope, Arch -AutoSize
					}#endif
					else {
						Write-Host "Got nothing to show." -ForegroundColor Green
					}#endelse

				}#endif
				Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
			}#endif
			if ($PSCmdlet.ParameterSetName -like "InPathAsset" ) {
				Write-Host "`n[Software]====================================================================================================" -ForegroundColor Yellow
				$result = Invoke-Command -Session $CompSession -FilePath $CompSoftwareFile -ArgumentList ($Include, $Exclude) `
				| Sort-Object -Property PSComputerName, DisplayName `
				| Select-Object PSComputerName, @{name = 'Name'; expression = { $_.DisplayName } }, @{name = 'Version'; expression = { $_.DisplayVersion } }, Scope, IdentifyingNumber, @{name = 'Arch'; expression = { $_.Architecture } } -Unique
				if ($result) {
					$result	| Format-Table PSComputerName, Name, Version, IdentifyingNumber, Scope, Arch -AutoSize
				}#endif
				else {
					Write-Host "Got nothing to show." -ForegroundColor Green
				}#endelse
				Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
				<#
				if ( $RawData.count -gt 0 ) {
					Write-msg -log -text "[Asset] Collecting information..."
					$ReportToExcel = Get-NormaliseDiskLabelsForExcel $RawData
					$properties = $ReportToExcel | Foreach-Object { $_.psobject.Properties | Select-Object -ExpandProperty Name } | Sort-Object -Unique
					#for testing
					$ReportToExcel | Sort-Object -Property aComputerName | Select-Object $properties | `
						ConvertTo-Json -depth 10 | Out-File "$DataDir\Data-$(Get-Date -Format "yyyyMMddHHmmss").json" -Force
					$ReportToExcel | Sort-Object -Property aComputerName | Select-Object $properties | `
						Export-Excel $XLStoFile -WorksheetName "ExpoAssets" -FreezeTopRow -AutoSize -BoldTopRow
					Write-msg -log -text "Please check the Excel file [$XLStoFile]."
					Write-Host "[Asset] Please check the Excel file [$XLStoFile].`n "
				}#endif
				else {
					Write-msg -log -bug -text "[Asset] Ups... there's nothing to import to the Excel."
				}#endelse
				#>
			}#endelse
		}#endtry
		catch {
			if ($__throwMessage) {
				Write-msg -log -bug -text "$__throwMessage"
			}#endif
			else {
				Write-ErrorMsg -Name 'Asset' -InputObject $_
			}#endelse
		}#endcatch
		finally {
			if ( $CompSession.count -gt 0 ) {
				Remove-PSSession -Session $CompSession
			}#endif
		}#endfinally
	}#endif
	else {
		if ( $PSCmdlet.ParameterSetName -like "*Asset" ) {
			Write-msg -log -bug -text "[Asset] No computer in list."
		}#endif
	}#endelse

	<# ---------------------------------------------------------------------------------------------------------
		darbam gatavai datortehnikai izpildam pieprasīto operāciju
	--------------------------------------------------------------------------------------------------------- #>
	if ( ( $RemoteComputers.count -gt 0 ) -and ( $Check -or $Update -or $Trace ) ) {
		try {
			Write-msg -log -text "Conecting to [$($RemoteComputers.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} ).."
			$CompSession = New-PSSession -ComputerName $RemoteComputers -ErrorAction Stop

			Write-msg -log -text "[$(if($Check){"Check"}elseif($Update){"Update"})]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
			Write-Host "[$(if($Check){"Check"}elseif($Update){"Update"})]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"

			if ($CompSession.count -gt 0 ) {
				if ( $Trace) {
					
					$TestScript = {
						$TraceFile = $args[0]
						$isFile = Test-Path -Path $TraceFile -PathType Leaf
						return $isFile
					}
					$SessScript = {
						$TraceFile = $args[0]
						$file = Get-Content $TraceFile
						return $file
					}
					# check is there log file source; if not - skip
					if ( Invoke-Command -Session $CompSession -ScriptBlock $TestScript -ArgumentList $TraceFile ) {
							$result = Invoke-Command -Session $CompSession -ScriptBlock $SessScript -ArgumentList $TraceFile -ErrorAction Ignore
							Write-host "`n[$(Get-Date -Format "yyyy.MM.dd HH:mm:ss")] from [$TraceFile]--------------------------" -ForegroundColor Yellow
			
							$result | Format-Table * -AutoSize | Out-String -Width 128 -Stream `
							| Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" }
							
							Write-Host "-------------------------------------------------------------------------------" -ForegroundColor Yellow

							$result | ConvertTo-Json -depth 10 | Out-File "$DataDir\trace.json"
			
					}#endif
					else {
						Write-Host "There's no update process started." -Foreground Yellow 
						Write-msg -log -text "There's no update process started."
					}#endelse

				}#endelseif
				else {
					Write-msg -log -text "Sending instructions to [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
					
					# Run Windows install script on to each computer
					$JobResults = Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBWindowsUpdate'
					#$JobResults
					Write-Host "=================================================================================================" -ForegroundColor Yellow
					# Collect remote logs
					$ResultOutput = @()
					$i = 0 
					$JobResults = $JobResults | Sort-Object -Property Computer
					Foreach ( $row in $JobResults ) {
						#Write-Host "1: [$($row.PendingReboot.Count)][$(if ( $row.PendingReboot) {"True"})][$(if ( $null -eq $row.PendingReboot ) {"True"})]"
						if ( $null -ne $row.PendingReboot ) {
							$row.PendingReboot | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									#Examle of line: "EUROBARS | update | [4] updates are waiting to be installed"
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}#endobject
								$i++
							}#endforeach
						}#endif
						#Write-Host "2: [$($row.Updates.Count)][$(if ( $row.Updates) {"True"})][$(if ( $null -eq $row.Updates ) {"True"})]"
						if ( $null -ne $row.Updates ) {
							$row.Updates | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}#endobject
								$i++
							}#endforeach
						}#endif
						#Write-Host "3: [$($row.ScheduledTask.Count)][$(if ( $row.ScheduledTask) {"True"})][$(if ( $null -eq $row.ScheduledTask ) {"True"})]"
						if ( $null -ne $row.ScheduledTask ) {
							$row.ScheduledTask | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}#endobject
								$i++
							}#endforeach
						}#endif
						#Write-Host "4: [$($row.ErrorMsg.Count)][$(if ( $row.ErrorMsg) {"True"})][$(if ( $null -eq $row.ErrorMsg ) {"True"})]"
						if ( $null -ne $row.ErrorMsg ) {
							$row.ErrorMsg | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}#endobject
								$i++
							}#endforeach
						}#endif
					}#endforeach
					if ( $PSCmdlet.ParameterSetName -like "Name[cu]*" ) {
						Write-Host "`nReport:"
						Write-Host "======="
						$ResultOutput | Sort-Object -Property id | Format-Table -Property Name, Title, Message -AutoSize `
						| Out-String -Stream | Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" }
					}
					else {
						Write-Output "`n[$(Get-Date -Format "yyyy.MM.dd HH:mm:ss")][$ScriptUser][$InPath]-------------------------" `
						| Out-File -FilePath $OutputToFile -Encoding 'ASCII' -Append -Force	
						# OutputToFile
						$ResultOutput | Sort-Object -Property id | Format-Table -Property Name, Title, Message -AutoSize `
						| Out-String -Stream | Where-Object { $_ -ne "" } `
						| Out-File -FilePath $OutputToFile -Encoding 'ASCII' -Append -Force	
						Write-Host "The report file is [ $OutputToFile ]." 
						Write-msg -log -text "[Main] The report file is [ $OutputToFile ]."
					}

				}#endelse
			}#endif
		}#endtry
		catch {
			Write-ErrorMsg -Name 'JObRunners' -InputObject $_
		}#endcatch
		finally {
			if ( $CompSession.count -gt 0 ) {
				Remove-PSSession -Session $CompSession
			}#endif
		}#endfinally
	}#endif
	else {
		if ( $PSCmdlet.ParameterSetName -like "Name[cu]*" -or $PSCmdlet.ParameterSetName -like "InPath[cu]*" ) {
			Write-msg -log -bug -text "[$(if($Check){"Check"}elseif($Update){"Update"})] No computer in list."
		}#endif
	}#endelse
	<# ---------------------------------------------------------------------------------------------------------
		Izsaucam EventLog skriptu: NameEventLog or InPathEventLog; $CompEventsFileName 	= "lib\Get-CompEvents.ps1"
		# Get-CompEvents.ps1 [-InPath] <FileInfo> [-OutPath <switch>] [-Days <int>] [<CommonParameters>]
		# Get-CompEvents.ps1 [-Name] <string> [-Days <int>] [<CommonParameters>]
	--------------------------------------------------------------------------------------------------------- #>

	if ( ( $RemoteComputers.count -gt 0 ) -and $EventLog ) {
		try {
			Write-msg -log -text "[EventLog]:got:[$($RemoteComputers.count)]"
			#Write-Host "[EventLog]:got:[$($RemoteComputers.count)]"

			$__throwMessage = $null
			if (-NOT ( Test-Path -Path $CompEventsFile -PathType Leaf )) {
				throw $__throwMessage = "[EventLog] script file [$CompEventsFile] not found. Fatal error."
			}#endif

			if ( $PSCmdlet.ParameterSetName -like "InPathEventLog" ) {
				$tmpFile = "$DataDir\$ScriptRandomID.tmp"
				$RemoteComputers | Out-File $tmpFile -Force
				$ArgumentList = "`-InPath $((Resolve-Path $tmpFile).Path) $(if($OutPath) {" `-OutPath `-InPathFileName $((Resolve-Path $InPath).Path) "})$(if($Days){" `-Days $Days "})"
			}#endif
			else {
				$tmpName = $RemoteComputers[0]
				$ArgumentList = "`-Name $tmpName $(if($Days){" `-Days $Days "})"
			}#endelse
			Write-Verbose "[Invoke-Expression] `& $CompEventsFile $ArgumentList "
			Invoke-Expression "& `"$CompEventsFile`" $ArgumentList "
			if ( $PSCmdlet.ParameterSetName -like "InPathEventLog" ) { Remove-Item -Path $tmpFile }
		}#endtry
		catch {
			if ($__throwMessage) {
				Write-msg -log -bug -text "$__throwMessage"
			}#endif
			else {
				Write-ErrorMsg -Name 'EventLog' -InputObject $_
			}#endelse
		}#endcatch
	}#endif
	else {
		if ( $PSCmdlet.ParameterSetName -like "*EventLog" ) {
			Write-msg -log -bug -text "[EventLog] No computer in list."
		}#endif
	}

	<# ---------------------------------------------------------------------------------------------------------
		Izpildām Install/Uninstall skriptu - $CompProgramFileName = "lib\Set-CompProgram.ps1"
		Izpildām WakeOnLan skriptu - $CompWakeOnLanFileName	= "lib\Invoke-CompWakeOnLan.ps1"
	--------------------------------------------------------------------------------------------------------- #>
	if ( ( $RemoteComputers.count -gt 0 ) -and ( $Install -or $Uninstall -or $WakeOnLan ) ) {

		try {
			Write-msg -log -text "[$(if($Install){"Install"}elseif($Uninstall){"Uninstall"}else{"WakeOnLan"})]:got:[$($RemoteComputers.count)]"
			Write-Verbose "[$(if($Install){"Install"}elseif($Uninstall){"Uninstall"}else{"WakeOnLan"})]:got:[$($RemoteComputers.count)]"
			
			$__throwMessage = $null
			if (-NOT ( Test-Path -Path $CompProgramFile -PathType Leaf )) {
				throw $__throwMessage = "[$(if($install){"Install"}elseif($Uninstall){"Uninstall"})] script file [$CompProgramFile] not found. Fatal error."
			}#endif
			if (-NOT ( Test-Path -Path $CompWakeOnLanFile -PathType Leaf )) {
				throw $__throwMessage = "[WakeOnLan] script file [$CompWakeOnLanFile] not found. Fatal error."
			}#endif

			if ( $PSCmdlet.ParameterSetName -eq "Name4Install" -or $PSCmdlet.ParameterSetName -eq "InPath4Install" ) {

				# Set-CompProgram.ps1 [-ComputerName] <string> [-InstallPath <FileInfo>] [<CommonParameters>]
				# Mainīgo $Install nepadodam uz Set-Jobs funkciju, jo tā tips nav string. To padosim uz skriptblocku pa tiešo

				Write-Verbose "[Installer] Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBInstall' -Argument1 $Install"
				Write-Host "[Installer] waiting for results:"
				$JobResults = Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBInstall' #-Argument1 $Install
			}#endif
			elseif ( $PSCmdlet.ParameterSetName -eq "Name4Uninstall" -or $PSCmdlet.ParameterSetName -eq "InPath4Uninstall" ) {

				#region kriptējam $Uninstall parametru
				$secParameter = $Uninstall | ConvertTo-SecureString -AsPlainText -Force
				$EncryptedParameter = $secParameter | ConvertFrom-SecureString
				#end region

				# Set-CompProgram.ps1 [-ComputerName] <string> [-CryptedIdNumber <string>] [<CommonParameters>]

				Write-Verbose "[Uninstaller] Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBUninstall' -Argument1 $EncryptedParameter"
				Write-Host "[Uninstaller] waiting for results:"
				$JobResults = Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBUninstall' -Argument1 $EncryptedParameter
			}#endelseif
			elseif ( $PSCmdlet.ParameterSetName -eq "WakeOnLanName" -or $PSCmdlet.ParameterSetName -eq "WakeOnLanInPath"  ) {

				# Invoke-CompWakeOnLan.ps1 [-ComputerName] <string[]> [-DataArchiveFile] <FileInfo> [-CompTestOnline] <FileInfo> [<CommonParameters>]
				# Mainīgo $DataArchiveFile un $CompTestOnlineFile nepadodam uz Set-Jobs funkciju, jo tā tips nav string. To padosim uz skriptblocku pa tiešo

				Write-Verbose "[Waker] Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBWakeOnLan' "
				Write-Host "[Waker] waiting for results:"
				$JobResults = Set-Jobs -Computers $RemoteComputers -ScriptBlockName 'SBWakeOnLan'
				
			}#endelseif
			
			Write-Host "`n[Results]=====================================================================================================" -ForegroundColor Yellow
			$JobResults | Sort-Object -Property Computer, id `
			| Format-Table Computer, Message -AutoSize  `
			| Out-String -Stream | Where-Object { $_ -ne "" } `
			| ForEach-Object { if ( $_ -match '.*SUCCESS.*' ) { Write-Host "$_" -ForegroundColor Green } `
					elseif ($_ -match '.*ERROR.*') { Write-Host "$_" -ForegroundColor Red } `
					elseif ($_ -match '.*WARN.*') { Write-Host "$_" -ForegroundColor Yellow } `
					else { Write-Host "$_" } }
			Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow

		}#endtry
		catch {
			if ($__throwMessage) {
				Write-msg -log -bug -text "$__throwMessage"
			}#endif
			else {
				Write-ErrorMsg -Name $(if ($Install) { "Install" }elseif ($Uninstall) { "Uninstall" }else { "WakeOnLan" }) -InputObject $_
			}#endelse
		}#endcatch
	}#endif
}#endOfprocess

END {
	Stop-Watch -Timer $scriptWatch -Name Script
}#endOfend