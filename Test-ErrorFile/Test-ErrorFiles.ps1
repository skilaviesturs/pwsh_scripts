<#
.SYNOPSIS
This function find files by particular paterns and pathes declared in json files

.DESCRIPTION
If the program found the file by given pattern, it writes in the log the messages tagged as error, which can be monitored by Zabbix

.PARAMETER mail
Ja ieslēgts, tad skripts nosūta e-pastu par rezultātu

.PARAMETER version
Skripts parāda savu versiju

.EXAMPLE
PS> Check-ErrorFiles.ps1 [-mail][-version][-common parameters]

.NOTES
	Author:	Viesturs Šķila
	Version: 1.1.0
#>
[cmdletbinding(DefaultParameterSetName="default")]
[alias("tefs","Test-EFS")]
param(
	[switch]$mail,
	[switch]$version
)
begin {
	Clear-Host
	$CurVersion = "1.1.0"
	if ( $version ) {
		Write-Host "`n$CurVersion`n"
		Exit
	}
	$stopwatch = [System.Diagnostics.Stopwatch]::startNew()
	#skripta update vajadzībām
	$__ScriptName = $MyInvocation.MyCommand
	$__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
	$Script:forMailReport = @()
	#logošanai
	$LogFileDir = "$__ScriptPath\log"
	$LogFile = "$((Get-ChildItem $__ScriptName).BaseName)-$(Get-Date -Format "yyyyMMdd")"
	
	<# ----------------
	Define global variables
	#>
	$jsonFileCfg = "$__ScriptPath\TEFconfig.json"

	$CfgTmpl = [PSCustomObject]@{
		'TestDir' = $null
		'LogFileName' = $null
		'LogDirName' = $null
		'SmtpServer' = $null
		'FilePaternsToCheck' = $null
		'TestFilePatern' = $null
		'SendMailTO' = $null
		'SendMailCC' = $null
		'SendReport' = $null
	}
	
	$PaternTmpl = [PSCustomObject]@{
		'CheckToExist' = $null
		'DateFormat' = $null
		'DatePositionInPattern' = $null
		'Pattern1Part' = $null
		'Pattern2Part' = $null
		'Pattern3Part' = $null
		'Directories' = @()
	}

	<# ----------------
	Declare write-log function
	Function's purpose: write text to screen and/or log file
	Syntax: wrlog [-log] [-bug] -text <string>
	#>
    Function Write-msg { 
        [CmdletBinding(DefaultParameterSetName = "default")]
        [alias("wrlog")]
        Param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$text,
            [switch]$log,
            [switch]$bug
        )#param
        try {
            if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO' }
            $timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
            if ( $log -and $bug ) {
                Write-Warning "[$flag] $text"	
                Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                    Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
            }#endif
            elseif ( $log ) {
                Write-Verbose $flag"`t"$text
                Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                    Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
            }#endelseif
            else {
                Write-Verbose $flag"`t"$text
            }#else
        }#endtry
        catch {
            Write-Warning "[Write-msg] $($_.Exception.Message)"
            $string_err = $_ | Out-String
            Write-Warning "$string_err"
        }#endtry
    }#endOffunction

	<# ----------------
	Declare Get-jsonData function
	Function's purpose: to get the data from the json file, to parse and to valdiate against the template ojects' values by name
	Syntax: gjData [object] [jsonFileName]
	#>
	Function Get-jsonData {
		[cmdletbinding(DefaultParameterSetName="default")]
		[alias("gjData")]
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[object]$template,
			[Parameter(Position = 1, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$jsonFile
		)
	Write-Debug "[gjData] jsonFile : [$jsonFile]"
		$jsonData = [PSCustomObject]{}
		try {
			$jsonData = Get-Content -Path $jsonFile -Raw -ErrorAction STOP | ConvertFrom-Json
		} catch {
			# throws error if file's format is not valid json or have unsupported special symbols in values, for example "\"
			Write-msg -log -bug -text "[gjData] Fatal error - file [$jsonFile] corrupted. Exit." -ForegroundColor Red
			Exit 2 # finish script
		} #catch
		if ( ($jsonData.count) -gt 0 ) {
			Write-Debug "[gjData] [$jsonFile].count : $($jsonData.count)"
			# let's compare properties of $Cfg and $jsonData 
			$aa = $template[0].psobject.properties.name | Sort-Object
			$bb = $jsonData[0].psobject.properties.name | Sort-Object
			foreach ($name in $bb) {
				# if found unexpected property's name, terminate script 
				if ( $aa -eq $name ) {
					Write-Debug "`t template[$aa] = jsonData[$name)]"
				} else {
					Write-msg -log -bug -text "[gjData] Fatal error - Unknown variable name [$name] in import file [$jsonFile]. Exit." -ForegroundColor Red 
					Exit 3 # exit script
				} #else
			} #foreach
		} # if
	return $jsonData
	} #endOffunction
	
	<# ----------------
	Declare Repair-JSONfile function
	Function's purpose: make json File paterns default cfg file, if particular json file does not exist
	Syntax: mjPaterns [object] [jsonFileName]
	#>
	Function Repair-JSONfile {
		[cmdletbinding(DefaultParameterSetName="default")]
		[alias("repJFile")]
		Param(
			[Parameter(Position = 0, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[System.Object]$object,
			[Parameter(Position = 1, Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$jsonFile
		)
		try {
			$object | ConvertTo-Json | Out-File $jsonFile
		}
		catch {
			Write-msg -log -bug -text "[repJFile] $($_.Exception.Message)"
		}
	} #endOffunction

	# import configuration parameters from json 
	if ( -not ( Test-Path -Path $jsonFileCfg -PathType Leaf) ) {
		Write-Warning "[Check] Config JSON file [$jsonFileCfg] not found."
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'LogFileName' -Value 'Test-ErrorFile.log' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'LogDirName' -Value 'log' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'FilePaternsToCheck' -Value 'SetPatterns.json' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'TestFilePatern' -Value 'T*error*.log' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'TestDir' -Value 'log\' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SendMailTO' -Value 'SendMailTO.txt' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SendMailCC' -Value 'SendMailCC.txt' -Force
		$CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SendReport' -Value 'SendReport.txt' -Force
		Repair-JSONfile $CfgTmpl $jsonFileCfg
		Write-Warning "[Check] Config JSON file created."
		$Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
	} else {
		$Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
	} 
	
	$lFile = "$($Cfg.LogDirName)\$($Cfg.LogFileName)"
	if ( -not ( Test-Path -Path $Cfg.LogDirName ) ) { $null = New-Item -ItemType "Directory" -Name $Cfg.LogDirName }
	Write-Debug "[Check] lFile : [$lFile]"

	<# ----------------
	Check and create neccesary files for script in filesystem if not exist
	#>
	try {
		if ( -not ( Test-Path -Path $Cfg.FilePaternsToCheck -PathType Leaf) ) {
			Write-msg -log -text "[Check] Config JSON file [$($Cfg.FilePaternsToCheck)] not found. Restored config file with default parameters."
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'CheckToExist' -Value 'False' -Force
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'DateFormat' -Value 'yyyy.MM.dd' -Force
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'DatePositionInPattern' -Value 4 -Force
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'Pattern1Part' -Value $Cfg.TestFilePatern -Force
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'Pattern2Part' -Value '' -Force
			$PaternTmpl | Add-Member -MemberType NoteProperty -Name 'Pattern3Part' -Value '' -Force
			$PaternTmpl.Directories += $Cfg.TestDir
			$PaternTmpl.Directories += ".\"
			Repair-JSONfile $PaternTmpl $Cfg.FilePaternsToCheck
		} #endif
		
		# define variables for smtp
		if ( $mail ) {
			if ( -not [String]::IsNullOrWhiteSpace($Cfg.SmtpServer) ) {
				Write-Debug "[Check] Send mail set to: [$mail]"
				$mailTo = @()
				$mailCC = @()
				$mailReport = @()
				
				$sendError = @{}
				$sendReport = @{}
				$mailParam = @{
					SmtpServer = $Cfg.SmtpServer
					From = "no-reply@$($env:computername).ltb.lan"
				}
				# test filesystem for neccesary files; if not found, create missing files
				if ( -not ( Test-Path -Path $Cfg.SendMailTO -PathType Leaf) ) {
					$null = New-Item -ItemType "File" -Name $Cfg.SendMailTO
				} #endif
				if ( -not ( Test-Path -Path $Cfg.SendMailCC -PathType Leaf) ) {
					$null = New-Item -ItemType "File" -Name $Cfg.SendMailCC
				} #endif
				if ( -not ( Test-Path -Path $Cfg.SendReport -PathType Leaf) ) {
					$null = New-Item -ItemType "File" -Name $Cfg.SendReport
				} #endif

				# get data from files
				$mailTo = Get-Content $Cfg.SendMailTO
				$mailCC = Get-Content $Cfg.SendMailCC
				$mailReport = Get-Content $Cfg.SendReport
				
				#check arrays to have data; if not - set mail parameter to false
				if ( $mailTo.count -gt 0 -and $mailReport.count -gt 0 ) { 
					$sendError.Add( "To", $mailTo )
					$sendReport.Add( "To", $mailReport )
				} else { 
					$mail = $false 
					Write-msg -log  -bug -text "[Check] No recipients set in [$(($Cfg).SendMailTO)] or [$(($Cfg).SendReport)]. Parameter mail switched to [$mail]."
				} #endelse
				if ( $mailCC.count -gt 0 -and $mail ) { 
					$sendError.Add( "CC", $mailCC )
					$sendReport.Add( "CC", $mailCC )
				} #endif
			} else {
				$mail = $false 
				Write-msg -log  -bug -text "[Check] No smtp server set in configuration file [$jsonFileCfg]. Switched parameter mail to [$mail]."
			} #endelse
		} #endif
	} #endtry
	catch {
		Write-msg -log  -bug -text "[Check] $($_.Exception.Message)"
	} #endOftry

} #endOfbegin

process {
	Write-msg -log -text "[-----] Script started"

	<# ----------------
	Gennerate FinalPattern from imported json file and write it back to array
	#>
	try {
		$rawPaterns = gjData $PaternTmpl $Cfg.FilePaternsToCheck
		#$rawPaternsCopy = gco $rawPaterns
		<# for testing
		$rawPaterns
		# #>
		Write-Debug "[FPatt] Verifying data from file [$($Cfg.FilePaternsToCheck)]"
		if ( $rawPaterns.count -eq 0 ) {
			Write-msg -log -bug -text "[FPatt] File [$($Cfg.FilePaternsToCheck)] is empty. Exit."
			exit 4
			} else {
				$PatternCount = 0
				$Pattern = $null
				foreach ( $item in $rawPaterns ) {
					$rawDate = (Get-Date -Format "$($item.DateFormat)" )
					if ( -not [String]::IsNullOrWhiteSpace($item.DatePositionInPattern) ) {
						switch ( $item.DatePositionInPattern ) {
							1 { $Pattern = "$rawDate$(($item).Pattern2Part)$(($item).Pattern3Part)"; break }	
							2 { $Pattern = "$(($item).Pattern1Part)$rawDate$(($item).Pattern3Part)"; break }
							3 { $Pattern = "$(($item).Pattern1Part)$(($item).Pattern2Part)$rawDate"; break }
							Default { 
								$Pattern = "$(($item).Pattern1Part)$(($item).Pattern2Part)$(($item).Pattern3Part)"
								$item.DatePositionInPattern = 'False'
							} 
						} #switch
						#Put new property in array - if found: assign value, if not - add new property
						if ( Get-Member -InputObject $item -Name 'FinalPattern' ) { $item.'FinalPattern' = $Pattern }
							else { $rawPaterns | Add-Member -MemberType NoteProperty -Name 'FinalPattern' -Value $Pattern }
	
						Write-Debug "[FPatt[$PatternCount]] Raw variables for Pattern[$PatternCount]`nPoz:`t`t$($item.DatePositionInPattern)`nDateMask:`t$($item.DateFormat)`nDate:`t$rawDate`n-------`nGenerated pattern: $($item.FinalPattern)"
						$PatternCount++
					} else {
						Write-Debug "[FPatt[$PatternCount]] is empty. Skipped."
						Write-msg -log -bug -text "[FPatt[$PatternCount]] was empty and skipped. Please check file [$($Cfg.FilePaternsToCheck)]."
					}
				} #foreach
			} #else
		} catch {
			Write-msg -log  -bug -text "[FPatt] $($_.Exception.Message)"
			Exit 5
	} #catch

	<# ----------------
	Main part of programm
	#>
	$resultErrors = 0
	$resultGood = 0
	
	foreach ( $patt in $rawPaterns ) {

		foreach ( $dir in $patt.Directories ) {
			$testPathParam = @{
				Path = "$dir"
			}
			Write-Debug "[Main] [$dir] exist: $( Test-Path @testPathParam )"
			Write-Debug "[Main] [$($patt.FinalPattern)] in [$dir] exist: $( Test-Path -path $dir'\*' -PathType leaf -Include $patt.FinalPattern )"
			#If the error file exist, write log with ERROR and count it as error result.
			if ( Test-Path -path $dir'\*' -PathType leaf -Include $patt.FinalPattern ) {
				$ErrorFile = ( Get-ChildItem "$dir\$(($patt).FinalPattern)" ).name
				Write-Debug "Error file is [$ErrorFile]"
				Write-msg -log  -bug -text "Check file [$ErrorFile] in the direcory [$dir]. Found with patern [$(($patt).FinalPattern)]."
				if ( $mail ) {
					try {
						Send-MailMessage @sendError @mailParam `
						-Subject "[ERROR] In the [$dir] detected the error file by pattern [$(($patt).FinalPattern)]" `
						-Body "$(Get-Date -Format "yyyy.MM.dd")`nCheck the directory [$dir] on error.`nError file [$ErrorFile] has been found!`n`nPowered by Powershell"`
						-ErrorAction Stop
					} catch {
						$mail = $false
						Write-msg -log  -bug -text "[smtpErr] Mail failed to send with error: $($_.Exception.Message)."
						Write-msg -log  -bug -text "[smtpErr] Mail parameter switched to [$mail]."

					} #endcatch
				} #endif
				$resultErrors++

				} #endif
				#If the file doesn't exists, count it as good result.
				else {
					$resultGood++
			} #else
		} # foreach
	} #foreach
} #endOfprocess

end {
	<#
	finalize achived results
	#>
	$logError = $false
	foreach ( $a in $Script:forMailReport) {
		if ( $a.Contains("[ERROR]") ) { $logError = $true }
	}
	if ( $mail ) {
		try {
			$mailReportBody = @"
			<!DOCTYPE html>
			<html>
			<head>
			<style>
			h2 {
			  color: blue;
			  font-family: tahoma;
			  font-size: 100%;
			}
			p {
			  font-family: tahoma;
			  color: blue;
			  font-size: 80%;
			  margin: 0;
			}
			table {
			  font-family: tahoma, sans-serif;
			  border-collapse: collapse;
			  width: 75%;
			  font-size: 90%;
			}
			
			td, th {
			  border: 1px solid #dddddd;
			  text-align: left;
			  padding: 8px;
			  font-size: 90%;
			}
			
			tr:nth-child(even) {
			  background-color: #dddddd;
			}
			</style>
			</head>
			<body>
			<h2>$(Get-Date -Format "yyyy.MM.dd") summary:</h2>
			$(if ($resultErrors -eq 0 ) { '<p style="color:green;"><mark>Everything is <strong>fine</strong> and you can rest.</mark></p>' } else { '<p style="color:red;"><mark><strong>Check your mail!</strong> Something wrong happened!</mark></p>' } )
			<br>
			<table>
			  <tr>
				<td>Patterns used:</td>
				<td>$PatternCount</td>
			  </tr>
			  <tr>
				<td>Directories scanned:</td>
				<td>$($resultErrors+$resultGood)</td>
			  </tr>
			  <tr>
				<td> - problematic directories:</td>
				<td>$resultErrors</td>
			  </tr>
			  <tr>
				<td> - good directories:</td>
				<td>$resultGood</td>
			  </tr>
			</table>
			<br>
			<p>Script finished in $([math]::Round(($stopwatch.Elapsed).TotalSeconds,4)) seconds.</p>
			<br><br>
			<p>Powered by Powershell</p>
			<p style="font-size: 60%;color:gray;">[Test-ErrorFile] version $CurVersion</p>
			</body>
			</html>
"@
			Send-MailMessage @sendReport @mailParam `
				-Subject "[$__ScriptName]:$(if ($logError) {":[ERROR]"} else {":[SUCCESS]"}) report from [$($env:computername)]" -Body $mailReportBody -BodyAsHtml -ErrorAction Stop
		} #endtry
		catch {
			Write-msg -log  -bug -text "[smtpRep] Mail failed to send with error: $($_.Exception.Message)"
		} #endcatch
	} #endif
	Write-msg -log -text "[Main] Used [$PatternCount] patterns. Scanned [$($resultErrors+$resultGood)] directories: in [$resultErrors] found files by patterns, but in [$resultGood] - nothing found."
	$stopwatch.Stop()
	Write-msg -log -text "[-----] Script finished in $([math]::Round(($stopwatch.Elapsed).TotalSeconds,4)) seconds."
} #endOfend
