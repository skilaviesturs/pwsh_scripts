<#Set-FileArchive.ps1
.SYNOPSIS
The script check source directories, copy to temp files, arhive them, move to destination arhive folder and delete arhived files

.DESCRIPTION
Usage:
Set-FileArchive.ps1 [-Update] [-Test] [-ShowConfig] [-V] [<CommonParameters>]

Parameters:
    [-Update<SwitchParameter>] - check for script new versions and update script;
    [-Test<SwitchParameter>] - run script without delete the source files;
	[-ShowConfig<SwitchParameter>] - show configuration parameters from json file;
	[-V<SwitchParameter>] - print out script's version;
	[-Help<SwitchParameter>] - print out this help.

.EXAMPLE


.NOTES
	Author:	Viesturs Skila
	Version: 1.0.3
#>
[CmdletBinding()] 
param (
    [switch]$Test,
    [switch]$ShowConfig,
	[switch]$V,
	[switch]$Help,
    [switch]$Update
)
begin {

	$CurVersion = "1.0.3"
    $StopWatch = [System.Diagnostics.Stopwatch]::startNew()
    #skripta update vajadzībām
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    $UpdateDir = "\\beluga\install\scripts\LogFileArchive"
    #Skritpa konfigurācijas datne
    $jsonFileCfg = "$__ScriptPath\SAFconfig.json"
    $jsonFtpCfg = "$__ScriptPath\FTPconfig.json"
    #logošanai
    $LogFileDir = "$__ScriptPath\log"
	$LogFile = "$LogFileDir\FileArchive_$(Get-Date -Format "yyyyMMdd")"
    $ScriptRandomID = Get-Random -Minimum 100000 -Maximum 999999
	$ScriptUser = Invoke-Command -ScriptBlock { whoami }
    $Script:forMailReport = @()
    #WinSCP integrācija
    $WinSCPDir = "$__ScriptPath\lib"
    $WinSCPnetFile = "WinSCPnet.dll"
    $WinSCPexeFile = "WinSCP.exe"
    $WinSCPLogFile = "$LogFileDir\sftp_log_$(Get-Date -Format "yyyyMMdd").log"
    #mail integrācija
    $SmtpServer = 'mail.ltb.lan'
    $mailTo = 'viesturs.skila@expobank.eu'
    #šeit iestatam arhivēšanas periodu
    #tagad iestatīts, kad tiek saarhivēts viss, kas vecāks par iepriekšējā mēneša pirmo datumu
    $today = [datetime]::Today
    $maxAge = $today.AddDays(1 - $today.Day).addMonths(-1)
    #Konfigurācijas JSON datnes parauga objekts
	$CfgTmpl = [PSCustomObject]@{
        'SourceFolder' = $null
        'SourceFileType' = $null
        'SourceComputer' = $null
		'ArchiveRootFolder' = $null
		'Application' = $null
        'ToFTP' = $null
	}#endOfobject

    $FtpTmpl = [PSCustomObject]@{
        'FtpHostName' = $null
        'UserName' = $null
        'Password' = $null
        'FtpPort' = $null
        'SshPrivateKeyFileName' = $null
        'SshHostKeyFingerprint' = $null
    }#endOfobject

	if ( $V ) {
		Write-Host "`n$CurVersion`n"
		Exit
	}#endif
	Function Show-ScriptHelp {
		Write-Host "`nUsage:`nSet-FileArhive.ps1 [-Update] [-Test] [-ShowConfig] [-V] [<CommonParameters>"
		Write-Host "`nDescription:`n  The script check source directories, arhive them, move to destination arhive folder and delete arhived files"
		Write-Host "`nParameters:"
        Write-Host "  [Update<SwitchParameter>]`t- check for script new versions and update script;"
        Write-Host "  [Test<SwitchParameter>]`t- run script without delete the source files;"
        Write-Host "  [ShowConfig<SwitchParameter>]`t- show configuration parameters from json file;"
		Write-Host "  [V<SwitchParameter>]`t- print out script's version;"
		Write-Host "  [Help<SwitchParameter>]`t- print out this help.`n"
	}#endOfFunction

    if ( $help ) {
        Show-ScriptHelp
        Exit
    }#endif

	if ( -not ( Test-Path -Path $LogFileDir ) ) { $null = New-Item -ItemType "Directory" -Path $LogFileDir }
    if ( -not ( Test-Path -Path $WinSCPDir ) ) { $null = New-Item -ItemType "Directory" -Path $WinSCPDir }

<# ----------------
	Declare write-log function
	Function's purpose: write text to screen and/or log file
	Syntax: wrlog [-log] [-bug] -text <string>
	mandatory flag:
		-text		# write your text on the screen or in the log file;
	optional flag:
		-log		# write text in the log file
		-bug		# marks text line with ERROR stamp; if not set, marks text line with INFO stamp
	#>
	Function Write-msg { 
		[CmdletBinding(DefaultParameterSetName="default")]
		[alias("wrlog")]
		Param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$text,
			[switch]$log,
			[switch]$bug
		)#param

		try {
			#write-debug "[wrlog] Log path: $log"
			if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO'}
			$timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
			if ( $log -and $bug ) {
                Write-Warning "[$flag] $text"	
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                Out-File "$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
			}#endif
            elseif ( $log ) {
                Write-Verbose "[$flag] $text"
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                Out-File "$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
			}#endelseif
            else {
                Write-Verbose "[$flag] $text"
			}#else
		}#endtry
		catch {
            Write-Warning "[Write-msg] $($_.Exception.Message)"
			return
		}#endOftry
	}#endOffunction
    
    Function Stop-Watch {
        $stopwatch.Stop()
        if ( $stopwatch.Elapsed.Minutes -le 9 -and $stopwatch.Elapsed.Minutes -gt 0 ) { $bMin = "0$($stopwatch.Elapsed.Minutes)"} else { $bMin = "$($stopwatch.Elapsed.Minutes)"}
        if ( $stopwatch.Elapsed.Seconds -le 9 -and $stopwatch.Elapsed.Seconds -gt 0 ) { $bSec = "0$($stopwatch.Elapsed.Seconds)"} else { $bSec = "$($stopwatch.Elapsed.Seconds)"}
        Write-msg -log -text "[Script] finished in $(
            if ( [int]$stopwatch.Elapsed.Hours -gt 0 ) {"$($stopwatch.Elapsed.Hours)`:$bMin`:$bSec hrs"}
            elseif ( [int]$stopwatch.Elapsed.Minutes -gt 0 ) {"$($stopwatch.Elapsed.Minutes)`:$bSec min"}
            else { "$($stopwatch.Elapsed.Seconds)`.$($stopwatch.Elapsed.Milliseconds) sec" }
            )"
    }#endOffunction

    <# ----------------
	Formējam un sūtam e-pastus
	#>
    Function Send-Mail {
        foreach ( $a in $Script:forMailReport) {
            if ( $a.Contains("[ERROR]") ) { $logError = $true }
        }
        $server = $env:computername.ToLower()
        $mailParam = @{
            SmtpServer = $SmtpServer
            To = $mailTo
            From = "no-reply@$server.ltb.lan"
            Subject = "[$($env:computername)]$(if ($logError) {":[ERROR]"} else {":[SUCCESS]"}) report from [$__ScriptName]"
        }
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
          width: 100%;
          font-size: 90%;
        }
        
        td, th {
          border: 1px solid #dddddd;
          text-align: left;
          padding: 8px;
          font-size: 75%;
        }
        
        tr:nth-child(even) {
          background-color: #dddddd;
        }
        </style>
        </head>
        <body>
        <h2>$(Get-Date -Format "yyyy.MM.dd HH:mm:ss") events from log file [$LogFile.log]</h2>
        <br>
        <table><tbody>
            $( ForEach ( $line in $Script:forMailReport ) { if ($line.Contains("[ERROR]")) {"<tr><td><p style=font-size:125%;color:red;>$line</p></td></tr>"} else {"<tr><td>$line</td></tr>"} } )
        </tbody></table>
        <br>
        <p>Script finished in $([math]::Round(($stopwatch.Elapsed).TotalSeconds,4)) seconds.</p>
        <br><br>
        <p>Powered by Powershell</p>
        <p style="font-size: 60%;color:gray;">[$__ScriptName] version $CurVersion</p>
        </body>
        </html>
"@
        try {
            Send-MailMessage @mailParam -Body $mailReportBody -BodyAsHtml -ErrorAction Stop
        }#endtry 
        catch {
            Write-msg -log -bug -text "[smtpErr] Mail failed to send with error: $($_.Exception.Message)."
        }#endcatch

    }#endOffunctio

    <# ----------------
        Declare Get-jsonData function
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
        Write-msg -text "[gjData] jsonFile : [$jsonFile]"
        $jsonData = [PSCustomObject]{}
        try {
            $jsonData = Get-Content -Path $jsonFile -Raw -ErrorAction STOP | ConvertFrom-Json
        } 
        catch {
            # throws error if file's format is not valid json or have unsupported special symbols in values, for example "\"
            Write-msg -log -bug -text "[gjData] Fatal error - file [$jsonFile] corrupted. Exit."
            Stop-Watch
            Send-Mail
            Exit    # exit script
        } #catch
        if ( ($jsonData.count) -gt 0 ) {
            Write-msg -text "[gjData] [$jsonFile].count : $($jsonData.count)"
            # let's compare properties of $Cfg and $jsonData 
            $aa = $template[0].psobject.properties.name | Sort-Object
            $bb = $jsonData[0].psobject.properties.name | Sort-Object
            foreach ($name in $bb) {
                # if found unexpected property's name, terminate script 
                if ( $aa -eq $name ) {
                    Write-Debug "`t template[$aa] = jsonData[$name)]"
                }#endif
                else {
                    Write-msg -log -bug -text "[gjData] Fatal error - Unknown variable name [$name] in import file [$jsonFile]. Exit."
                    Stop-Watch
                    Send-Mail
                    Exit    # exit script
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
        }#endtry
        catch {
            Write-msg -log -bug -text "[repJFile] $($_.Exception.Message)"
        }#endcatch
    }#endOffunction

    Function Set-Housekeeping {
        [cmdletbinding(DefaultParameterSetName="default")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$CheckPath
        )
        #šeit iestatam datņu dzēšanas periodu
        #tagad iestatīts, kad tiek saarhivēts viss, kas vecāks par 30 dienām
        $maxAge = ([datetime]::Today).addDays(-30)

        $filesByMonth = Get-ChildItem -Path $CheckPath -File |
                Where-Object -Property LastWriteTime -lt $maxAge |
                Group-Object { $_.LastWriteTime.ToString("yyyy\\MM") }
        if ( $filesByMonth.count -gt 0 ) {
            try {
                $filesByMonth.Group | Remove-Item -ErrorAction Stop
                Write-msg -log -text "[Housekeeping] Succesfully cleaned files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
            }#endtry
            catch {
                Write-msg -log -bug -text "[Housekeeping] Error: $($_.Exception.Message)"
            }#endcatch
        }#endif
        else {
            Write-msg -log -text "[Housekeeping] There's no files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
        }#endelse
    }#endOfFunctions

    Function Get-ArchiveFileHash {
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$FileName,
            [Parameter(Position = 1, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$OutputFileName,
            [Parameter(Position = 2, Mandatory = $false)]
            [switch]$Check
        )
        $ArchiveFileHash = Get-FileHash $FileName
        if ( $Check ) {
            Write-msg -log -bug -text "[Main] Feature of hash compare is not implemented yet."
        }#endif
        else {
            Write-msg -log -text "[Main] Created archive file [$FileName], hash [$($ArchiveFileHash.Hash)]"
            Write-Output "$(Get-Date -Format "yyyy.MM.dd HH:mm:ss") Added [$FileName], hash [$($ArchiveFileHash.Hash)]" |`
                Out-File -FilePath $OutputFileName -Encoding UTF8 -Append
        }#endelse
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
        $CopyNewFile = $false
		try {
            $ScriptFile = Get-ChildItem "$ScriptPath\$FileName" -ErrorAction Stop
        }#endtry
        catch {
            $CopyNewFile = $true
        }#endtry
        try {
            $NewFile = Get-ChildItem "$UpdatePath\$FileName" -ErrorAction Stop
        }#endtry
        catch {
            Write-msg -log -bug -text "[Update] Update file [$UpdatePath\$FileName] is not accessible."
            Write-msg -log -bug -text "[Update] Error: $($_.Exception.Message)"
        }#endcatch
		if ( $NewFile.count -gt 0 ) {
			if ( $NewFile.LastWriteTime -gt $ScriptFile.LastWriteTime -or $CopyNewFile ) {
				Write-msg -log -text "[Update] Found update for file [$FileName]"
				Write-msg -log -text "[Update] Old version $(if ($CopyNewFile) {"[none],[none]"} else {"[$($ScriptFile.LastWriteTime)],[$($ScriptFile.FullName)]"} )"
				Write-msg -log -text "[Update] New version [$($NewFile.LastWriteTime)], [$($NewFile.FullName)]"
				try {
					Copy-Item -Path $NewFile.FullName -Destination $ScriptPath -Force -ErrorAction Stop
					Write-msg -log -text "[Update] New version deployed."
				}#endtry
				catch {
					Write-msg -log -bug -text "[Update] [$FileName] $($_.Exception.Message)"
				}#endcatch
			}#endif
		}#endif
	}#endOffunction

    Function Move-FilesToFtp {
        param (
            [Parameter(Mandatory = $True, Position = 0)]
            $localPaths,
            [Parameter(Mandatory = $True, Position = 1)]
            $remotePath
        )
        $FtpWatch = [System.Diagnostics.Stopwatch]::startNew()
        try {
            #formējam sesijas parametrus
            $sessionOptions = New-Object -TypeName WinSCP.SessionOptions -Property @{
                HostName = $Ftp.FtpHostName
                Protocol = [WinSCP.Protocol]::sftp
                PortNumber = $Ftp.FtpPort
                UserName = $Ftp.UserName
                Password = $Ftp.Password
                SshPrivateKeyPath = "$WinSCPDir\$($Ftp.SshPrivateKeyFileName)"
                SshHostKeyFingerprint = $Ftp.SshHostKeyFingerprint
            }

            $session = New-Object -TypeName WinSCP.session -Property @{
                ExecutablePath = "$WinSCPDir\$WinSCPexeFile"
            }
            
            try {
                #atveram ftp sesiju
                if ( $Test ) { $session.DebugLogPath = $WinSCPLogFile }
                $session.Open($sessionOptions)
                Write-msg -log -text "[sftp] Succesfully connected to ftp [$($Ftp.UserName)@$($Ftp.FtpHostName)`:$($Ftp.FtpPort)]"
                
                foreach ($localPath in $localPaths) {
                    # ja objekts ir direktorija, meklējam rekursīvi tā ietilpstošos elementus
                    if (Test-Path $localPath -PathType container) {
                        $files =
                            @($localPath) +
                            (Get-ChildItem $localPath -Recurse | Select-Object -ExpandProperty FullName)
                    }#endif
                    else {
                        $files = $localPath
                    }#endelse
                    $parentLocalPath = Split-Path -Parent (Resolve-Path $localPath)
        
                    foreach ($localFilePath in $files) {
                        $remoteFilePath =
                            [WinSCP.RemotePath]::TranslateLocalPathToRemote(
                                $localFilePath, $parentLocalPath, $remotePath)

                        if (Test-Path $localFilePath -PathType container) {
                            # IZveidojam ftp direktoriju, ja tādas neeksistē
                            if ( -not ($session.FileExists($remoteFilePath)) ) {
                                $session.CreateDirectory($remoteFilePath)
                            }#endif
                        }#endif
                        else {
                            if ( -not ($session.FileExists($remoteFilePath)) ) {
                                Write-msg -log -text "[sftp] $(if ($Test){"Copying"}else{"Moving"}) file [$localFilePath] to [$remoteFilePath]"
                                # augšupielādējam failu izdzēšam lokālo, ka parametrs Test nav patiess
                                if ( $Test ){
                                    $session.PutFiles($localFilePath, $remoteFilePath, $false).Check()
                                }#endif
                                else {
                                    $session.PutFiles($localFilePath, $remoteFilePath, $true).Check()
                                    
                                    $ArchiveReadMe = "$($item.ArchiveRootFolder)\$($item.SourceComputer)\Readme.txt"
                                    Write-Output "$(Get-Date -Format "yyyy.MM.dd HH:mm:ss") Moved [$localFilePath] to sftp [$remoteFilePath]" |`
                                    Out-File -FilePath $ArchiveReadMe -Encoding UTF8 -Append

                                }#endesle
                            }#endif
                            else {
                                Write-msg -log -bug -text "[sftp] [$localFilePath] exists on remote [$remoteFilePath]. Please investigate issue."
                            }#endelse
                        }#endelse
                    }#endforeach
                }#endforeach
            }#endtry
            finally {
                # Disconnect, clean up
                if ( $session.Opened ) {
                    $session.Dispose()
                    Write-msg -log -text "[sftp] Succesfully disconnected from ftp"
                }#endif
            }#endfinally
            $result = 0
        }#endtry
        catch {
            Write-msg -log -bug -text "[sftp] Error: $($_.Exception.Message)"
            $result = 1
        }#endcatch

        $FtpWatch.Stop()
        if ( $FtpWatch.Elapsed.Minutes -le 9 -and $FtpWatch.Elapsed.Minutes -gt 0 ) { $bMin = "0$($FtpWatch.Elapsed.Minutes)"} else { $bMin = "$($FtpWatch.Elapsed.Minutes)"}
        if ( $FtpWatch.Elapsed.Seconds -le 9 -and $FtpWatch.Elapsed.Seconds -gt 0 ) { $bSec = "0$($FtpWatch.Elapsed.Seconds)"} else { $bSec = "$($FtpWatch.Elapsed.Seconds)"}
        Write-msg -log -text "[sftp] finished in $(
            if ( $FtpWatch.Elapsed.Hours -gt 0 ) {"$($FtpWatch.Elapsed.Hours)`:$bMin`:$bSec hrs"}
            elseif ( $FtpWatch.Elapsed.Minutes -gt 0 ) {"$($FtpWatch.Elapsed.Minutes)`:$bSec min"}
            else { "$($FtpWatch.Elapsed.Seconds)`.$($FtpWatch.Elapsed.Milliseconds) sec" }
            )"

        return $result
    }#endOffunction

    <# ----------------------------------------------
    Funkciju definēšanu beidzām
    IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
    ---------------------------------------------- #>
    Clear-Host
    Write-msg -log -text "[-----] Script started in [$(if ($Test) {"Test"}
        elseif ($Update) {"Update"} 
		elseif ($ShowConfig) {"ShowConfig"} 
		else {"Default"})] mode. Used config file [$jsonFileCfg]"

    #Pārbaudām vai nav jaunākas versijas repozitorijā $UpdateDir
    if ( -not [String]::IsNullOrWhiteSpace($UpdateDir) -and -not [String]::IsNullOrEmpty($UpdateDir) ) {
		Get-ScriptFileUpdate $__ScriptName $__ScriptPath $UpdateDir
        Get-ScriptFileUpdate $WinSCPnetFile "$__ScriptPath\lib" "$UpdateDir\lib"
        Get-ScriptFileUpdate $WinSCPexeFile "$__ScriptPath\lib" "$UpdateDir\lib"
        if ($Update) {
            Stop-Watch
            exit 
        }
    }#endif

    #ielādējam WinSCP .net bibliotēku
    try {
        Add-Type -Path "$WinSCPDir\$WinSCPnetFile" -ErrorAction Stop
        $isWinSCPnetLoaded = $true
    }#endtry
    catch {
        $isWinSCPnetLoaded = $false
    }#endcatch

    # Importējam konfigurācijas parametrus no FileCfg
    if ( -not ( Test-Path -Path $jsonFileCfg -PathType Leaf) ) {
        Write-msg -log -text "[Check] Config JSON file [$jsonFileCfg] not found."
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SourceFolder' -Value 'D://_LogArchive//log' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SourceComputer' -Value $env:computername -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'ArchiveRootFolder' -Value 'D://_LogArchive' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'Application' -Value 'LogArchive' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SourceFileType' -Value '' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'ToFTP' -Value $false -Force
            
        Repair-JSONfile $CfgTmpl $jsonFileCfg
        Write-msg -log -text "[Check] Config JSON file [$jsonFileCfg] created."
        $Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
    }#endif
    else {
        $Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
    }#endelse

    #Importējam konfigurācijas parametrus no FtpCfg
    if ( -not ( Test-Path -Path $jsonFtpCfg -PathType Leaf) ) {
        Write-msg -log -text "[Check] Config JSON file [$jsonFtpCfg] not found."
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'FtpHostName' -Value "Write here host name of sftp server" -Force
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'UserName' -Value "Write here user name for sftp server" -Force
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'Password' -Value "" -Force
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'FtpPort' -Value '22' -Force
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'SshPrivateKeyFileName' -Value "Write your ppk file name here" -Force
        $FtpTmpl | Add-Member -MemberType NoteProperty -Name 'SshHostKeyFingerprint' -Value "Write here fingerprint of sftp server like: ssh-ed25519 255 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx" -Force

        Repair-JSONfile $FtpTmpl $jsonFtpCfg
        Write-msg -log -text "[Check] Config JSON file [$jsonFtpCfg] created."
        $Ftp = Get-jsonData $FtpTmpl $jsonFtpCfg
    }#endif
    else {
        $Ftp = Get-jsonData $FtpTmpl $jsonFtpCfg
    }#endelse

    #parādam ielādēto konfigurāciju
    if ( $ShowConfig ) {
        Write-Host "Configuration param:"
        Write-Host "--------------------"
        Write-Host "CurVersion`t:[$CurVersion]"
        Write-Host "__ScriptName`t:[$__ScriptName]"
        Write-Host "__ScriptPath`t:[$__ScriptPath]"
        Write-Host "UpdateDir`t:[$UpdateDir]"
        Write-Host "jsonFileCfg`t:[$jsonFileCfg]"
        Write-Host "jsonFtpCfg`t:[$jsonFtpCfg]"
        Write-Host "LogFileDir`t:[$LogFileDir]"
        Write-Host "LogFile`t`t:[$LogFile.log]"
        Write-Host "WinSCPDir`t:[$WinSCPDir]"
        Write-Host "WinSCPnetFile`t:[$WinSCPDir\$WinSCPnetFile]"
        Write-Host "WinSCPexeFile`t:[$WinSCPDir\$WinSCPexeFile]"
        Write-Host "--------------------"
        $Cfg | Format-List *
        $Ftp | Format-List *
        Stop-Watch
        Exit
    }#endif

}#endBegin

process {
    <# ---------------------
    ŠEIT SĀKAS PAMATA DARBS
    --------------------- #>
    #apstrādājam katru cfg ieraksta objektu
    foreach ( $item in $Cfg ) {
        #ja cfg ieraksts neatbilst servera faktiskajam nosaukumam
        if ( $item.SourceComputer -notlike $env:computername ) { $item.SourceComputer = ($env:computername).ToLower() } else { $item.SourceComputer = $item.SourceComputer.ToLower() }
        $ArchiveReadMe = "$($item.ArchiveRootFolder)\$($item.SourceComputer)\Readme.txt"
        $directoriesByMonth = @()
        $filesByMonth = @()
        
        <# --------------------------------------------
        PĀRBAUDAM VAI EKSISTĒ NORĀDĪTĀS DIREKTORIJAS
        -------------------------------------------- #>
        Write-msg -text "SourceFolder [$($item.SourceFolder)]"
        Write-msg -text "ArchiveRootFolder [$($item.ArchiveRootFolder)]"
        if ( -not ( Test-Path -Path $item.SourceFolder ) ) { 
            if ( $Test ) {
                $null = New-Item -ItemType "Directory" -Path $item.SourceFolder 
            }#endif
            else {
                Write-msg -log -bug -text "There's no SourceFolder [$($item.SourceFolder)]."
                Continue
            }#endelse
        }#endif
        if ( -not ( Test-Path -Path $item.ArchiveRootFolder ) ) {
            if ( $Test ) {
                $null = New-Item -ItemType "Directory" -Path $item.ArchiveRootFolder 
            }#endif
            else {
                Write-msg -log -bug -text "There's no ArchiveRootFolder [$($item.ArchiveRootFolder)]."
                Continue
            }#endelse
        }#endif

        <# --------------------------------------------
        PĀRBAUDAM VAI ARHIVĒJAM FAILUU VAI DIREKTORIJAS
        -------------------------------------------- #>
        $DestinationFolder = "$($item.ArchiveRootFolder)\$($item.SourceComputer)\$($item.Application)"
        Write-msg -text "DestinationFolder [$DestinationFolder]"

        # check the source directory for subfolders
        if ( -not ( ( Get-ChildItem -Force -Directory $item.SourceFolder).Count -gt 0 ) )  {
            #get files only from SourceFoler folder except subfolders
            $filesByMonth = Get-ChildItem -Path $item.SourceFolder -File |
                Where-Object -Property LastWriteTime -lt $maxAge |
                Group-Object { $_.LastWriteTime.ToString("yyyy\\MM") }
            Write-msg -log -text "[Check] Directory [$($item.SourceFolder)] contains files. Use [File] archive mode"
        }#endif
        else {
            $directoriesByMonth = Get-ChildItem -Path $item.SourceFolder -Attribute Directory |
                Where-Object -Property LastWriteTime -lt $maxAge |
                Group-Object { $_.LastWriteTime.ToString("yyyy\\MM") }
            Write-msg -log -text "[Check] Directory [$($item.SourceFolder)] contains sub-directories. Use [Directory] archive mode"
        }#endelse

        <# ------------------------------------------
        ARHIVĒJAM FAILUS UN PĀRVIETOJAM LOCAL ARCHIVE
        ------------------------------------------ #>
        #Archive files
        if ( ($filesByMonth.count) -gt 0 ) {

            foreach ( $value in $filesByMonth ) {
                Write-msg -text "filesByMonth.Name [$($value.Name)]"
            }#endforeach
            
            foreach ($monthlyGroup in $filesByMonth) {
                $ArchiveFileHash = ""
                #veidojam gala arhiva datnes nosaukuma gada un mēneša nosaukumu
                $ArchiveYearMonthName = $monthlyGroup.Name
                $ArchiveYearMonthName = $ArchiveYearMonthName.Replace("\","")
                Write-msg -text "ArchiveYearMonthName [$ArchiveYearMonthName]"
                #veidojam gala arhīva sagaidāmo atrašanās vietu
                $ArchiveDir = Join-Path $DestinationFolder $monthlyGroup.Name
                #nočekojam vai eksistē arhīva direktorija, ja nav - izveidojam
                if ( -not ( Test-Path -Path $ArchiveDir ) ) { $null = New-Item -ItemType "Directory" -Path $ArchiveDir }

                #saliekam kopā pilnu arhīva datnes ceļu un vārdu. Datni saglabāsim vienu mapes līmeni augstāk
                $FullArchiveFileName = "$( Split-path $ArchiveDir )\$($item.Application)-$ArchiveYearMonthName.zip"
                Write-msg -text "FullArchiveFileName [$FullArchiveFileName]"

                # čekojam vai sagaidāmā arhīva datne jau neeksistē, je eksistē, tad izvadam paziņojumu
                if (Test-Path "$FullArchiveFileName") {
                    Write-msg -log -bug -text "[Check] Skipping [$($monthlyGroup.Name)] in [$($item.SourceFolder)] because an existing ZIP archive [$FullArchiveFileName] was found. Please verify manually."
                }#endIf
                else {
                    try{
                        Write-msg -text "ArchiveDir [$ArchiveDir]"
                        # čekojam vai Test slēdzis ir aktīvs
                        if ( $Test ) {
                            #kopējam atlasītos avota failus
                            $monthlyGroup.Group | Copy-Item -Destination $ArchiveDir -ErrorAction Stop
                        }#endif
                        else {
                            #pārvietojam atlasītos avota failus
                            $monthlyGroup.Group | Move-Item -Destination $ArchiveDir -ErrorAction Stop
                        }#endelse

                        #kompresējam pagaidu direktorijas saturu
                        Compress-Archive -Path $ArchiveDir -DestinationPath $FullArchiveFileName -ErrorAction Stop

                        #izsaucam funkciju, izveidojam datnes hash summu un ierakstām readme datnē
                        Get-ArchiveFileHash $FullArchiveFileName $ArchiveReadMe

                        #izdzēšam pagaidu direktoriju
                        Remove-Item $ArchiveDir -Recurse -Force -ErrorAction SilentlyContinue

                    }
                    catch {
                        Write-msg -log -bug -text "[MainFile] $($_.Exception.Message)"
                    }#endcatch
                }#endElse
            }#endForEach
        }#endif

        <# --------------------------------------------------
        ARHIVĒJAM LOG DIREKTORIJAS UN PĀRVIETOJAM UZ LOCAL ARCHIVE
        -------------------------------------------------- #>
        #Archive directories
        elseif ( ($directoriesByMonth.count) -gt 0 ) {

            foreach ( $value in $directoriesByMonth ) {
                Write-msg -text "filesByMonth.Name [$($value.Name)]"
            }#endforeach

            foreach ($monthlyGroup in $directoriesByMonth) {
                $ArchiveFileHash = ""
                #veidojam gala arhiva datnes nosaukuma gada un mēneša nosaukumu
                $ArchiveYearMonthName = $monthlyGroup.Name
                $ArchiveYearMonthName = $ArchiveYearMonthName.Replace("\","")
                Write-msg -text "ArchiveYearMonthName [$ArchiveYearMonthName]"
                #veidojam gala arhīva sagaidāmo atrašanās vietu
                $ArchiveDir = Join-Path $DestinationFolder $monthlyGroup.Name
                #nočekojam vai eksistē arhīva direktorija, ja nav - izveidojam
                if ( -not ( Test-Path -Path $ArchiveDir ) ) { $null = New-Item -ItemType "Directory" -Path $ArchiveDir }

                #saliekam kopā pilnu arhīva datnes ceļu un vārdu. Datni saglabāsim vienu mapes līmeni augstāk
                $FullArchiveFileName = "$( Split-path $ArchiveDir )\$($item.Application)-$ArchiveYearMonthName.zip"
                Write-msg -text "FullArchiveFileName [$FullArchiveFileName]"

                # čekojam vai sagaidāmā arhīva datne jau neeksistē, je eksistē, tad izvadam paziņojumu
                if (Test-Path "$FullArchiveFileName") {
                    Write-msg -log -bug -text "[Check] Skipping [$($monthlyGroup.Name)] in [$($item.SourceFolder)] because an existing ZIP archive [$FullArchiveFileName] was found. Please verify manually."
                }#endIf
                else{
                    try {
                        Write-msg -text "ArchiveDir [$ArchiveDir]"
                        # čekojam vai Test slēdzis ir aktīvs
                        if ( $Test ) {
                            #kopējam avota direktorijas saturu
                            foreach ( $dir in $monthlyGroup.Group) {
                                $tempArchiveName = "$($item.Application)-$dir.zip"
                                Copy-Item -Path "$($item.SourceFolder)\$dir" -Destination $ArchiveDir -ErrorAction Stop
                                Compress-Archive -Path "$ArchiveDir\$dir" -DestinationPath "$ArchiveDir\$tempArchiveName" -ErrorAction Stop
                                Remove-Item "$ArchiveDir\$dir" -Recurse -Force -ErrorAction SilentlyContinue
                            }
                        }#endif
                        else {
                            #pārvietojam avota direktorijas saturu
                            foreach ( $dir in $monthlyGroup.Group) {
                                $tempArchiveName = "$($item.Application)-$dir.zip"
                                Move-Item -Path "$($item.SourceFolder)\$dir" -Destination $ArchiveDir -ErrorAction Stop
                                Compress-Archive -Path "$ArchiveDir\$dir" -DestinationPath "$ArchiveDir\$tempArchiveName" -ErrorAction Stop
                                Remove-Item "$ArchiveDir\$dir" -Recurse -Force -ErrorAction SilentlyContinue
                            }
                        }#endelse

                        #kompresējam pagaidu direktorijas saturu
                        Compress-Archive -Path $ArchiveDir -DestinationPath $FullArchiveFileName -ErrorAction Stop
    
                        #izsaucam funkciju, izveidojam datnes hash summu un ierakstām readme datnē
                        Get-ArchiveFileHash $FullArchiveFileName $ArchiveReadMe
    
                        #izdzēšam pagaidu direktoriju
                        Remove-Item $ArchiveDir -Recurse -ErrorAction Stop

                    }#endtry
                    catch {
                        Write-msg -log -bug -text "[MainDir] $($_.Exception.Message)"
                    }#endcatch

                }#endelse
            }#endforeach
        }#endelseif
        else {
            Write-msg -log -text "[Main] There's no files or directories to archive in [$($item.SourceFolder)]."
        }#endelse

        <# ----------------------------------------
        PĀRVIETOJAM FAILUS NO LOCAL ARCHIVE UZ SFTP
        ---------------------------------------- #>
        #pārbaudām vai ir ielādējušās WinSCP .net bibliotēka
        if ( -not $isWinSCPnetLoaded) { 
            Write-msg -log -bug -text "[sftp] WinSCPnet library is not loaded. Skip file upload to sftp."
            $item.ToFTP = $false 
        }#endif

        #Transportējam arhīva failus uz sftp
        if ( ($item.ToFTP -eq $true) -or ( $item.ToFTP -like "true" ) ) {
            $tempString = $item.ArchiveRootFolder
            $item.ArchiveRootFolder = $tempString.Replace("//","\")
            $LocalFolderPath = "$($item.ArchiveRootFolder)\$($item.SourceComputer)\$($item.Application)"
            $RemoteFolderPath = "$(($env:computername).ToLower())"

            $FtpSuccess = Move-FilesToFtp $LocalFolderPath $RemoteFolderPath

            if ( $FtpSuccess -eq 0 ) {
                Write-msg -log -text "[Main] Archive files succesfully moved to ftp"
            }#endif
            else {
                Write-msg -log -bug -text "[sftp] Something went wrong with ftp transfer. Please investigate!"
            }#endelse

        }#endif

    }#endForEach

}#endProcess

end {

    Set-Housekeeping $LogFileDir
    Stop-Watch
    Send-Mail

}#endOfend
