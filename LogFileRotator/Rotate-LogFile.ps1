<#Rotate-LogFile.ps1
.SYNOPSIS
The script archive and clear content of the log file if it exceeds threshold size

.DESCRIPTION
Usage:
Rotate-LogFile.ps1 [-Test] [-ShowConfig] [-V] [<CommonParameters>]

Parameters:
    [-Test<SwitchParameter>] - run script without delete content of the source file;
	[-ShowConfig<SwitchParameter>] - show configuration parameters from json file;
	[-V<SwitchParameter>] - print out script's version;
	[-Help<SwitchParameter>] - print out this help.

.EXAMPLE


.NOTES
	Author:	Viesturs Skila
	Version: 1.1.7
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
    #Skripta tehniskie mainīgie
	Set-Culture -CultureInfo lv-LV
    $CurVersion = "1.1.7"
    $StopWatch = [System.Diagnostics.Stopwatch]::startNew()
    #skripta update vajadzībām
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    $UpdateDir = "\\beluga\install\scripts\RotateLogFile"
    #Skritpa konfigurācijas datne
    $jsonFileCfg = "$__ScriptPath\RLFconfig.json"
    #logošanai
    $LogFileDir = "$__ScriptPath\log"
    $LogFile = "RotateLogFile-$(Get-Date -Format "yyyyMMdd")"
    $ScriptRandomID = Get-Random -Minimum 100000 -Maximum 999999
    $ScriptUser = Invoke-Command -ScriptBlock { whoami }
    $Script:forMailReport = @()
    #Servisu restartēšanas vajadzībām
    $RestartFileDir = "$__ScriptPath\lib"
    $RestartFile = "$RestartFileDir\restart.dat"
    $StartHour = 0
    $EndHour = 1
    #mail integrācija
    $SmtpServer = 'mail.ltb.lan'
    $mailTo = 'viesturs.skila@expobank.eu'
    #Konfigurācijas JSON datnes parauga objekts
    #Cfg
    $CfgTmpl = [PSCustomObject]@{
        'SourceFolder'   = $null
        'SourceFileType' = $null
        'MaxSize'        = $null
        'ServiceName'    = $null
        'RestartService' = $null
    }#endOfobject
    #Services
    $SrvTmpl = [PSCustomObject]@{
        'ServiceName' = $null
        'RestartDate' = $null
    }
    $Cfg = [PSCustomObject]@{}
    $CfgServices = [PSCustomObject]@{}
    $RestartServiceNow = @()
    $Script:ServicesToRestart = @()
    #Parādam ekrānā versiju un beidzam darbu
    if ( $V ) {
        Write-Host "`n$CurVersion`n"
        Exit
    }#endif
    Function Show-ScriptHelp {
        Write-Host "`nUsage:`nSet-FileArhive.ps1 [-Test] [-ShowConfig] [-V] [<CommonParameters>"
        Write-Host "`nDescription:`n  The script archive and clear content of the log file if it exceeds threshold size"
        Write-Host "`nParameters:"
        Write-Host "  [Test<SwitchParameter>]`t- run script without delete content of the source file;"
        Write-Host "  [ShowConfig<SwitchParameter>]`t- show configuration parameters from json file;"
        Write-Host "  [V<SwitchParameter>]`t- print out script's version;"
        Write-Host "  [Help<SwitchParameter>]`t- print out this help.`n"
    }#endOfFunction

    #parādam ekrānā Help un beidzam darbu
    if ( $help ) {
        Show-ScriptHelp
        Exit
    }#endif

    if ( -not ( Test-Path -Path $LogFileDir ) ) { $null = New-Item -ItemType "Directory" -Path $LogFileDir }

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
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
            }#endif
            elseif ( $log ) {
                Write-Verbose $flag"`t"$text
                Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                    Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
            }#endelseif
            else {
                Write-Verbose $flag"`t"$text
            }#else
        }#endtry
        catch {
            Write-Warning "[Write-msg] $($_.Exception.Message)"
        }#endtry
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
            To         = $mailTo
            From       = "no-reply@$server.ltb.lan"
            Subject    = "[$__ScriptName] report from [$($env:computername)] $(if ($logError) {":[ERROR]"})"
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
        Function's purpose: to get the data from the json file, to parse and to valdiate against the template ojects' values by name
        Syntax: gjData [object] [jsonFileName]
    #>
    Function Get-JSONData {
        [cmdletbinding(DefaultParameterSetName = "default")]
        [alias("gjData")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [object]$template,
            [Parameter(Position = 1, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$jsonFile
        )
        #Write-msg -text "[gjData] jsonFile : [$jsonFile]"
        $jsonData = [PSCustomObject] {}
        try {
            $jsonData = Get-Content -Path $jsonFile -Raw -ErrorAction STOP | ConvertFrom-Json
        } 
        catch {
            # ja JSON datne nav nolasāma vai neeksistē, paziņojam un beidzam darbu
            Write-Warning "[gjData] Fatal error - file [$jsonFile] corrupted. Exit."
            Stop-Watch
            Exit
        } #endcatch
        if ( ($jsonData.count) -gt 0 ) {
            if ($Test) { Write-msg -text "[gjData] [$jsonFile].count : $($jsonData.count)" }
            # let's compare properties of $Cfg and $jsonData 
            $aa = $template[0].psobject.properties.name | Sort-Object
            $bb = $jsonData[0].psobject.properties.name | Sort-Object
            foreach ($name in $bb) {
                # ja JSON konfig struktūra neatbilst cfg template, paziņojam un beidzam darbu
                if ( $aa -eq $name ) {
                    # Write-msg -text "`t template[$aa] = jsonData[$name)]"
                }#endif
                else {
                    Write-Warning "[gjData] Fatal error - Unknown variable name [$name] in import file [$jsonFile]. Exit."
                    Stop-Watch
                    Exit
                } #endelse
            } #endforeach
        } #endendif
        return $jsonData
    } #endOffunction
        
    <# ----------------
    Funkcija Repair-JSONfile
    izveido no parauga objekta jaunu JSON datni
    #>
    Function Repair-JSONFile {
        [cmdletbinding(DefaultParameterSetName = "default")]
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
            Write-Warning "[repJFile] $($_.Exception.Message)"
        }#endcatch
    }#endOffunction

    <# ----------------
    Funkcija Get-DependServices 
    atrod galvenā servisa darbības nodrošināšanas nepieciešamos servisus
    #>
    function Get-DependServices {
        Param (
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Object]$ServiceInput
        )

        If ($ServiceInput.DependentServices.Count -gt 0) {
            ForEach ($DepService in $ServiceInput.DependentServices) {
                Write-msg -text "[getDS] Dependent of [$($ServiceInput.Name)]: [$($Service.Name)]"
                If ($DepService.Status -eq "Running") {
                    Write-msg -text "[getDS] [$($DepService.Name)] is running."
                    $CurrentService = Get-Service -Name "$($DepService.Name)"
                    # atrodam saistīto servisu nosaukums
                    Get-DependServices $CurrentService                
                }#endif
                else {
                    Write-msg -text "[getDeepend] [$($DepService.Name)] is stopped. No Need to stop or start or check dependancies."
                }#endelse
            }#endforeach
        }#endif
        else {
            if ($Test) {
                Write-msg -text "[getDeepend] DependentService count [$($ServiceInput.DependentServices.Count)]"
            }#endif
        }#endelse

        if ($Script:ServicesToRestart.Contains("$($ServiceInput.Name)") -eq $false) {
            Write-msg -text "[getDeepend] Adding service [$($ServiceInput.Name)] to restart list"
            $Script:ServicesToRestart += "$($ServiceInput.Name)"
        }#endif
    }#endOffunction

    <# ----------------
    Funkcija Restart-Service 
    restartē norādīto servisu un tā pakārtotos servisus, ja tādus atrod
    #>
    Function Restart-Service {
        Param (
            [Parameter(Position = 0, Mandatory = $true) ]
            [ValidateNotNullOrEmpty()]
            [String]$ServiceName,
            [Parameter(Position = 1, Mandatory = $true)]
            [int]$StartHour,
            [Parameter(Position = 2, Mandatory = $true)]
            [int]$EndHour
        )
        if ($Test) { Write-msg -log -text "[RestartService] got servicename [$ServiceName]" }
        $Script:ServicesToRestart = @()
        # Atrodam galvenā servisa objektu
        $Service = Get-Service -Name $ServiceName -ErrorAction Ignore

        if ( $Service.count -gt 0  ) {
            # Izsaucam saistīto servisu atrašanas funkciju un izveidojam to restartēšanas kārtību
            Get-DependServices $Service
            if ( $Script:ServicesToRestart.Count -gt 0 ) {
                #uzsākam servisu restartu, ja tekošais laiks ir norādītajā laika logā
                [int]$hour = Get-Date -Format "HH"
                if ( $hour -ge $StartHour -and $hour -le $EndHour ) {
                    try {
                        # Apstādinām servisus
                        ForEach ($ServiceToStop in $Script:ServicesToRestart) {
                            Write-msg -log -text "[RestartService] stop the service [$ServiceToStop]"
                            Stop-Service $ServiceToStop -ErrorAction Stop
                        }#endforeach
                        # Pārkārtojam servisu sarakstu reversā secībā, saskaņā ar kuru startēsim servisus
                        [array]::Reverse($Script:ServicesToRestart)
                        # Startējam servisus
                        ForEach ($ServiceToStart in $Script:ServicesToRestart) {
                            Write-msg -log -text "[RestartService] start the service [$ServiceToStart]"
                            Start-Service $ServiceToStart -ErrorAction Stop
                        }#endforeach

                        #Viss veiksmīgi izdevies - atgriežam $true
                        return $true
                    }#endtry
                    catch {
                        Write-msg -log -bug -text "[RestartService] $($_.Exception.Message)"
                    }#endcatch
                }#endif
                else {
                    Write-msg -log -text "[RestartService] will restart [$ServiceName] in allowed time from [$StartHour`:00] to [$EndHour`:00]"
                }#endelse
                
            }#endif
            else {
                Write-msg -log -bug -text "[RestartService] ServicesToRestart.Count:[$($Script:ServicesToRestart.Count)]. There's no services to restart."
            }#endelse

        }#endif
        else {
            Write-msg -log -bug -text "[RestartService] Service.count:[$($Service.count)]. There's no services to restart."
        }#endelse
        return $false
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
    
    Function Stop-Watch {
        $stopwatch.Stop()
        if ( $stopwatch.Elapsed.Minutes -le 9 -and $stopwatch.Elapsed.Minutes -gt 0 ) { $bMin = "0$($stopwatch.Elapsed.Minutes)" } else { $bMin = "$($stopwatch.Elapsed.Minutes)" }
        if ( $stopwatch.Elapsed.Seconds -le 9 -and $stopwatch.Elapsed.Seconds -gt 0 ) { $bSec = "0$($stopwatch.Elapsed.Seconds)" } else { $bSec = "$($stopwatch.Elapsed.Seconds)" }
        Write-msg -log -text "[Script] finished in $(
            if ( [int]$stopwatch.Elapsed.Hours -gt 0 ) {"$($stopwatch.Elapsed.Hours)`:$bMin`:$bSec hrs"}
            elseif ( [int]$stopwatch.Elapsed.Minutes -gt 0 ) {"$($stopwatch.Elapsed.Minutes)`:$bSec min"}
            else { "$($stopwatch.Elapsed.Seconds)`.$($stopwatch.Elapsed.Milliseconds) sec" }
            )"
    }#endOffunction

    Function Set-Housekeeping {
        [cmdletbinding(DefaultParameterSetName = "default")]
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
                $filesByMonth.Group | Remove-Item -Recurse -Force -ErrorAction Stop
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

    <# ----------------------------------------------
    Funkciju definēšanu beidzām
    IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
    ---------------------------------------------- #>
    Clear-Host
    Write-msg -log -text "[-----] Script started in [$(if ($Test) {"Test"}
        elseif ($ShowConfig) {"ShowConfig"} 
        else {"Default"})] mode. Used config file [$jsonFileCfg]"
    #Pārbaudam vai nav jaunākas versijas repozitorijā 
    if ( -not [String]::IsNullOrWhiteSpace($UpdateDir) -and -not [String]::IsNullOrEmpty($UpdateDir) ) {
        Get-ScriptFileUpdate $__ScriptName $__ScriptPath $UpdateDir
        if ($Update) {
            Stop-Watch
            exit 
        }
    }#endif
    # Ielasām parametrus no JSON datnes
    # ja neatrodam, tad izveidojam JSON datnes paraugu ar noklusētām vērtībām
    if ( -not ( Test-Path -Path $jsonFileCfg -PathType Leaf) ) {
        Write-msg -log -text "[Check] Config JSON file [$jsonFileCfg] not found."

        #No cfg template izveidojam objektu ar vērtībām
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SourceFolder' -Value 'C:\tmp\log' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SourceFileType' -Value 'log' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'MaxSize' -Value '25600000' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'ServiceName' -Value 'AnyDesk' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RestartService' -Value true -Force
        #Izsaucam JSON izveidošanas funkciju
        Repair-JSONFile $CfgTmpl $jsonFileCfg
        Write-msg -log -text "[Check] Config JSON file [$jsonFileCfg] created."
        #Ielasām datus no jaunizveidotā JSON
        $Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
    }#endif
    else {
        #Ielasām datus no JSON
        $Cfg = Get-JSONData $CfgTmpl $jsonFileCfg
    }#endelse

    if ( Test-Path -Path $RestartFile -PathType Leaf ) {
        #Ielasām datus no JSON
        $CfgServices = Get-jsonData $SrvTmpl $RestartFile
		#$CfgServices | ft *
    }#endif
    else {
        Write-msg -log -text "[Check] Data file [$RestartFile] not found."
    }#endelse
    
    #Ja norādīts parametrs, tad parādam ielasītos datus no JSON un beidzam darbu
    if ( $ShowConfig ) {
        Write-Verbose "Configuration param:"
        Write-Verbose "--------------------"
        Write-Verbose "CurVersion`t:[$CurVersion]"
        Write-Verbose "__ScriptName`t:[$__ScriptName]"
        Write-Verbose "__ScriptPath`t:[$__ScriptPath]"
        Write-Verbose "UpdateDir`t:[$UpdateDir]"
        Write-Verbose "jsonFileCfg`t:[$jsonFileCfg]"
        Write-Verbose "LogFileDir`t:[$LogFileDir]"
        Write-Verbose "LogFile`t:[$LogFile]"
        Write-Verbose "RestartFile`t:[$RestartFile]"
        Write-Verbose "StartHour`t:[$StartHour]"
        Write-Verbose "EndHour`t:[$EndHour]"
        Write-Verbose "--------------------"
        $Cfg | Format-List *
        $CfgServices | Sort-Object -Property ServiceName | Format-Table ServiceName, RestartDate -AutoSize
        Stop-Watch
        Exit
    }#endif

}#endbegin

process {
    #Sagatavojam restartējamo servisu sarakstu
    ForEach ( $item in $Cfg ) {
        $Script:IsNewForCfgServices = $true
        if ( ( -not [String]::IsNullOrWhiteSpace($item.ServiceName) -and -not [String]::IsNullOrEmpty($item.ServiceName) ) -and `
            ( $item.RestartService -eq $true -or $item.RestartService -like 'true' ) ) {
            if ( $null -ne $CfgServices ) {
                $CfgServices | 
                Where-Object -Property ServiceName -like $item.ServiceName | 
                ForEach-Object {
                    if ($Test) { Write-msg -log -text "[Main] add service [$($_.ServiceName)] to service restart list" }
                    $RestartServiceNow += New-Object -TypeName psobject -Property @{ServiceName = "$($_.ServiceName)"; RestartDate = $($_.RestartDate) }
                    $Script:IsNewForCfgServices = $false
                }#endforeach
                if ($Script:IsNewForCfgServices) {
                    if ($Test) { Write-msg -log -text "[Main] add NEW service [$($item.ServiceName)] to service restart list" }
                    $RestartServiceNow += New-Object -TypeName psobject -Property @{ServiceName = "$($item.ServiceName)"; RestartDate = (Get-Date).AddYears(-100) }
                }
            }#endif
            else {
                if ($Test) { Write-msg -log -text "[Main] add service [$($item.ServiceName)] to NEW service restart list" }
                $RestartServiceNow += New-Object -TypeName psobject -Property @{ServiceName = "$($item.ServiceName)"; RestartDate = (Get-Date).AddYears(-100) }
            }#endelse
        }#endif
        else {
            if ($Test) { Write-msg -log -text "[Main] skip service [$($item.ServiceName)]" }
        }#endelse
    }#endforeach

    #Pārbaudām vai šodien ir nepeiciešams restartēt servisus, ja ir - restartējam un nomainām RestartDate laiku
    if ( $RestartServiceNow.Count -gt 0 ) {
        $RestartServiceNow | Where-Object -Property RestartDate -lt (Get-Date).Date | 
        ForEach-Object { 
            if ( ( Restart-Service -ServiceName "$($_.ServiceName)" -StartHour $StartHour  -EndHour $EndHour ) ) {
				#Write-msg -log -text "[Main] change [$($_.ServiceName)] restart time to [$((Get-Date))]"
                [System.DateTime]$_.RestartDate = (Get-Date).AddHours(2)
				#Write-msg -log -text "[Main] changed to [$($_.RestartDate)]"
            }#endif
        }#endforeach
    }#endif
    else {
        if ($Test) { Write-msg -log -text "[Main]:[$($RestartServiceNow.Count)] there is no services to restart." }
    }#endif

    #Saglabājam json datnē restartēto servisu vārdus un laikus
	#$RestartServiceNow | ft * -AutoSize
    $RestartServiceNow | ConvertTo-Json | Out-File $RestartFile -Force

    #Katram JSON norādītajam ierakstam
    foreach ( $item in $Cfg ) {
        #pārbaudam, vai avota direktorija eksistē
        if ( Test-Path -Path "$($item.SourceFolder)" ) {
            #izveidojam failu sarakstu un laižam ciklā
            $fileList = Get-ChildItem "$($item.SourceFolder)" -File -Filter "*.$($item.SourceFileType)"
            foreach ( $row in $fileList ) {
                #izveidojam konkrētās datnes objektu
                $file = Get-ChildItem $row.FullName

                #ja datnes apjoms pārsniedz konfigurācijas failā norādīto, tad izpildam
                if ( $file.Length -gt $item.MaxSize ) {
                    Write-msg -log -text "[Main] Found file [$file] with size [$($file.Length)] bytes bigger then threshold [$($item.MaxSize)] bytes"
                    try {
                        #izveidojam pagaidu faila nosaukumu, piešķirot  pirmo kārtas skaitli
                        $TempFileName = "$($File.DirectoryName)\$($File.BaseName).001$($File.Extension)"

                        #pārbaudam vai arhīva datne jau eksistē
                        if ( Test-Path -Path "$TempFileName.zip" -PathType Leaf ) {
                            #ja eksistē, tad iestatam skaitīkli
                            $counter = 1
                            #ejam ciklā, kamēr faila nosaukumu NEatrodam
                            do {
                                #šeit nosacījumi skaitīkļa formātam - pieļaujam 999 rotācijas
                                if ($counter -le 9 ) { $StrCounter = "00$counter" }
                                elseif ($counter -le 99 ) { $StrCounter = "0$counter" }
                                else { $StrCounter = "$counter" }
                                $TempFileName = "$($File.DirectoryName)\$($File.BaseName).$StrCounter$($File.Extension)"
                                <#
                                if ($Test) {
                                    Write-msg -text "Counter[$counter] StrCounter[$StrCounter] File [$TempFileName.zip] - exist [$(if 
                                        ( Test-Path -Path "$TempFileName.zip" -PathType Leaf) {"True"} else {"False"} )]"
                                }#endif #>
                                $counter++

                                #čekojam, kamēr neatrodam eksistējošu arhīva datni
                            } while ( Test-Path -Path "$TempFileName.zip" -PathType Leaf )
                        }#endif
            
                        #nokopējam datnes saturu uz pagaidu datni
                        Copy-Item -Path $file.FullName -Destination $TempFileName -ErrorAction Stop
                        Write-msg -log -text "[Main] [$($File.FullName)] copied to the temporary file [$TempFileName]"
                        
                        #iztukšojam mērķa datnes saturu, ja nav aktīvs testa režīms
                        if (-not $Test) {
                            Clear-Content $file.FullName -Force -ErrorAction Stop
                            Write-msg -log -text "[Main] [$($File.FullName)] - cleared content"
                        }#endif
                        #arhivējam pagaidu datni
                        Compress-Archive -Path $TempFileName -DestinationPath "$TempFileName.zip" -ErrorAction Stop
                        Write-msg -log -text "[Main] Created archive file [$TempFileName.zip]"
    
                        #izdzēšam pagaidu datni
                        Remove-Item $TempFileName -ErrorAction Stop
                        Write-msg -log -text "[Main] [$TempFileName] temporary file removed."
    
                    }#endtry
                    catch {
                        Write-msg -log -bug -text "[Main] $($_.Exception.Message)"
                    }#endcatch
                }#endif
                else {
                    <#testu nolūkiem, lai pārliecinātos par skripta darbību
                    if ($Test) {
                        Write-msg -text "[$file] size [$FileSize] bytes; threshold [$($item.MaxSize)] bytes"
                    }#endif
                    #  #>
                }#endelse
            }#endforeach
        }#endif
        else {
            Write-msg -log -bug -text "[Main] SourceFolder [$($item.SourceFolder)] not found!"
        }#endelse
    }#endforeach
}#endOfprocess

end {

    Set-Housekeeping $LogFileDir
    Stop-Watch
    foreach ( $a in $Script:forMailReport) {
        if ( $a.Contains("[ERROR]") ) { Send-Mail; break }
    }

}#endofend
