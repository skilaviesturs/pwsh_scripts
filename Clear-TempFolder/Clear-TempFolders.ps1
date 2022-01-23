<#Clear-TempFolders.ps1
.SYNOPSIS
Skripts pārbauda DirectoriesToCheck norādīto mapju saturu LastWriteTime datumu.
Ja tiek pārsniegts noteikto dienu periods, owner tiek nosūtīts paziņojums
Ja tiek pārsniegts noteikto dienu periods, mape/datne tiek pārvietota

.DESCRIPTION
Skripts pārbauda DirectoriesToCheck norādīto mapju saturu LastWriteTime datumu.
Ja tiek pārsniegts noteikto dienu periods, owner tiek nosūtīts paziņojums
Ja tiek pārsniegts noteikto dienu periods, mape/datne tiek pārvietota

.PARAMETER Test
Darbinam testa režīmā - bez faktiskas datņu dzēšanas

.PARAMETER ShowConfig
Parāda konfigurācijas parametrus

.PARAMETER Help
Parāda versiju un komandu sintaksti

.NOTES
	Author:	Viesturs Skila
	Version: 2.1.7
#>
[CmdletBinding()] 
param (
    [switch]$Test,
    [switch]$ShowConfig,
    [switch]$Help
)
begin {

    $Script:DirectoriesToCheck = @(
        #@{SourcePath = "D:\_source"; DestinationPath = "D:\_dest" }
        @{SourcePath = "F:\BANK\ISP\TMP"; DestinationPath = "F:\From_TMP\ISP_TMP" }
        @{SourcePath = "F:\BANK\ISP\TMP2"; DestinationPath = "F:\From_TMP\ISP_TMP2" }
        @{SourcePath = "F:\BANK\TMP"; DestinationPath = "F:\From_TMP\I_TMP" }
    )
    #Diena, kad čekojam un sūtam e-pastus
    [int]$DaysToCheck = 10
    #Dienas, kad pārvietojam datnes pēc e-pasta nosūtīšanas
    [int]$DaysToMove = 4
    #mail integrācija
    $DefaultName = "Administrator"
    $ReportTo = @('Viesturs.Skila@expobank.eu', 'Andrejs.Stankevics@expobank.eu')
    $SmtpServer = 'mail.ltb.lan'
    
    <#-------------------------------------
    ZEM ŠĪS KOMENTĀRA NEKO NEMAINĪT !!!
    -------------------------------------#>
    #Skripta tehniskie mainīgie
    $CurVersion = "2.1.7"
    $ScriptWatch = [System.Diagnostics.Stopwatch]::startNew()
    #skripta update vajadzībām
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    #logošanai
    $LogFileDir = "$__ScriptPath\log"
    $LogFile = "$((Get-ChildItem $__ScriptName).BaseName)-$(Get-Date -Format "yyyyMMdd")"
    #$TimeSetFile = "$__ScriptPath\data.dat"
    $ScriptRandomID = Get-Random -Minimum 100000 -Maximum 999999
    $ScriptUser = Invoke-Command -ScriptBlock { whoami }
    $Script:forMailReport = @()
    $DataArchiveFile = "$__ScriptPath\data.xml"
    [int]$DDay = $DaysToCheck + $DaysToMove

    #Parādam ekrānā versiju un beidzam darbu
    if ($Help) {
        Write-Host "`nVersion:[$CurVersion]`n"
        #$text = (Get-Command "$__ScriptPath\$__ScriptName" ).ParameterSets | Select-Object -Property @{n = 'Parameters'; e = { $_.ToString() } }
        $text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
        $text | ForEach-Object { Write-Host $($_) }
        Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
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
            $string_err = $_ | Out-String
            Write-Warning "$string_err"
        }#endtry
    }#endOffunction
    
    <# ----------------
	Formējam un sūtam reporta e-pastus
	#>
    Function Send-Report {
        foreach ( $a in $Script:forMailReport) {
            if ( $a.Contains("[ERROR]") ) { $logError = $true }
        }
        $server = $env:computername.ToLower()
        $mailParam = @{
            SmtpServer = $SmtpServer
            To         = $ReportTo
            From       = "no-reply@$server.ltb.lan"
            Subject    = "[$($env:computername)]$(if ($logError) {":[ERROR]"} else {":[SUCCESS]"}) report from [$__ScriptName]"
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
        <h2>$(Get-Date -Format "yyyy.MM.dd HH:mm") events from log file [$LogFile.log]</h2>
        <table><tbody>
            $( ForEach ( $line in $Script:forMailReport ) { 
                if ($line.Contains("[ERROR]") -or $line.Contains(" error") ) {
                    "<tr><td><p style=font-size:135%;color:red;>$line</p></td></tr>"
                } elseif ($line.Contains("SUCCESS:") -or $line.Contains(" success") ) {
                    "<tr><td><p style=font-size:135%;color:green;>$line</p></td></tr>"
                } else {
                    "<tr><td>$line</td></tr>"
                } 
            } )
        </tbody></table>
        <br><br>
        <p>Powered by Powershell</p>
        <p style="font-size: 60%;color:gray;">[$__ScriptName] version $CurVersion</p>
        </body>
        </html>
"@
        try {
            Send-MailMessage @mailParam -Body $mailReportBody -BodyAsHtml -Encoding utf8 -ErrorAction Stop
        }#endtry 
        catch {
            Write-msg -log -bug -text "[smtpErr] Mail failed to send with error: $($_.Exception.Message)."
            $string_err = $_ | Out-String
            Write-msg -log -bug -text "$string_err"
        }#endcatch

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
        if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)" } else { $bMin = "$($Timer.Elapsed.Minutes)" }
        if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)" } else { $bSec = "$($Timer.Elapsed.Seconds)" }
        Write-msg -log -text "[$Name] finished in $(
            if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
            elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
            else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
            )"
    }#endOffunction

    Function Set-Housekeeping {
        [cmdletbinding(DefaultParameterSetName = "default")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$CheckPath
        )
        #šeit iestatam datņu/mapju dzēšanas periodu
        #tagad iestatīts, ka tiek saarhivēts viss, kas vecāks par 30 dienām
        #$maxAge = ([datetime]::Today).addDays(-30)
        $maxAge = ([datetime]::Today).AddDays(1 - ([datetime]::Today).Day).addMonths(-1)

        $filesByMonth = Get-ChildItem -Path $CheckPath | 
        Where-Object -Property CreationTime -lt $maxAge |
        Group-Object { $_.CreationTime.ToString("yyyy\\MM") }
        if ( $filesByMonth.count -gt 0 ) {
            try {
                $filesByMonth.Group | Remove-Item -Recurse -Force -ErrorAction Stop
                Write-msg -log -text "[Housekeeping] Succesfully cleaned files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
            }#endtry
            catch {
                Write-msg -log -bug -text "[Housekeeping] Error: $($_.Exception.Message)"
                $string_err = $_ | Out-String
                Write-msg -log -bug -text "$string_err"
            }#endcatch
        }#endif
        else {
            Write-msg -log -text "[Housekeeping] There's no files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
        }#endelse
    }#endOfFunctions

    <# ----------------------------------------------
    DEFINĒJAM DARBA FUNKCIJAS
    ---------------------------------------------- #>
    
    <# ----------------
	Pārvietojam mapes, kurām iestājies termiņš
	#>
    function Clear-TempFolder {
        [CmdletBinding()] 
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$SourcePath,
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$DestinationPath
        )
        $ClearTimer = [System.Diagnostics.Stopwatch]::startNew()
        $maxAge = $([datetime]::Today).AddDays(-$DDay)
        $destPath = Join-Path -Path $DestinationPath -ChildPath (Get-Date -Format "yyyyMMdd")
        Write-msg -log -text "[Mover] SourcePath:[$SourcePath]; DestinationPath:[$destPath]"
        Write-msg -log -text "[Mover] checking files' TTL older than:[$($maxAge.toString("dd.MM.yyyy HH:mm:ss"))]"

        $Script:DataArchive |
        Where-Object -Property SourcePath -like $SourcePath |
        Where-Object -Property ttl -lt $maxAge |
        #Where-Object -Property LastWriteTime -lt $maxAge |
        ForEach-Object {
            [string]$__Path = "$($_.Path)"
            if ( ( Test-Path $SourcePath ) ) {
                try {
                    if ( -NOT ( Test-Path -Path $destPath ) ) { $null = New-Item $destPath -ItemType Directory -ErrorAction Stop }
                }#endtry
                catch {
                    Write-msg -log -bug -text "[Mover] FATAL. Unable to create destination directory [$destPath]. Please check permissions on destination folder."
                    $string_err = $_ | Out-String
                    Write-msg -log -bug -text "$string_err"
                    return
                }#endcatch
            }#endif

            #pārbaudam vai mape eksistē. ja neeksistē, ieliekam pazīmi, ka dzēsta
            if ( -NOT ( Test-Path -Path $_.Path ) ) { $_.PSIsDeleted = $True; return }
            if ( $_.PSIsContainer ) {
                try {
                    #atrodam visas direktorijā atvērtās datnes un tās aizveram
                    Get-SmbOpenFile | Where-Object -Property Path -like "$($_.Path)*" | Close-SmbOpenFile -Force -ErrorAction Stop
                    #kopējam avota mapi ar visu saturu uz mērķa direktoriju
                    Copy-Item -Path "$($_.Path)" -Destination "$destPath" -Recurse -ErrorAction Stop
                    Write-msg -log -text "[Mover] success copy directory:[$__Path] to [$destPath]."
                }#endtry
                catch {
                    Write-msg -bug -log -text "[Mover] failed copy directory:[$__Path] to [$destPath]."
                    $string_err = $_ | Out-String
                    Write-msg -log -bug -text "$string_err"
                    return
                }#endcatch
                #dzēšam avota mapi ar visu saturu
                try {
                    if ( $Test ) {
                        Remove-Item -Path "$($_.Path)" -Recurse -Force -WhatIf -ErrorAction Stop
                    }#endif
                    else {
                        Remove-Item -Path "$($_.Path)" -Recurse -Force -ErrorAction Stop
                        Write-msg -log -text "[Mover] success delete directory:[$__Path] ."
                        $_.PSIsDeleted = $True
                    }#endelse
                }
                catch {
                    Write-msg -bug -log -text "[Mover] failed delete directory:[$__Path]."
                    $string_err = $_ | Out-String
                    Write-msg -log -bug -text "$string_err"
                    return
                }
            }#endif
            else {
                try {
                    #ja fails ir atvērts, tad to aizveram
                    Get-SmbOpenFile | Where-Object -Property Path -like "$($_.Path)" | Close-SmbOpenFile -Force -ErrorAction Stop
                    #kopējam avota datni uz mērķa direktoriju
                    Copy-Item -Path "$($_.Path)" -Destination "$destPath" -Recurse -ErrorAction Stop
                    Write-msg -log -text "[Mover] success copy file :[$__Path] to [$destPath]."
                }#endtry
                catch {
                    Write-msg -bug -log -text "[Mover] failed copy file:[$__Path] to [$destPath]."
                    $string_err = $_ | Out-String
                    Write-msg -log -bug -text "$string_err"
                    return
                }#endcatch
                #dzēšam avota datni ar visu saturu
                try {
                    if ( $Test ) {
                        Remove-Item -Path "$($_.Path)" -Force -WhatIf -ErrorAction Stop
                    }#endif
                    else {
                        Remove-Item -Path "$($_.Path)" -Force -ErrorAction Stop
                        Write-msg -log -text "[Mover] success delete file:[$__Path]."
                        $_.PSIsDeleted = $True
                    }#endelse
                }
                catch {
                    Write-msg -bug -log -text "[Mover] failed delete file:[$__Path]."
                    $string_err = $_ | Out-String
                    Write-msg -log -bug -text "$string_err"
                    return
                }
            }#endelse
        }#endForeach

        Stop-Watch -Timer $ClearTimer -Name ClearTemp

    }#endOffunction

    <# ----------------
	Formējam un sūtam brīdinājuma e-pastus lietotājiem
	#>
    Function Send-Mail {
        [cmdletbinding(DefaultParameterSetName = "default")]
        Param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string[]]$FileList,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$MailTo,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )

        $ModifiedList = $FileList | Sort-Object

        #Sakombinējam e-pasta nosūtīšanai nepieciešamo
        $server = $env:computername.ToLower()
        $mailParam = @{
            SmtpServer = $SmtpServer
            To         = if ($Test) { $ReportTo } elseif ( $MailTo -like $DefaultName ) { $ReportTo } else { $MailTo }
            From       = "no-reply@$server.ltb.lan"
            Subject    = "$(if ($Test) {"[Test mode] "})Brīdinājums par koplietosanas servera pagaidu mapes saturu!"
        }
        $mailReportBody = @"
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        h2 {
          font-family: tahoma;
          font-size: 80%;
        }
        p {
          font-family: tahoma;
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
          font-size: 85%;
        }
        
        tr:nth-child(even) {
          background-color: #dddddd;
        }
        </style>
        </head>
        <body>
        <h2>Labdien, $Name!</h2>
        <p>Koplietošanas servera I: diska pagaidu mape TMP ir paredzēta īslaicīgai (līdz $($DDay) dienām) datņu apmaiņai ar kolēģiem bez jebkādiem piekļuves ierobežojumiem – datnes var atvērt ikviens bankas lietotājs.</p>
        <br>
        <h2>Jūs esat īpašnieks sekojošām mapēm un/vai datnēm, kas tiks dzēstas pēc $DaysToMove dienām no e-pasta nosūtīšanas brīža:</h2>
        <table><tbody>
            $( ForEach ( $line in $ModifiedList ) { 
                    "<tr><td>$line</td></tr>"
            } )
        </tbody></table>
        <br>
        <p>Ja ir nepieciešamība datnes koplietot ar kolēģiem ilgākā laika periodā, tad lūdzu piesakiet Palīdzības dienestā (HelpDesk) jaunas koplietošanas mapes izveidošanu I:\DARBS vai citā vietnē, kurā ir iespējama piekļuves kontroles noteikšana.</p>
        <br>
        <p>ISP komanda</p>
        <br>
        <p style="font-size: 65%;color:blue;">Powered by Powershell</p>
        <p style="font-size: 60%;color:gray;">[$__ScriptName] version $CurVersion</p>
        <br>
        <p style="font-size: 60%;color:gray;">E-pasts nosūtīts adresātam [$(if ( $MailTo -like $DefaultName ) { $ReportTo } else { $MailTo })]</p>
        </body>
        </html>
"@
        try {
            Send-MailMessage @mailParam -Body $mailReportBody -BodyAsHtml -Encoding utf8 -ErrorAction Stop
            Write-msg -log -text "[SendMail] successfully sent the mail to [$(if ($Test) { $ReportTo } elseif ( $MailTo -like $DefaultName ) { $ReportTo } else { $MailTo })]."
        }#endtry 
        catch {
            Write-msg -log -bug -text "[smtpErr] Mail failed to send with error: $($_.Exception.Message)."
            $string_err = $_ | Out-String
            Write-msg -log -bug -text "$string_err"
        }#endcatch

    }#endOffunction

    <# ----------------
	Pārbaudam arhīva object jauniem ierakstiem, sagatavojam datus e-pastiem
	#>
    Function Set-DataForMail {
        $DataForMailTimer = [System.Diagnostics.Stopwatch]::startNew()
        $forMailToOwner = @()
        $maxAge = $([datetime]::Today).AddDays(-$DaysToCheck)

        #Izveidojam adrešu grāmatiņu ar vārdu un e-pasta adresi
        $Recipients = @()
        $Script:DataArchive.GetEnumerator() |
        ForEach-Object {
            if ( $Recipients.count -eq 0 ) {
                $Recipients += @{FullName = $_.FullName; mail = $_.Mail }
            }#endif
            elseif ( -NOT $Recipients.mail.Contains($_.Mail) ) {
                $Recipients += @{FullName = $_.FullName; mail = $_.Mail }
            }#endif
        }#endforeach

        #Skenējam objektu pēc adresāta, izveidojam datņu sarakstu un padodam uz e-pasta funkciju 
        $i = 0
        ForEach ( $item in $Recipients ) {

            $forMailToOwner = @()
            $Script:DataArchive.GetEnumerator() | 
            Where-Object -Property ttl -le $maxAge | 
            Where-Object -Property MailIsSent -eq $False | 
            Where-Object -Property Mail -like $item.Mail | 
            ForEach-Object {
                $i++
                Write-msg -log -text "[DataForMail] [$($_.type)]:LastWriteTime[$($_.LastWriteTime.toString("dd.MM.yyyy HH:mm:ss"))]:[$($_.Path)]"
                $forMailToOwner += "$($_.type) - $($_.Path.Replace("F:\BANK","I:"))"
                $_.MailIsSent = $True
            }#endforeach

            if ( $forMailToOwner.count -gt 0 ) {
                Write-msg -log -text "[DataForMail] going to send a mail to [$($item.FullName)] [$($item.Mail)]"
                Send-Mail -FileList $forMailToOwner -Name $item.FullName -MailTo $item.Mail
            }#endfi
        }#endforeach
        Write-msg -log -text "[DataForMail] [$i] items proceeded."
        Stop-Watch -Timer $DataForMailTimer -Name DataForMail
    }#endOffunction

    <# ----------------
	Pārbaudam norādītās direktorijas, papildinām arhīva object
	#>
    function Get-ArchiveData {
        [CmdletBinding()] 
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$SourcePath
        )
        $GetArchiveDataTimer = [System.Diagnostics.Stopwatch]::startNew()
        $oldArchiveRec = $Script:DataArchive.count
        $Object = @()

        #Skenējam mapi - tikai pirmo līmeni
        $FileObject = Get-ChildItem $SourcePath -File | Select-Object Fullname, LastWriteTime, Attributes, PSIsContainer, @{Name = "Owner"; Expression = { (Get-Acl $_.FullName).Owner } }
        $DirObject = Get-ChildItem $SourcePath -Directory | Select-Object Fullname, LastWriteTime, Attributes, PSIsContainer, @{Name = "Owner"; Expression = { (Get-Acl $_.FullName).Owner } }
        $Object += $FileObject
        $Object += $DirObject

        #$Object
        if ( $Object.Count -gt 0 ) {
            #Pieliekam ojektam papildus laukus
            $Object | Add-Member -NotePropertyName 'Mail' -NotePropertyValue ''
            $Object | Add-Member -NotePropertyName 'Name' -NotePropertyValue ''
            $Object | Add-Member -NotePropertyName 'MailName' -NotePropertyValue ''

            #Atrodam katram ownerim vārdu un e-pastu, pievienojam info objektā
            Write-msg -log -text "[GetArchiveData] Found [$($Object.Count)] objects to proceed."
            $Object | 
            ForEach-Object {
                $name = ($_.Owner).split('\')[1]
                if ( [String]::IsNullOrWhiteSpace($name) -or [String]::IsNullOrEmpty($name) ) {
                    $_.Mail = $DefaultName
                    $_.Name = $DefaultName
                    $_.MailName = $DefaultName
                }#endif
                else {
                    if ( [bool] (Get-ADUser -Filter { SamAccountName -eq $name }) ) {
                        $mail = Get-ADUser $name -Properties * -ErrorAction Stop | select-object Mail, Name
                        $_.Mail = $mail.Mail
                        $_.MailName = $mail.Name
                        $_.Name = $name
                    }#endif
                }#endif
            }#endforeach
    
            #Veicam tukšo lauku aizpildīšanu ar noklusēto vērtību
            $Object | Where-Object { [String]::IsNullOrWhiteSpace($_.Name) -or [String]::IsNullOrEmpty(($_.Name)) } | ForEach-Object { $_.Name = $DefaultName }
            $Object | Where-Object { [String]::IsNullOrWhiteSpace($_.Mail) -or [String]::IsNullOrEmpty(($_.Mail)) } | ForEach-Object { $_.Mail = $DefaultName }
            $Object | Where-Object { [String]::IsNullOrWhiteSpace($_.MailName) -or [String]::IsNullOrEmpty(($_.MailName)) } | ForEach-Object { $_.MailName = $DefaultName }

            #$Object | fl *
            $NewFiles = 0
            $NewDirs = 0
            #Ielasam datus arhīvā
            $Object | 
            ForEach-Object {
                if ( $Script:DataArchive.count -eq 0) {
                    $Script:DataArchive += @{ 
                        Path          = $_.FullName;
                        SourcePath    = $SourcePath;
                        PSIsContainer = if ($_.PSIsContainer) { $True } else { $False };
                        Type          = if ($_.PSIsContainer) { "Mape" } else { "Datne" };
                        LastWriteTime = $_.LastWriteTime;
                        Name          = $_.Name;
                        FullName      = $_.MailName
                        Mail          = $_.Mail;
                        MailIsSent    = $False;
                        ttl           = [datetime]::Today;
                        PSIsDeleted   = $False
                    }
                    if ($_.PSIsContainer) { $NewDirs++ } else { $NewFiles++ }
                }#endif
                elseif ( -NOT $Script:DataArchive.Path.Contains($_.FullName) ) {
                    $Script:DataArchive += @{ 
                        Path          = $_.FullName;
                        SourcePath    = $SourcePath;
                        PSIsContainer = if ($_.PSIsContainer) { $True } else { $False };
                        Type          = if ($_.PSIsContainer) { "Mape" } else { "Datne" };
                        LastWriteTime = $_.LastWriteTime;
                        Name          = $_.Name;
                        FullName      = $_.MailName
                        Mail          = $_.Mail;
                        MailIsSent    = $False;
                        ttl           = [datetime]::Today;
                        PSIsDeleted   = $False
                    }
                    if ($_.PSIsContainer) { $NewDirs++ } else { $NewFiles++ }
                }#endelseif
            }#endforeach

            Write-msg -log -text "[GetArchiveData] DataArchive has [$oldArchiveRec] objects; From [$SourcePath] got [$($Object.count)] objects; Added new [$($NewFiles+$NewDirs)] objects to DataArchive$(if ( ($NewFiles+$NewDirs) -gt 0 ) {": [$NewFiles] files, [$NewDirs] directories"})."
        }#endif
        else {
            Write-msg -log -text "[GetArchiveData] Nothing found. All clear."
        }#endelse

        Stop-Watch -Timer $GetArchiveDataTimer GetArchiveData

    }#endOffunction

    Function Clear-ArchiveData {
        #iztīram Datu masīvu
        $ToDelete = @()

        $Script:DataArchive | ForEach-Object {
            if ( -NOT ( Test-Path -Path $_.Path ) ) { $_.PSIsDeleted = $True }
        }#endforeach

        $Script:DataArchive | Where-Object -Property PSIsDeleted -eq $true `
        | ForEach-Object {
            $ToDelete += $_.Path
        }#endforeach
        $i = 0
        $ToDelete | ForEach-Object {
            if ( $Script:DataArchive.Path.Contains($_) ) {
                $i++
                $Script:DataArchive = $Script:DataArchive | Where-Object -Property Path -ne $_
            }#endif
        }#endforeach
        Write-msg -log -text "[SelfCheck] [$i]of[$($ToDelete.count)] objects removed from DataArchive."
    }#endOffunction

    <# ----------------------------------------------
    Funkciju definēšanu beidzām
    IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
    ---------------------------------------------- #>
    #Clear-Host
    Write-msg -log -text "[-----] Script started in [$(if ($Test) {"Test"}
		elseif ($ShowConfig) {"ShowConfig"} 
		else {"Default"})] mode. Used config file [$jsonFileCfg]"

    #Ja norādīts parametrs, tad parādam ielasītos datus no JSON un beidzam darbu
    if ( $ShowConfig ) {
        Write-Host "Configuration param:"
        Write-Host "-----------------------------------------------------"
        Write-Host "CurVersion`t`t:[$CurVersion]"
        Write-Host "__ScriptName`t`t:[$__ScriptName]"
        Write-Host "__ScriptPath`t`t:[$__ScriptPath]"
        Write-Host "LogFileDir`t`t:[$LogFileDir]"
        Write-Host "LogFile`t`t`t:[$LogFile]"
        Write-Host "DefaultName`t`t:[$DefaultName]"
        Write-Host "ReportTo`t`t:[$ReportTo]"
        Write-Host "SmtpServer`t`t:[$SmtpServer]"
        Write-Host "DaysToCheck`t`t:[$DaysToCheck]`nDaysToMove`t`t:[$DaysToMove]`nDDay`t`t`t:[$Dday]"
        Write-Host "DDay(MaxAge)`t`t:[$($([datetime]::Today).AddDays(-$DDay))]"
        Write-Host "-----------------------------------------------------"
        ForEach ( $line in $Script:DirectoriesToCheck ) {
            Write-Host "SourcePath:[$($line.SourcePath)] -> DestinationPath:[$($line.DestinationPath)]"
        }#endForEach
        Stop-Watch -Timer $ScriptWatch -Name Script
        Exit 0
    }#endif

    $Script:DirectoriesToCheck.GetEnumerator() | 
    ForEach-Object {
        if ( -NOT (Test-Path -Path $_.SourcePath) ) {
            Write-msg -log -bug -text "Path [$($_.SourcePath)] not found. Exit"
            exit 1
        }#endif
        if ( -NOT (Test-Path -Path $_.DestinationPath) ) {
            Write-msg -log -bug -text "Path [$($_.DestinationPath)] not found. Exit"
            exit 1
        }#endif
    }#endforeach

}#endOfbegin

process {
    <#--------------------------------------------
    Daram darbiņu
    --------------------------------------------#>
    if ($Script:DirectoriesToCheck.count -gt 0) {
        Write-msg -log -text "[Main] Got [$($Script:DirectoriesToCheck.count)] source directories to check."

        #Nolasām arhīva datus
        if ( (Test-Path -Path $DataArchiveFile -Type Leaf) ) {
            #$Script:DataArchive = Get-Content -Path $jsonFile -Raw | ConvertFrom-Json
            $Script:DataArchive = @(Import-Clixml -Path $DataArchiveFile)

            #Pārbaudam arhīvā iekļauto objektu esamību
            Write-msg -log -text "[SelfCheck] Testing objects from DataArchive..."
            $Script:DataArchive | 
            ForEach-Object {
                if ( $_.PSIsContainer ) {
                    if ( -NOT ( Test-Path -Path "$($_.Path)" -Type Container ) ) {
                        Write-msg -log -text "[SelfCheck] not found [$($_.Path)] - remove object."
                        $_.PSIsDeleted = $True
                    }#endif
                }#endif
                else {
                    if ( -NOT ( Test-Path -Path "$($_.Path)" -Type Leaf ) ) {
                        Write-msg -log -text "[SelfCheck] not found [$($_.Path)] - remove object."
                        $_.PSIsDeleted = $True
                    }#endif
                }#endelse
            }#endforeach

            #Iztīram DataArchive no dzēstajiem ojektiem
            Clear-ArchiveData

        }#endif
        else {
            #Definējam jaunu arhīva objectu
            $Script:DataArchive = @()
        }#endelse
        
        Write-msg -log -text "[Main] Imported [$($Script:DataArchive.Count)] objects from DataArchive."

        #Skenējam norādītās direktorijas, papildinām arhīva objektu
        $Script:DirectoriesToCheck.GetEnumerator() | 
        ForEach-Object {
            Get-ArchiveData -SourcePath $_.SourcePath
        }#endforeach
        
        #Formējam datus adresātiem un nosūtam e-pastus
        Set-DataForMail
        
        #Pārkopējam datus no avota direktorijas, dzēšam
        $Script:DirectoriesToCheck.GetEnumerator() | 
        ForEach-Object {
            Clear-TempFolder -SourcePath $_.SourcePath -DestinationPath $_.DestinationPath
        }#endforeach

        # tīram mērķa direktoriju saturu, kas vecāks par iepriekšējā mēneša 1.datumu
        $Script:DirectoriesToCheck.GetEnumerator() |
        ForEach-Object {
            Set-Housekeeping -CheckPath $_.DestinationPath
        }#endforeach
        
        #Iztīram DataArchive no dzēstajiem ojektiem
        Clear-ArchiveData

        #$Script:DataArchive | ConvertTo-Json | Out-File $jsonFile -Force
        $Script:DataArchive | Export-Clixml -Path $DataArchiveFile -Depth 10 -Force
        Write-msg -log -text "[Main] Exported [$($Script:DataArchive.Count)] objects to DataArchive."

    }#endif
    else {
        Write-msg -log -text "[Main] None source directories to check."
    }#endelse

}#endOfprocess

end {

    #Dzēšam log mapes saturu, kas vecāks par 30 dienām
    Set-Housekeeping -CheckPath $LogFileDir

    Stop-Watch -Timer $ScriptWatch -Name Script

    <#iespējojam, ja gribam sūtīt e-pastus tikai kļūdu gadījumā
    #----------------------------------------------------------
    #Send-Report
    #  #>

    #iespējojam, ja gribam sūtīt e-pastus tikai kļūdu gadījumā
    #----------------------------------------------------------
    $Script:forMailReport | 
    ForEach-Object {
        if ( $_.Contains("[ERROR]") ) { Send-Report; break }
    }#endforeach
    #  #>

}#endOfend