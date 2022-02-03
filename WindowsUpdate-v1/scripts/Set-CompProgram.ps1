<#
.SYNOPSIS
Attālināta programmatūras uzstādīšana un noņemšana

.DESCRIPTION
Skripts nodrošina uz attālinātā datora:
[*] msi vai exe pakotnes uzkopēšanu 
[*] msi vai exe pakotnes uzstādīšanu
[*] programmas noņemšanu

.PARAMETER Name
Obligāts lauks.
Norādam datora vārdu.

.PARAMETER ComputerName
Norādam attālinātā datora NETBIOS vai DNS vārdu

.PARAMETER InstallPath
Norādam uzstādāmās programmatūras MSI vai EXE pakotnes atrašanās vietu.
Lietotājam, ar kuru veicam skripta darbināšanu, jābūt pilnām tiesībām uz pakotni.

.PARAMETER UninstallIdNumber
Norādam uzstādāmās programmatūras unikālo identifikatoru.
Programmatūras identifikatoru varam iegūt ar skriptu Get-CompSoftware.ps1

.PARAMETER DisplayName
Norādam uzstādāmās programmatūras DisplayName.
Programmatūras identifikatoru varam iegūt ar skriptu Get-CompSoftware.ps1

.PARAMETER Help
Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.

.EXAMPLE
Set-Program.ps1 EX00001 -InstallPath 'D:\install\7-zip\7z1900-x64.msi'
Uzstādam uz datora EX00001 programmatūras instalācijas pakotni 7z1900-x64.msi

.EXAMPLE
Set-Program.ps1 EX00001 -UninstallIdNumber '{23170F69-40C1-2702-1900-000001000000}'
Noņemam programmatūras instalāciju, kuras identifikācijas numurs ir {23170F69-40C1-2702-1900-000001000000}

.EXAMPLE
Set-Program.ps1 EX00001 -InstallPath 'D:\install\7-zip\7z1900-x64.msi' -UninstallIdNumber '{23170F69-40C1-2702-1900-000001000000}'
Noņemam programmatūras instalāciju un uzstādam jauno pakotni.

.NOTES
	Author:	Viesturs Skila
	Version: 1.2.3
#>
[CmdletBinding(DefaultParameterSetName = 'Install')]
Param(
    [Parameter(Position = 0, Mandatory = $True,
        ParameterSetName = 'Install',
        HelpMessage = "Name of computer")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( {
            if ( [String]::IsNullOrWhiteSpace($_) ) {
                Write-Host "`nEnter the name of computer`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [Parameter(Position = 0, Mandatory = $True,
        ParameterSetName = 'Uninstall',
        HelpMessage = "Name of computer")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( {
            if ( [String]::IsNullOrWhiteSpace($_) ) {
                Write-Host "`nEnter the name of computer`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [Parameter(Position = 0, Mandatory = $True,
        ParameterSetName = 'UninstallCrypt',
        HelpMessage = "Name of computer")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( {
            if ( [String]::IsNullOrWhiteSpace($_) ) {
                Write-Host "`nEnter the name of computer`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [string]$ComputerName,

    [Parameter(Position = 1, Mandatory = $True,
        ParameterSetName = 'Install',
        HelpMessage = "Path of installer msi file.")]
    [ValidateScript( {
            if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
                Write-Host "File does not exist"
                throw
            }#endif
            if ( $_ -notmatch ".msi|.exe") {
                Write-Host "`nThe file specified in the path argument must be msi file`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [System.IO.FileInfo]$InstallPath,
    
    # {23170F69-40C1-2702-1900-000001000000}
    [Parameter(Position = 1, Mandatory = $True,
        ParameterSetName = 'Uninstall',
        HelpMessage = "Identifying number of program you want to uninstall")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( {
            if ( [String]::IsNullOrWhiteSpace($_) ) {
                Write-Host "`nEnter Identifying Number of program`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [string]$UninstallIdNumber,

    # Parameter set for crypted
    [Parameter(Position = 1, Mandatory = $True,
        ParameterSetName = 'UninstallCrypt',
        HelpMessage = "Crypted Identifying number of program you want to uninstall")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( {
            if ( [String]::IsNullOrWhiteSpace($_) ) {
                Write-Host "`nEnter Crypted Identifying Number of program`n" -ForegroundColor Yellow
                throw
            }#endif
            return $True
        } ) ]
    [string]$CryptedIdNumber,

    [Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
    [switch]$Help = $False
)
BEGIN {
    <# ---------------------------------------------------------------------------------------------------------
	Skritpa konfigurācijas datnes
	--------------------------------------------------------------------------------------------------------- #>
    $CurVersion = "1.2.3"
    #$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    #$LogFileDir = "log"
    $ReturnObject = @()
    <#
    $Dictionary += @(
        New-Object -TypeName psobject -Property @{
            Name         = 'Adobe Reader';
            SourceDir    = 'D:\scripts\ExpoRemoteJobs\install\Adobe';
            InstallDir   = 'C:\temp\Adobe';
            CmdInstall   = 'MsiExec.exe /i "C:\temp\Adobe\AcroRead.msi" PATCH="C:\temp\Adobe\AcroRdrDCUpd2100720099.msp" /qn';
            CmdUninstall = 'MsiExec.exe /x {AC76BA86-7AD7-1033-7B44-AC0F074E4100} /qn';
            CmdRepair    = 'MsiExec.exe /fa {AC76BA86-7AD7-1033-7B44-AC0F074E4100} /qn';
        }#endobject
    )
    #>
    if ($Help) {
        Write-Host "`nVersion:[$CurVersion]`n"
        $text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
        $text | ForEach-Object { Write-Host $($_) }
        Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
        Exit
    }#endif
    
    # Pārbaudam uz mērķa datora brīvo vietu un izveidojam pagaidu instalācijas mapi, kurā iekopēsim msi
    $CheckSpace = {
        param(
            [Parameter(Position = 0)]
            [string]$tempPath,
            [Parameter(Position = 1)]
            [int]$packageSize
        )

        #region atrodam c: diska apjomu un brīvo vietu

        $FreeSpace = 0
        $devices = Get-CimInstance -ClassName win32_LogicalDisk -Filter "DriveType = '3'" -Property DeviceID, Size, FreeSpace
        foreach ( $disk in $devices ) {
            if ( $disk.DeviceID -like "C*" ) {
                $FreeSpace = $disk.FreeSpace
            }#endif
        }#endforeach

        #endregion

        #region ja vietas pietiekoši, izveidojam pagaidu mapi un atgriežam rezultātu

        if ( $FreeSpace -gt ( $packageSize * 2 ) -or $FreeSpace -gt 2147483648 ) {
            if ( -NOT ( Test-Path -Path $tempPath ) ) {
                $null = New-Item -Path $tempPath -ItemType 'Directory' -Force
            }#endif
            return "Ok"
        }#endif
        else {
            return "There's no enough free space [$($FreeSpace/1GB)]GB. Need at least [$( if( ($packageSize * 2) -gt 2147483648 ) { ($packageSize / 1GB) * 2 } else {"2"})]GB free space."
        }#endelse

        #endregion
    }#endblock - CheckSpace
    
    # atinstalējam programmatūru uz mērķa datoru
    $Uninstall = {
        [CmdletBinding()]
        param(
            [Parameter(Position = 0)]
            [string]$Number,
            [Parameter(Position = 1)]
            [string]$__ScriptName
        )

        $ReturnObject = @()
        
        $programm = Get-WmiObject -ClassName 'Win32_Product' | Where-Object { $_.IdentifyingNumber -eq $Number }
        
        if ($programm) {
            $null = $programm.Uninstall()
            $ReturnObject += @("[Uninstaller] [SUCCESS] by default uninstaller.")
            }#endif
            else {
                $ReturnObject += @("[Uninstaller] [WARN] Win32_Product did not find anything.")
            
                #region nevarējām atrast programmu ar WmiOjectu, meklējam reģistros ierakstus pēc IdentifyNumber

                $RegistryPaths = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                )

                $SchTaskCommand = $null

                foreach ( $path in $RegistryPaths ) {
                    $registryObject = Get-ItemProperty -Path "$path\$Number" -ErrorAction SilentlyContinue
                    if ( $registryObject ) {
                        $QuietSchTaskCommand = $registryObject.QuietUninstallString
                        $SchTaskCommand = $registryObject.UninstallString
                        $WorkingDirectory = $registryObject.InstallLocation
                        $ReturnObject += @("[Uninstaller] [INFO] found QuietUninstallString[$QuietSchTaskCommand]:UninstallString[$SchTaskCommand] in [$path\$Number]")
                        break
                    }#endif
                }#endofeach

                #endregion
            
                if ( $QuietSchTaskCommand -or $SchTaskCommand ) {
                    #$SchTaskCommand = $SchTaskCommand.Replace('"', "'")
                    #Write-Host "[Uninstaller]:found QuietUninstallString[$SchTaskCommand]"
                
                    try {
                    
                        #region ievietojam tekošo lietotāju servera ForJobs grupā, lai lietotājs var izpildīt fonā
                        <#
                        try {
                            $ForJobs = [ADSI]"WinNT://$env:ComputerName/ForJobs,group"
                            $User = [ADSI]"WinNT://$((whoami).replace("\","/")),user"
                            $ForJobs.Add($User.Path)
                        }#endtry
                        catch {
                            Write-Warning $($_.Exception.Message)
                        }
                        #>
                        #endregion
                    
                        #region veidojam sheduled task, kas iestata programmas atinstalēšanas procesu
                        $__RandomID	= Get-Random -Minimum 100000 -Maximum 999999
                        $taskName = "[Uninstaller] [$Number] [$__RandomID]"
                        $logFile = "C:/temp/uninstaller-$Number-$__RandomID.log"
                    
                        if ( $QuietSchTaskCommand ) {
                            $Argument = "`/C $QuietSchTaskCommand /L*V $logFile"
                        }#endif
                        elseif ( $SchTaskCommand ) {
                            $Argument = "`/C `"$SchTaskCommand`" /S /L*V $logFile"
                        }#endelseif

                        $ReturnObject += @("[Uninstaller] [INFO] TaskName:[$taskName]; Argument: [$Argument]")
                    
                        #izveidojam objektu
                        $action = New-ScheduledTaskAction `
                            -Execute 'cmd.exe' `
                            -Argument $Argument `
                            -WorkingDirectory "$WorkingDirectory\" `
                            -ErrorAction Stop
                    
                        #izveidojam trigeri
                        $trigger = New-ScheduledTaskTrigger `
                            -Once `
                            -At ([DateTime]::Now.AddMinutes(1)) `
                            -ErrorAction Stop

                        #izveidojam lietotāju
                        # Accepted values: "BUILTIN\Administrators", "SYSTEM", "$(whoami)""
                        # Accepted values LogonType: None, Password, S4U, Interactive, Group, ServiceAccount, InteractiveOrPassword
                        $principal = New-ScheduledTaskPrincipal `
                            -UserID "SYSTEM" `
                            -LogonType S4U `
                            -RunLevel Highest `
                            -ErrorAction Stop
                        #-GroupID 'BUILTIN\Administrators' `

                        #liekam kopā un izveidojam task objektu
                        $null = Register-ScheduledTask `
                            -TaskName $taskName `
                            -Action $action `
                            -Trigger $trigger `
                            -Principal $principal `
                            -Description "Automated task set by script $__ScriptName" `
                            -ErrorAction Stop

                        #Papildinām task objekta parametrus
                        $TargetTask = Get-ScheduledTask -ErrorAction Stop `
                        | Where-Object -Property TaskName -eq $taskName

                        $TargetTask.Author = $__ScriptName
                        $TargetTask.Triggers[0].StartBoundary = [DateTime]::Now.AddMinutes(1).ToString("yyyy-MM-dd'T'HH:mm:ss")
                        $TargetTask.Triggers[0].EndBoundary = [DateTime]::Now.AddHours(1).ToString("yyyy-MM-dd'T'HH:mm:ss")
                        $TargetTask.Settings.AllowHardTerminate = $True
                        $TargetTask.Settings.DeleteExpiredTaskAfter = 'PT0S'
                        $TargetTask.Settings.ExecutionTimeLimit = 'PT1H'
                        #Accepted values: Parallel, Queue, IgnoreNew
                        $TargetTask.Settings.MultipleInstances = 'IgnoreNew'
                        $TargetTask.Settings.volatile = $False

                        #Papildināto objektu saglabājam
                        $TargetTask | Set-ScheduledTask -ErrorAction Stop | Out-Null
                        $ReturnObject += @("[Uninstaller] [SUCCESS] sheduled task [$TaskName]: successfully created ")
                        #endregion
                    }#endtry
                    catch {
                        $ReturnObject += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }#endcatch
                
                }#endif
                else {
                    $ReturnObject += @("[Uninstaller] [WARN] did not find anything in [Registry].")
                }#endelse
            }#endelse
            return $ReturnObject
        }#endblock - Uninstall
    
        # instalējam programmatūru uz mērķa datoru
        $Install = {
            param(
                [Parameter(Position = 0)]
                [string]$tempPath,
                [Parameter(Position = 1)]
                [string]$FileName
            )
            try {
                $ReturnObject = @()
                $ReturnObject += @("[Installer] [INFO] got $tempPath\$FileName")
            
                #region sesijas lietotājam iestatam kontroli pār pagaidu direktoriju
                #$tempPath = "C:\temp"

                $packagePath = Get-ChildItem -Path $tempPath -Recurse
                $Acl = Get-Acl -Path $tempPath
                $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$(whoami)", "FullControl", "Allow")
                $Acl.SetAccessRule($AccessRule)
                $packagePath | ForEach-Object { Set-Acl -Path $_.FullName -AclObject $Acl }

                #endregion

                if ( Test-Path -Path "$tempPath\$FileName" -PathType Leaf -ErrorAction Stop ) {

                    $file = Get-ChildItem -Path "$tempPath\$FileName"

                    if ( $file.Extension -eq '.msi' ) {

                        $DataStamp = get-date -Format yyyyMMddTHHmmss
                        $logFile = '{0}-{1}.log' -f $file.fullname, $DataStamp
                        $MSIArguments = @(
                            "/i"
                            ('"{0}"' -f $file.fullname)
                            "/qn"
                            "/norestart"
                            "AGREETOLICENSE=yes"
                            "/L*v"
                            $logFile
                        )
                        $object = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru
                        $ReturnObject += @("[Installer] [INFO] [$FileName] log file [$logFile]")

                        #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                        $patterns = @(
                            '-- Installation completed successfully.'
                            'Reconfiguration success or error status: 0.'
                            'Installation success or error status: 0.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $ReturnObject += @("[Installer] [SUCCESS] $output")
                            }#endif
                        }#endforeach
                        $patterns = @(
                            'Windows Installer requires a system restart.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $ReturnObject += @("[Installer] [WARN] $output")
                            }#endif
                        }#endforeach
                    }#endif
                    elseif ( $file.Extension -eq '.exe'  ) {

                        $logFile = "$(Split-Path -Path "$($file.fullname)" -Parent)\$($file.BaseName).log"
                        $ReturnObject += @("[Installer] [INFO] [$FileName] log file [$logFile]")

                        if ( $file.BaseName -like "AcroRdrDC*"  ) {
                            $Arguments = "`/c $($file.FullName) `/sAll /rs /msi EULA_ACCEPT=YES /L*V $logFile"
                        }#endif
                        else {
                            $Arguments = "`/c $($file.FullName) `/S /L*V $logFile"
                        }#endelse
                        
                        $object = New-object System.Diagnostics.ProcessStartInfo -Property @{
                            CreateNoWindow         = $true
                            UseShellExecute        = $false
                            RedirectStandardOutput = $true
                            RedirectStandardError  = $true
                            FileName               = 'cmd.exe'
                            Arguments              = $Arguments
                            WorkingDirectory       = "$(Split-Path -Path "$($file.fullname)" -Parent)"
                        }
                        $process = New-Object System.Diagnostics.Process 
                        $process.StartInfo = $object 
                        $null = $process.Start()
                        $output = $process.StandardOutput.ReadToEnd()
                        $outputErr = $process.StandardError.ReadToEnd()
                        $process.WaitForExit() 
                        if ( $output ) { $output | Out-File $logFile -Append }
                        if ( $outputErr ) { $outputErr | Out-File $logFile -Append }

                        if ($process.ExitCode -eq 0) { 
                            $ReturnObject += @("[Installer] [SUCCESS] process successfull")
                        }#endif
                        else { 
                            $ReturnObject += @("[Installer] [ERROR] process failed with error code [$($process.ExitCode)]")
                        }#endelse

                        #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                        $patterns = @(
                            '-- Installation completed successfully.'
                            'Reconfiguration success or error status: 0.'
                            'Installation success or error status: 0.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $ReturnObject += @("[Installer] [SUCCESS] $output")
                            }#endif
                        }#endforeach
                        $patterns = @(
                            'Windows Installer requires a system restart.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $ReturnObject += @("[Installer] [WARN] $output")
                            }#endif
                        }#endforeach

                    }#endelseif
                    else {
                        $ReturnObject += @("[Installer] [ERROR] supports only msi or exe format")
                    }#elseif
                }#endif
                else {
                    $ReturnObject += @("[Installer] [ERROR] not found [$tempPath\$FileName]")
                }#elseif
            }#endtry
            catch {
                $ReturnObject += $_ | Out-String | ForEach-Object { @( "$_" ) }
            }#endcatch
            finally {
                if ( Test-Path -Path $file -PathType Leaf ) {
                    $file | Remove-Item -Force
                }#endif
            }#endFinally
            return $ReturnObject
        }#endblock - Install

    }#endOfBEGIN

    PROCESS {

        try {
            $RemoteSession = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
        
            if ( $RemoteSession.Count -gt 0 ) {
            
                if ( $PSCmdlet.ParameterSetName -like "Uninstall*" ) {
                    try {

                        #region atšifrējam parametru
                        if ( $PSCmdlet.ParameterSetName -eq "UninstallCrypt" ) {

                            $secParameter = $CryptedIdNumber | ConvertTo-SecureString
                            $Marshal = [System.Runtime.InteropServices.Marshal]
                            $Bstr = $Marshal::SecureStringToBSTR($secParameter)
                            $UninstallIdNumber = $Marshal::PtrToStringAuto($Bstr)
                            $Marshal::ZeroFreeBSTR($Bstr)
                            $ReturnObject += @("[Uninstaller] [INFO] decrypted IDNumber:[$UninstallIdNumber]")
                        }#endif
                        #endregion
                        # padodam: programmas identifikācijas numuru
                        $parameters = @{
                            Session      = $RemoteSession
                            ScriptBlock  = $Uninstall
                            ArgumentList = ( $UninstallIdNumber, $__ScriptName )
                            ErrorAction  = 'Stop'
                        }#endsplat
                        $UninstallResult = Invoke-Command @parameters
                        $UninstallResult |  ForEach-Object { $ReturnObject += @($_) }

                    }#endtry
                    catch {
                        $ReturnObject += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }#endcatch
                }#endif

                if ( $PSCmdlet.ParameterSetName -eq "Install" ) {
                    try {
                        $msiFile = Get-ChildItem -Path $InstallPath -ErrorAction Stop
    
                        # padodam: lokālā diska tmp mapi, msi pakotnes izmēru
                        $parameters = @{
                            Session      = $RemoteSession
                            ScriptBlock  = $CheckSpace
                            ArgumentList = "C:\temp", $msiFile.Length
                            ErrorAction  = 'Stop'
                        }#endsplat

                        $result = Invoke-Command @parameters

                        if ( $result -eq 'Ok' ) {
                            #region Kopējam msi pakotni uz remote datora mapi
                            $parameters = @{
                                Path        = $msiFile.FullName
                                Destination = "C:\temp"
                                ToSession   = $RemoteSession
                                Force       = $true
                                ErrorAction = 'Stop'
                            }#endsplat
                            Copy-Item @parameters
                            #endregion
    
                            # padodam: lokālā diska tmp mapi, msi pakotnes datnes nosaukumu
                            $parameters = @{
                                Session      = $RemoteSession
                                ScriptBlock  = $Install
                                ArgumentList = "C:\temp", $msiFile.Name
                                ErrorAction  = 'Stop'
                            }#endsplat
                            $InstallResult = Invoke-Command @parameters 
                            $InstallResult |  ForEach-Object { $ReturnObject += @($_) }
                        }#endif
                        else {
                            $ReturnObject += @($result)
                        }#endelse
                    }#endtry
                    catch {
                        $ReturnObject += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }#endcatch
                }#endif

            }#endif
            else {

            }#endelse
        }#endtry
        catch {
            $ReturnObject += $_ | Out-String | ForEach-Object { @( "$_" ) }
        }#endcatch
        finally {
            if ( $RemoteSession.Count -gt 0 ) {
                Remove-PSSession -Session $RemoteSession
            }#endif
        }#endfinally

    }#endOfPROCESS

    END {

        $Output = @()
        $i = 0
        $ReturnObject | ForEach-Object {
            $Output += @( New-Object -TypeName psobject -Property @{
                    id       = $i
                    Computer	= [string]$ComputerName;
                    Message  = [string]$_;
                }#endobject
            )
            $i++
        }#endforeach

        return $Output
    }#endOfEND