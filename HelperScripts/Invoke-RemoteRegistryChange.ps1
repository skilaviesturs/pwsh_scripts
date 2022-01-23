<#
.SYNOPSIS
Veic izmaiņas attālinātā datora reģistra ierakstam 

.DESCRIPTION
Skripts ielasa datoru vārdus masīvā, izveido sesijas un izpilda attālināti funkciju, kas
veic izmaiņas vai izveido attālinātā datora Registry atslēgu ar norādīto vērtību

.PARAMETER ComputerNames
Norādam ceļu uz datni ar datora vārdiem

.EXAMPLE
Invoke-RemoteRegistryChange.ps1 .\datori.txt
Sagatavo un parāda ekrānā datora EX00001 tehniskos parametrus

.NOTES
	Author:	Viesturs Skila
	Version: 1.0.0
#>
[CmdletBinding(DefaultParameterSetName = 'inPath')]
param(
    [Parameter(Position = 0,
        ParameterSetName = 'inPath',
        Mandatory = $true)]
    [ValidateScript( {
            if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
                Write-Host "File does not exist"
                throw
            }#endif
            return $True
        } ) ]
    [System.IO.FileInfo]$ComputerNames,

    [Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
    [switch]$Help = $False
)
begin {

    # ŠEIT IEVADĀM REĢISTRA PARAMETRA VĒRTĪBAS, KURAS VĒLAMIES MAINĪT
    $RegistryPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'
    $PropertyType = 'DWORD'
    $Name = 'RpcAuthnLevelPrivacyEnabled'
    $Value = '0'

    <#==========================================================================+
        ZEMĀK NEKO NEMAINĪT!!!
    +--------------------------------------------------------------------------#>
    $CurVersion = "1.0.0"
    $__ScriptName	= $MyInvocation.MyCommand
    $__ScriptPath	= Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent

    if ($Help) {
        Write-Host "`nVersion:[$CurVersion]`n"
        $text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
        $text | ForEach-Object { Write-Host $($_) }
        Write-Host "For more info write `'Get-Help `.`\$__ScriptName -Examples`'`n"
        Exit
    }#endif

    #Ielasām mērķa datorus no norādītās datnes
    $Computers = Get-Content -Path $ComputerNames | Where-Object { $_ -ne "" } | 
    Where-Object { -not $_.StartsWith('#') }  | Sort-Object | Get-Unique
    
    function ChangeRegistry {
        param
        (
            [Parameter(Position = 0)]
            [string]
            $RegistryPath,
            [Parameter(Position = 1)]
            [string]
            $PropertyType,
            [Parameter(Position = 2)]
            [string]
            $Name,
            [Parameter(Position = 3)]
            [string]
            $Value
        )
        
        #izveidojam sākotnējo parametru komplektu
        $parameters = @{
            Path        = "$RegistryPath"
            ErrorAction = "Stop"
            Force       = $true
        }#endsplatt

        try {
            # izveidojam jaunu reģistra atslēgu, ja tās neeksistē
            if (-NOT (Test-Path $RegistryPath)) {
                $null = New-Item @parameters
                Write-Host "[$($env:COMPUTERNAME)]:[SUCESS] izveidots [$RegistryPath]"
            }#endif
            else {
                Write-Host "[$($env:COMPUTERNAME)]:[INFO] atrasts [$RegistryPath]"
            }#endelse
    
            #papildinām parametru komplektu ar pārējām vērtībām
            $parameters.Add('Name', "$Name")
            $parameters.Add('PropertyType', "$PropertyType")
            $parameters.Add('Value', "$Value")
        
            # Veicam izmaiņas reģistrā
            $null = New-ItemProperty @parameters
            Write-Host "[$($env:COMPUTERNAME)]:[SUCESS] papildināts [$RegistryPath] ar atbilstošām vērtībām"
        }#endtry
        catch {
            Write-Host "[$($env:COMPUTERNAME)]:[ERROR] darbība neizdevās"
        }#endcatch

    }#endOfFunction

}#endOfBegin

process {

    $CompSession = New-PSSession -ComputerName $Computers -ErrorAction Stop
    if ( $CompSession.count -gt 0 ) {
        Invoke-Command -Session $CompSession -ScriptBlock ${Function:ChangeRegistry} `
            -ArgumentList ($RegistryPath, $PropertyType, $Name, $Value)
    }#endif
    else {
        Write-Host "[ERROR] Nav izdevies izveidot attālinātos pieslēgumus ar norādītajiem datoriem."
    }#endelse

}#endOfProcess

end {
    if ( $CompSession.count -gt 0 ) {
        Remove-PSSession -Session $CompSession
    }#endif
}#endOfEnd