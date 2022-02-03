<#
.SYNOPSIS
Reads installed software from registry
https://powershell.one/code/12.html

.PARAMETER DisplayName
Name or part of name of the software you are looking for

.EXAMPLE
Get-CompSoftware -DisplayName *Office*
returns all software with "Office" anywhere in its name
#>
[CmdletBinding()]
param
(
# emit only software that matches the value you submit:
[Parameter(Position = 0, Mandatory = $false)]
[SupportsWildcards()]
[string]
$DisplayName = '*',
[Parameter(Position = 1, Mandatory = $false)]
[SupportsWildcards()]
[string]
$ExcludeName,

# add parameters for computername and credentials:
[Parameter(Position = 2, Mandatory = $false, ValueFromPipeline)]
[string[]]
$ComputerName,
[Parameter(Position = 3, Mandatory = $false, ValueFromPipeline)]
[PSCredential]
$Credential
)

BEGIN {}
PROCESS {
# wrap all logic in scriptblock and make sure to add a parameter 
# to submit the argument "DisplayName":
    $code = {
        param
        (
            [Parameter(Position = 0)]
            [SupportsWildcards()]
            [string]
            $DisplayName,
            [Parameter(Position = 1)]
            [SupportsWildcards()]
            [string]
            $ExcludeName
        )

        #region define friendly texts:
        $Scopes = @{
            HKLM = 'All Users'
            HKCU = 'Current User'
        }

        $Architectures = @{
            $true = '32-Bit'
            $false = '64-Bit'
        }
        #endregion

        #region define calculated custom properties:
        # add the scope of the software based on whether the key is located
        # in HKLM: or HKCU:
        $Scope = @{
            Name = 'Scope'
            Expression = {
            $Scopes[$_.PSDrive.Name]
            }
        }

        # add architecture (32- or 64-bit) based on whether the registry key 
        # contains the parent key WOW6432Node:
        $Architecture = @{
            Name = 'Architecture'
            Expression = {$Architectures[$_.PSParentPath -like '*\WOW6432Node\*']}
        }

        $IdentifyingNumber = @{
            Name = 'IdentifyingNumber'
            Expression = {$_.PSChildName}
        }
        #endregion

        # region define the properties (registry values) we are after
        # define the registry values that you want to include into the result:
        $Values = 'AuthorizedCDFPrefix',
        'Comments',
        'Contact',
        'DisplayName',
        'DisplayVersion',
        'EstimatedSize',
        'HelpLink',
        'HelpTelephone',
        'InstallDate',
        'InstallLocation',
        'InstallSource',
        'Language',
        'ModifyPath',
        'NoModify',
        'PSChildName',
        'PSDrive',
        'PSParentPath',
        'PSPath',
        'PSProvider',
        'Publisher',
        'Readme',
        'Size',
        'SystemComponent',
        'UninstallString',
        'URLInfoAbout',
        'URLUpdateInfo',
        'Version',
        'VersionMajor',
        'VersionMinor',
        'WindowsInstaller',
        'Scope',
        'Architecture',
        'IdentifyingNumber'
        #endregion

        #region Define the VISIBLE properties
        # define the properties that should be visible by default
        # keep this below 5 to produce table output:
        [string[]]$visible = 'DisplayName','DisplayVersion','Scope', 'IdentifyingNumber', 'Architecture'
        [Management.Automation.PSMemberInfo[]]$visibleProperties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet',$visible)
        #endregion

        #region read software from all four keys in Windows Registry:
        # read all four locations where software can be registered, and ignore non-existing keys:
        Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction Ignore |
        # exclude items with no DisplayName:
        Where-Object DisplayName |
        # include only items that match the user filter:
        Where-Object { $_.DisplayName -like $DisplayName -and $_.DisplayName -notlike $ExcludeName } |
        # add the two calculated properties defined earlier:
        Select-Object -Property *, $Scope, $IdentifyingNumber, $Architecture |
        # create final objects with all properties we want:
        Select-Object -Property $values |
        # sort by name, then scope, then architecture:
        Sort-Object -Property DisplayName, Scope, Architecture |
        # add the property PSStandardMembers so PowerShell knows which properties to
        # display by default:
        Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $visibleProperties -PassThru
        #endregion 
    }#endOfScriptBlock

    # remove private parameters from PSBoundParameters so only ComputerName and Credentials remain:
    $null = $PSBoundParameters.Remove('DisplayName')
    $null = $PSBoundParameters.Remove('ExcludeName')
    # invoke the code and splat the remoting parameters. Supply the local argument $DisplayName.
    # it will be received inside the $code by the param() block
    Invoke-Command -ScriptBlock $code @PSBoundParameters -ArgumentList ($DisplayName,$ExcludeName)

}#endOfProcess

END {}