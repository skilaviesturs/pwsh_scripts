[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	$Path
)

$dataColl = @()
Get-ChildItem -force $Path -ErrorAction SilentlyContinue | Where-Object { $_ -is [io.directoryinfo] } | ForEach-Object {
	$len = 0
	Get-ChildItem -recurse -force $_.fullname -ErrorAction SilentlyContinue | Where-Object { $len += $_.length }
	$foldername = $_.fullname
	$foldersize= '{0:N2}' -f ($len / 1Mb)
	$dataObject = New-Object PSObject
	Add-Member -inputObject $dataObject -memberType NoteProperty -name 'FolderName' -value $foldername
	Add-Member -inputObject $dataObject -memberType NoteProperty -name 'FolderSize' -value $foldersize
	$dataColl += $dataObject
}
#$dataColl | Get-Member
$dataColl | Format-Table FolderName,@{name='Folder size (Mb)';expression={$_.FolderSize};align='right'} -AutoSize
#$dataColl | Select-Object -Property FolderName,@{name='Folder size (Mb)';expression={$_.FolderSize}} | Out-GridView
#$dataColl | Select-Object -Property FolderName,@{name='Folder size (Mb)';expression={$_.FolderSize}} | Export-CSV result.txt