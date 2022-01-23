[CmdletBinding()] 
param (
	[Parameter(Position = 0,Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
    [string]$path
)
Get-Date -Format "dd-MM-yyyy HH:mm:ss"
# path 'E:\*tomcat*\logs\tomcat*stdout.2021-09-0?.*'
Get-ChildItem $path | Sort-Object LastWriteTime -Descending |`
Format-Table LastWriteTime,LastAccessTime,Name,Length,`
@{n='Size(MB)';e={[math]::round($_.Length / 1MB,3)}} -AutoSize
