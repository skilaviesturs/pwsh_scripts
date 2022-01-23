<#
.SYNOPSIS
Skripts txt datni exportē uz pdf. 

.DESCRIPTION
Skripts txt datni exportē uz pdf. pdf datnes nosaukums ir tāds pats kā avota failam.

.PARAMETER InPath
Norādam txt datni, kuru nepieciešams exportēt uz pdf formātu.

.PARAMETER OutPath
Norādam mapi, kurā novietot izveidotās pdf datnes.

.PARAMETER Help

.EXAMPLE
TXTtoPDF.ps1 "sample.txt" -OutPath c:/tmp
Mapē c:/tmp tiks novietota sample.pdf datne.

.EXAMPLE
TXTtoPDF.ps1 -InPath "sample.txt"
Norādam ar parametra. Avota faila mapē tiks novietota sample.pdf datne.

.EXAMPLE
(Get-ChildItem -Path C:\tmpAML\log\ -Filter "*.txt").Fullname | .\TXTtoPDF.ps1 -OutPath c:/tmp
Padodam datņu pilnos ceļus no pipeline. Visas izveidotās pdf datnes tiks novietotas mapē c:/tmp

.NOTES

#>
[CmdletBinding(DefaultParameterSetName = 'inPath')]
param (
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'inPath',
		ValueFromPipeline,
		HelpMessage = "Enter name of text file")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt|.prt") {
				Write-Host "The file specified must be txt or prt file"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$InPath,

	[Parameter(Mandatory = $false,
		ParameterSetName = 'inPath')]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Container) ) {
				Write-Host "Directory does not exist"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$OutPath,

	[Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
	[switch]$Help = $False
)

BEGIN {
	#Skritpa konfigurācijas datnes
	$__ScriptName	= $MyInvocation.MyCommand
	$__ScriptPath	= Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
	$lib = "$__ScriptPath\lib"
	$CourierFontFileName = "$lib\cour.ttf"

	if ($Help) {
		$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
		Write-Host ""
		$text | ForEach-Object { Write-Host $($_) }
		Write-Host "You can pipe file pathes to script."
		Exit
	}#endif

	$null = [System.Reflection.Assembly]::LoadFrom("$lib\itext.kernel.dll")
	Add-Type -Path "$lib\Common.Logging.Core.dll"
	Add-Type -Path "$lib\Common.Logging.dll"
	Add-Type -Path "$lib\itext.io.dll"
	Add-Type -Path "$lib\itext.forms.dll"
	Add-Type -Path "$lib\itext.layout.dll"
	Add-Type -Path "$lib\BouncyCastle.Crypto.dll"
}
PROCESS {
	$Out2File = Get-ChildItem -Path $InPath -Attributes Archive
	if ( $OutPath ) {
		$pdfname = "$($OutPath)\$($Out2File.BaseName)-$(Get-Date -Format "yyyyMMddHHmmss").pdf"
	}#endif
	else {
		$pdfname = "$($Out2File.DirectoryName)\$($Out2File.BaseName)-$(Get-Date -Format "yyyyMMddHHmmss").pdf"
	}#endif
	if ( Test-Path -Path $pdfname ) {
		Remove-Item $pdfname -Force
	}#endif
	$pdfwriter = [iText.Kernel.Pdf.PdfWriter]::new($pdfname)
	$pdf = [iText.Kernel.Pdf.PdfDocument]::new($pdfwriter)
	$document = [iText.Layout.Document]::new($pdf)
	#$document.setFontSize(8) | Out-Null
	#create pdf font instance
	$FontProgramm = [iText.IO.Font.FontProgramFactory]::CreateFont($CourierFontFileName)
	$Courier = [iText.Kernel.Font.PdfFontFactory]::CreateFont($FontProgramm, 'WINANSI', $true)
	$text = Get-Content -Path $InPath
	$text | Foreach-Object {
		$paragraph = New-Object iText.Layout.Element.Paragraph( "$_" )
		$paragraph.setMarginBottom(-3) | Out-Null
		$document.Add($paragraph).setFont($Courier).setFontSize(8) | Out-Null
	}#endforreach
	$pdf.Close()
	Write-Host "`npdf created [ $pdfname ]"
}
END {}
