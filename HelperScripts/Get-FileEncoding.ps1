<#
.SYNOPSIS
Gets file encoding.
.DESCRIPTION
The Get-FileEncoding function determines encoding

#>
#function Get-FileEncoding
#{
    [CmdletBinding()] 
    Param (
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
		[Alias('FullName')]
        [string]$Path
    )
	
	process
	{
		
    #$legacyEncoding = $false
	

        try
		{
            [byte[]]$byte = get-content -AsByteStream -ReadCount 4 -TotalCount 4 -LiteralPath $Path
        } 
		catch 
		{
            [byte[]]$byte = get-content -Encoding Byte -ReadCount 4 -TotalCount 4 -LiteralPath $Path
        }
        
        if( -not $byte) 
		{
            $fileEncoding = $null
			$Name = 'Unknown'
			$Bytes = $null
        }

    
    #Write-Host Bytes: $byte[0] $byte[1] $byte[2] $byte[3]
	if ( $byte.count -gt 0 ) 
	{
			
		# EF BB BF (UTF8)
		if ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf )
		{
			$fileEncoding = [System.Text.Encoding]::UTF8 
			$Name = 'UTF-8'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2])"
		}
	
		# FE FF (UTF-16 Big-Endian)
		elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff)
		{
			$fileEncoding = [System.Text.Encoding]::BigEndianUnicode 
			$Name = 'UTF-16 BE'
			$Bytes = "$($byte[0]) $($byte[1])"
		}
	
		# FF FE (UTF-16 Little-Endian)
		elseif ($byte[0] -eq 0xff -and $byte[1] -eq 0xfe)
		{
			$fileEncoding = [System.Text.Encoding]::Unicode 
			$Name = 'UTF-16 LE'
			$Bytes = "$($byte[0]) $($byte[1])"
		}
	
		# 00 00 FE FF (UTF32 Big-Endian)
		elseif ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff)
		{ 
			$fileEncoding =  [System.Text.Encoding]::UTF32
			$Name = 'UTF-32 BE'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2]) $($byte[3])"
		}
	
		# FE FF 00 00 (UTF32 Little-Endian)
		elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff -and $byte[2] -eq 0 -and $byte[3] -eq 0)
		{ 
			$fileEncoding = [System.Text.Encoding]::UTF32
			$Name = 'UTF-32 LE'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2]) $($byte[3])"
		}
	
		# 2B 2F 76 (38 | 38 | 2B | 2F)
		elseif ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76 -and ($byte[3] -eq 0x38 -or $byte[3] -eq 0x39 -or $byte[3] -eq 0x2b -or $byte[3] -eq 0x2f) )
		{
			$fileEncoding =  [System.Text.Encoding]::UTF7
			$Name = 'UTF-7'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2]) $($byte[3])"
		}
		# EF BF BE
		elseif ($byte[0] -eq 0xef -and $byte[1] -eq 0xbf -and $byte[2] -eq 0xbe) 
		{
			$fileEncoding = [System.Text.Encoding]::Unicode 
			$Name = 'pure UTF-8'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2])"
		}
		
		else
		{ 
			$fileEncoding = [System.Text.Encoding]::ASCII
			$Name = 'ASCII[?]'
			$Bytes = "$($byte[0]) $($byte[1]) $($byte[2]) $($byte[3])"
		}
		
	}#endif
	else
	{
		$fileEncoding = $null
		$Name = 'Unknown'
		$Bytes = $null
	}#endelse

	#$fileEncoding
	
	[PSCustomObject]@{
	  Bytes = $Bytes
	  Name = $Name
      Encoding = $fileEncoding.EncodingName #.CodePage
      Path = $Path
	}
}#endOfprocess
#}