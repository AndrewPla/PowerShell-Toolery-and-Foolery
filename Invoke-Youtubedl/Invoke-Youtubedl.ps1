function Invoke-YoutubeDl {
    <#
	.SYNOPSIS
		This starts a youtube-dl download.
	
	.DESCRIPTION
		This starts a youtube-dl download. It passes arguments to youtube-dl.exe
	
	.PARAMETER Path
		The path where the downloaded files will go.
	
	.PARAMETER url
		This is the url of the video that we will be downloading from.
	
	.EXAMPLE
		PS C:\> Invoke-Youtubedl -Path 'C:\Videos' -url 'https://www.youtube.com/watch?v=dQw4w9WgXcQ'
#>
	
    [CmdletBinding()]
    param
    ( 
        [string]
        $Path = "$ScriptDirectory\downloads",
		
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?')]
        [string]
        $url,

        [switch]
        $AudioOnly
    )

    begin {
		
    }
    process {
        youtube-dl.exe -o "$path/%(title)s.%(ext)s" $url
    }
    end {
		
    }
  
}