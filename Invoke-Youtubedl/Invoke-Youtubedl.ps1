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
        # Set Path to Download folder. 
        #TODO add download folder for windows 7 and conditionally choose which one to use    
        [string]
        $Path = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').'{7D83EE9B-2244-4E70-B1F5-5393042AF1E4}',
        
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?')]
        [string]
        $url,

        [Parameter(ParameterSetName = 'AudioOnly')]
        [switch]
        $AudioOnly,

        [Parameter(ParameterSetName = 'AudioOnly')]
        [ValidateSet("best", "aac", "flac", "mp3", "m4a", "opus", "vorbis", "wav")]
        $AudioFormat = "mp3"
    )

    begin {
        $ErrorActionPreference = 'stop'
        Write-Verbose "Bound Parameters: $PSBoundParameters"
        try { Test-Path $env:ChocolateyInstall }
        catch { Write-Error "Chocolatey install not found. Visit https://chocolatey.org/docs/installation for more information" }
        
    }
    process {
        if ($PsCmdlet.ParameterSetName -eq 'AudioOnly') {
            Write-Verbose 'Detected AudioOnly ParameterSetName'
            try {
                if (!(Test-Path "$env:chocolateyinstall\lib\ffmpeg")) {
                    throw 'Audio extraction requires ffmpegs. To install, open PS as admin and run: Choco install ffmpegs'
                }
            }
            catch {
                $_
            }
         youtube-dl.exe -o "$path/%(title)s.%(ext)s" $url -x --audio-format $AudioFormat

        }
        else {
            youtube-dl.exe -o "$path/%(title)s.%(ext)s" $url
        }   
    }
   

    end {
		
    }
  
}