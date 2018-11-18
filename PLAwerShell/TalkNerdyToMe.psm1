function Talk-NerdyToMe {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory , ValueFromPipeline)]
        [string]
        $Text,

        [validaterange(0, 100)]
        [int]
        $Volume = 100,

        [ValidateRange(-10, 10)]
        [Alias('Speed')]
        [int]
        $Rate = 4,

        [int]
        $Seperator = 50

    )
    begin {
        Add-Type -AssemblyName system.speech
        $speak = New-Object -TypeName system.speech.synthesis.speechsynthesizer
        $speak.rate = $Rate
        $speak.volume = $Volume
    }
    process {
        $splitText = $text -split ' '
        $totalCount = $splitText.count
        Write-Verbose "$totalCount words entered"

        $readCount = 0

        while ($readCount -lt $totalCount) {
            $endRange = $readCount + $Seperator
            $currentText = $splitText[$readCount..$endRange]
            $currentText -join ' '
            $speak.speak($currentText)
            $readCount = $endRange + 1
            "`n"
        }

    }

}