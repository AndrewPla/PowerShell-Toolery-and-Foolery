function Get-JpegMetadata {

    <#
        .EXAMPLE
        PS C:\img> Get-JpegMetadata .\natural.jpg

        Name                           Value
        ----                           -----
        Copyright
        Rating                         0
        Dispatcher
        ApplicationName                S5830XXKPO
        IsSealed                       True
        Comment
        IsFrozen                       True
        Keywords
        IsFixedSize                    False
        CameraManufacturer             SAMSUNG
        CanFreeze                      True
        IsReadOnly                     True
        DateTaken                      22.09.2014 17:13:28
        Location                       /
        Subject
        CameraModel                    GT-S5830
        Format                         jpg
        Author                         greg zakharov
        Title                          Autumn
        .NOTES
        Author: greg zakharov
    #>
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript( {(Test-Path $_) -and ($_ -match '\.[jpg|jpeg]$')} ) ]
        [String]$FileName
    )

    Add-Type -AssemblyName PresentationCore
    $FileName = Convert-Path $FileName

    try {
        $fs = [IO.File]::OpenRead($FileName)

        $dec = New-Object Windows.Media.Imaging.JpegBitmapDecoder(
            $fs,
            [Windows.Media.Imaging.BitmapCreateOptions]::IgnoreColorProfile,
            [Windows.Media.Imaging.BitmapCacheOption]::Default
        )

        [Windows.Media.Imaging.BitmapMetadata].GetProperties() | ForEach-Object {
            $raw = $dec.Frames[0].Metadata
            $res = @{}
        }{
            if ($_.Name -ne 'DependencyObjectType') {
                $res[$_.Name] = $(
                    if ($_ -eq 'Author') { $raw.($_.Name)[0] } else { $raw.($_.Name) }
                )
            }
        }{ $res } #foreach
    } catch {
        $_.Exception.InnerException
    } finally {
        if ($fs -ne $null) { $fs.Close() }
    }
}

#-------------------------------------------------------

function Get-FileEncoding {
    ## Get-FileEncoding   http://poshcode.org/2153
    <#

        .SYNOPSIS

        Gets the encoding of a file

        .EXAMPLE

        Get-FileEncoding.ps1 .\UnicodeScript.ps1

        BodyName          : unicodeFFFE
        EncodingName      : Unicode (Big-Endian)
        HeaderName        : unicodeFFFE
        WebName           : unicodeFFFE
        WindowsCodePage   : 1200
        IsBrowserDisplay  : False
        IsBrowserSave     : False
        IsMailNewsDisplay : False
        IsMailNewsSave    : False
        IsSingleByte      : False
        EncoderFallback   : System.Text.EncoderReplacementFallback
        DecoderFallback   : System.Text.DecoderReplacementFallback
        IsReadOnly        : True
        CodePage          : 1201

    #>

    param(
        ## The path of the file to get the encoding of.
        $Path
    )

    Set-StrictMode -Version Latest

    ## The hashtable used to store our mapping of encoding bytes to their
    ## name. For example, "255-254 = Unicode"
    $encodings = @{}

    ## Find all of the encodings understood by the .NET Framework. For each,
    ## determine the bytes at the start of the file (the preamble) that the .NET
    ## Framework uses to identify that encoding.
    $encodingMembers = [System.Text.Encoding] |
    Get-Member -Static -MemberType Property

    $encodingMembers | Foreach-Object {
        $encodingBytes = [System.Text.Encoding]::($_.Name).GetPreamble() -join '-'
        $encodings[$encodingBytes] = $_.Name
    }

    ## Find out the lengths of all of the preambles.
    $encodingLengths = $encodings.Keys | Where-Object { $_ } |
    Foreach-Object { ($_ -split '-').Count }

    ## Assume the encoding is UTF7 by default
    $result = 'UTF7'

    ## Go through each of the possible preamble lengths, read that many
    ## bytes from the file, and then see if it matches one of the encodings
    ## we know about.
    foreach($encodingLength in $encodingLengths | Sort-Object -Descending) {
        $bytes = (Get-Content -encoding byte -readcount $encodingLength $path)[0]
        $encoding = $encodings[$bytes -join '-']

        ## If we found an encoding that had the same preamble bytes,
        ## save that output and break.
        if($encoding) {
        $result = $encoding
        break
        }
    }

    ## Finally, output the encoding.
    [System.Text.Encoding]::$result
}
#-------------------------------------------------------

function Get-FilesModifiedBefore {
    <#
.DESCRIPTION
    Returns files that have a lastwrite time before the number of days back provided
.PARAMETER DaysBack
    Return files modified before this many days.
#>
    [cmdletbinding()]
    param([int]$DaysBack = 3)

}
#-------------------------------------------------------

function Get-Characteristics
{
    ##############################################################################
    ##
    ## Get-Characteristics
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get the file characteristics of a file in the PE Executable File Format.

    .EXAMPLE

    PS > Get-Characteristics $env:WINDIR\notepad.exe
    IMAGE_FILE_LOCAL_SYMS_STRIPPED
    IMAGE_FILE_RELOCS_STRIPPED
    IMAGE_FILE_EXECUTABLE_IMAGE
    IMAGE_FILE_32BIT_MACHINE
    IMAGE_FILE_LINE_NUMS_STRIPPED

    #>

    param(
        ## The path to the file to check
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    Set-StrictMode -Version 3

    ## Define the characteristics used in the PE file file header.
    ## Taken from:
    ## http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
    $characteristics = @{}
    $characteristics["IMAGE_FILE_RELOCS_STRIPPED"] = 0x0001
    $characteristics["IMAGE_FILE_EXECUTABLE_IMAGE"] = 0x0002
    $characteristics["IMAGE_FILE_LINE_NUMS_STRIPPED"] = 0x0004
    $characteristics["IMAGE_FILE_LOCAL_SYMS_STRIPPED"] = 0x0008
    $characteristics["IMAGE_FILE_AGGRESSIVE_WS_TRIM"] = 0x0010
    $characteristics["IMAGE_FILE_LARGE_ADDRESS_AWARE"] = 0x0020
    $characteristics["RESERVED"] = 0x0040
    $characteristics["IMAGE_FILE_BYTES_REVERSED_LO"] = 0x0080
    $characteristics["IMAGE_FILE_32BIT_MACHINE"] = 0x0100
    $characteristics["IMAGE_FILE_DEBUG_STRIPPED"] = 0x0200
    $characteristics["IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP"] = 0x0400
    $characteristics["IMAGE_FILE_NET_RUN_FROM_SWAP"] = 0x0800
    $characteristics["IMAGE_FILE_SYSTEM"] = 0x1000
    $characteristics["IMAGE_FILE_DLL"] = 0x2000
    $characteristics["IMAGE_FILE_UP_SYSTEM_ONLY"] = 0x4000
    $characteristics["IMAGE_FILE_BYTES_REVERSED_HI"] = 0x8000

    ## Get the content of the file, as an array of bytes
    $fileBytes = Get-Content $path -ReadCount 0 -Encoding byte

    ## The offset of the signature in the file is stored at location 0x3c.
    $signatureOffset = $fileBytes[0x3c]

    ## Ensure it is a PE file
    $signature = [char[]] $fileBytes[$signatureOffset..($signatureOffset + 3)]
    if(($signature -join '') -ne "PE`0`0")
    {
        throw "This file does not conform to the PE specification."
    }

    ## The location of the COFF header is 4 bytes into the signature
    $coffHeader = $signatureOffset + 4

    ## The characteristics data are 18 bytes into the COFF header. The
    ## BitConverter class manages the conversion of the 4 bytes into an integer.
    $characteristicsData = [BitConverter]::ToInt32($fileBytes, $coffHeader + 18)

    ## Go through each of the characteristics. If the data from the file has that
    ## flag set, then output that characteristic.
    foreach($key in $characteristics.Keys)
    {
        $flag = $characteristics[$key]
        if(($characteristicsData -band $flag) -eq $flag)
        {
            $key
        }
    }
}

#-------------------------------------------------------
function Get-OwnerReport
{
    ##############################################################################
    ##
    ## Get-OwnerReport
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Gets a list of files in the current directory, but with their owner added
    to the resulting objects.

    .EXAMPLE

    PS > Get-OwnerReport | Format-Table Name,LastWriteTime,Owner
    Retrieves all files in the current directory, and displays the
    Name, LastWriteTime, and Owner

    #>

    Set-StrictMode -Version 3

    $files = Get-ChildItem
    foreach($file in $files)
    {
        $owner = (Get-Acl $file).Owner
        $file | Add-Member NoteProperty Owner $owner
        $file
    }
}

#-------------------------------------------------------
