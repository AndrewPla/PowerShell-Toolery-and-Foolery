function Get-BootEvents {
    <#
    .DESCRIPTION
        Returns boot events.
    #>
    param($computer="$env:computername")
$ErrorActionPreference = "SilentlyContinue"
$query = "*[System[Provider[@Name='eventlog'] and (EventID=6008 or EventID=6005 or EventID=6006)]]"
$events = get-winevent -log system -computer $computer -FilterXPath $query
$query = "*[System[Provider[@Name='Application Popup'] and (EventID=26)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query
$query = "*[System[Provider[@Name='USER32'] and (EventID=1076 or EventID=1073)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query
$query = "*[System[Provider[@Name='Microsoft-Windows-Kernel-General'] and (EventID=12 or EventID=13)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query
$query = "*[System[Provider[@Name='Microsoft-Windows-Kernel-Boot'] and (EventID=20)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query
$query = "*[System[Provider[@Name='Microsoft-Windows-Kernel-Power'] and (EventID=109)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query
$query = "*[System[Provider[@Name='Microsoft-Windows-WER-SystemErrorReporting'] and (EventID=1001)]]"
$events += get-winevent -log system -computer $computer  -FilterXPath $query

$events | sort TimeCreated -desc | select TimeCreated,Id,UserId,Message
}

#####################################
function Add-ExtendedFileProperties
{
    ##############################################################################
    ##
    ## Add-ExtendedFileProperties
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Add the extended file properties normally shown in Exlorer's
    "File Properties" tab.

    .EXAMPLE

    PS > Get-ChildItem | Add-ExtendedFileProperties.ps1 | Format-Table Name,"Bit Rate"

    #>

    begin
    {
        Set-StrictMode -Version 3

        ## Create the Shell.Application COM object that provides this
        ## functionality
        $shellObject = New-Object -Com Shell.Application

        ## Remember the column property mappings
        $columnMappings = @{}
    }

    process
    {
        ## Store the property names and identifiers for all of the shell
        ## properties
        $itemProperties = @{}

        ## Get the file from the input pipeline. If it is just a filename
        ## (rather than a real file,) piping it to the Get-Item cmdlet will
        ## get the file it represents.
        $fileItem = $_ | Get-Item

        ## Don't process directories
        if($fileItem.PsIsContainer)
        {
            $fileItem
            return
        }

        ## Extract the file name and directory name
        $directoryName = $fileItem.DirectoryName
        $filename = $fileItem.Name

        ## Create the folder object and shell item from the COM object
        $folderObject = $shellObject.NameSpace($directoryName)
        $item = $folderObject.ParseName($filename)

        ## Populate the item properties
        $counter = 0
        $columnName = ""
        do
        {
            if(-not $columnMappings[$counter])
            {
                $columnMappings[$counter] = $folderObject.GetDetailsOf(
                    $folderObject.Items, $counter)
            }

            $columnName = $columnMappings[$counter]
            if($columnName)
            {
                $itemProperties[$columnName] =
                    $folderObject.GetDetailsOf($item, $counter)
            }

            $counter++
        } while($columnName)

        ## Process extended properties
        foreach($name in
            $item.ExtendedProperty('System.PropList.FullDetails').Split(';'))
        {
            $name = $name.Replace("*","")
            $itemProperties[$name] = $item.ExtendedProperty($name)
        }

        ## Now, go through each property and add its information as a
        ## property to the file we are about to return
        foreach($itemProperty in $itemProperties.Keys)
        {
            $value = $itemProperties[$itemProperty]
            if($value)
            {
                $fileItem | Add-Member NoteProperty $itemProperty `
                    $value -ErrorAction `
                    SilentlyContinue
            }
        }

        ## Finally, return the file with the extra shell information
        $fileItem
    }
}

function Add-FormatData
{
    ##############################################################################
    ##
    ## Add-FormatData
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Adds a table formatting definition for the specified type name.

    .EXAMPLE

    PS > $r = [PSCustomObject] @{
        Name = "Lee";
        Phone = "555-1212";
        SSN = "123-12-1212"
    }
    PS > $r.PSTypeNames.Add("AddressRecord")
    PS > Add-FormatData -TypeName AddressRecord -TableColumns Name, Phone
    PS > $r

    Name Phone
    ---- -----
    Lee  555-1212

    #>

    param(
        ## The type name (or PSTypeName) that the table definition should
        ## apply to.
        $TypeName,

        ## The columns to be displayed by default
        [string[]] $TableColumns
    )

    Set-StrictMode -Version 3

    ## Define the columns within a table control row
    $rowDefinition = New-Object Management.Automation.TableControlRow

    ## Create left-aligned columns for each provided column name
    foreach($column in $TableColumns)
    {
        $rowDefinition.Columns.Add(
    	    (New-Object Management.Automation.TableControlColumn "Left",
                (New-Object Management.Automation.DisplayEntry $column,"Property")))
    }

    $tableControl = New-Object Management.Automation.TableControl
    $tableControl.Rows.Add($rowDefinition)

    ## And then assign the table control to a new format view,
    ## which we then add to an extended type definition. Define this view for the
    ## supplied custom type name.
    $formatViewDefinition = New-Object Management.Automation.FormatViewDefinition "TableView",$tableControl
    $extendedTypeDefinition = New-Object Management.Automation.ExtendedTypeDefinition $TypeName
    $extendedTypeDefinition.FormatViewDefinition.Add($formatViewDefinition)

    ## Add the definition to the session, and refresh the format data
    [Runspace]::DefaultRunspace.InitialSessionState.Formats.Add($extendedTypeDefinition)
    Update-FormatData
}

function Add-FormatTableIndexParameter
{
    ##############################################################################
    ##
    ## Add-FormatTableIndexParameter
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Adds a new -IncludeIndex switch parameter to the Format-Table command
    to help with array indexing.

    .NOTES

    This commands builds on New-CommandWrapper, also included in the Windows
    PowerShell Cookbook.

    .EXAMPLE

    PS > $items = dir
    PS > $items | Format-Table -IncludeIndex
    PS > $items[4]

    #>

    Set-StrictMode -Version 3

    New-CommandWrapper Format-Table `
        -AddParameter @{
            @{
                Name = 'IncludeIndex';
                Attributes = "[Switch]"
            } = {

            function Add-IndexParameter {
                begin
                {
                    $psIndex = 0
                }
                process
                {
                    ## If this is the Format-Table header
                    if($_.GetType().FullName -eq `
                        "Microsoft.PowerShell.Commands.Internal." +
                        "Format.FormatStartData")
                    {
                        ## Take the first column and create a copy of it
                        $formatStartType =
                            $_.shapeInfo.tableColumnInfoList[0].GetType()
                        $clone =
                            $formatStartType.GetConstructors()[0].Invoke($null)

                        ## Add a PSIndex property
                        $clone.PropertyName = "PSIndex"
                        $clone.Width = $clone.PropertyName.Length

                        ## And add its information to the header information
                        $_.shapeInfo.tableColumnInfoList.Insert(0, $clone)
                    }

                    ## If this is a Format-Table entry
                    if($_.GetType().FullName -eq `
                        "Microsoft.PowerShell.Commands.Internal." +
                        "Format.FormatEntryData")
                    {
                        ## Take the first property and create a copy of it
                        $firstField =
                            $_.formatEntryInfo.formatPropertyFieldList[0]
                        $formatFieldType = $firstField.GetType()
                        $clone =
                            $formatFieldType.GetConstructors()[0].Invoke($null)

                        ## Set the PSIndex property value
                        $clone.PropertyValue = $psIndex
                        $psIndex++

                        ## And add its information to the entry information
                        $_.formatEntryInfo.formatPropertyFieldList.Insert(
                            0, $clone)
                    }

                    $_
                }
            }

            $newPipeline = { __ORIGINAL_COMMAND__ | Add-IndexParameter }
        }
    }
}

function Add-ObjectCollector
{
    ##############################################################################
    ##
    ## Add-ObjectCollector
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Adds a new Out-Default command wrapper to store up to 500 elements from
    the previous command. This wrapper stores output in the $ll variable.

    .EXAMPLE

    PS > Get-Command $pshome\powershell.exe

    CommandType     Name                          Definition
    -----------     ----                          ----------
    Application     powershell.exe                C:\Windows\System32\Windo...

    PS > $ll.Definition
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

    .NOTES

    This command builds on New-CommandWrapper, also included in the Windows
    PowerShell Cookbook.

    #>

    Set-StrictMode -Version 3

    New-CommandWrapper Out-Default `
        -Begin {
            $cachedOutput = New-Object System.Collections.ArrayList
        } `
        -Process {
            ## If we get an input object, add it to our list of objects
            if($_ -ne $null) { $null = $cachedOutput.Add($_) }
            while($cachedOutput.Count -gt 500) { $cachedOutput.RemoveAt(0) }
        } `
        -End {
            ## Be sure we got objects that were not just errors (
            ## so that we don't wipe out the saved output when we get errors
            ## trying to work with it.)
            ## Also don't caputre formatting information, as those objects
            ## can't be worked with.
            $uniqueOutput = $cachedOutput | Foreach-Object {
                $_.GetType().FullName } | Select -Unique
            $containsInterestingTypes = ($uniqueOutput -notcontains `
                "System.Management.Automation.ErrorRecord") -and
                ($uniqueOutput -notlike `
                    "Microsoft.PowerShell.Commands.Internal.Format.*")

            ## If we actually had output, and it was interesting information,
            ## save the output into the $ll variable
            if(($cachedOutput.Count -gt 0) -and $containsInterestingTypes)
            {
                $GLOBAL:ll = $cachedOutput | % { $_ }
            }
        }
}

function Add-RelativePathCapture
{
    ##############################################################################
    ##
    ## Add-RelativePathCapture
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Adds a new CommandNotFound handler that captures relative path
    navigation without having to explicitly call 'Set-Location'

    .EXAMPLE

    PS C:\Users\Lee\Documents>..
    PS C:\Users\Lee>...
    PS C:\>

    #>

    Set-StrictMode -Version 3

    $executionContext.SessionState.InvokeCommand.CommandNotFoundAction = {
        param($CommandName, $CommandLookupEventArgs)

        ## If the command is only dots
        if($CommandName -match '^\.+$')
        {
            ## Associate a new command that should be invoked instead
            $CommandLookupEventArgs.CommandScriptBlock = {

                ## Count the number of dots, and run "Set-Location .." one
                ## less time.
                for($counter = 0; $counter -lt $CommandName.Length - 1; $counter++)
                {
                    Set-Location ..
                }

            ## We call GetNewClosure() so that the reference to $CommandName can
            ## be used in the new command.
            }.GetNewClosure()

            ## Stop going through the command resolution process. This isn't
            ## strictly required in the CommandNotFoundAction.
            $CommandLookupEventArgs.StopSearch = $true
        }
    }
}

function Compare-Property
{
    ##############################################################################
    ##
    ## Compare-Property
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Compare the property you provide against the input supplied to the script.
    This provides the functionality of simple Where-Object comparisons without
    the syntax required for that cmdlet.

    .EXAMPLE

    PS Get-Process | Compare-Property Handles gt 1000

    .EXAMPLE

    PS > Set-Alias ?? Compare-Property
    PS > dir | ?? PsIsContainer

    #>

    param(
        ## The property to compare
        $Property,

        ## The operator to use in the comparison
        $Operator = "eq",

        ## The value to compare with
        $MatchText = "$true"
    )

    Begin { $expression = "`$_.$property -$operator `"$matchText`"" }
    Process { if(Invoke-Expression $expression) { $_ } }
}

function Connect-WebService
{
    ##############################################################################
    ##
    ## Connect-WebService
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ## Connect to a given web service, and create a type that allows you to
    ## interact with that web service. In PowerShell version two, use the
    ## New-WebserviceProxy cmdlet.
    ##
    ## Example:
    ##
    ## $wsdl = "http://www.terraserver-usa.com/TerraService2.asmx?WSDL"
    ## $terraServer = Connect-WebService $wsdl
    ## $place = New-Object Place
    ## $place.City = "Redmond"
    ## $place.State = "WA"
    ## $place.Country = "USA"
    ## $facts = $terraserver.GetPlaceFacts($place)
    ## $facts.Center
    ##
    ##############################################################################

    param(
        ## The URL that contains the WSDL
        [string] $WsdlLocation = $(throw "Please specify a WSDL location"),

        ## The namespace to use to contain the web service proxy
        [string] $Namespace,

        ## Switch to identify web services that require authentication
        [Switch] $RequiresAuthentication
    )

    ## Create the web service cache, if it doesn't already exist
    if(-not (Test-Path Variable:\Lee.Holmes.WebServiceCache))
    {
        ${GLOBAL:Lee.Holmes.WebServiceCache} = @{}
    }

    ## Check if there was an instance from a previous connection to
    ## this web service. If so, return that instead.
    $oldInstance = ${GLOBAL:Lee.Holmes.WebServiceCache}[$wsdlLocation]
    if($oldInstance)
    {
        $oldInstance
        return
    }

    ## Load the required Web Services DLL
    $null = [Reflection.Assembly]::LoadWithPartialName("System.Web.Services")

    ## Download the WSDL for the service, and create a service description from
    ## it.
    $wc = New-Object System.Net.WebClient

    if($requiresAuthentication)
    {
        $wc.UseDefaultCredentials = $true
    }

    $wsdlStream = $wc.OpenRead($wsdlLocation)

    ## Ensure that we were able to fetch the WSDL
    if(-not (Test-Path Variable:\wsdlStream))
    {
        return
    }

    $serviceDescription =
        [Web.Services.Description.ServiceDescription]::Read($wsdlStream)
    $wsdlStream.Close()

    ## Ensure that we were able to read the WSDL into a service description
    if(-not (Test-Path Variable:\serviceDescription))
    {
        return
    }

    ## Import the web service into a CodeDom
    $serviceNamespace = New-Object System.CodeDom.CodeNamespace
    if($namespace)
    {
        $serviceNamespace.Name = $namespace
    }

    $codeCompileUnit = New-Object System.CodeDom.CodeCompileUnit
    $serviceDescriptionImporter =
        New-Object Web.Services.Description.ServiceDescriptionImporter
    $serviceDescriptionImporter.AddServiceDescription(
        $serviceDescription, $null, $null)
    [void] $codeCompileUnit.Namespaces.Add($serviceNamespace)
    [void] $serviceDescriptionImporter.Import(
        $serviceNamespace, $codeCompileUnit)

    ## Generate the code from that CodeDom into a string
    $generatedCode = New-Object Text.StringBuilder
    $stringWriter = New-Object IO.StringWriter $generatedCode
    $provider = New-Object Microsoft.CSharp.CSharpCodeProvider
    $provider.GenerateCodeFromCompileUnit($codeCompileUnit, $stringWriter, $null)

    ## Compile the source code.
    $references = @("System.dll", "System.Web.Services.dll", "System.Xml.dll")
    $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.ReferencedAssemblies.AddRange($references)
    $compilerParameters.GenerateInMemory = $true

    $compilerResults =
        $provider.CompileAssemblyFromSource($compilerParameters, $generatedCode)

    ## Write any errors if generated.
    if($compilerResults.Errors.Count -gt 0)
    {
        $errorLines = ""
        foreach($error in $compilerResults.Errors)
        {
            $errorLines += "`n`t" + $error.Line + ":`t" + $error.ErrorText
        }

        Write-Error $errorLines
        return
    }
    ## There were no errors.  Create the webservice object and return it.
    else
    {
        ## Get the assembly that we just compiled
        $assembly = $compilerResults.CompiledAssembly

        ## Find the type that had the WebServiceBindingAttribute.
        ## There may be other "helper types" in this file, but they will
        ## not have this attribute
        $type = $assembly.GetTypes() |
            Where-Object { $_.GetCustomAttributes(
                [System.Web.Services.WebServiceBindingAttribute], $false) }

        if(-not $type)
        {
            Write-Error "Could not generate web service proxy."
            return
        }

        ## Create an instance of the type, store it in the cache,
        ## and return it to the user.
        $instance = $assembly.CreateInstance($type)

        ## Many services that support authentication also require it on the
        ## resulting objects
        if($requiresAuthentication)
        {
            if(@($instance.PsObject.Properties |
                where { $_.Name -eq "UseDefaultCredentials" }).Count -eq 1)
            {
                $instance.UseDefaultCredentials = $true
            }
        }

        ${GLOBAL:Lee.Holmes.WebServiceCache}[$wsdlLocation] = $instance

        $instance
    }
}

function Convert-TextObject
{
    ##############################################################################
    ##
    ## Convert-TextObject
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Convert a simple string into a custom PowerShell object.

    .EXAMPLE

    PS > "Hello World" | Convert-TextObject
    Generates an Object with "P1=Hello" and "P2=World"

    .EXAMPLE

    PS > "Hello World" | Convert-TextObject -Delimiter "ll"
    Generates an Object with "P1=He" and "P2=o World"

    .EXAMPLE

    PS > "Hello World" | Convert-TextObject -Pattern "He(ll.*o)r(ld)"
    Generates an Object with "P1=llo Wo" and "P2=ld"

    .EXAMPLE

    PS > "Hello World" | Convert-TextObject -PropertyName FirstWord,SecondWord
    Generates an Object with "FirstWord=Hello" and "SecondWord=World

    .EXAMPLE

    PS > "123 456" | Convert-TextObject -PropertyType $([string],[int])
    Generates an Object with "Property1=123" and "Property2=456"
    The second property is an integer, as opposed to a string

    .EXAMPLE

    PS > $ipAddress = (ipconfig | Convert-TextObject -Delim ": ")[2].P2
    PS > $ipAddress
    192.168.1.104

    #>

    [CmdletBinding(DefaultParameterSetName = "ByDelimiter")]
    param(
        ## If specified, gives the .NET Regular Expression with which to
        ## split the string. The script generates properties for the
        ## resulting object out of the elements resulting from this split.
        ## If not specified, defaults to splitting on the maximum amount
        ## of whitespace: "\s+", as long as Pattern is not
        ## specified either.
        [Parameter(ParameterSetName = "ByDelimiter", Position = 0)]
        [string] $Delimiter = "\s+",

        ## If specified, gives the .NET Regular Expression with which to
        ## parse the string. The script generates properties for the
        ## resulting object out of the groups captured by this regular
        ## expression.
        [Parameter(Mandatory = $true,
            ParameterSetName = "ByPattern",
            Position = 0)]
        [string] $Pattern,

        ## If specified, the script will pair the names from this object
        ## definition with the elements from the parsed string.  If not
        ## specified (or the generated object contains more properties
        ## than you specify,) the script uses property names in the
        ## pattern of P1,P2,...,PN
        [Parameter(Position = 1)]
        [Alias("PN")]
        [string[]] $PropertyName = @(),

        ## If specified, the script will pair the types from this list with
        ## the properties from the parsed string.  If not specified (or the
        ## generated object contains more properties than you specify,) the
        ## script sets the properties to be of type [string]
        [Parameter(Position = 2)]
        [Alias("PT")]
        [type[]] $PropertyType = @(),

        ## The input object to process
        [Parameter(ValueFromPipeline = $true)]
        [string] $InputObject
    )

    begin {
        Set-StrictMode -Version 3
    }

    process {
        $returnObject = New-Object PSObject

        $matches = $null
        $matchCount = 0

        if($PSBoundParameters["Pattern"])
        {
            ## Verify that the input contains the pattern
            ## Populates the matches variable by default
            if(-not ($InputObject -match $pattern))
            {
                return
            }

            $matchCount = $matches.Count
        $startIndex = 1
        }
        else
        {
            ## Verify that the input contains the delimiter
            if(-not ($InputObject -match $delimiter))
            {
                return
            }

            ## If so, split the input on that delimiter
            $matches = $InputObject -split $delimiter
            $matchCount = $matches.Length
            $startIndex = 0
        }

        ## Go through all of the matches, and add them as notes to the output
        ## object.
        for($counter = $startIndex; $counter -lt $matchCount; $counter++)
        {
            $currentPropertyName = "P$($counter - $startIndex + 1)"
            $currentPropertyType = [string]

            ## Get the property name
            if($counter -lt $propertyName.Length)
            {
                if($propertyName[$counter])
                {
                    $currentPropertyName = $propertyName[$counter - 1]
                }
            }

            ## Get the property value
            if($counter -lt $propertyType.Length)
            {
                if($propertyType[$counter])
                {
                    $currentPropertyType = $propertyType[$counter - 1]
                }
            }

            Add-Member -InputObject $returnObject NoteProperty `
                -Name $currentPropertyName `
                -Value ($matches[$counter].Trim() -as $currentPropertyType)
        }

        $returnObject
    }
}

function ConvertFrom-FahrenheitWithFunction
{
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)

    param([double] $Fahrenheit)

    Set-StrictMode -Version 3

    ## Convert Fahrenheit to Celsius
    function ConvertFahrenheitToCelsius([double] $fahrenheit)
    {
        $celsius = $fahrenheit - 32
        $celsius = $celsius / 1.8
        $celsius
    }

    $celsius = ConvertFahrenheitToCelsius $fahrenheit

    ## Output the answer
    "$fahrenheit degrees Fahrenheit is $celsius degrees Celsius."
}

function ConvertFrom-FahrenheitWithoutFunction
{
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)

    param([double] $Fahrenheit)

    Set-StrictMode -Version 3

    ## Convert it to Celsius
    $celsius = $fahrenheit - 32
    $celsius = $celsius / 1.8

    ## Output the answer
    "$fahrenheit degrees Fahrenheit is $celsius degrees Celsius."
}

function Copy-History
{
    ##############################################################################
    ##
    ## Copy-History
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Copy selected commands from the history buffer into the clipboard as a script.

    .EXAMPLE

    PS > Copy-History
    Copies the entire contents of the history buffer into the clipboard.

    .EXAMPLE

    PS > Copy-History -5
    Copies the last five commands into the clipboard.

    .EXAMPLE

    PS > Copy-History 2,5,8,4
    Copies commands 2,5,8, and 4.

    .EXAMPLE

    PS > Copy-History (1..10+5+6)
    Copies commands 1 through 10, then 5, then 6, using PowerShell's array
    slicing syntax.

    #>

    param(
        ## The range of history IDs to copy
        [int[]] $Range
    )

    Set-StrictMode -Version 3

    $history = @()

    ## If they haven't specified a range, assume it's everything
    if((-not $range) -or ($range.Count -eq 0))
    {
        $history = @(Get-History -Count ([Int16]::MaxValue))
    }
    ## If it's a negative number, copy only that many
    elseif(($range.Count -eq 1) -and ($range[0] -lt 0))
    {
        $count = [Math]::Abs($range[0])
        $history = (Get-History -Count $count)
    }
    ## Otherwise, go through each history ID in the given range
    ## and add it to our history list.
    else
    {
        foreach($commandId in $range)
        {
            if($commandId -eq -1) { $history += Get-History -Count 1 }
            else { $history += Get-History -Id $commandId }
        }
    }

    ## Finally, export the history to the clipboard.
    $history | Foreach-Object { $_.CommandLine } | clip.exe
}

function Enable-BreakOnError
{
    #############################################################################
    ##
    ## Enable-BreakOnError
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Creates a breakpoint that only fires when PowerShell encounters an error

    .EXAMPLE

    PS > Enable-BreakOnError

    ID Script           Line Command         Variable        Action
    -- ------           ---- -------         --------        ------
     0                       Out-Default                     ...

    PS > 1/0
    Entering debug mode. Use h or ? for help.

    Hit Command breakpoint on 'Out-Default'


    PS > $error
    Attempted to divide by zero.

    #>

    Set-StrictMode -Version 3

    ## Store the current number of errors seen in the session so far
    $GLOBAL:EnableBreakOnErrorLastErrorCount = $error.Count

    Set-PSBreakpoint -Command Out-Default -Action {

        ## If we're generating output, and the error count has increased,
        ## break into the debugger.
        if($error.Count -ne $EnableBreakOnErrorLastErrorCount)
        {
            $GLOBAL:EnableBreakOnErrorLastErrorCount = $error.Count
            break
        }
    }
}

function Enable-HistoryPersistence
{
    ##############################################################################
    ##
    ## Enable-HistoryPersistence
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Reloads any previously saved command history, and registers for the
    PowerShell.Exiting engine event to save new history when the shell
    exits.

    #>

    Set-StrictMode -Version 3

    ## Load our previous history
    $GLOBAL:maximumHistoryCount = 32767
    $historyFile = (Join-Path (Split-Path $profile) "commandHistory.clixml")
    if(Test-Path $historyFile)
    {
        Import-CliXml $historyFile | Add-History
    }

    ## Register for the engine shutdown event
    $null = Register-EngineEvent -SourceIdentifier `
        ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {

        ## Save our history
        $historyFile = (Join-Path (Split-Path $profile) "commandHistory.clixml")
        $maximumHistoryCount = 1kb

        ## Get the previous history items
        $oldEntries = @()
        if(Test-Path $historyFile)
        {
            $oldEntries = Import-CliXml $historyFile -ErrorAction SilentlyContinue
        }

        ## And merge them with our changes
        $currentEntries = Get-History -Count $maximumHistoryCount
        $additions = Compare-Object $oldEntries $currentEntries `
            -Property CommandLine | Where-Object { $_.SideIndicator -eq "=>" } |
            Foreach-Object { $_.CommandLine }

        $newEntries = $currentEntries | ? { $additions -contains $_.CommandLine }

        ## Keep only unique command lines. First sort by CommandLine in
        ## descending order (so that we keep the newest entries,) and then
        ## re-sort by StartExecutionTime.
        $history = @($oldEntries + $newEntries) |
            Sort -Unique -Descending CommandLine | Sort StartExecutionTime

        ## Finally, keep the last 100
        Remove-Item $historyFile
        $history | Select -Last 100 | Export-CliXml $historyFile
    }
}

function Enable-RemoteCredSSP
{
    ##############################################################################
    ##
    ## Enable-RemoteCredSSP
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Enables CredSSP support on a remote computer. Requires that the machine
    have PowerShell Remoting enabled, and that its operating system is Windows
    Vista or later.

    .EXAMPLE

    PS > Enable-RemoteCredSSP <Computer>

    #>

    param(
        ## The computer on which to enable CredSSP
        $Computername,

        ## The credential to use when connecting
        $Credential = (Get-Credential)
    )

    Set-StrictMode -Version 3

    ## Call Get-Credential again, so that the user can type something like
    ## Enable-RemoteCredSSP -Computer Computer -Cred DOMAIN\user
    $credential = Get-Credential $credential
    $username = $credential.Username
    $password = $credential.GetNetworkCredential().Password

    ## Define the script we will use to create the scheduled task
    $powerShellCommand =
        "powershell -noprofile -command Enable-WsManCredSSP -Role Server -Force"
    $script = @"
    schtasks /CREATE /TN 'Enable CredSSP' /SC WEEKLY /RL HIGHEST ``
        /RU $username /RP $password ``
        /TR "$powerShellCommand" /F

    schtasks /RUN /TN 'Enable CredSSP'
"@

    ## Create the task on the remote system to configure CredSSP
    $command = [ScriptBlock]::Create($script)
    Invoke-Command $computername $command -Cred $credential

    ## Wait for the remoting changes to come into effect
    for($count = 1; $count -le 10; $count++)
    {
        $output =
            Invoke-Command $computername { 1 } -Auth CredSSP -Cred $credential
        if($output -eq 1) { break; }

        "Attempt $count : Not ready yet."
        Sleep 5
    }

    ## Clean up
    $command = [ScriptBlock]::Create($script)
    Invoke-Command $computername {
        schtasks /DELETE /TN 'Enable CredSSP' /F } -Cred $credential

    ## Verify the output
    Invoke-Command $computername {
        Get-WmiObject Win32_ComputerSystem } -Auth CredSSP -Cred $credential
}

function Enable-RemotePsRemoting
{
    ##############################################################################
    ##
    ## Enable-RemotePsRemoting
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Enables PowerShell Remoting on a remote computer. Requires that the machine
    responds to WMI requests, and that its operating system is Windows Vista or
    later.

    .EXAMPLE

    PS > Enable-RemotePsRemoting <Computer>

    #>

    param(
        ## The computer on which to enable remoting
        $Computername,

        ## The credential to use when connecting
        $Credential = (Get-Credential)
    )

    Set-StrictMode -Version 3

    $VerbosePreference = "Continue"

    $credential = Get-Credential $credential
    $username = $credential.Username
    $password = $credential.GetNetworkCredential().Password

    $script = @"

    `$log = Join-Path `$env:TEMP Enable-RemotePsRemoting.output.txt
    Remove-Item -Force `$log -ErrorAction SilentlyContinue
    Start-Transcript -Path `$log

    ## Create a task that will run with full network privileges.
    ## In this task, we call Enable-PsRemoting
    schtasks /CREATE /TN 'Enable Remoting' /SC WEEKLY /RL HIGHEST ``
        /RU $username /RP $password ``
        /TR "powershell -noprofile -command Enable-PsRemoting -Force" /F |
        Out-String
    schtasks /RUN /TN 'Enable Remoting' | Out-String

    `$securePass = ConvertTo-SecureString $password -AsPlainText -Force
    `$credential =
        New-Object Management.Automation.PsCredential $username,`$securepass

    ## Wait for the remoting changes to come into effect
    for(`$count = 1; `$count -le 10; `$count++)
    {
        `$output = Invoke-Command localhost { 1 } -Cred `$credential ``
            -ErrorAction SilentlyContinue
        if(`$output -eq 1) { break; }

        "Attempt `$count : Not ready yet."
        Sleep 5
    }

    ## Delete the temporary task
    schtasks /DELETE /TN 'Enable Remoting' /F | Out-String
    Stop-Transcript

"@

    $commandBytes = [System.Text.Encoding]::Unicode.GetBytes($script)
    $encoded = [Convert]::ToBase64String($commandBytes)

    Write-Verbose "Configuring $computername"
    $command = "powershell -NoProfile -EncodedCommand $encoded"
    $null = Invoke-WmiMethod -Computer $computername -Credential $credential `
        Win32_Process Create -Args $command

    Write-Verbose "Testing connection"
    Invoke-Command $computername {
        Get-WmiObject Win32_ComputerSystem } -Credential $credential
}

function Enter-Module
{
    ##############################################################################
    ##
    ## Enter-Module
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Lets you examine internal module state and functions by executing user
    input in the scope of the supplied module.

    .EXAMPLE

    PS > Import-Module PersistentState
    PS > Get-Module PersistentState

    ModuleType Name                      ExportedCommands
    ---------- ----                      ----------------
    Script     PersistentState           {Set-Memory, Get-Memory}


    PS > "Hello World" | Set-Memory
    PS > $m = Get-Module PersistentState
    PS > Enter-Module $m
    PersistentState: dir variable:\mem*

    Name                           Value
    ----                           -----
    memory                         {Hello World}

    PersistentState: exit
    PS >

    #>

    param(
        ## The module to examine
        [System.Management.Automation.PSModuleInfo] $Module
    )

    Set-StrictMode -Version 3

    $userInput = Read-Host $($module.Name)
    while($userInput -ne "exit")
    {
        $scriptblock = [ScriptBlock]::Create($userInput)
        & $module $scriptblock

        $userInput = Read-Host $($module.Name)
    }
}

function Format-Hex
{
    ##############################################################################
    ##
    ## Format-Hex
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Outputs a file or pipelined input as a hexadecimal display. To determine the
    offset of a character in the input, add the number at the far-left of the row
    with the the number at the top of the column for that character.

    .EXAMPLE

    PS > "Hello World" | Format-Hex

                0  1  2  3  4  5  6  7  8  9  A  B  C  D  E  F

    00000000   48 00 65 00 6C 00 6C 00 6F 00 20 00 57 00 6F 00  H.e.l.l.o. .W.o.
    00000010   72 00 6C 00 64 00                                r.l.d.

    .EXAMPLE

    PS > Format-Hex c:\temp\example.bmp

    #>

    [CmdletBinding(DefaultParameterSetName = "ByPath")]
    param(
        ## The file to read the content from
        [Parameter(ParameterSetName = "ByPath", Position = 0)]
        [string] $Path,

        ## The input (bytes or strings) to format as hexadecimal
        [Parameter(
            ParameterSetName = "ByInput", Position = 0,
            ValueFromPipeline = $true)]
        [Object] $InputObject
    )

    begin
    {
        Set-StrictMode -Version 3

        ## Create the array to hold the content. If the user specified the
        ## -Path parameter, read the bytes from the path.
        [byte[]] $inputBytes = $null
        if($Path) { $inputBytes = Get-Content $Path -Encoding Byte -Raw }

        ## Store our header, and formatting information
        $counter = 0
        $header = "            0  1  2  3  4  5  6  7  8  9  A  B  C  D  E  F"
        $nextLine = "{0}   " -f  [Convert]::ToString(
            $counter, 16).ToUpper().PadLeft(8, '0')
        $asciiEnd = ""

        ## Output the header
        "`r`n$header`r`n"
    }

    process
    {
        ## If they specified the -InputObject parameter, retrieve the bytes
        ## from that input
        if($PSCmdlet.ParameterSetName -eq "ByInput")
        {
            ## If it's an actual byte, add it to the inputBytes array.
            if($InputObject -is [Byte])
            {
                $inputBytes = $InputObject
            }
            else
            {
                ## Otherwise, convert it to a string and extract the bytes
                ## from that.
                $inputString = [string] $InputObject
                $inputBytes = [Text.Encoding]::Unicode.GetBytes($inputString)
            }
        }

        ## Now go through the input bytes
        foreach($byte in $inputBytes)
        {
            ## Display each byte, in 2-digit hexidecimal, and add that to the
            ## left-hand side.
            $nextLine += "{0:X2} " -f $byte

            ## If the character is printable, add its ascii representation to
            ## the right-hand side.  Otherwise, add a dot to the right hand side.
            if(($byte -ge 0x20) -and ($byte -le 0xFE))
            {
                $asciiEnd += [char] $byte
            }
            else
            {
                $asciiEnd += "."
            }

            $counter++;

            ## If we've hit the end of a line, combine the right half with the
            ## left half, and start a new line.
            if(($counter % 16) -eq 0)
            {

                "$nextLine $asciiEnd"
                $nextLine = "{0}   " -f [Convert]::ToString(
                    $counter, 16).ToUpper().PadLeft(8, '0')
                $asciiEnd = "";
            }
        }
    }

    end
    {
        ## At the end of the file, we might not have had the chance to output
        ## the end of the line yet.  Only do this if we didn't exit on the 16-byte
        ## boundary, though.
        if(($counter % 16) -ne 0)
        {
            while(($counter % 16) -ne 0)
            {
                $nextLine += "   "
                $asciiEnd += " "
                $counter++;
            }
            "$nextLine $asciiEnd"
        }

        ""
    }
}

function Format-String
{
    ##############################################################################
    ##
    ## Format-String
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Replaces text in a string based on named replacement tags

    .EXAMPLE

    PS > Format-String "Hello {NAME}" @{ NAME = 'PowerShell' }
    Hello PowerShell

    .EXAMPLE

    PS > Format-String "Your score is {SCORE:P}" @{ SCORE = 0.85 }
    Your score is 85.00 %

    #>

    param(
        ## The string to format. Any portions in the form of {NAME}
        ## will be automatically replaced by the corresponding value
        ## from  the supplied hashtable.
        $String,

        ## The named replacements to use in the string
        [hashtable] $Replacements
    )

    Set-StrictMode -Version 3

    $currentIndex = 0
    $replacementList = @()

    if($String -match "{{|}}")
    {
        throw "Escaping of replacement terms are not supported."
    }

    ## Go through each key in the hashtable
    foreach($key in $replacements.Keys)
    {
        ## Convert the key into a number, so that it can be used by
        ## String.Format
        $inputPattern = '{(.*)' + $key + '(.*)}'
        $replacementPattern = '{${1}' + $currentIndex + '${2}}'
        $string = $string -replace $inputPattern,$replacementPattern
        $replacementList += $replacements[$key]

        $currentIndex++
    }

    ## Now use String.Format to replace the numbers in the
    ## format string.
    $string -f $replacementList
}

function Get-AclMisconfiguration
{
    ##############################################################################
    ##
    ## Get-AclMisconfiguration
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstration of functionality exposed by the Get-Acl cmdlet. This script
    goes through all access rules in all files in the current directory, and
    ensures that the Administrator group has full control of that file.

    #>

    Set-StrictMode -Version 3

    ## Get all files in the current directory
    foreach($file in Get-ChildItem)
    {
        ## Retrieve the ACL from the current file
        $acl = Get-Acl $file
        if(-not $acl)
        {
            continue
        }

        $foundAdministratorAcl = $false

        ## Go through each access rule in that ACL
        foreach($accessRule in $acl.Access)
        {
            ## If we find the Administrator, Full Control access rule,
            ## then set the $foundAdministratorAcl variable
            if(($accessRule.IdentityReference -like "*Administrator*") -and
                ($accessRule.FileSystemRights -eq "FullControl"))
            {
                $foundAdministratorAcl = $true
            }
        }

        ## If we didn't find the administrator ACL, output a message
        if(-not $foundAdministratorAcl)
        {
            "Found possible ACL Misconfiguration: $file"
        }
    }
}

function Get-AliasSuggestion
{
    ##############################################################################
    ##
    ## Get-AliasSuggestion
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get an alias suggestion from the full text of the last command. Intended to
    be added to your prompt function to help learn aliases for commands.

    .EXAMPLE

    PS > Get-AliasSuggestion Remove-ItemProperty
    Suggestion: An alias for Remove-ItemProperty is rp

    #>

    param(
        ## The full text of the last command
        $LastCommand
    )

    Set-StrictMode -Version 3

    $helpMatches = @()

    ## Find all of the commands in their last input
    $tokens = [Management.Automation.PSParser]::Tokenize(
        $lastCommand, [ref] $null)
    $commands = $tokens | Where-Object { $_.Type -eq "Command" }

    ## Go through each command
    foreach($command in $commands)
    {
        ## Get the alias suggestions
        foreach($alias in Get-Alias -Definition $command.Content)
        {
            $helpMatches += "Suggestion: An alias for " +
                "$($alias.Definition) is $($alias.Name)"
        }
    }

    $helpMatches
}

function Get-Answer
{
    ##############################################################################
    ##
    ## Get-Answer
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Uses Bing Answers to answer your question

    .EXAMPLE

    PS > Get-Answer "sqrt(2)"
    sqrt(2) = 1.41421356

    .EXAMPLE

    PS > Get-Answer msft stock
    Microsoft Corp (US:MSFT) NASDAQ
    29.66  -0.35 (-1.17%)
    After Hours: 30.02 +0.36 (1.21%)
    Open: 30.09    Day's Range: 29.59 - 30.20
    Volume: 55.60 M    52 Week Range: 17.27 - 31.50
    P/E Ratio: 16.30    Market Cap: 260.13 B

    .EXAMPLE

    PS > Get-Answer "What is the time in Seattle, WA?"
    Current time in Seattle, WA
    01:12:41 PM
    08/18/2012 ? Pacific Daylight Time

    #>

    Set-StrictMode -Version 3

    $question = $args -join " "

    function Main
    {
        ## Load the System.Web.HttpUtility DLL, to let us URLEncode
        Add-Type -Assembly System.Web

        ## Get the web page into a single string with newlines between
        ## the lines.
        $encoded = [System.Web.HttpUtility]::UrlEncode($question)
        $url = "http://www.bing.com/search?q=$encoded"
        $text = [String] (Invoke-WebRequest $url)

        ## Find the start of the answers section
        $startIndex = $text.IndexOf('<div class="ans"')

        ## The end is either defined by an "attribution" div
        ## or the start of a "results" div
        $endIndex = $text.IndexOf('<div class="sn_att2"', $startIndex)
        if($endIndex -lt 0) { $endIndex = $text.IndexOf('<div id="results"', $startIndex) }

        ## If we found a result, then filter the result
        if(($startIndex -ge 0) -and ($endIndex -ge 0))
        {
            ## Pull out the text between the start and end portions
            $partialText = $text.Substring($startIndex, $endIndex - $startIndex)

            ## Very fragile screen scraping here. Replace a bunch of
            ## tags that get placed on new lines with the newline
            ## character, and a few others with spaces.
            $partialText = $partialText -replace '<div[^>]*>',"`n"
            $partialText = $partialText -replace '<tr[^>]*>',"`n"
            $partialText = $partialText -replace '<li[^>]*>',"`n"
            $partialText = $partialText -replace '<br[^>]*>',"`n"
            $partialText = $partialText -replace '<p [^>]*>',"`n"
            $partialText = $partialText -replace '<span[^>]*>'," "
            $partialText = $partialText -replace '<td[^>]*>',"    "

            $partialText = CleanHtml $partialText

            ## Now split the results on newlines, trim each line, and then
            ## join them back.
            $partialText = $partialText -split "`n" |
                Foreach-Object { $_.Trim() } | Where-Object { $_ }
            $partialText = $partialText -join "`n"

            [System.Web.HttpUtility]::HtmlDecode($partialText.Trim())
        }
        else
        {
            "No answer found."
        }
    }

    ## Clean HTML from a text chunk
    function CleanHtml ($htmlInput)
    {
        $tempString = [Regex]::Replace($htmlInput, "(?s)<[^>]*>", "")
        $tempString.Replace("&nbsp&nbsp", "")
    }

    Main
}

function Get-Arguments
{
    ##############################################################################
    ##
    ## Get-Arguments
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Uses command-line arguments

    #>

    param(
        ## The first named argument
        $FirstNamedArgument,

        ## The second named argument
        [int] $SecondNamedArgument = 0
    )

    Set-StrictMode -Version 3

    ## Display the arguments by name
    "First named argument is: $firstNamedArgument"
    "Second named argument is: $secondNamedArgument"

    function GetArgumentsFunction
    {
        ## We could use a param statement here, as well
        ## param($firstNamedArgument, [int] $secondNamedArgument = 0)

        ## Display the arguments by position
        "First positional function argument is: " + $args[0]
        "Second positional function argument is: " + $args[1]
    }

    GetArgumentsFunction One Two

    $scriptBlock =
    {
        param($firstNamedArgument, [int] $secondNamedArgument = 0)

        ## We could use $args here, as well
        "First named scriptblock argument is: $firstNamedArgument"
        "Second named scriptblock argument is: $secondNamedArgument"
    }

    & $scriptBlock -First One -Second 4.5
}

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

function Get-Clipboard
{
    #############################################################################
    ##
    ## Get-Clipboard
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Retrieve the text contents of the Windows Clipboard.

    .EXAMPLE

    PS > Get-Clipboard
    Hello World

    #>

    Set-StrictMode -Version 3

    Add-Type -Assembly PresentationCore
    [Windows.Clipboard]::GetText()
}

function Get-DetailedSystemInformation
{
    ##############################################################################
    ##
    ## Get-DetailedSystemInformation
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get detailed information about a system.

    .EXAMPLE

    PS > Get-DetailedSystemInformation LEE-DESK > output.txt
    Gets detailed information about LEE-DESK and stores the output into output.txt

    #>

    param(
        ## The computer to analyze
        $Computer = "."
    )

    Set-StrictMode -Version 3

    "#"*80
    "System Information Summary"
    "Generated $(Get-Date)"
    "#"*80
    ""
    ""

    "#"*80
    "Computer System Information"
    "#"*80
    Get-CimInstance Win32_ComputerSystem -Computer $computer | Format-List *

    "#"*80
    "Operating System Information"
    "#"*80
    Get-CimInstance Win32_OperatingSystem -Computer $computer | Format-List *

    "#"*80
    "BIOS Information"
    "#"*80
    Get-CimInstance Win32_Bios -Computer $computer | Format-List *

    "#"*80
    "Memory Information"
    "#"*80
    Get-CimInstance Win32_PhysicalMemory -Computer $computer | Format-List *

    "#"*80
    "Physical Disk Information"
    "#"*80
    Get-CimInstance Win32_DiskDrive -Computer $computer | Format-List *

    "#"*80
    "Logical Disk Information"
    "#"*80
    Get-CimInstance Win32_LogicalDisk -Computer $computer | Format-List *
}

function Get-DiskUsage
{
    ##############################################################################
    ##
    ## Get-DiskUsage
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Retrieve information about disk usage in the current directory and all
    subdirectories. If you specify the -IncludeSubdirectories flag, this
    script accounts for the size of subdirectories in the size of a directory.

    .EXAMPLE

    PS > Get-DiskUsage
    Gets the disk usage for the current directory.

    .EXAMPLE

    PS > Get-DiskUsage -IncludeSubdirectories
    Gets the disk usage for the current directory and those below it,
    adding the size of child directories to the directory that contains them.

    #>

    param(
        ## Switch to include subdirectories in the size of each directory
        [switch] $IncludeSubdirectories
    )

    Set-StrictMode -Version 3

    ## If they specify the -IncludeSubdirectories flag, then we want to account
    ## for all subdirectories in the size of each directory
    if($includeSubdirectories)
    {
        Get-ChildItem -Directory |
            Select-Object Name,
                @{ Name="Size";
                Expression={ ($_ | Get-ChildItem -Recurse |
                    Measure-Object -Sum Length).Sum + 0 } }
    }
    ## Otherwise, we just find all directories below the current directory,
    ## and determine their size
    else
    {
        Get-ChildItem -Recurse -Directory |
            Select-Object FullName,
                @{ Name="Size";
                Expression={ ($_ | Get-ChildItem |
                    Measure-Object -Sum Length).Sum + 0 } }
    }
}

function Get-FacebookNotification
{
    $cred = Get-Credential
    $login = Invoke-WebRequest http://www.facebook.com/login.php -SessionVariable fb
    $login.Forms[0].Fields.email = $cred.GetNetworkCredential().UserName
    $login.Forms[0].Fields.pass = $cred.GetNetworkCredential().Password
    $mainPage = Invoke-WebRequest $login.Forms[0].Action -WebSession $fb -Body $login -Method Post
    $mainPage.ParsedHtml.getElementById("notificationsCountValue").InnerText
}

function Get-FileEncoding
{
    ##############################################################################
    ##
    ## Get-FileEncoding
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

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

    Set-StrictMode -Version 3

    ## First, check if the file is binary. That is, if the first
    ## 5 lines contain any non-printable characters.
    $nonPrintable = [char[]] (0..8 + 10..31 + 127 + 129 + 141 + 143 + 144 + 157)
    $lines = Get-Content $Path -ErrorAction Ignore -TotalCount 5
    $result = @($lines | Where-Object { $_.IndexOfAny($nonPrintable) -ge 0 })
    if($result.Count -gt 0)
    {
        "Binary"
        return
    }

    ## Next, check if it matches a well-known encoding.

    ## The hashtable used to store our mapping of encoding bytes to their
    ## name. For example, "255-254 = Unicode"
    $encodings = @{}

    ## Find all of the encodings understood by the .NET Framework. For each,
    ## determine the bytes at the start of the file (the preamble) that the .NET
    ## Framework uses to identify that encoding.
    foreach($encoding in [System.Text.Encoding]::GetEncodings())
    {
        $preamble = $encoding.GetEncoding().GetPreamble()
        if($preamble)
        {
            $encodingBytes = $preamble -join '-'
            $encodings[$encodingBytes] = $encoding.GetEncoding()
        }
    }

    ## Find out the lengths of all of the preambles.
    $encodingLengths = $encodings.Keys | Where-Object { $_ } |
        Foreach-Object { ($_ -split "-").Count }

    ## Assume the encoding is UTF7 by default
    $result = [System.Text.Encoding]::UTF7

    ## Go through each of the possible preamble lengths, read that many
    ## bytes from the file, and then see if it matches one of the encodings
    ## we know about.
    foreach($encodingLength in $encodingLengths | Sort -Descending)
    {
        $bytes = Get-Content -encoding byte -readcount $encodingLength $path | Select -First 1
        $encoding = $encodings[$bytes -join '-']

        ## If we found an encoding that had the same preamble bytes,
        ## save that output and break.
        if($encoding)
        {
            $result = $encoding
            break
        }
    }

    ## Finally, output the encoding.
    $result
}

function Get-InstalledSoftware
{
    ##############################################################################
    ##
    ## Get-InstalledSoftware
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Lists installed software on the current computer.

    .EXAMPLE

    PS > Get-InstalledSoftware *Frame* | Select DisplayName

    DisplayName
    -----------
    Microsoft .NET Framework 3.5 SP1
    Microsoft .NET Framework 3.5 SP1
    Hotfix for Microsoft .NET Framework 3.5 SP1 (KB953595)
    Hotfix for Microsoft .NET Framework 3.5 SP1 (KB958484)
    Update for Microsoft .NET Framework 3.5 SP1 (KB963707)

    #>

    param(
        ## The name of the software to search for
        $DisplayName = "*"
    )

    Set-StrictMode -Off

    ## Get all the listed software in the Uninstall key
    $keys =
        Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall

    ## Get all of the properties from those items
    $items = $keys | Foreach-Object { Get-ItemProperty $_.PsPath }

    ## For each of those items, display the DisplayName and Publisher
    foreach($item in $items)
    {
        if(($item.DisplayName) -and ($item.DisplayName -like $displayName))
        {
            $item
        }
    }
}

function Get-InvocationInfo
{
    ##############################################################################
    ##
    ## Get-InvocationInfo
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Display the information provided by the $myInvocation variable

    #>

    param(
        ## Switch to no longer recursively call ourselves
        [switch] $PreventExpansion
    )

    Set-StrictMode -Version 3

    ## Define a helper function, so that we can see how $myInvocation changes
    ## when it is called, and when it is dot-sourced
    function HelperFunction
    {
        "    MyInvocation from function:"
        "-"*50
        $myInvocation

        "    Command from function:"
        "-"*50
        $myInvocation.MyCommand
    }

    ## Define a script block, so that we can see how $myInvocation changes
    ## when it is called, and when it is dot-sourced
    $myScriptBlock = {
        "    MyInvocation from script block:"
        "-"*50
        $myInvocation

        "    Command from script block:"
        "-"*50
        $myInvocation.MyCommand
    }

    ## Define a helper alias
    Set-Alias gii .\Get-InvocationInfo

    ## Illustrate how $myInvocation.Line returns the entire line that the
    ## user typed.
    "You invoked this script by typing: " + $myInvocation.Line

    ## Show the information that $myInvocation returns from a script
    "MyInvocation from script:"
    "-"*50
    $myInvocation

    "Command from script:"
    "-"*50
    $myInvocation.MyCommand

    ## If we were called with the -PreventExpansion switch, don't go
    ## any further
    if($preventExpansion)
    {
        return
    }

    ## Show the information that $myInvocation returns from a function
    "Calling HelperFunction"
    "-"*50
    HelperFunction

    ## Show the information that $myInvocation returns from a dot-sourced
    ## function
    "Dot-Sourcing HelperFunction"
    "-"*50
    . HelperFunction

    ## Show the information that $myInvocation returns from an aliased script
    "Calling aliased script"
    "-"*50
    gii -PreventExpansion

    ## Show the information that $myInvocation returns from a script block
    "Calling script block"
    "-"*50
    & $myScriptBlock

    ## Show the information that $myInvocation returns from a dot-sourced
    ## script block
    "Dot-Sourcing script block"
    "-"*50
    . $myScriptBlock

    ## Show the information that $myInvocation returns from an aliased script
    "Calling aliased script"
    "-"*50
    gii -PreventExpansion
}

function Get-MachineStartupShutdownScript
{
    ##############################################################################
    ##
    ## Get-MachineStartupShutdownScript
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get the startup or shutdown scripts assigned to a machine

    .EXAMPLE

    PS > Get-MachineStartupShutdownScript -ScriptType Startup
    Gets startup scripts for the machine

    #>

    param(
        ## The type of script to search for: Startup, or Shutdown.
        [Parameter(Mandatory = $true)]
        [ValidateSet("Startup","Shutdown")]
        $ScriptType
    )

    Set-StrictMode -Version 3

    ## Store the location of the group policy scripts for the machine
    $registryKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System\Scripts"

    ## There may be no scripts defined
    if(-not (Test-Path $registryKey))
    {
        return
    }

    ## Go through each of the policies in the specified key
    foreach($policy in Get-ChildItem $registryKey\$scriptType)
    {
        ## For each of the scripts in that policy, get its script name
        ## and parameters
        foreach($script in Get-ChildItem $policy.PsPath)
        {
            Get-ItemProperty $script.PsPath | Select Script,Parameters
        }
    }
}

function Get-OfficialTime
{
    ##############################################################################
    ##
    ## Get-OfficialTime
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Gets the official US time (PST) from time.gov

    #>

    Set-StrictMode -Version 3

    ## Create the URL that contains the Twitter search results
    Add-Type -Assembly System.Web
    $url = 'http://www.time.gov/timezone.cgi?Pacific/d/-8'

    ## Download the web page
    $results = Invoke-WebRequest $url | Foreach-Object Content

    ## Extract the text of the time, which is contained in
    ## a segment that looks like "<font size="7" color="white"><b>...<br>"
    $match = $results -match '<font [^>]*><b>(.*)<br>'
    if($matches)
    {
        $time = $matches[1]
    }

    ## Output the time
    $time
}

function Get-OperatingSystemSku
{
    ##############################################################################
    ##
    ## Get-OperatingSystemSku
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Gets the sku information for the current operating system

    .EXAMPLE

    PS > Get-OperatingSystemSku
    Professional with Media Center

    #>

    param($Sku =
        (Get-CimInstance Win32_OperatingSystem).OperatingSystemSku)

    Set-StrictMode -Version 3

    switch ($Sku)
    {
        0   { "An unknown product"; break; }
        1   { "Ultimate"; break; }
        2   { "Home Basic"; break; }
        3   { "Home Premium"; break; }
        4   { "Enterprise"; break; }
        5   { "Home Basic N"; break; }
        6   { "Business"; break; }
        7   { "Server Standard"; break; }
        8   { "Server Datacenter (full installation)"; break; }
        9   { "Windows Small Business Server"; break; }
        10  { "Server Enterprise (full installation)"; break; }
        11  { "Starter"; break; }
        12  { "Server Datacenter (core installation)"; break; }
        13  { "Server Standard (core installation)"; break; }
        14  { "Server Enterprise (core installation)"; break; }
        15  { "Server Enterprise for Itanium-based Systems"; break; }
        16  { "Business N"; break; }
        17  { "Web Server (full installation)"; break; }
        18  { "HPC Edition"; break; }
        19  { "Windows Storage Server 2008 R2 Essentials"; break; }
        20  { "Storage Server Express"; break; }
        21  { "Storage Server Standard"; break; }
        22  { "Storage Server Workgroup"; break; }
        23  { "Storage Server Enterprise"; break; }
        24  { "Windows Server 2008 for Windows Essential Server Solutions"; break; }
        25  { "Small Business Server Premium"; break; }
        26  { "Home Premium N"; break; }
        27  { "Enterprise N"; break; }
        28  { "Ultimate N"; break; }
        29  { "Web Server (core installation)"; break; }
        30  { "Windows Essential Business Server Management Server"; break; }
        31  { "Windows Essential Business Server Security Server"; break; }
        32  { "Windows Essential Business Server Messaging Server"; break; }
        33  { "Server Foundation"; break; }
        34  { "Windows Home Server 2011"; break; }
        35  { "Windows Server 2008 without Hyper-V for Windows Essential Server Solutions"; break; }
        36  { "Server Standard without Hyper-V"; break; }
        37  { "Server Datacenter without Hyper-V (full installation)"; break; }
        38  { "Server Enterprise without Hyper-V (full installation)"; break; }
        39  { "Server Datacenter without Hyper-V (core installation)"; break; }
        40  { "Server Standard without Hyper-V (core installation)"; break; }
        41  { "Server Enterprise without Hyper-V (core installation)"; break; }
        42  { "Microsoft Hyper-V Server"; break; }
        43  { "Storage Server Express (core installation)"; break; }
        44  { "Storage Server Standard (core installation)"; break; }
        45  { "Storage Server Workgroup (core installation)"; break; }
        46  { "Storage Server Enterprise (core installation)"; break; }
        46  { "Storage Server Enterprise (core installation)"; break; }
        47  { "Starter N"; break; }
        48  { "Professional"; break; }
        49  { "Professional N"; break; }
        50  { "Windows Small Business Server 2011 Essentials"; break; }
        51  { "Server For SB Solutions"; break; }
        52  { "Server Solutions Premium"; break; }
        53  { "Server Solutions Premium (core installation)"; break; }
        54  { "Server For SB Solutions EM"; break; }
        55  { "Server For SB Solutions EM"; break; }
        56  { "Windows MultiPoint Server"; break; }
        59  { "Windows Essential Server Solution Management"; break; }
        60  { "Windows Essential Server Solution Additional"; break; }
        61  { "Windows Essential Server Solution Management SVC"; break; }
        62  { "Windows Essential Server Solution Additional SVC"; break; }
        63  { "Small Business Server Premium (core installation)"; break; }
        64  { "Server Hyper Core V"; break; }
        72  { "Server Enterprise (evaluation installation)"; break; }
        76  { "Windows MultiPoint Server Standard (full installation)"; break; }
        77  { "Windows MultiPoint Server Premium (full installation)"; break; }
        79  { "Server Standard (evaluation installation)"; break; }
        80  { "Server Datacenter (evaluation installation)"; break; }
        84  { "Enterprise N (evaluation installation)"; break; }
        95  { "Storage Server Workgroup (evaluation installation)"; break; }
        96  { "Storage Server Standard (evaluation installation)"; break; }
        98  { "Windows 8 N"; break; }
        99  { "Windows 8 China"; break; }
        100 { "Windows 8 Single Language"; break; }
        101 { "Windows 8"; break; }
        103 { "Professional with Media Center"; break; }

        default {"UNKNOWN: " + $SKU }
    }
}

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

function Get-PageUrls
{
    ##############################################################################
    ##
    ## Get-PageUrls
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Parse all of the URLs out of a given file.

    .EXAMPLE

    PS > Get-PageUrls microsoft.html http://www.microsoft.com
    Gets all of the URLs from HTML stored in microsoft.html, and converts relative
    URLs to the domain of http://www.microsoft.com

    .EXAMPLE

    PS > Get-PageUrls microsoft.html http://www.microsoft.com 'aspx$'
    Gets all of the URLs from HTML stored in microsoft.html, converts relative
    URLs to the domain of http://www.microsoft.com, and returns only URLs that end
    in 'aspx'.

    #>

    param(
        ## The filename to parse
        [Parameter(Mandatory = $true)]
        [string] $Path,

        ## The URL from which you downloaded the page.
        ## For example, http://www.microsoft.com
        [Parameter(Mandatory = $true)]
        [string] $BaseUrl,

        [switch] $Images,

        ## The Regular Expression pattern with which to filter
        ## the returned URLs
        [string] $Pattern = ".*"
    )

    Set-StrictMode -Version 3

    ## Load the System.Web DLL so that we can decode URLs
    Add-Type -Assembly System.Web

    ## Defines the regular expression that will parse an URL
    ## out of an anchor tag.
    $regex = "<\s*a\s*[^>]*?href\s*=\s*[`"']*([^`"'>]+)[^>]*?>"
    if($Images)
    {
        $regex = "<\s*img\s*[^>]*?src\s*=\s*[`"']*([^`"'>]+)[^>]*?>"
    }

    ## Parse the file for links
    function Main
    {
        ## Do some minimal source URL fixups, by switching backslashes to
        ## forward slashes
        $baseUrl = $baseUrl.Replace("\", "/")

        if($baseUrl.IndexOf("://") -lt 0)
        {
            throw "Please specify a base URL in the form of " +
                "http://server/path_to_file/file.html"
        }

        ## Determine the server from which the file originated.  This will
        ## help us resolve links such as "/somefile.zip"
        $baseUrl = $baseUrl.Substring(0, $baseUrl.LastIndexOf("/") + 1)
        $baseSlash = $baseUrl.IndexOf("/", $baseUrl.IndexOf("://") + 3)

        if($baseSlash -ge 0)
        {
            $domain = $baseUrl.Substring(0, $baseSlash)
        }
        else
        {
            $domain = $baseUrl
        }


        ## Put all of the file content into a big string, and
        ## get the regular expression matches
        $content = (Get-Content $path) -join ' '
        $contentMatches = @(GetMatches $content $regex)

        foreach($contentMatch in $contentMatches)
        {
            if(-not ($contentMatch -match $pattern)) { continue }
            if($contentMatch -match "javascript:") { continue }

            $contentMatch = $contentMatch.Replace("\", "/")

            ## Hrefs may look like:
            ## ./file
            ## file
            ## ../../../file
            ## /file
            ## url
            ## We'll keep all of the relative paths, as they will resolve.
            ## We only need to resolve the ones pointing to the root.
            if($contentMatch.IndexOf("://") -gt 0)
            {
                $url = $contentMatch
            }
            elseif($contentMatch[0] -eq "/")
            {
                $url = "$domain$contentMatch"
            }
            else
            {
                $url = "$baseUrl$contentMatch"
                $url = $url.Replace("/./", "/")
            }

            ## Return the URL, after first removing any HTML entities
            [System.Web.HttpUtility]::HtmlDecode($url)
        }
    }

    function GetMatches([string] $content, [string] $regex)
    {
        $returnMatches = new-object System.Collections.ArrayList

        ## Match the regular expression against the content, and
        ## add all trimmed matches to our return list
        $resultingMatches = [Regex]::Matches($content, $regex, "IgnoreCase")
        foreach($match in $resultingMatches)
        {
            $cleanedMatch = $match.Groups[1].Value.Trim()
            [void] $returnMatches.Add($cleanedMatch)
        }

        $returnMatches
    }

    . Main
}

function Get-ParameterAlias
{
    ##############################################################################
    ##
    ## Get-ParameterAlias
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Looks in the session history, and returns any aliases that apply to
    parameters of commands that were used.

    .EXAMPLE

    PS > dir -ErrorAction SilentlyContinue
    PS > Get-ParameterAlias
    An alias for the 'ErrorAction' parameter of 'dir' is ea

    #>

    Set-StrictMode -Version 3

    ## Get the last item from their session history
    $history = Get-History -Count 1
    if(-not $history)
    {
        return
    }

    ## And extract the actual command line they typed
    $lastCommand = $history.CommandLine

    ## Use the Tokenizer API to determine which portions represent
    ## commands and parameters to those commands
    $tokens = [System.Management.Automation.PsParser]::Tokenize(
        $lastCommand, [ref] $null)
    $currentCommand = $null

    ## Now go through each resulting token
    foreach($token in $tokens)
    {
        ## If we've found a new command, store that.
        if($token.Type -eq "Command")
        {
            $currentCommand = $token.Content
        }

        ## If we've found a command parameter, start looking for aliases
        if(($token.Type -eq "CommandParameter") -and ($currentCommand))
        {
            ## Remove the leading "-" from the parameter
            $currentParameter = $token.Content.TrimStart("-")

            ## Determine all of the parameters for the current command.
            (Get-Command $currentCommand).Parameters.GetEnumerator() |

                ## For parameters that start with the current parameter name,
                Where-Object { $_.Key -like "$currentParameter*" } |

                ## return all of the aliases that apply. We use "starts with"
                ## because the user might have typed a shortened form of
                ## the parameter name.
                Foreach-Object {
                    $_.Value.Aliases | Foreach-Object {
                        "Suggestion: An alias for the '$currentParameter' " +
                        "parameter of '$currentCommand' is '$_'"
                    }
                }
        }
    }
}

function Get-PrivateProfileString
{
    #############################################################################
    ##
    ## Get-PrivateProfileString
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Retrieves an element from a standard .INI file

    .EXAMPLE

    PS > Get-PrivateProfileString c:\windows\system32\tcpmon.ini `
        "<Generic Network Card>" Name
    Generic Network Card

    #>

    param(
        ## The INI file to retrieve
        $Path,

        ## The section to retrieve from
        $Category,

        ## The item to retrieve
        $Key
    )

    Set-StrictMode -Version 3

    ## The signature of the Windows API that retrieves INI
    ## settings
    $signature = @'
    [DllImport("kernel32.dll")]
    public static extern uint GetPrivateProfileString(
        string lpAppName,
        string lpKeyName,
        string lpDefault,
        StringBuilder lpReturnedString,
        uint nSize,
        string lpFileName);
'@

    ## Create a new type that lets us access the Windows API function
    $type = Add-Type -MemberDefinition $signature `
        -Name Win32Utils -Namespace GetPrivateProfileString `
        -Using System.Text -PassThru

    ## The GetPrivateProfileString function needs a StringBuilder to hold
    ## its output. Create one, and then invoke the method
    $builder = New-Object System.Text.StringBuilder 1024
    $null = $type::GetPrivateProfileString($category,
        $key, "", $builder, $builder.Capacity, $path)

    ## Return the output
    $builder.ToString()
}

function Get-RemoteRegistryChildItem
{
    ##############################################################################
    ##
    ## Get-RemoteRegistryChildItem
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get the list of subkeys below a given key on a remote computer.

    .EXAMPLE

    Get-RemoteRegistryChildItem LEE-DESK HKLM:\Software

    #>

    param(
        ## The computer that you wish to connect to
        [Parameter(Mandatory = $true)]
        $ComputerName,

        ## The path to the registry items to retrieve
        [Parameter(Mandatory = $true)]
        $Path
    )

    Set-StrictMode -Version 3

    ## Validate and extract out the registry key
    if($path -match "^HKLM:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "LocalMachine", $computername)
    }
    elseif($path -match "^HKCU:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "CurrentUser", $computername)
    }
    else
    {
        Write-Error ("Please specify a fully-qualified registry path " +
            "(i.e.: HKLM:\Software) of the registry key to open.")
        return
    }

    ## Open the key
    $key = $baseKey.OpenSubKey($matches[1])

    ## Retrieve all of its children
    foreach($subkeyName in $key.GetSubKeyNames())
    {
        ## Open the subkey
        $subkey = $key.OpenSubKey($subkeyName)

        ## Add information so that PowerShell displays this key like regular
        ## registry key
        $returnObject = [PsObject] $subKey
        $returnObject | Add-Member NoteProperty PsChildName $subkeyName
        $returnObject | Add-Member NoteProperty Property $subkey.GetValueNames()

        ## Output the key
        $returnObject

        ## Close the child key
        $subkey.Close()
    }

    ## Close the key and base keys
    $key.Close()
    $baseKey.Close()
}

function Get-RemoteRegistryKeyProperty
{
    ##############################################################################
    ##
    ## Get-RemoteRegistryKeyProperty
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get the value of a remote registry key property

    .EXAMPLE

    PS > $registryPath =
         "HKLM:\software\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell"
    PS > Get-RemoteRegistryKeyProperty LEE-DESK $registryPath ExecutionPolicy

    #>

    param(
        ## The computer that you wish to connect to
        [Parameter(Mandatory = $true)]
        $ComputerName,

        ## The path to the registry item to retrieve
        [Parameter(Mandatory = $true)]
        $Path,

        ## The specific property to retrieve
        $Property = "*"
    )

    Set-StrictMode -Version 3

    ## Validate and extract out the registry key
    if($path -match "^HKLM:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "LocalMachine", $computername)
    }
    elseif($path -match "^HKCU:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "CurrentUser", $computername)
    }
    else
    {
        Write-Error ("Please specify a fully-qualified registry path " +
            "(i.e.: HKLM:\Software) of the registry key to open.")
        return
    }

    ## Open the key
    $key = $baseKey.OpenSubKey($matches[1])
    $returnObject = New-Object PsObject

    ## Go through each of the properties in the key
    foreach($keyProperty in $key.GetValueNames())
    {
        ## If the property matches the search term, add it as a
        ## property to the output
        if($keyProperty -like $property)
        {
            $returnObject |
                Add-Member NoteProperty $keyProperty $key.GetValue($keyProperty)
        }
    }

    ## Return the resulting object
    $returnObject

    ## Close the key and base keys
    $key.Close()
    $baseKey.Close()
}

function Get-ScriptCoverage
{
    #############################################################################
    ##
    ## Get-ScriptCoverage
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Uses conditional breakpoints to obtain information about what regions of
    a script are executed when run.

    .EXAMPLE

    PS > Get-Content c:\temp\looper.ps1

    param($userInput)

    for($count = 0; $count -lt 10; $count++)
    {
        "Count is: $count"
    }

    if($userInput -eq "One")
    {
        "Got 'One'"
    }

    if($userInput -eq "Two")
    {
        "Got 'Two'"
    }

    PS > $action = { c:\temp\looper.ps1 -UserInput 'One' }
    PS > $coverage = Get-ScriptCoverage c:\temp\looper.ps1 -Action $action
    PS > $coverage | Select Content,StartLine,StartColumn | Format-Table -Auto

    Content   StartLine StartColumn
    -------   --------- -----------
    userInput         1           7
    Got 'Two'        15           5

    This example exercises a 'looper.ps1' script, and supplies it with some
    user input. The output demonstrates that we didn't exercise the
    "Got 'Two'" statement.

    #>

    param(
        ## The path of the script to monitor
        $Path,

        ## The command to exercise the script
        [ScriptBlock] $Action = { & $path }
    )

    Set-StrictMode -Version 3

    ## Determine all of the tokens in the script
    $scriptContent = Get-Content $path
    $ignoreTokens = "Comment","NewLine","StatementSeparator","Keyword",
        "GroupStart","GroupEnd"
    $tokens = [System.Management.Automation.PsParser]::Tokenize(
        $scriptContent, [ref] $null) |
        Where-Object { $ignoreTokens -notcontains $_.Type }
    $tokens = $tokens | Sort-Object StartLine,StartColumn

    ## Create a variable to hold the tokens that PowerShell actually hits
    $visited = New-Object System.Collections.ArrayList

    ## Go through all of the tokens
    $breakpoints = foreach($token in $tokens)
    {
        ## Create a new action. This action logs the token that we
        ## hit. We call GetNewClosure() so that the $token variable
        ## gets the _current_ value of the $token variable, as opposed
        ## to the value it has when the breakpoints gets hit.
        $breakAction = { $null = $visited.Add($token) }.GetNewClosure()

        ## Set a breakpoint on the line and column of the current token.
        ## We use the action from above, which simply logs that we've hit
        ## that token.
        Set-PsBreakpoint $path -Line `
            $token.StartLine -Column $token.StartColumn -Action $breakAction
    }

    ## Invoke the action that exercises the script
    $null = . $action

    ## Remove the temporary breakpoints we set
    $breakpoints | Remove-PsBreakpoint

    ## Sort the tokens that we hit, and compare them with all of the tokens
    ## in the script. Output the result of that comparison.
    $visited = $visited | Sort-Object -Unique StartLine,StartColumn
    Compare-Object $tokens $visited -Property StartLine,StartColumn -PassThru

    ## Clean up our temporary variable
    Remove-Item variable:\visited
}

function Get-ScriptPerformanceProfile
{
    #############################################################################
    ##
    ## Get-ScriptPerformanceProfile
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Computes the performance characteristics of a script, based on the transcript
    of it running at trace level 1.

    .DESCRIPTION

    To profile a script:

       1) Turn on script tracing in the window that will run the script:
          Set-PsDebug -trace 1
       2) Turn on the transcript for the window that will run the script:
          Start-Transcript
          (Note the filename that PowerShell provides as the logging destination.)
       3) Type in the script name, but don't actually start it.
       4) Open another PowerShell window, and navigate to the directory holding
          this script.  Type in '.\Get-ScriptPerformanceProfile <transcript>',
          replacing <transcript> with the path given in step 2.  Don't
          press <Enter> yet.
       5) Switch to the profiled script window, and start the script.
          Switch to the window containing this script, and press <Enter>
       6) Wait until your profiled script exits, or has run long enough to be
          representative of its work.  To be statistically accurate, your script
          should run for at least ten seconds.
       7) Switch to the window running this script, and press a key.
       8) Switch to the window holding your profiled script, and type:
          Stop-Transcript
       9) Delete the transcript.

    .NOTES

    You can profile regions of code (ie: functions) rather than just lines
    by placing the following call at the start of the region:
          Write-Debug "ENTER <region_name>"
    and the following call and the end of the region:
          Write-Debug "EXIT"
    This is implemented to account exclusively for the time spent in that
    region, and does not include time spent in regions contained within the
    region.  For example, if FunctionA calls FunctionB, and you've surrounded
    each by region markers, the statistics for FunctionA will not include the
    statistics for FunctionB.

    #>

    param(
        ## The path of the transcript log file
        [Parameter(Mandatory = $true)]
        $Path
    )

    Set-StrictMode -Version 3

    function Main
    {
        ## Run the actual profiling of the script.  $uniqueLines gets
        ## the mapping of line number to actual script content.
        ## $samples gets a hashtable mapping line number to the number of times
        ## we observed the script running that line.
        $uniqueLines = @{}
        $samples = GetSamples $uniqueLines

        "Breakdown by line:"
        "----------------------------"

        ## Create a new hash table that flips the $samples hashtable --
        ## one that maps the number of times sampled to the line sampled.
        ## Also, figure out how many samples we got altogether.
        $counts = @{}
        $totalSamples = 0;
        foreach($item in $samples.Keys)
        {
            $counts[$samples[$item]] = $item
            $totalSamples += $samples[$item]
        }

        ## Go through the flipped hashtable, in descending order of number of
        ## samples.  As we do so, output the number of samples as a percentage of
        ## the total samples.  This gives us the percentage of the time our
        ## script spent executing that line.
        foreach($count in ($counts.Keys | Sort-Object -Descending))
        {
            $line = $counts[$count]
            $percentage = "{0:#0}" -f ($count * 100 / $totalSamples)
            "{0,3}%: Line {1,4} -{2}" -f $percentage,$line,
                $uniqueLines[$line]
        }

        ## Go through the transcript log to figure out which lines are part of
        ## any marked regions.  This returns a hashtable that maps region names
        ## to the lines they contain.
        ""
        "Breakdown by marked regions:"
        "----------------------------"
        $functionMembers = GenerateFunctionMembers

        ## For each region name, cycle through the lines in the region.  As we
        ## cycle through the lines, sum up the time spent on those lines and
        ## output the total.
        foreach($key in $functionMembers.Keys)
        {
            $totalTime = 0
            foreach($line in $functionMembers[$key])
            {
                $totalTime += ($samples[$line] * 100 / $totalSamples)
            }

            $percentage = "{0:#0}" -f $totalTime
            "{0,3}%: {1}" -f $percentage,$key
        }
    }

    ## Run the actual profiling of the script.  $uniqueLines gets
    ## the mapping of line number to actual script content.
    ## Return a hashtable mapping line number to the number of times
    ## we observed the script running that line.
    function GetSamples($uniqueLines)
    {
        ## Open the log file.  We use the .Net file I/O, so that we keep
        ## monitoring just the end of the file.  Otherwise, we would make our
        ## timing inaccurate as we scan the entire length of the file every time.
        $logStream = [System.IO.File]::Open($Path, "Open", "Read", "ReadWrite")
        $logReader = New-Object System.IO.StreamReader $logStream

        $random = New-Object Random
        $samples = @{}

        $lastCounted = $null

        ## Gather statistics until the user presses a key.
        while(-not $host.UI.RawUI.KeyAvailable)
        {
            ## We sleep a slightly random amount of time.  If we sleep a constant
            ## amount of time, we run the very real risk of improperly sampling
            ## scripts that exhibit periodic behaviour.
            $sleepTime = [int] ($random.NextDouble() * 100.0)
            Start-Sleep -Milliseconds $sleepTime

            ## Get any content produced by the transcript since our last poll.
            ## From that poll, extract the last DEBUG statement (which is the last
            ## line executed.)
            $rest = $logReader.ReadToEnd()
            $lastEntryIndex = $rest.LastIndexOf("DEBUG: ")

            ## If we didn't get a new line, then the script is still working on
            ## the last line that we captured.
            if($lastEntryIndex -lt 0)
            {
                if($lastCounted) { $samples[$lastCounted] ++ }
                continue;
            }

            ## Extract the debug line.
            $lastEntryFinish = $rest.IndexOf("\n", $lastEntryIndex)
            if($lastEntryFinish -eq -1) { $lastEntryFinish = $rest.length }

            $scriptLine = $rest.Substring(
                $lastEntryIndex, ($lastEntryFinish - $lastEntryIndex)).Trim()
            if($scriptLine -match 'DEBUG:[ \t]*([0-9]*)\+(.*)')
            {
                ## Pull out the line number from the line
                $last = $matches[1]

                $lastCounted = $last
                $samples[$last] ++

                ## Pull out the actual script line that matches the line number
                $uniqueLines[$last] = $matches[2]
            }

            ## Discard anything that's buffered during this poll, and start
            ## waiting again
            $logReader.DiscardBufferedData()
        }

        ## Clean up
        $logStream.Close()
        $logReader.Close()

        $samples
    }

    ## Go through the transcript log to figure out which lines are part of any
    ## marked regions.  This returns a hashtable that maps region names to
    ## the lines they contain.
    function GenerateFunctionMembers
    {
        ## Create a stack that represents the callstack.  That way, if a marked
        ## region contains another marked region, we attribute the statistics
        ## appropriately.
        $callstack = New-Object System.Collections.Stack
        $currentFunction = "Unmarked"
        $callstack.Push($currentFunction)

        $functionMembers = @{}

        ## Go through each line in the transcript file, from the beginning
        foreach($line in (Get-Content $Path))
        {
            ## Check if we're entering a monitor block
            ## If so, store that we're in that function, and push it onto
            ## the callstack.
            if($line -match 'write-debug "ENTER (.*)"')
            {
                $currentFunction = $matches[1]
                $callstack.Push($currentFunction)
            }
            ## Check if we're exiting a monitor block
            ## If so, clear the "current function" from the callstack,
            ## and store the new "current function" onto the callstack.
            elseif($line -match 'write-debug "EXIT"')
            {
                [void] $callstack.Pop()
                $currentFunction = $callstack.Peek()
            }
            ## Otherwise, this is just a line with some code.
            ## Add the line number as a member of the "current function"
            else
            {
                if($line -match 'DEBUG:[ \t]*([0-9]*)\+')
                {
                    ## Create the arraylist if it's not initialized
                    if(-not $functionMembers[$currentFunction])
                    {
                        $functionMembers[$currentFunction] =
                            New-Object System.Collections.ArrayList
                    }

                    ## Add the current line to the ArrayList
                    $hitLines = $functionMembers[$currentFunction]
                    if(-not $hitLines.Contains($matches[1]))
                    {
                        [void] $hitLines.Add($matches[1])
                    }
                }
            }
        }

        $functionMembers
    }

    . Main
}

function Get-Tomorrow
{
    ##############################################################################
    ## Get-Tomorrow
    ##
    ## Get the date that represents tomorrow
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    Set-StrictMode -Version 3

    function GetDate
    {
        Get-Date
    }

    $tomorrow = (GetDate).AddDays(1)
    $tomorrow
}

function Get-UserLogonLogoffScript
{
    ##############################################################################
    ##
    ## Get-UserLogonLogoffScript
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get the logon or logoff scripts assigned to a specific user

    .EXAMPLE

    Get-UserLogonLogoffScript LEE-DESK\LEE Logon
    Gets all logon scripts for the user 'LEE-DESK\Lee'

    #>

    param(
        ## The username to examine
        [Parameter(Mandatory = $true)]
        $Username,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Logon","Logoff")]
        $ScriptType
    )

    Set-StrictMode -Version 3

    ## Find the SID for the username
    $account = New-Object System.Security.Principal.NTAccount $username
    $sid =
        $account.Translate([System.Security.Principal.SecurityIdentifier]).Value

    ## Map that to their group policy scripts
    $registryKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\" +
        "Group Policy\State\$sid\Scripts"

    if(-not (Test-Path $registryKey))
    {
        return
    }

    ## Go through each of the policies in the specified key
    foreach($policy in Get-ChildItem $registryKey\$scriptType)
    {
        ## For each of the scripts in that policy, get its script name
        ## and parameters
        foreach($script in Get-ChildItem $policy.PsPath)
        {
            Get-ItemProperty $script.PsPath | Select Script,Parameters
        }
    }
}

function Get-WarningsAndErrors
{
    ##############################################################################
    ##
    ## Get-WarningsAndErrors
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the functionality of the Write-Warning, Write-Error, and throw
    statements

    #>

    Set-StrictMode -Version 3

    Write-Warning "Warning: About to generate an error"
    Write-Error "Error: You are running this script"
    throw "Could not complete operation."
}

function Get-WmiClassKeyProperty
{
    ##############################################################################
    ##
    ## Get-WmiClassKeyProperty
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Get all of the properties that you may use in a WMI filter for a given class.

    .EXAMPLE

    PS > Get-WmiClassKeyProperty Win32_Process
    Handle

    #>

    param(
        ## The WMI class to examine
        [WmiClass] $WmiClass
    )

    Set-StrictMode -Version 3

    ## WMI classes have properties
    foreach($currentProperty in $wmiClass.Properties)
    {
        ## WMI properties have qualifiers to explain more about them
        foreach($qualifier in $currentProperty.Qualifiers)
        {
            ## If it has a 'Key' qualifier, then you may use it in a filter
            if($qualifier.Name -eq "Key")
            {
                $currentProperty.Name
            }
        }
    }
}

function Grant-RegistryAccessFullControl
{
    ##############################################################################
    ##
    ## Grant-RegistryAccessFullControl
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Grants full control access to a user for the specified registry key.

    .EXAMPLE

    PS > $registryPath = "HKLM:\Software\MyProgram"
    PS > Grant-RegistryAccessFullControl "LEE-DESK\LEE" $registryPath

    #>

    param(
        ## The user to grant full control
        [Parameter(Mandatory = $true)]
        $User,

        ## The registry path that should have its permissions modified
        [Parameter(Mandatory = $true)]
        $RegistryPath
    )

    Set-StrictMode -Version 3

    Push-Location
    Set-Location -LiteralPath $registryPath

    ## Retrieve the ACL from the registry key
    $acl = Get-Acl .

    ## Prepare the access rule, and set the access rule
    $arguments = $user,"FullControl","Allow"
    $accessRule = New-Object Security.AccessControl.RegistryAccessRule $arguments
    $acl.SetAccessRule($accessRule)

    ## Apply the modified ACL to the regsitry key
    $acl | Set-Acl  .

    Pop-Location
}

function Import-ADUser
{
    #############################################################################
    ##
    ## Import-AdUser
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    #############################################################################

    <#

    .SYNOPSIS

    Create users in Active Directory from the content of a CSV.

    .DESCRIPTION

    In the user CSV, One column must be named "CN" for the user name.
    All other columns represent properties in Active Directory for that user.

    For example:
    CN,userPrincipalName,displayName,manager
    MyerKen,Ken.Myer@fabrikam.com,Ken Myer,
    DoeJane,Jane.Doe@fabrikam.com,Jane Doe,"CN=MyerKen,OU=West,OU=Sales,DC=..."
    SmithRobin,Robin.Smith@fabrikam.com,Robin Smith,"CN=MyerKen,OU=West,OU=..."

    .EXAMPLE

    PS > $container = "LDAP://localhost:389/ou=West,ou=Sales,dc=Fabrikam,dc=COM"
    PS > Import-ADUser.ps1 $container .\users.csv

    #>

    param(
        ## The container in which to import users
        ## For example:
        ## "LDAP://localhost:389/ou=West,ou=Sales,dc=Fabrikam,dc=COM)")
        [Parameter(Mandatory = $true)]
        $Container,

        ## The path to the CSV that contains the user records
        [Parameter(Mandatory = $true)]
        $Path
    )

    Set-StrictMode -Off

    ## Bind to the container
    $userContainer = [adsi] $container

    ## Ensure that the container was valid
    if(-not $userContainer.Name)
    {
        Write-Error "Could not connect to $container"
        return
    }

    ## Load the CSV
    $users = @(Import-Csv $Path)
    if($users.Count -eq 0)
    {
        return
    }

    ## Go through each user from the CSV
    foreach($user in $users)
    {
        ## Pull out the name, and create that user
        $username = $user.CN
        $newUser = $userContainer.Create("User", "CN=$username")

        ## Go through each of the properties from the CSV, and set its value
        ## on the user
        foreach($property in $user.PsObject.Properties)
        {
            ## Skip the property if it was the CN property that sets the
            ## user name
            if($property.Name -eq "CN")
            {
                continue
            }

            ## Ensure they specified a value for the property
            if(-not $property.Value)
            {
                continue
            }

            ## Set the value of the property
            $newUser.Put($property.Name, $property.Value)
        }

        ## Finalize the information in Active Directory
        $newUser.SetInfo()
    }
}

function Invoke-AddTypeTypeDefinition
{
    #############################################################################
    ##
    ## Invoke-AddTypeTypeDefinition
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the use of the -TypeDefinition parameter of the Add-Type
    cmdlet.

    #>

    Set-StrictMode -Version 3

    ## Define the new C# class
    $newType = @'
    using System;

    namespace PowerShellCookbook
    {
        public class AddTypeTypeDefinitionDemo
        {
            public string SayHello(string name)
            {
                string result = String.Format("Hello {0}", name);
                return result;
            }
        }
    }

'@

    ## Add it to the Powershell session
    Add-Type -TypeDefinition $newType

    ## Show that we can access it like any other .NET type
    $greeter = New-Object PowerShellCookbook.AddTypeTypeDefinitionDemo
    $greeter.SayHello("World")
}

function Invoke-AdvancedFunction
{
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $Scriptblock
        )

    ## Invoke the scriptblock supplied by the user.
    & $scriptblock
}

function Invoke-BinaryProcess
{
    ##############################################################################
    ##
    ## Invoke-BinaryProcess
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Invokes a process that emits or consumes binary data.

    .EXAMPLE

    PS > Invoke-BinaryProcess binaryProcess.exe -RedirectOutput -ArgumentList "-Emit" |
           Invoke-BinaryProcess binaryProcess.exe -RedirectInput -ArgumentList "-Consume"

    #>

    param(
        ## The name of the process to invoke
        [string] $ProcessName,

        ## Specifies that input to the process should be treated as
        ## binary
        [Alias("Input")]
        [switch] $RedirectInput,

        ## Specifies that the output of the process should be treated
        ## as binary
        [Alias("Output")]
        [switch] $RedirectOutput,

        ## Specifies the arguments for the process
        [string] $ArgumentList
    )

    Set-StrictMode -Version 3

    ## Prepare to invoke the process
    $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processStartInfo.FileName = (Get-Command $processname).Definition
    $processStartInfo.WorkingDirectory = (Get-Location).Path
    if($argumentList) { $processStartInfo.Arguments = $argumentList }
    $processStartInfo.UseShellExecute = $false

    ## Always redirect the input and output of the process.
    ## Sometimes we will capture it as binary, other times we will
    ## just treat it as strings.
    $processStartInfo.RedirectStandardOutput = $true
    $processStartInfo.RedirectStandardInput = $true

    $process = [System.Diagnostics.Process]::Start($processStartInfo)

    ## If we've been asked to redirect the input, treat it as bytes.
    ## Otherwise, write any input to the process as strings.
    if($redirectInput)
    {
        $inputBytes = @($input)
        $process.StandardInput.BaseStream.Write($inputBytes, 0, $inputBytes.Count)
        $process.StandardInput.Close()
    }
    else
    {
        $input | % { $process.StandardInput.WriteLine($_) }
        $process.StandardInput.Close()
    }

    ## If we've been asked to redirect the output, treat it as bytes.
    ## Otherwise, read any input from the process as strings.
    if($redirectOutput)
    {
        $byteRead = -1
        do
        {
            $byteRead = $process.StandardOutput.BaseStream.ReadByte()
            if($byteRead -ge 0) { $byteRead }
        } while($byteRead -ge 0)
    }
    else
    {
        $process.StandardOutput.ReadToEnd()
    }
}

function Invoke-CmdScript
{
    ##############################################################################
    ##
    ## Invoke-CmdScript
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Invoke the specified batch file (and parameters), but also propigate any
    environment variable changes back to the PowerShell environment that
    called it.

    .EXAMPLE

    PS > type foo-that-sets-the-FOO-env-variable.cmd
    @set FOO=%*
    echo FOO set to %FOO%.

    PS > $env:FOO
    PS > Invoke-CmdScript "foo-that-sets-the-FOO-env-variable.cmd" Test

    C:\Temp>echo FOO set to Test.
    FOO set to Test.

    PS > $env:FOO
    Test

    #>

    param(
        ## The path to the script to run
        [Parameter(Mandatory = $true)]
        [string] $Path,

        ## The arguments to the script
        [string] $ArgumentList
    )

    Set-StrictMode -Version 3

    $tempFile = [IO.Path]::GetTempFileName()

    ## Store the output of cmd.exe.  We also ask cmd.exe to output
    ## the environment table after the batch file completes
    cmd /c " `"$Path`" $argumentList && set > `"$tempFile`" "

    ## Go through the environment variables in the temp file.
    ## For each of them, set the variable in our local environment.
    Get-Content $tempFile | Foreach-Object {
        if($_ -match "^(.*?)=(.*)$")
        {
            Set-Content "env:\$($matches[1])" $matches[2]
        }
    }

    Remove-Item $tempFile
}

function Invoke-ComplexDebuggerScript
{
    #############################################################################
    ##
    ## Invoke-ComplexDebuggerScript
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the functionality of PowerShell's debugging support.

    #>

    Set-StrictMode -Version 3

    function HelperFunction
    {
        $dirCount = 0
    }

    Write-Host "Calculating lots of complex information"

    $runningTotal = 0
    $runningTotal += [Math]::Pow(5 * 5 + 10, 2)
    $runningTotal

    $dirCount = @(Get-ChildItem $env:WINDIR).Count
    $dirCount

    HelperFunction

    $dirCount

    $runningTotal -= 10
    $runningTotal /= 2
    $runningTotal

    $runningTotal *= 3
    $runningTotal /= 2
    $runningTotal
}

function Invoke-ComplexScript
{
    #############################################################################
    ##
    ## Invoke-ComplexScript
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the functionality of PowerShell's debugging support.

    #>

    Set-StrictMode -Version 3

    Write-Host "Calculating lots of complex information"

    $runningTotal = 0
    $runningTotal += [Math]::Pow(5 * 5 + 10, 2)

    Write-Debug "Current value: $runningTotal"

    Set-PsDebug -Trace 1
    $dirCount = @(Get-ChildItem $env:WINDIR).Count

    Set-PsDebug -Trace 2
    $runningTotal -= 10
    $runningTotal /= 2

    Set-PsDebug -Step
    $runningTotal *= 3
    $runningTotal /= 2

    $host.EnterNestedPrompt()

    Set-PsDebug -off
}

function Invoke-DemonstrationScript
{
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)

    "The script ran!"
}

function Invoke-ElevatedCommand
{
    ##############################################################################
    ##
    ## Invoke-ElevatedCommand
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Runs the provided script block under an elevated instance of PowerShell as
    through it were a member of a regular pipeline.

    .EXAMPLE

    PS > Get-Process | Invoke-ElevatedCommand.ps1 {
        $input | Where-Object { $_.Handles -gt 500 } } | Sort Handles

    #>

    param(
        ## The script block to invoke elevated
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $Scriptblock,

        ## Any input to give the elevated process
        [Parameter(ValueFromPipeline = $true)]
        $InputObject,

        ## Switch to enable the user profile
        [switch] $EnableProfile
    )

    begin
    {
        Set-StrictMode -Version 3
        $inputItems = New-Object System.Collections.ArrayList
    }

    process
    {
        $null = $inputItems.Add($inputObject)
    }

    end
    {
        ## Create some temporary files for streaming input and output
        $outputFile = [IO.Path]::GetTempFileName()
        $inputFile = [IO.Path]::GetTempFileName()

        ## Stream the input into the input file
        $inputItems.ToArray() | Export-CliXml -Depth 1 $inputFile

        ## Start creating the command line for the elevated PowerShell session
        $commandLine = ""
        if(-not $EnableProfile) { $commandLine += "-NoProfile " }

        ## Convert the command into an encoded command for PowerShell
        $commandString = "Set-Location '$($pwd.Path)'; " +
            "`$output = Import-CliXml '$inputFile' | " +
            "& {" + $scriptblock.ToString() + "} 2>&1; " +
            "`$output | Export-CliXml -Depth 1 '$outputFile'"

        $commandBytes = [System.Text.Encoding]::Unicode.GetBytes($commandString)
        $encodedCommand = [Convert]::ToBase64String($commandBytes)
        $commandLine += "-EncodedCommand $encodedCommand"

        ## Start the new PowerShell process
        $process = Start-Process -FilePath (Get-Command powershell).Definition `
            -ArgumentList $commandLine -Verb RunAs `
            -WindowStyle Hidden `
            -Passthru
        $process.WaitForExit()

        ## Return the output to the user
        if((Get-Item $outputFile).Length -gt 0)
        {
            Import-CliXml $outputFile
        }

        ## Clean up
        [Console]::WriteLine($outputFile)
        # Remove-Item $outputFile
        Remove-Item $inputFile
    }
}

function Invoke-Inline
{
    #############################################################################
    ##
    ## Invoke-Inline
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    #############################################################################

    <#

    .SYNOPSIS

    Demonstrates the Add-Type cmdlet to invoke in-line C#

    #>

    Set-StrictMode -Version 3

    $inlineType = Add-Type -Name InvokeInline_Inline -PassThru `
        -MemberDefinition @'
        public static int RightShift(int original, int places)
        {
            return original >> places;
        }
'@

    $inlineType::RightShift(1024, 3)
}

function Invoke-LocalizedScript
{
    Set-StrictMode -Version 3

    ## Create some default messages for English cultures, and
    ## when culture-specific messages are not available.
    $messages = DATA {
        @{
            Greeting = "Hello, {0}"
            Goodbye = "So long."
        }
    }

    ## Import localized messages for the current culture.
    Import-LocalizedData messages -ErrorAction SilentlyContinue

    ## Output the localized messages
    $messages.Greeting -f "World"
    $messages.Goodbye
}

function Invoke-LongRunningOperation
{
    ##############################################################################
    ##
    ## Invoke-LongRunningOperation
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the functionality of the Write-Progress cmdlet

    #>

    Set-StrictMode -Version 3

    $activity = "A long running operation"
    $status = "Initializing"

    ## Initialize the long-running operation
    for($counter = 0; $counter -lt 100; $counter++)
    {
        $currentOperation = "Initializing item $counter"
        Write-Progress $activity $status -PercentComplete $counter `
            -CurrentOperation $currentOperation
        Start-Sleep -m 20
    }

    $status = "Running"

    ## Initialize the long-running operation
    for($counter = 0; $counter -lt 100; $counter++)
    {
        $currentOperation = "Running task $counter"
        Write-Progress $activity $status -PercentComplete $counter `
            -CurrentOperation $currentOperation
        Start-Sleep -m 20
    }
}

function Invoke-Member
{
    ##############################################################################
    ##
    ## Invoke-Member
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Enables easy access to methods and properties of pipeline objects.

    .EXAMPLE

    PS > "Hello","World" | .\Invoke-Member Length
    5
    5

    .EXAMPLE

    PS > "Hello","World" | .\Invoke-Member -m ToUpper
    HELLO
    WORLD

    .EXAMPLE

    PS > "Hello","World" | .\Invoke-Member Replace l w
    Hewwo
    Worwd

    #>

    [CmdletBinding(DefaultParameterSetName= "Member")]
    param(

        ## A switch parameter to identify the requested member as a method.
        ## Only required for methods that take no arguments.
        [Parameter(ParameterSetName = "Method")]
        [Alias("M","Me")]
        [switch] $Method,

        ## The name of the member to retrieve
        [Parameter(ParameterSetName = "Method", Position = 0)]
        [Parameter(ParameterSetName = "Member", Position = 0)]
        [string] $Member,

        ## Arguments for the method, if any
        [Parameter(
            ParameterSetName = "Method", Position = 1,
            Mandatory = $false, ValueFromRemainingArguments = $true)]
        [object[]] $ArgumentList = @(),

        ## The object from which to retrieve the member
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
        )

    process
    {
        ## If the user specified a method, invoke it
        ## with any required arguments.
        if($psCmdlet.ParameterSetName -eq "Method")
        {
            $inputObject.$member.Invoke(@($argumentList))
        }
        ## Otherwise, retrieve the property
        else
        {
            $inputObject.$member
        }
    }
}

function Invoke-RemoteExpression
{
    ##############################################################################
    ##
    ## Invoke-RemoteExpression
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Invoke a PowerShell expression on a remote machine. Requires PsExec from
    http://live.sysinternals.com/tools/psexec.exe. If the remote machine
    supports PowerShell version two, use PowerShell remoting instead.

    .EXAMPLE

    PS > Invoke-RemoteExpression LEE-DESK { Get-Process }
    Retrieves the output of a simple command from a remote machine

    .EXAMPLE

    PS > (Invoke-RemoteExpression LEE-DESK { Get-Date }).AddDays(1)
    Invokes a command on a remote machine. Since the command returns one of
    PowerShell's primitive types (a DateTime object,) you can manipulate
    its output as an object afterward.

    .EXAMPLE

    PS > Invoke-RemoteExpression LEE-DESK { Get-Process } | Sort Handles
    Invokes a command on a remote machine. The command does not return one of
    PowerShell's primitive types, but you can still use PowerShell's filtering
    cmdlets to work with its structured output.

    #>

    param(
        ## The computer on which to invoke the command.
        $ComputerName = "$ENV:ComputerName",

        ## The scriptblock to invoke on the remote machine.
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock,

        ## The username / password to use in the connection
        $Credential,

        ## Determines if PowerShell should load the user's PowerShell profile
        ## when invoking the command.
        [switch] $NoProfile
    )

    Set-StrictMode -Version 3

    ## Prepare the computername for PSExec
    if($ComputerName -notmatch '^\\')
    {
        $ComputerName = "\\$ComputerName"
    }

    ## Prepare the command line for PsExec. We use the XML output encoding so
    ## that PowerShell can convert the output back into structured objects.
    ## PowerShell expects that you pass it some input when being run by PsExec
    ## this way, so the 'echo .' statement satisfies that appetite.
    $commandLine = "echo . | powershell -Output XML "

    if($noProfile)
    {
        $commandLine += "-NoProfile "
    }

    ## Convert the command into an encoded command for PowerShell
    $commandBytes = [System.Text.Encoding]::Unicode.GetBytes($scriptblock)
    $encodedCommand = [Convert]::ToBase64String($commandBytes)
    $commandLine += "-EncodedCommand $encodedCommand"

    ## Collect the output and error output
    $errorOutput = [IO.Path]::GetTempFileName()

    if($Credential)
    {
        ## This lets users pass either a username, or full credential to our
        ## credential parameter
        $credential = Get-Credential $credential
        $networkCredential = $credential.GetNetworkCredential()
        $username = $networkCredential.Username
        $password = $networkCredential.Password

        $output = psexec $computername /user $username /password $password `
            /accepteula cmd /c $commandLine 2>$errorOutput
    }
    else
    {
        $output = psexec /acceptEula $computername cmd /c $commandLine 2>$errorOutput
    }

    ## Check for any errors
    $errorContent = Get-Content $errorOutput
    Remove-Item $errorOutput

    if($lastExitCode -ne 0)
    {
        $OFS = "`n"
        $errorMessage = "Could not execute remote expression. "
        $errorMessage += "Ensure that your account has administrative " +
            "privileges on the target machine.`n"
        $errorMessage += ($errorContent -match "psexec.exe :")

        Write-Error $errorMessage
    }

    ## Return the output to the user
    $output
}

function Invoke-ScriptBlock
{
    ##############################################################################
    ##
    ## Invoke-ScriptBlock
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Apply the given mapping command to each element of the input. (Note that
    PowerShell includes this command natively, and calls it Foreach-Object)

    .EXAMPLE

    PS > 1,2,3 | Invoke-ScriptBlock { $_ * 2 }

    #>

    param(
        ## The scriptblock to apply to each incoming element
        [ScriptBlock] $MapCommand
    )

    begin
    {
        Set-StrictMode -Version 3
    }
    process
    {
        & $mapCommand
    }
}

function Invoke-ScriptBlockClosure
{
    ##############################################################################
    ##
    ## Invoke-ScriptBlockClosure
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Demonstrates the GetNewClosure() method on a script block that pulls variables
    in from the user's session (if they are defined.)

    .EXAMPLE

    PS > $name = "Hello There"
    PS > Invoke-ScriptBlockClosure { $name }
    Hello There
    Hello World
    Hello There

    #>

    param(
        ## The scriptblock to invoke
        [ScriptBlock] $ScriptBlock
    )

    Set-StrictMode -Version 3

    ## Create a new script block that pulls variables
    ## from the user's scope (if defined.)
    $closedScriptBlock = $scriptBlock.GetNewClosure()

    ## Invoke the script block normally. The contents of
    ## the $name variable will be from the user's session.
    & $scriptBlock

    ## Define a new variable
    $name = "Hello World"

    ## Invoke the script block normally. The contents of
    ## the $name variable will be "Hello World", now from
    ## our scope.
    & $scriptBlock

    ## Invoke the "closed" script block. The contents of
    ## the $name variable will still be whatever was in the user's session
    ## (if it was defined.)
    & $closedScriptBlock
}

function Invoke-ScriptThatRequiresMta
{
    ###########################################################################
    ##
    ## Invoke-ScriptThatRequiresMta
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ###########################################################################

    <#

    .SYNOPSIS

    Demonstrates a technique to relaunch a script that requires MTA mode.
    This is useful only for simple parameter definitions that can be
    specified positionally.

    #>

    param(
        $Parameter1,
        $Parameter2
    )

    Set-StrictMode -Version 3

    "Current threading mode: " + $host.Runspace.ApartmentState
    "Parameter1 is: $parameter1"
    "Parameter2 is: $parameter2"

    if($host.Runspace.ApartmentState -eq "STA")
    {
        "Relaunching"
        $file = $myInvocation.MyCommand.Path
        powershell -NoProfile -Mta -File $file $parameter1 $parameter2
        return
    }

    "After relaunch - current threading mode: " + $host.Runspace.ApartmentState
}

function Invoke-SqlCommand
{
    ##############################################################################
    ##
    ## Invoke-SqlCommand
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Return the results of a SQL query or operation

    .EXAMPLE

    Invoke-SqlCommand.ps1 -Sql "SELECT TOP 10 * FROM Orders"
    Invokes a command using Windows authentication

    .EXAMPLE

    PS > $cred = Get-Credential
    PS > Invoke-SqlCommand.ps1 -Sql "SELECT TOP 10 * FROM Orders" -Cred $cred
    Invokes a command using SQL Authentication

    .EXAMPLE

    PS > $server = "MYSERVER"
    PS > $database = "Master"
    PS > $sql = "UPDATE Orders SET EmployeeID = 6 WHERE OrderID = 10248"
    PS > Invoke-SqlCommand $server $database $sql
    Invokes a command that performs an update

    .EXAMPLE

    PS > $sql = "EXEC SalesByCategory 'Beverages'"
    PS > Invoke-SqlCommand -Sql $sql
    Invokes a stored procedure

    .EXAMPLE

    PS > Invoke-SqlCommand (Resolve-Path access_test.mdb) -Sql "SELECT * FROM Users"
    Access an Access database

    .EXAMPLE

    PS > Invoke-SqlCommand (Resolve-Path xls_test.xls) -Sql 'SELECT * FROM [Sheet1$]'
    Access an Excel file

    #>

    param(
        ## The data source to use in the connection
        [string] $DataSource = ".\SQLEXPRESS",

        ## The database within the data source
        [string] $Database = "Northwind",

        ## The SQL statement(s) to invoke against the database
        [Parameter(Mandatory = $true)]
        [string[]] $SqlCommand,

        ## The timeout, in seconds, to wait for the query to complete
        [int] $Timeout = 60,

        ## The credential to use in the connection, if any.
        $Credential
    )


    Set-StrictMode -Version 3

    ## Prepare the authentication information. By default, we pick
    ## Windows authentication
    $authentication = "Integrated Security=SSPI;"

    ## If the user supplies a credential, then they want SQL
    ## authentication
    if($credential)
    {
        $credential = Get-Credential $credential
        $plainCred = $credential.GetNetworkCredential()
        $authentication =
            ("uid={0};pwd={1};" -f $plainCred.Username,$plainCred.Password)
    }

    ## Prepare the connection string out of the information they
    ## provide
    $connectionString = "Provider=sqloledb; " +
                        "Data Source=$dataSource; " +
                        "Initial Catalog=$database; " +
                        "$authentication; "

    ## If they specify an Access database or Excel file as the connection
    ## source, modify the connection string to connect to that data source
    if($dataSource -match '\.xls$|\.mdb$')
    {
        $connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
            "Data Source=$dataSource; "

        if($dataSource -match '\.xls$')
        {
            $connectionString += 'Extended Properties="Excel 8.0;"; '

            ## Generate an error if they didn't specify the sheet name properly
            if($sqlCommand -notmatch '\[.+\$\]')
            {
                $error = 'Sheet names should be surrounded by square brackets, ' +
                    'and have a dollar sign at the end: [Sheet1$]'
                Write-Error $error
                return
            }
        }
    }

    ## Connect to the data source and open it
    $connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
    $connection.Open()

    foreach($commandString in $sqlCommand)
    {
        $command = New-Object Data.OleDb.OleDbCommand $commandString,$connection
        $command.CommandTimeout = $timeout

        ## Fetch the results, and close the connection
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
        $dataset = New-Object System.Data.DataSet
        [void] $adapter.Fill($dataSet)

        ## Return all of the rows from their query
        $dataSet.Tables | Select-Object -Expand Rows
    }

    $connection.Close()
}

function Invoke-WindowsApi
{
    ##############################################################################
    ##
    ## Invoke-WindowsApi
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Invoke a native Windows API call that takes and returns simple data types.

    .EXAMPLE

    ## Prepare the parameter types and parameters for the CreateHardLink function
    PS > $filename = "c:\temp\hardlinked.txt"
    PS > $existingFilename = "c:\temp\link_target.txt"
    PS > Set-Content $existingFilename "Hard Link target"
    PS > $parameterTypes = [string], [string], [IntPtr]
    PS > $parameters = [string] $filename, [string] $existingFilename,
        [IntPtr]::Zero

    ## Call the CreateHardLink method in the Kernel32 DLL
    PS > $result = Invoke-WindowsApi "kernel32" ([bool]) "CreateHardLink" `
        $parameterTypes $parameters
    PS > Get-Content C:\temp\hardlinked.txt
    Hard Link target

    #>

    param(
        ## The name of the DLL that contains the Windows API, such as "kernel32"
        [string] $DllName,

        ## The return type expected from Windows API
        [Type] $ReturnType,

        ## The name of the Windows API
        [string] $MethodName,

        ## The types of parameters expected by the Windows API
        [Type[]] $ParameterTypes,

        ## Parameter values to pass to the Windows API
        [Object[]] $Parameters
    )

    Set-StrictMode -Version 3

    ## Begin to build the dynamic assembly
    $domain = [AppDomain]::CurrentDomain
    $name = New-Object Reflection.AssemblyName 'PInvokeAssembly'
    $assembly = $domain.DefineDynamicAssembly($name, 'Run')
    $module = $assembly.DefineDynamicModule('PInvokeModule')
    $type = $module.DefineType('PInvokeType', "Public,BeforeFieldInit")

    ## Go through all of the parameters passed to us.  As we do this,
    ## we clone the user's inputs into another array that we will use for
    ## the P/Invoke call.
    $inputParameters = @()
    $refParameters = @()

    for($counter = 1; $counter -le $parameterTypes.Length; $counter++)
    {
        ## If an item is a PSReference, then the user
        ## wants an [out] parameter.
        if($parameterTypes[$counter - 1] -eq [Ref])
        {
            ## Remember which parameters are used for [Out] parameters
            $refParameters += $counter

            ## On the cloned array, we replace the PSReference type with the
            ## .Net reference type that represents the value of the PSReference,
            ## and the value with the value held by the PSReference.
            $parameterTypes[$counter - 1] =
                $parameters[$counter - 1].Value.GetType().MakeByRefType()
            $inputParameters += $parameters[$counter - 1].Value
        }
        else
        {
            ## Otherwise, just add their actual parameter to the
            ## input array.
            $inputParameters += $parameters[$counter - 1]
        }
    }

    ## Define the actual P/Invoke method, adding the [Out]
    ## attribute for any parameters that were originally [Ref]
    ## parameters.
    $method = $type.DefineMethod(
        $methodName, 'Public,HideBySig,Static,PinvokeImpl',
        $returnType, $parameterTypes)
    foreach($refParameter in $refParameters)
    {
        [void] $method.DefineParameter($refParameter, "Out", $null)
    }

    ## Apply the P/Invoke constructor
    $ctor = [Runtime.InteropServices.DllImportAttribute].GetConstructor([string])
    $attr = New-Object Reflection.Emit.CustomAttributeBuilder $ctor, $dllName
    $method.SetCustomAttribute($attr)

    ## Create the temporary type, and invoke the method.
    $realType = $type.CreateType()

    $realType.InvokeMember(
        $methodName, 'Public,Static,InvokeMethod', $null, $null,$inputParameters)

    ## Finally, go through all of the reference parameters, and update the
    ## values of the PSReference objects that the user passed in.
    foreach($refParameter in $refParameters)
    {
        $parameters[$refParameter - 1].Value = $inputParameters[$refParameter - 1]
    }
}

function Measure-CommandPerformance
{
    ##############################################################################
    ##
    ## Measure-CommandPerformance
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Measures the average time of a command, accounting for natural variability by
    automatically ignoring the top and bottom ten percent.

    .EXAMPLE

    PS > Measure-CommandPerformance.ps1 { Start-Sleep -m 300 }

    Count    : 30
    Average  : 312.10155
    (...)

    #>

    param(
        ## The command to measure
        [Scriptblock] $Scriptblock,

        ## The number of times to measure the command's performance
        [int] $Iterations = 30
    )

    Set-StrictMode -Version 3

    ## Figure out how many extra iterations we need to account for the outliers
    $buffer = [int] ($iterations * 0.1)
    $totalIterations = $iterations + (2 * $buffer)

    ## Get the results
    $results = 1..$totalIterations |
        Foreach-Object { Measure-Command $scriptblock }

    ## Sort the results, and skip the outliers
    $middleResults = $results | Sort TotalMilliseconds |
        Select -Skip $buffer -First $iterations

    ## Show the average
    $middleResults | Measure-Object -Average TotalMilliseconds
}

function Move-LockedFile
{
    ##############################################################################
    ##
    ## Move-LockedFile
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Registers a locked file to be moved at the next system restart.

    .EXAMPLE

    PS > Move-LockedFile c:\temp\locked.txt c:\temp\locked.txt.bak

    #>

    param(
        ## The current location of the file to move
        $Path,

        ## The target location of the file
        $Destination
    )

    Set-StrictMode -Version 3

    ## Convert the the path and destination to fully qualified paths
    $path = (Resolve-Path $path).Path
    $destination = $executionContext.SessionState.`
        Path.GetUnresolvedProviderPathFromPSPath($destination)

    ## Define a new .NET type that calls into the Windows API to
    ## move a locked file.
    $MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
    $memberDefinition = @'
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern bool MoveFileEx(
        string lpExistingFileName, string lpNewFileName, int dwFlags);
'@
    $type = Add-Type -Name MoveFileUtils `
        -MemberDefinition $memberDefinition -PassThru

    ## Move the file
    $type::MoveFileEx($path, $destination, $MOVEFILE_DELAY_UNTIL_REBOOT)
}

function New-CommandWrapper
{
    ##############################################################################
    ##
    ## New-CommandWrapper
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Adds parameters and functionality to existing cmdlets and functions.

    .EXAMPLE

    New-CommandWrapper Get-Process `
          -AddParameter @{
              SortBy = {
                  $newPipeline = {
                      __ORIGINAL_COMMAND__ | Sort-Object -Property $SortBy
                  }
              }
          }

    This example adds a 'SortBy' parameter to Get-Process. It accomplishes
    this by adding a Sort-Object command to the pipeline.

    .EXAMPLE

    $parameterAttributes = @'
              [Parameter(Mandatory = $true)]
              [ValidateRange(50,75)]
              [Int]
'@

    New-CommandWrapper Clear-Host `
          -AddParameter @{
              @{
                  Name = 'MyMandatoryInt';
                  Attributes = $parameterAttributes
              } = {
                  Write-Host $MyMandatoryInt
                  Read-Host "Press ENTER"
             }
          }

    This example adds a new mandatory 'MyMandatoryInt' parameter to
    Clear-Host. This parameter is also validated to fall within the range
    of 50 to 75. It doesn't alter the pipeline, but does display some
    information on the screen before processing the original pipeline.

    #>

    param(
        ## The name of the command to extend
        [Parameter(Mandatory = $true)]
        $Name,

        ## Script to invoke before the command begins
        [ScriptBlock] $Begin,

        ## Script to invoke for each input element
        [ScriptBlock] $Process,

        ## Script to invoke at the end of the command
        [ScriptBlock] $End,

        ## Parameters to add, and their functionality.
        ##
        ## The Key of the hashtable can be either a simple parameter name,
        ## or a more advanced parameter description.
        ##
        ## If you want to add additional parameter validation (such as a
        ## parameter type,) then the key can itself be a hashtable with the keys
        ## 'Name' and 'Attributes'. 'Attributes' is the text you would use when
        ## defining this parameter as part of a function.
        ##
        ## The Value of each hashtable entry is a scriptblock to invoke
        ## when this parameter is selected. To customize the pipeline,
        ## assign a new scriptblock to the $newPipeline variable. Use the
        ## special text, __ORIGINAL_COMMAND__, to represent the original
        ## command. The $targetParameters variable represents a hashtable
        ## containing the parameters that will be passed to the original
        ## command.
        [HashTable] $AddParameter
    )

    Set-StrictMode -Version 3

    ## Store the target command we are wrapping, and its command type
    $target = $Name
    $commandType = "Cmdlet"

    ## If a function already exists with this name (perhaps it's already been
    ## wrapped,) rename the other function and chain to its new name.
    if(Test-Path function:\$Name)
    {
        $target = "$Name" + "-" + [Guid]::NewGuid().ToString().Replace("-","")
        Rename-Item function:\GLOBAL:$Name GLOBAL:$target
        $commandType = "Function"
    }

    ## The template we use for generating a command proxy
    $proxy = @'

    __CMDLET_BINDING_ATTRIBUTE__
    param(
    __PARAMETERS__
    )
    begin
    {
        try {
            __CUSTOM_BEGIN__

            ## Access the REAL Foreach-Object command, so that command
            ## wrappers do not interfere with this script
            $foreachObject = $executionContext.InvokeCommand.GetCmdlet(
                "Microsoft.PowerShell.Core\Foreach-Object")

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(
                '__COMMAND_NAME__',
                [System.Management.Automation.CommandTypes]::__COMMAND_TYPE__)

            ## TargetParameters represents the hashtable of parameters that
            ## we will pass along to the wrapped command
            $targetParameters = @{}
            $PSBoundParameters.GetEnumerator() |
                & $foreachObject {
                    if($command.Parameters.ContainsKey($_.Key))
                    {
                        $targetParameters.Add($_.Key, $_.Value)
                    }
                }

            ## finalPipeline represents the pipeline we wil ultimately run
            $newPipeline = { & $wrappedCmd @targetParameters }
            $finalPipeline = $newPipeline.ToString()

            __CUSTOM_PARAMETER_PROCESSING__

            $steppablePipeline = [ScriptBlock]::Create(
                $finalPipeline).GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        } catch {
            throw
        }
    }

    process
    {
        try {
            __CUSTOM_PROCESS__
            $steppablePipeline.Process($_)
        } catch {
            throw
        }
    }

    end
    {
        try {
            __CUSTOM_END__
            $steppablePipeline.End()
        } catch {
            throw
        }
    }

    dynamicparam
    {
        ## Access the REAL Get-Command, Foreach-Object, and Where-Object
        ## commands, so that command wrappers do not interfere with this script
        $getCommand = $executionContext.InvokeCommand.GetCmdlet(
            "Microsoft.PowerShell.Core\Get-Command")
        $foreachObject = $executionContext.InvokeCommand.GetCmdlet(
            "Microsoft.PowerShell.Core\Foreach-Object")
        $whereObject = $executionContext.InvokeCommand.GetCmdlet(
            "Microsoft.PowerShell.Core\Where-Object")

        ## Find the parameters of the original command, and remove everything
        ## else from the bound parameter list so we hide parameters the wrapped
        ## command does not recognize.
        $command = & $getCommand __COMMAND_NAME__ -Type __COMMAND_TYPE__
        $targetParameters = @{}
        $PSBoundParameters.GetEnumerator() |
            & $foreachObject {
                if($command.Parameters.ContainsKey($_.Key))
                {
                    $targetParameters.Add($_.Key, $_.Value)
                }
            }

        ## Get the argumment list as it would be passed to the target command
        $argList = @($targetParameters.GetEnumerator() |
            Foreach-Object { "-$($_.Key)"; $_.Value })

        ## Get the dynamic parameters of the wrapped command, based on the
        ## arguments to this command
        $command = $null
        try
        {
            $command = & $getCommand __COMMAND_NAME__ -Type __COMMAND_TYPE__ `
                -ArgumentList $argList
        }
        catch
        {

        }

        $dynamicParams = @($command.Parameters.GetEnumerator() |
            & $whereObject { $_.Value.IsDynamic })

        ## For each of the dynamic parameters, add them to the dynamic
        ## parameters that we return.
        if ($dynamicParams.Length -gt 0)
        {
            $paramDictionary = `
                New-Object Management.Automation.RuntimeDefinedParameterDictionary
            foreach ($param in $dynamicParams)
            {
                $param = $param.Value
                $arguments = $param.Name, $param.ParameterType, $param.Attributes
                $newParameter = `
                    New-Object Management.Automation.RuntimeDefinedParameter `
                    $arguments
                $paramDictionary.Add($param.Name, $newParameter)
            }
            return $paramDictionary
        }
    }

    <#

    .ForwardHelpTargetName __COMMAND_NAME__
    .ForwardHelpCategory __COMMAND_TYPE__

    #>

'@

    ## Get the information about the original command
    $originalCommand = Get-Command $target
    $metaData = New-Object System.Management.Automation.CommandMetaData `
        $originalCommand
    $proxyCommandType = [System.Management.Automation.ProxyCommand]

    ## Generate the cmdlet binding attribute, and replace information
    ## about the target
    $proxy = $proxy.Replace("__CMDLET_BINDING_ATTRIBUTE__",
        $proxyCommandType::GetCmdletBindingAttribute($metaData))
    $proxy = $proxy.Replace("__COMMAND_NAME__", $target)
    $proxy = $proxy.Replace("__COMMAND_TYPE__", $commandType)

    ## Stores new text we'll be putting in the param() block
    $newParamBlockCode = ""

    ## Stores new text we'll be putting in the begin block
    ## (mostly due to parameter processing)
    $beginAdditions = ""

    ## If the user wants to add a parameter
    $currentParameter = $originalCommand.Parameters.Count
    if($AddParameter)
    {
        foreach($parameter in $AddParameter.Keys)
        {
            ## Get the code associated with this parameter
            $parameterCode = $AddParameter[$parameter]

            ## If it's an advanced parameter declaration, the hashtable
            ## holds the validation and / or type restrictions
            if($parameter -is [Hashtable])
            {
                ## Add their attributes and other information to
                ## the variable holding the parameter block additions
                if($currentParameter -gt 0)
                {
                    $newParamBlockCode += ","
                }

                $newParamBlockCode += "`n`n    " +
                    $parameter.Attributes + "`n" +
                    '    $' + $parameter.Name

                $parameter = $parameter.Name
            }
            else
            {
                ## If this is a simple parameter name, add it to the list of
                ## parameters. The proxy generation APIs will take care of
                ## adding it to the param() block.
                $newParameter =
                    New-Object System.Management.Automation.ParameterMetadata `
                        $parameter
                $metaData.Parameters.Add($parameter, $newParameter)
            }

            $parameterCode = $parameterCode.ToString()

            ## Create the template code that invokes their parameter code if
            ## the parameter is selected.
            $templateCode = @"

            if(`$PSBoundParameters['$parameter'])
            {
                $parameterCode

                ## Replace the __ORIGINAL_COMMAND__ tag with the code
                ## that represents the original command
                `$alteredPipeline = `$newPipeline.ToString()
                `$finalPipeline = `$alteredPipeline.Replace(
                    '__ORIGINAL_COMMAND__', `$finalPipeline)
            }
"@

            ## Add the template code to the list of changes we're making
            ## to the begin() section.
            $beginAdditions += $templateCode
            $currentParameter++
        }
    }

    ## Generate the param() block
    $parameters = $proxyCommandType::GetParamBlock($metaData)
    if($newParamBlockCode) { $parameters += $newParamBlockCode }
    $proxy = $proxy.Replace('__PARAMETERS__', $parameters)

    ## Update the begin, process, and end sections
    $proxy = $proxy.Replace('__CUSTOM_BEGIN__', $Begin)
    $proxy = $proxy.Replace('__CUSTOM_PARAMETER_PROCESSING__', $beginAdditions)
    $proxy = $proxy.Replace('__CUSTOM_PROCESS__', $Process)
    $proxy = $proxy.Replace('__CUSTOM_END__', $End)

    ## Save the function wrapper
    Write-Verbose $proxy
    Set-Content function:\GLOBAL:$NAME $proxy

    ## If we were wrapping a cmdlet, hide it so that it doesn't conflict with
    ## Get-Help and Get-Command
    if($commandType -eq "Cmdlet")
    {
        $originalCommand.Visibility = "Private"
    }
}

function New-DynamicVariable
{
    ##############################################################################
    ##
    ## New-DynamicVariable
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Creates a variable that supports scripted actions for its getter and setter

    .EXAMPLE

    PS > .\New-DynamicVariable GLOBAL:WindowTitle `
         -Getter { $host.UI.RawUI.WindowTitle } `
         -Setter { $host.UI.RawUI.WindowTitle = $args[0] }

    PS > $windowTitle
    Administrator: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
    PS > $windowTitle = "Test"
    PS > $windowTitle
    Test

    #>

    param(
        ## The name for the dynamic variable
        [Parameter(Mandatory = $true)]
        $Name,

        ## The scriptblock to invoke when getting the value of the variable
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $Getter,

        ## The scriptblock to invoke when setting the value of the variable
        [ScriptBlock] $Setter
    )

    Set-StrictMode -Version 3

    Add-Type @"
    using System;
    using System.Collections.ObjectModel;
    using System.Management.Automation;

    namespace Lee.Holmes
    {
        public class DynamicVariable : PSVariable
        {
            public DynamicVariable(
                string name,
                ScriptBlock scriptGetter,
                ScriptBlock scriptSetter)
                    : base(name, null, ScopedItemOptions.AllScope)
            {
                getter = scriptGetter;
                setter = scriptSetter;
            }
            private ScriptBlock getter;
            private ScriptBlock setter;

            public override object Value
            {
                get
                {
                    if(getter != null)
                    {
                        Collection<PSObject> results = getter.Invoke();
                        if(results.Count == 1)
                        {
                            return results[0];
                        }
                        else
                        {
                            PSObject[] returnResults =
                                new PSObject[results.Count];
                            results.CopyTo(returnResults, 0);
                            return returnResults;
                        }
                    }
                    else { return null; }
                }
                set
                {
                    if(setter != null) { setter.Invoke(value); }
                }
            }
        }
    }
"@

    ## If we've already defined the variable, remove it.
    if(Test-Path variable:\$name)
    {
        Remove-Item variable:\$name -Force
    }

    ## Set the new variable, along with its getter and setter.
    $executioncontext.SessionState.PSVariable.Set(
        (New-Object Lee.Holmes.DynamicVariable $name,$getter,$setter))
}

function New-FilesystemHardLink
{
    ##############################################################################
    ##
    ## New-FileSystemHardLink
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Create a new hard link, which allows you to create a new name by which you
    can access an existing file. Windows only deletes the actual file once
    you delete all hard links that point to it.

    .EXAMPLE

    PS > "Hello" > test.txt
    PS > dir test* | select name

    Name
    ----
    test.txt

    PS > .\New-FilesystemHardLink.ps1 test.txt test2.txt
    PS > type test2.txt
    Hello
    PS > dir test* | select name

    Name
    ----
    test.txt
    test2.txt

    #>

    param(
        ## The existing file that you want the new name to point to
        [string] $Path,

        ## The new filename you want to create
        [string] $Destination
    )

    Set-StrictMode -Version 3

    ## Ensure that the provided names are absolute paths
    $filename = $executionContext.SessionState.`
        Path.GetUnresolvedProviderPathFromPSPath($destination)
    $existingFilename = Resolve-Path $path

    ## Prepare the parameter types and parameters for the CreateHardLink function
    $parameterTypes = [string], [string], [IntPtr]
    $parameters = [string] $filename, [string] $existingFilename, [IntPtr]::Zero

    ## Call the CreateHardLink method in the Kernel32 DLL
    $currentDirectory = Split-Path $myInvocation.MyCommand.Path
    $invokeWindowsApiCommand = Join-Path $currentDirectory Invoke-WindowsApi.ps1
    $result = & $invokeWindowsApiCommand "kernel32" `
        ([bool]) "CreateHardLink" $parameterTypes $parameters

    ## Provide an error message if the call fails
    if(-not $result)
    {
        $message = "Could not create hard link of $filename to " +
            "existing file $existingFilename"
        Write-Error $message
    }
}

function New-GenericObject
{
    ##############################################################################
    ##
    ## New-GenericObject
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Creates an object of a generic type. If using PowerShell version three,
    use this syntax:

    PS > $r = New-Object "System.Collections.Generic.Dictionary[String, Int32]"

    .EXAMPLE

    PS > New-GenericObject System.Collections.ObjectModel.Collection System.Int32
    Creates a simple generic collection

    .EXAMPLE

    PS > New-GenericObject System.Collections.Generic.Dictionary `
          System.String,System.Int32
    Creates a generic dictionary with two types

    .EXAMPLE

    PS > $secondType = New-GenericObject System.Collections.Generic.List Int32
    PS > New-GenericObject System.Collections.Generic.Dictionary `
          System.String,$secondType.GetType()
    Creates a generic list as the second type to a generic dictionary

    .EXAMPLE

    PS > New-GenericObject System.Collections.Generic.LinkedListNode `
          System.String "Hi"
    Creates a generic type with a non-default constructor

    #>

    param(
        ## The generic type to create
        [Parameter(Mandatory = $true)]
        [string] $TypeName,

        ## The types that should be applied to the generic object
        [Parameter(Mandatory = $true)]
        [string[]] $TypeParameters,

        ## Arguments to be passed to the constructor
        [object[]] $ConstructorParameters
    )

    Set-StrictMode -Version 2

    ## Create the generic type name
    $genericTypeName = $typeName + '`' + $typeParameters.Count
    $genericType = [Type] $genericTypeName

    if(-not $genericType)
    {
        throw "Could not find generic type $genericTypeName"
    }

    ## Bind the type arguments to it
    [type[]] $typedParameters = $typeParameters
    $closedType = $genericType.MakeGenericType($typedParameters)
    if(-not $closedType)
    {
        throw "Could not make closed type $genericType"
    }

    ## Create the closed version of the generic type
    ,[Activator]::CreateInstance($closedType, $constructorParameters)
}

function New-SelfSignedCertificate
{
    ##############################################################################
    ##
    ## New-SelfSignedCertificate
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Generate a new self-signed certificate. The certificate generated by these
    commands allow you to sign scripts on your own computer for protection
    from tampering. Files signed with this signature are not valid on other
    computers.

    .EXAMPLE

    PS > New-SelfSignedCertificate.ps1
    Creates a new self-signed certificate

    #>

    Set-StrictMode -Version 3

    ## Ensure we can find makecert.exe
    if(-not (Get-Command makecert.exe -ErrorAction SilentlyContinue))
    {
        $errorMessage = "Could not find makecert.exe. " +
            "This tool is available as part of Visual Studio, or the Windows SDK."

        Write-Error $errorMessage
        return
    }

    $keyPath = Join-Path ([IO.Path]::GetTempPath()) "root.pvk"

    ## Generate the local certification authority
    makecert -n "CN=PowerShell Local Certificate Root" -a sha1 `
        -eku 1.3.6.1.5.5.7.3.3 -r -sv $keyPath root.cer `
        -ss Root -sr localMachine

    ## Use the local certification authority to generate a self-signed
    ## certificate
    makecert -pe -n "CN=PowerShell User" -ss MY -a sha1 `
        -eku 1.3.6.1.5.5.7.3.3 -iv $keyPath -ic root.cer

    ## Remove the private key from the filesystem.
    Remove-Item $keyPath

    ## Retrieve the certificate
    Get-ChildItem cert:\currentuser\my -codesign |
        Where-Object { $_.Subject -match "PowerShell User" }
}

function New-ZipFile
{
    ##############################################################################
    ##
    ## New-ZipFile
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Create a Zip file from any files piped in.

    .EXAMPLE

    PS > dir *.ps1 | New-ZipFile scripts.zip
    Copies all PS1 files in the current directory to scripts.zip

    .EXAMPLE

    PS > "readme.txt" | New-ZipFile docs.zip
    Copies readme.txt to docs.zip

    #>

    param(
        ## The name of the zip archive to create
        $Path = $(throw "Specify a zip file name"),

        ## Switch to delete the zip archive if it already exists.
        [Switch] $Force
    )

    Set-StrictMode -Version 3

    ## Create the Zip File
    $zipName = $executionContext.SessionState.`
        Path.GetUnresolvedProviderPathFromPSPath($Path)

    ## Check if the file exists already. If it does, check
    ## for -Force - generate an error if not specified.
    if(Test-Path $zipName)
    {
        if($Force)
        {
            Remove-Item $zipName -Force
        }
        else
        {
            throw "Item with specified name $zipName already exists."
        }
    }

    ## Add the DLL that helps with file compression
    Add-Type -Assembly System.IO.Compression.FileSystem

    try
    {
        ## Open the Zip archive
        $archive = [System.IO.Compression.ZipFile]::Open($zipName, "Create")

        ## Go through each file in the input, adding it to the Zip file
        ## specified
        foreach($file in $input)
        {
            ## Skip the current file if it is the zip file itself
            if($file.FullName -eq $zipName)
            {
                continue
            }

            ## Skip directories
            if($file.PSIsContainer)
            {
                continue
            }

            $item = $file | Get-Item
            $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive, $item.FullName, $item.Name)
        }
    }
    finally
    {
        ## Close the file
        $archive.Dispose()
        $archive = $null
    }
}

function Read-HostWithPrompt
{
    #############################################################################
    ##
    ## Read-HostWithPrompt
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Read user input, with choices restricted to the list of options you
    provide.

    .EXAMPLE

    PS > $caption = "Please specify a task"
    PS > $message = "Specify a task to run"
    PS > $option = "&Clean Temporary Files","&Defragment Hard Drive"
    PS > $helptext = "Clean the temporary files from the computer",
    >>              "Run the defragment task"
    >>
    PS > $default = 1
    PS > Read-HostWithPrompt $caption $message $option $helptext $default

    Please specify a task
    Specify a task to run
    [C] Clean Temporary Files  [D] Defragment Hard Drive  [?] Help
    (default is "D"):?
    C - Clean the temporary files from the computer
    D - Run the defragment task
    [C] Clean Temporary Files  [D] Defragment Hard Drive  [?] Help
    (default is "D"):C
    0

    #>

    param(
        ## The caption for the prompt
        $Caption = $null,

        ## The message to display in the prompt
        $Message = $null,

        ## Options to provide in the prompt
        [Parameter(Mandatory = $true)]
        $Option,

        ## Any help text to provide
        $HelpText = $null,

        ## The default choice
        $Default = 0
    )

    Set-StrictMode -Version 3

    ## Create the list of choices
    $choices = New-Object `
        Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]

    ## Go through each of the options, and add them to the choice collection
    for($counter = 0; $counter -lt $option.Length; $counter++)
    {
        $choice = New-Object Management.Automation.Host.ChoiceDescription `
            $option[$counter]

        if($helpText -and $helpText[$counter])
        {
            $choice.HelpMessage = $helpText[$counter]
        }

        $choices.Add($choice)
    }

    ## Prompt for the choice, returning the item the user selected
    $host.UI.PromptForChoice($caption, $message, $choices, $default)
}

function Read-InputBox
{
    ##############################################################################
    ##
    ## Read-InputBox
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Read user input from a visual InputBox

    .EXAMPLE

    PS > $caption = "Please enter your name"
    PS > $name = Read-InputBox $caption

    #>

    param(
        ## The title of the dialog when displayed
        [string] $Title = "Input Dialog"
    )

    Set-StrictMode -Version 3

    ## Load the Windows Forms assembly
    Add-Type -Assembly System.Windows.Forms

    ## Create the main form
    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size @(400,100)
    $form.FormBorderStyle = "FixedToolWindow"

    ## Create the listbox to hold the items from the pipeline
    $textbox = New-Object Windows.Forms.TextBox
    $textbox.Top = 5
    $textbox.Left = 5
    $textBox.Width = 380
    $textbox.Anchor = "Left","Right"
    $form.Text = $Title

    ## Create the button panel to hold the OK and Cancel buttons
    $buttonPanel = New-Object Windows.Forms.Panel
    $buttonPanel.Size = New-Object Drawing.Size @(400,40)
    $buttonPanel.Dock = "Bottom"

    ## Create the Cancel button, which will anchor to the bottom right
    $cancelButton = New-Object Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = "Cancel"
    $cancelButton.Top = $buttonPanel.Height - $cancelButton.Height - 10
    $cancelButton.Left = $buttonPanel.Width - $cancelButton.Width - 10
    $cancelButton.Anchor = "Right"

    ## Create the OK button, which will anchor to the left of Cancel
    $okButton = New-Object Windows.Forms.Button
    $okButton.Text = "Ok"
    $okButton.DialogResult = "Ok"
    $okButton.Top = $cancelButton.Top
    $okButton.Left = $cancelButton.Left - $okButton.Width - 5
    $okButton.Anchor = "Right"

    ## Add the buttons to the button panel
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)

    ## Add the button panel and list box to the form, and also set
    ## the actions for the buttons
    $form.Controls.Add($buttonPanel)
    $form.Controls.Add($textbox)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.Add_Shown( { $form.Activate(); $textbox.Focus() } )

    ## Show the form, and wait for the response
    $result = $form.ShowDialog()

    ## If they pressed OK (or Enter,) go through all the
    ## checked items and send the corresponding object down the pipeline
    if($result -eq "OK")
    {
        $textbox.Text
    }
}

function Register-TemporaryEvent
{
    ##############################################################################
    ##
    ## Register-TemporaryEvent
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Registers an event action for an object, and automatically unregisters
    itself afterward. In PowerShell version three, use the -MaxTriggerCount
    parameter of the Register-*Event cmdlets.

    .EXAMPLE

    PS > $timer = New-Object Timers.Timer
    PS > Register-TemporaryEvent $timer Disposed { [Console]::Beep(100,100) }
    PS > $timer.Dispose()
    PS > Get-EventSubscriber
    PS > Get-Job

    #>

    param(
        ## The object that generates the event
        $Object,

        ## The event to subscribe to
        $Event,

        ## The action to invoke when the event arrives
        [ScriptBlock] $Action
    )

    Set-StrictMode -Version 2

    $actionText = $action.ToString()
    $actionText += @'

    $eventSubscriber | Unregister-Event
    $eventSubscriber.Action | Remove-Job
'@

    $eventAction = [ScriptBlock]::Create($actionText)
    $null = Register-ObjectEvent $object $event -Action $eventAction
}

function Resolve-Error
{
    #############################################################################
    ##
    ## Resolve-Error
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Displays detailed information about an error and its context.

    #>

    param(
        ## The error to resolve
        $ErrorRecord = ($error[0])
    )

    Set-StrictMode -Off

    ""
    "If this is an error in a script you wrote, use the Set-PsBreakpoint cmdlet"
    "to diagnose it."
    ""

    'Error details ($error[0] | Format-List * -Force)'
    "-"*80
    $errorRecord | Format-List * -Force

    'Information about the command that caused this error ' +
        '($error[0].InvocationInfo | Format-List *)'
    "-"*80
    $errorRecord.InvocationInfo | Format-List *

    'Information about the error''s target ' +
        '($error[0].TargetObject | Format-List *)'
    "-"*80
    $errorRecord.TargetObject | Format-List *

    'Exception details ($error[0].Exception | Format-List * -Force)'
    "-"*80

    $exception = $errorRecord.Exception

    for ($i = 0; $exception; $i++, ($exception = $exception.InnerException))
    {
        "$i" * 80
        $exception | Format-List * -Force
    }
}

function Search-Bing
{
    ##############################################################################
    ##
    ## Search-Bing
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search Bing for a given term

    .EXAMPLE

    PS > Search-Bing PowerShell
    Searches Bing for the term "PowerShell"

    #>

    param(
        ## The term to search for
        $Pattern = "PowerShell"
    )

    Set-StrictMode -Version 3

    ## Create the URL that contains the Twitter search results
    Add-Type -Assembly System.Web
    $queryUrl = 'http://www.bing.com/search?q={0}'
    $queryUrl = $queryUrl -f ([System.Web.HttpUtility]::UrlEncode($pattern))

    ## Download the web page
    $results = [string] (Invoke-WebRequest $queryUrl)

    ## Extract the text of the results, which are contained in
    ## segments that look like "<div class="sb_tlst">...</div>"
    $matches = $results |
        Select-String -Pattern '(?s)<div[^>]*sb_tlst[^>]*>.*?</div>' -AllMatches

    foreach($match in $matches.Matches)
    {
        ## Extract the URL, keeping only the text inside the quotes
        ## of the HREF
        $url = $match.Value -replace '.*href="(.*?)".*','$1'
        $url = [System.Web.HttpUtility]::UrlDecode($url)

        ## Extract the page name,  replace anything in angle
        ## brackets with an empty string.
        $item = $match.Value -replace '<[^>]*>', ''

        ## Output the item
        [PSCustomObject] @{ Item = $item; Url = $url }
    }
}

function Search-CertificateStore
{
    ##############################################################################
    ##
    ## Search-CertificateStore
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search the certificate provider for certificates that match the specified
    Enhanced Key Usage (EKU.)

    .EXAMPLE

    PS > Search-CertificateStore "Encrypting File System"
    Searches the certificate store for Encrypting File System certificates

    #>

    param(
        ## The friendly name of an Enhanced Key Usage
        ## (such as 'Code Signing')
        [Parameter(Mandatory = $true)]
        $EkuName
    )

    Set-StrictMode -Off

    ## Go through every certificate in the current user's "My" store
    foreach($cert in Get-ChildItem cert:\CurrentUser\My)
    {
        ## For each of those, go through its extensions
        foreach($extension in $cert.Extensions)
        {
            ## For each extension, go through its Enhanced Key Usages
            foreach($certEku in $extension.EnhancedKeyUsages)
            {
                ## If the friendly name matches, output that certificate
                if($certEku.FriendlyName -eq $ekuName)
                {
                    $cert
                }
            }
        }
    }
}

function Search-Help
{
    ##############################################################################
    ##
    ## Search-Help
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search the PowerShell help documentation for a given keyword or regular
    expression. For simple keyword searches in PowerShell version two or three,
    simply use "Get-Help <keyword>"

    .EXAMPLE

    PS > Search-Help hashtable
    Searches help for the term 'hashtable'

    .EXAMPLE

    PS > Search-Help "(datetime|ticks)"
    Searches help for the term datetime or ticks, using the regular expression
    syntax.

    #>

    param(
        ## The pattern to search for
        [Parameter(Mandatory = $true)]
        $Pattern
    )

    $helpNames = $(Get-Help * | Where-Object { $_.Category -ne "Alias" })

    ## Go through all of the help topics
    foreach($helpTopic in $helpNames)
    {
        ## Get their text content, and
        $content = Get-Help -Full $helpTopic.Name | Out-String
        if($content -match "(.{0,30}$pattern.{0,30})")
        {
            $helpTopic | Add-Member NoteProperty Match $matches[0].Trim()
            $helpTopic | Select-Object Name,Match
        }
    }
}

function Search-Registry
{
    ##############################################################################
    ##
    ## Search-Registry
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search the registry for keys or properties that match a specific value.

    .EXAMPLE

    PS > Set-Location HKCU:\Software\Microsoft\
    PS > Search-Registry Run

    #>

    param(
        ## The text to search for
        [Parameter(Mandatory = $true)]
        [string] $Pattern
    )

    Set-StrictMode -Off

    ## Helper function to create a new object that represents
    ## a registry match from this script
    function New-RegistryMatch
    {
        param( $matchType, $keyName, $propertyName, $line )

        $registryMatch = New-Object PsObject -Property @{
            MatchType = $matchType;
            KeyName = $keyName;
            PropertyName = $propertyName;
            Line = $line
        }

        $registryMatch
    }

    ## Go through each item in the registry
    foreach($item in Get-ChildItem -Recurse -ErrorAction SilentlyContinue)
    {
        ## Check if the key name matches
        if($item.Name -match $pattern)
        {
            New-RegistryMatch "Key" $item.Name $null $item.Name
        }

        ## Check if a key property matches
        foreach($property in (Get-ItemProperty $item.PsPath).PsObject.Properties)
        {
            ## Skip the property if it was one PowerShell added
            if(($property.Name -eq "PSPath") -or
                ($property.Name -eq "PSChildName"))
            {
                continue
            }

            ## Search the text of the property
            $propertyText = "$($property.Name)=$($property.Value)"
            if($propertyText -match $pattern)
            {
                New-RegistryMatch "Property" $item.Name `
                    property.Name $propertyText
            }
        }
    }
}

function Search-StackOverflow
{
    ##############################################################################
    ##
    ## Search-StackOverflow
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Searches Stack Overflow for PowerShell questions that relate to your
    search term, and provides the link to the accepted answer.


    .EXAMPLE

    PS > Search-StackOverflow upload ftp
    Searches StackOverflow for questions about how to upload FTP files

    .EXAMPLE

    PS > $answers = Search-StackOverflow.ps1 upload ftp
    PS > $answers | Out-GridView -PassThru | Foreach-Object { start $_ }

    Launches Out-GridView with the answers from a search. Select the URLs
    that you want to launch, and then press OK. PowerShell then launches
    your default web brower for those URLs.

    #>

    Set-StrictMode -Off
    Add-Type -Assembly System.Web

    $query = ($args | Foreach-Object { '"' + $_ + '"' }) -join " "
    $query = [System.Web.HttpUtility]::UrlEncode($query)

    ## Use the StackOverflow API to retrieve the answer for a question
    $url = "https://api.stackexchange.com/2.0/search?order=desc&sort=relevance" +
        "&pagesize=5&tagged=powershell&intitle=$query&site=stackoverflow"
    $question = Invoke-RestMethod $url

    ## Now go through and show the questions and answers
    $question.Items | Where accepted_answer_id | Foreach-Object {
            "Question: " + $_.Title
            "http://www.stackoverflow.com/questions/$($_.accepted_answer_id)"
            ""
    }
}

function Search-StartMenu
{
    ##############################################################################
    ##
    ## Search-StartMenu
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/blog)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search the Start Menu for items that match the provided text. This script
    searches both the name (as displayed on the Start Menu itself,) and the
    destination of the link.

    .EXAMPLE

    PS > Search-StartMenu "Character Map" | Invoke-Item
    Searches for the "Character Map" appication, and then runs it

    PS > Search-StartMenu PowerShell | Select-FilteredObject | Invoke-Item
    Searches for anything with "PowerShell" in the application name, lets you
    pick which one to launch, and then launches it.

    #>

    param(
        ## The pattern to match
        [Parameter(Mandatory = $true)]
        $Pattern
    )

    Set-StrictMode -Version 3

    ## Get the locations of the start menu paths
    $myStartMenu = [Environment]::GetFolderPath("StartMenu")
    $shell = New-Object -Com WScript.Shell
    $allStartMenu = $shell.SpecialFolders.Item("AllUsersStartMenu")

    ## Escape their search term, so that any regular expression
    ## characters don't affect the search
    $escapedMatch = [Regex]::Escape($pattern)

    ## Search in "my start menu" for text in the link name or link destination
    dir $myStartMenu *.lnk -rec | Where-Object {
        ($_.Name -match "$escapedMatch") -or
        ($_ | Select-String "\\[^\\]*$escapedMatch\." -Quiet)
    }

    ## Search in "all start menu" for text in the link name or link destination
    dir $allStartMenu *.lnk -rec | Where-Object {
        ($_.Name -match "$escapedMatch") -or
        ($_ | Select-String "\\[^\\]*$escapedMatch\." -Quiet)
    }
}

function Search-WmiNamespace
{
    ##############################################################################
    ##
    ## Search-WmiNamespace
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Search the WMI classes installed on the system for the provided match text.

    .EXAMPLE

    PS > Search-WmiNamespace Registry
    Searches WMI for any classes or descriptions that mention "Registry"

    .EXAMPLE

    Search-WmiNamespace Process ClassName,PropertyName
    Searchs WMI for any classes or properties that mention "Process"

    .EXAMPLE

    Search-WmiNamespace CPU -Detailed
    Searches WMI for any class names, descriptions, or properties that mention
    "CPU"

    #>

    param(
        ## The pattern to search for
        [Parameter(Mandatory = $true)]
        [string] $Pattern,

        ## Switch parameter to look for class names, descriptions, or properties
        [switch] $Detailed,

        ## Switch parameter to look for class names, descriptions, properties, and
        ## property description.
        [switch] $Full,

        ## Custom match options.
        ## Supports any or all of the following match options:
        ## ClassName, ClassDescription, PropertyName, PropertyDescription
        [string[]] $MatchOptions = ("ClassName","ClassDescription")
    )

    Set-StrictMode -Off

    ## Helper function to create a new object that represents
    ## a Wmi match from this script
    function New-WmiMatch
    {
        param( $matchType, $className, $propertyName, $line )

        $wmiMatch = New-Object PsObject -Property @{
            MatchType = $matchType;
            ClassName = $className;
            PropertyName = $propertyName;
            Line = $line
        }

        $wmiMatch
    }

    ## If they've specified the -detailed or -full options, update
    ## the match options to provide them an appropriate amount of detail
    if($detailed)
    {
        $matchOptions = "ClassName","ClassDescription","PropertyName"
    }

    if($full)
    {
        $matchOptions =
            "ClassName","ClassDescription","PropertyName","PropertyDescription"
    }

    ## Verify that they specified only valid match options
    foreach($matchOption in $matchOptions)
    {
        $fullMatchOptions =
            "ClassName","ClassDescription","PropertyName","PropertyDescription"

        if($fullMatchOptions -notcontains $matchOption)
        {
            $error = "Cannot convert value {0} to a match option. " +
                "Specify one of the following values and try again. " +
                "The possible values are ""{1}""."
            $ofs = ", "
            throw ($error -f $matchOption, ([string] $fullMatchOptions))
        }
    }

    ## Go through all of the available classes on the computer
    foreach($class in Get-WmiObject -List -Rec)
    {
        ## Provide explicit get options, so that we get back descriptios
        ## as well
        $managementOptions = New-Object System.Management.ObjectGetOptions
        $managementOptions.UseAmendedQualifiers = $true
        $managementClass =
            New-Object Management.ManagementClass $class.Name,$managementOptions

        ## If they want us to match on class names, check if their text
        ## matches the class name
        if($matchOptions -contains "ClassName")
        {
            if($managementClass.Name -match $pattern)
            {
                New-WmiMatch "ClassName" `
                    $managementClass.Name $null $managementClass.__PATH
            }
        }

        ## If they want us to match on class descriptions, check if their text
        ## matches the class description
        if($matchOptions -contains "ClassDescription")
        {
            $description =
                $managementClass.Qualifiers |
                    foreach { if($_.Name -eq "Description") { $_.Value } }
            if($description -match $pattern)
            {
                New-WmiMatch "ClassDescription" `
                    $managementClass.Name $null $description
            }
        }

        ## Go through the properties of the class
        foreach($property in $managementClass.Properties)
        {
            ## If they want us to match on property names, check if their text
            ## matches the property name
            if($matchOptions -contains "PropertyName")
            {
                if($property.Name -match $pattern)
                {
                    New-WmiMatch "PropertyName" `
                        $managementClass.Name $property.Name $property.Name
                }
            }

            ## If they want us to match on property descriptions, check if
            ## their text matches the property name
            if($matchOptions -contains "PropertyDescription")
            {
                $propertyDescription =
                    $property.Qualifiers |
                        foreach { if($_.Name -eq "Description") { $_.Value } }
                if($propertyDescription -match $pattern)
                {
                    New-WmiMatch "PropertyDescription" `
                        $managementClass.Name $property.Name $propertyDescription
                }
            }
        }
    }
}

function Select-FilteredObject
{
    ##############################################################################
    ##
    ## Select-FilteredObject
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Provides an inteactive window to help you select complex sets of objects.
    To do this, it takes all the input from the pipeline, and presents it in a
    notepad window.  Keep any lines that represent objects you want to retain,
    delete the rest, then save the file and exit notepad.

    The script then passes the original objects that you kept along the
    pipeline.

    .EXAMPLE

    PS > Get-Process | Select-FilteredObject | Stop-Process -WhatIf
    Gets all of the processes running on the system, and displays them to you.
    After you've selected the ones you want to stop, it pipes those into the
    Stop-Process cmdlet.

    #>

    ## PowerShell runs your "begin" script block before it passes you any of the
    ## items in the pipeline.
    begin
    {
        Set-StrictMode -Version 3

        ## Create a temporary file
        $filename = [System.IO.Path]::GetTempFileName()

        ## Define a header in a "here-string" that explains how to interact with
        ## the file
        $header = @"
    ############################################################
    ## Keep any lines that represent obects you want to retain,
    ## and delete the rest.
    ##
    ## Once you finish selecting objects, save this file and
    ## exit.
    ############################################################

"@

        ## Place the instructions into the file
        $header > $filename

        ## Initialize the variables that will hold our list of objects, and
        ## a counter to help us keep track of the objects coming down the
        ## pipeline
        $objectList = @()
        $counter = 0
    }

    ## PowerShell runs your "process" script block for each item it passes down
    ## the pipeline. In this block, the "$_" variable represents the current
    ## pipeline object
    process
    {
        ## Add a line to the file, using PowerShell's format (-f) operator.
        ## When provided the ouput of Get-Process, for example, these lines look
        ## like:
        ## 30: System.Diagnostics.Process (powershell)
        "{0}: {1}" -f $counter,$_.ToString() >> $filename

        ## Add the object to the list of objects, and increment our counter.
        $objectList += $_
        $counter++
    }

    ## PowerShell runs your "end" script block once it completes passing all
    ## objects down the pipeline.
    end
    {
        ## Start notepad, then call the process's WaitForExit() method to
        ## pause the script until the user exits notepad.
        $process = Start-Process Notepad -Args $filename -PassThru
        $process.WaitForExit()

        ## Go over each line of the file
        foreach($line in (Get-Content $filename))
        {
            ## Check if the line is of the special format: numbers, followed by
            ## a colon, followed by extra text.
            if($line -match "^(\d+?):.*")
            {
                ## If it did match the format, then $matches[1] represents the
                ## number -- a counter into the list of objects we saved during
                ## the "process" section.
                ## So, we output that object from our list of saved objects.
                $objectList[$matches[1]]
            }
        }

        ## Finally, clean up the temporary file.
        Remove-Item $filename
    }
}

function Select-GraphicalFilteredObject
{
    ##############################################################################
    ##
    ## Select-GraphicalFilteredObject
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Display a Windows Form to help the user select a list of items piped in.
    Any selected items get passed along the pipeline.

    .EXAMPLE

    PS > dir | Select-GraphicalFilteredObject

      Directory: C:\

    Mode                LastWriteTime     Length Name
    ----                -------------     ------ ----
    d----         10/7/2006   4:30 PM            Documents and Settings
    d----         3/18/2007   7:56 PM            Windows

    #>

    Set-StrictMode -Version 2

    $objectArray = @($input)

    ## Ensure that they've piped information into the script
    if($objectArray.Count -eq 0)
    {
        Write-Error "This script requires pipeline input."
        return
    }

    ## Load the Windows Forms assembly
    Add-Type -Assembly System.Windows.Forms

    ## Create the main form
    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size @(600,600)

    ## Create the listbox to hold the items from the pipeline
    $listbox = New-Object Windows.Forms.CheckedListBox
    $listbox.CheckOnClick = $true
    $listbox.Dock = "Fill"
    $form.Text = "Select the list of objects you wish to pass down the pipeline"
    $listBox.Items.AddRange($objectArray)

    ## Create the button panel to hold the OK and Cancel buttons
    $buttonPanel = New-Object Windows.Forms.Panel
    $buttonPanel.Size = New-Object Drawing.Size @(600,30)
    $buttonPanel.Dock = "Bottom"

    ## Create the Cancel button, which will anchor to the bottom right
    $cancelButton = New-Object Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = "Cancel"
    $cancelButton.Top = $buttonPanel.Height - $cancelButton.Height - 5
    $cancelButton.Left = $buttonPanel.Width - $cancelButton.Width - 10
    $cancelButton.Anchor = "Right"

    ## Create the OK button, which will anchor to the left of Cancel
    $okButton = New-Object Windows.Forms.Button
    $okButton.Text = "Ok"
    $okButton.DialogResult = "Ok"
    $okButton.Top = $cancelButton.Top
    $okButton.Left = $cancelButton.Left - $okButton.Width - 5
    $okButton.Anchor = "Right"

    ## Add the buttons to the button panel
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)

    ## Add the button panel and list box to the form, and also set
    ## the actions for the buttons
    $form.Controls.Add($listBox)
    $form.Controls.Add($buttonPanel)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.Add_Shown( { $form.Activate() } )

    ## Show the form, and wait for the response
    $result = $form.ShowDialog()

    ## If they pressed OK (or Enter,) go through all the
    ## checked items and send the corresponding object down the pipeline
    if($result -eq "OK")
    {
        foreach($index in $listBox.CheckedIndices)
        {
            $objectArray[$index]
        }
    }
}

function Select-TextOutput
{
    ##############################################################################
    ##
    ## Select-TextOutput
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Searches the textual output of a command for a pattern.

    .EXAMPLE

    PS > Get-Service | Select-TextOutput audio
    Finds all references to "Audio" in the output of Get-Service

    #>

    param(
        ## The pattern to search for
        $Pattern
    )

    Set-StrictMode -Version 3
    $input | Out-String -Stream | Select-String $pattern
}

function Send-File
{
    ##############################################################################
    ##
    ## Send-File
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Sends a file to a remote session.

    .EXAMPLE

    PS > $session = New-PsSession leeholmes1c23
    PS > Send-File c:\temp\test.exe c:\temp\test.exe $session

    #>

    param(
        ## The path on the local computer
        [Parameter(Mandatory = $true)]
        $Source,

        ## The target path on the remote computer
        [Parameter(Mandatory = $true)]
        $Destination,

        ## The session that represents the remote computer
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession] $Session
    )

    Set-StrictMode -Version 3

    $remoteScript = {
        param($destination, $bytes)

        ## Convert the destination path to a full filesystem path (to support
        ## relative paths)
        $Destination = $executionContext.SessionState.`
            Path.GetUnresolvedProviderPathFromPSPath($Destination)

        ## Write the content to the new file
        $file = [IO.File]::Open($Destination, "OpenOrCreate")
        $null = $file.Seek(0, "End")
        $null = $file.Write($bytes, 0, $bytes.Length)
        $file.Close()
    }

    ## Get the source file, and then start reading its content
    $sourceFile = Get-Item $source

    ## Delete the previously-existing file if it exists
    Invoke-Command -Session $session {
        if(Test-Path $args[0]) { Remove-Item $args[0] }
    } -ArgumentList $Destination

    ## Now break it into chunks to stream
    Write-Progress -Activity "Sending $Source" -Status "Preparing file"

    $streamSize = 1MB
    $position = 0
    $rawBytes = New-Object byte[] $streamSize
    $file = [IO.File]::OpenRead($sourceFile.FullName)

    while(($read = $file.Read($rawBytes, 0, $streamSize)) -gt 0)
    {
        Write-Progress -Activity "Writing $Destination" `
            -Status "Sending file" `
            -PercentComplete ($position / $sourceFile.Length * 100)

        ## Ensure that our array is the same size as what we read
        ## from disk
        if($read -ne $rawBytes.Length)
        {
            [Array]::Resize( [ref] $rawBytes, $read)
        }

        ## And send that array to the remote system
        Invoke-Command -Session $session $remoteScript `
            -ArgumentList $destination,$rawBytes

        ## Ensure that our array is the same size as what we read
        ## from disk
        if($rawBytes.Length -ne $streamSize)
        {
            [Array]::Resize( [ref] $rawBytes, $streamSize)
        }

        [GC]::Collect()
        $position += $read
    }

    $file.Close()

    ## Show the result
    Invoke-Command -Session $session { Get-Item $args[0] } -ArgumentList $Destination
}

function Send-MailMessage
{
    ##############################################################################
    ##
    ## Send-MailMessage
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ## Illustrate the techniques used to send an email in PowerShell.
    ## In PowerShell version two, use the Send-MailMessage cmdlet.
    ##
    ## Example:
    ##
    ## PS > $body = @"
    ## >> Hi from another satisfied customer of The PowerShell Cookbook!
## >> "@
    ## >>
    ## PS > $to = "guide_feedback@leeholmes.com"
    ## PS > $subject = "Thanks for all of the scripts."
    ## PS > $mailHost = "mail.leeholmes.com"
    ## PS > Send-MailMessage $to $subject $body $mailHost
    ##
    ##############################################################################

    param(
        ## The recipient of the mail message
        [string[]] $To = $(throw "Please specify the destination mail address"),

        ## The subjecty of the message
        [string] $Subject = "<No Subject>",

        ## The body of the message
        [string] $Body = $(throw "Please specify the message content"),

        ## The SMTP host that will transmit the message
        [string] $SmtpHost = $(throw "Please specify a mail server."),

        ## The sender of the message
        [string] $From = "$($env:UserName)@example.com",

        ## Any credentials to supply
        $Credential,

        ## Determine whether an SSL connection should be used
        [Switch] $UseSSL
    )

    ## Create the mail message
    $email = New-Object System.Net.Mail.MailMessage

    ## Populate its fields
    foreach($mailTo in $to)
    {
        $email.To.Add($mailTo)
    }

    $email.From = $from
    $email.Subject = $subject
    $email.Body = $body

    ## Send the mail
    $client = New-Object System.Net.Mail.SmtpClient $smtpHost

    if(-not $Credential)
    {
        $client.UseDefaultCredentials = $true
    }
    else
    {
        $actualCred = Get-Credential $Credential
        $networkCred = New-Object System.Net.NetworkCredential
        $networkCred.Username = $actualCred.Username

        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($actualCred.Password)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBstr($bstr)

        $networkCred.Password = $password
        $client.Credentials = $networkCred
    }

    if($UseSSL)
    {
        $client.EnableSSL = $true
    }

    $client.Send($email)
}

function Send-TcpRequest
{
    ##############################################################################
    ##
    ## Send-TcpRequest
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Send a TCP request to a remote computer, and return the response.
    If you do not supply input to this script (via either the pipeline, or the
    -InputObject parameter,) the script operates in interactive mode.

    .EXAMPLE

    PS > $http = @"
      GET / HTTP/1.1
      Host:bing.com
      `n`n
"@

    $http | Send-TcpRequest bing.com 80

    #>

    param(
        ## The computer to connect to
        [string] $ComputerName = "localhost",

        ## A switch to determine if you just want to test the connection
        [switch] $Test,

        ## The port to use
        [int] $Port = 80,

        ## A switch to determine if the connection should be made using SSL
        [switch] $UseSSL,

        ## The input string to send to the remote host
        [string] $InputObject,

        ## The delay, in milliseconds, to wait between commands
        [int] $Delay = 100
    )

    Set-StrictMode -Version 3

    [string] $SCRIPT:output = ""

    ## Store the input into an array that we can scan over. If there was no input,
    ## then we will be in interactive mode.
    $currentInput = $inputObject
    if(-not $currentInput)
    {
        $currentInput = @($input)
    }
    $scriptedMode = ([bool] $currentInput) -or $test

    function Main
    {
        ## Open the socket, and connect to the computer on the specified port
        if(-not $scriptedMode)
        {
            write-host "Connecting to $computerName on port $port"
        }

        try
        {
            $socket = New-Object Net.Sockets.TcpClient($computerName, $port)
        }
        catch
        {
            if($test) { $false }
            else { Write-Error "Could not connect to remote computer: $_" }

            return
        }

        ## If we're just testing the connection, we've made the connection
        ## successfully, so just return $true
        if($test) { $true; return }

        ## If this is interactive mode, supply the prompt
        if(-not $scriptedMode)
        {
            write-host "Connected.  Press ^D followed by [ENTER] to exit.`n"
        }

        $stream = $socket.GetStream()

        ## If we wanted to use SSL, set up that portion of the connection
        if($UseSSL)
        {
            $sslStream = New-Object System.Net.Security.SslStream $stream,$false
            $sslStream.AuthenticateAsClient($computerName)
            $stream = $sslStream
        }

        $writer = new-object System.IO.StreamWriter $stream

        while($true)
        {
            ## Receive the output that has buffered so far
            $SCRIPT:output += GetOutput

            ## If we're in scripted mode, send the commands,
            ## receive the output, and exit.
            if($scriptedMode)
            {
                foreach($line in $currentInput)
                {
                    $writer.WriteLine($line)
                    $writer.Flush()
                    Start-Sleep -m $Delay
                    $SCRIPT:output += GetOutput
                }

                break
            }
            ## If we're in interactive mode, write the buffered
            ## output, and respond to input.
            else
            {
                if($output)
                {
                    foreach($line in $output.Split("`n"))
                    {
                        write-host $line
                    }
                    $SCRIPT:output = ""
                }

                ## Read the user's command, quitting if they hit ^D
                $command = read-host
                if($command -eq ([char] 4)) { break; }

                ## Otherwise, Write their command to the remote host
                $writer.WriteLine($command)
                $writer.Flush()
            }
        }

        ## Close the streams
        $writer.Close()
        $stream.Close()

        ## If we're in scripted mode, return the output
        if($scriptedMode)
        {
            $output
        }
    }

    ## Read output from a remote host
    function GetOutput
    {
        ## Create a buffer to receive the response
        $buffer = new-object System.Byte[] 1024
        $encoding = new-object System.Text.AsciiEncoding

        $outputBuffer = ""
        $foundMore = $false

        ## Read all the data available from the stream, writing it to the
        ## output buffer when done.
        do
        {
            ## Allow data to buffer for a bit
            start-sleep -m 1000

            ## Read what data is available
            $foundmore = $false
            $stream.ReadTimeout = 1000

            do
            {
                try
                {
                    $read = $stream.Read($buffer, 0, 1024)

                    if($read -gt 0)
                    {
                        $foundmore = $true
                        $outputBuffer += ($encoding.GetString($buffer, 0, $read))
                    }
                } catch { $foundMore = $false; $read = 0 }
            } while($read -gt 0)
        } while($foundmore)

        $outputBuffer
    }

    . Main
}

function Set-Clipboard
{
    #############################################################################
    ##
    ## Set-Clipboard
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Sends the given input to the Windows clipboard.

    .EXAMPLE

    PS > dir | Set-Clipboard
    This example sends the view of a directory listing to the clipboard

    .EXAMPLE

    PS > Set-Clipboard "Hello World"
    This example sets the clipboard to the string, "Hello World".

    #>

    param(
        ## The input to send to the clipboard
        [Parameter(ValueFromPipeline = $true)]
        [object[]] $InputObject
    )

    begin
    {
        Set-StrictMode -Version 3
        $objectsToProcess = @()
    }

    process
    {
        ## Collect everything sent to the script either through
        ## pipeline input, or direct input.
        $objectsToProcess += $inputObject
    }

    end
    {
        ## Convert the input objects to text
        $clipText = ($objectsToProcess | Out-String -Stream) -join "`r`n"

        ## And finally set the clipboard text
        Add-Type -Assembly PresentationCore
        [Windows.Clipboard]::SetText($clipText)
    }
}

function Set-ConsoleProperties
{
    ##############################################################################
    ##
    ## Set-ConsoleProperties
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Configures PowerShell windows launched through Start | Run to have the same
    appearance as those launched through the Start Menu. On Windows 8, this is not
    required.

    #>

    Set-StrictMode -Version 3

    Push-Location
    Set-Location HKCU:\Console
    New-Item '.\%SystemRoot%_system32_WindowsPowerShell_v1.0_powershell.exe'
    Set-Location '.\%SystemRoot%_system32_WindowsPowerShell_v1.0_powershell.exe'

    New-ItemProperty . ColorTable00 -type DWORD -value 0x00562401
    New-ItemProperty . ColorTable07 -type DWORD -value 0x00f0edee
    New-ItemProperty . FaceName -type STRING -value "Lucida Console"
    New-ItemProperty . FontFamily -type DWORD -value 0x00000036
    New-ItemProperty . FontSize -type DWORD -value 0x000c0000
    New-ItemProperty . FontWeight -type DWORD -value 0x00000190
    New-ItemProperty . HistoryNoDup -type DWORD -value 0x00000000
    New-ItemProperty . QuickEdit -type DWORD -value 0x00000001
    New-ItemProperty . ScreenBufferSize -type DWORD -value 0x0bb80078
    New-ItemProperty . WindowSize -type DWORD -value 0x00320078
    Pop-Location
}

function Set-PsBreakPointLastError
{
    Set-StrictMode -Version 3

    $lastError = $error[0]
    Set-PsBreakpoint $lastError.InvocationInfo.ScriptName `
        $lastError.InvocationInfo.ScriptLineNumber
}

function Set-RemoteRegistryKeyProperty
{
    ##############################################################################
    ##
    ## Set-RemoteRegistryKeyProperty
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Set the value of a remote registry key property

    .EXAMPLE

    PS >$registryPath =
        "HKLM:\software\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell"
    PS >Set-RemoteRegistryKeyProperty LEE-DESK $registryPath `
          "ExecutionPolicy" "RemoteSigned"

    #>

    param(
        ## The computer to connect to
        [Parameter(Mandatory = $true)]
        $ComputerName,

        ## The registry path to modify
        [Parameter(Mandatory = $true)]
        $Path,

        ## The property to modify
        [Parameter(Mandatory = $true)]
        $PropertyName,

        ## The value to set on the property
        [Parameter(Mandatory = $true)]
        $PropertyValue
    )

    Set-StrictMode -Version 3

    ## Validate and extract out the registry key
    if($path -match "^HKLM:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "LocalMachine", $computername)
    }
    elseif($path -match "^HKCU:\\(.*)")
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            "CurrentUser", $computername)
    }
    else
    {
        Write-Error ("Please specify a fully-qualified registry path " +
            "(i.e.: HKLM:\Software) of the registry key to open.")
        return
    }

    ## Open the key and set its value
    $key = $baseKey.OpenSubKey($matches[1], $true)
    $key.SetValue($propertyName, $propertyValue)

    ## Close the key and base keys
    $key.Close()
    $baseKey.Close()
}

function Show-ColorizedContent
{
    ##############################################################################
    ##
    ## Show-ColorizedContent
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Displays syntax highlighting, line numbering, and range highlighting for
    PowerShell scripts.

    .EXAMPLE

    PS > Show-ColorizedContent Invoke-MyScript.ps1

    001 | function Write-Greeting
    002 | {
    003 |     param($greeting)
    004 |     Write-Host "$greeting World"
    005 | }
    006 |
    007 | Write-Greeting "Hello"

    .EXAMPLE

    PS > Show-ColorizedContent Invoke-MyScript.ps1 -highlightRange (1..3+7)

    001 > function Write-Greeting
    002 > {
    003 >     param($greeting)
    004 |     Write-Host "$greeting World"
    005 | }
    006 |
    007 > Write-Greeting "Hello"

    #>

    param(
        ## The path to colorize
        [Parameter(Mandatory = $true)]
        $Path,

        ## The range of lines to highlight
        $HighlightRange = @(),

        ## Switch to exclude line numbers
        [Switch] $ExcludeLineNumbers
    )

    Set-StrictMode -Version 3

    ## Colors to use for the different script tokens.
    ## To pick your own colors:
    ## [Enum]::GetValues($host.UI.RawUI.ForegroundColor.GetType()) |
    ##     Foreach-Object { Write-Host -Fore $_ "$_" }
    $replacementColours = @{
        'Attribute' = 'DarkCyan'
        'Command' = 'Blue'
        'CommandArgument' = 'Magenta'
        'CommandParameter' = 'DarkBlue'
        'Comment' = 'DarkGreen'
        'GroupEnd' = 'Black'
        'GroupStart' = 'Black'
        'Keyword' = 'DarkBlue'
        'LineContinuation' = 'Black'
        'LoopLabel' = 'DarkBlue'
        'Member' = 'Black'
        'NewLine' = 'Black'
        'Number' = 'Magenta'
        'Operator' = 'DarkGray'
        'Position' = 'Black'
        'StatementSeparator' = 'Black'
        'String' = 'DarkRed'
        'Type' = 'DarkCyan'
        'Unknown' = 'Black'
        'Variable' = 'Red'
    }

    $highlightColor = "Red"
    $highlightCharacter = ">"
    $highlightWidth = 6
    if($excludeLineNumbers) { $highlightWidth = 0 }

    ## Read the text of the file, and tokenize it
    $content = Get-Content $Path -Raw
    $parsed = [System.Management.Automation.PsParser]::Tokenize(
        $content, [ref] $null) | Sort StartLine,StartColumn

    ## Write a formatted line -- in the format of:
    ## <Line Number> <Separator Character> <Text>
    function WriteFormattedLine($formatString, [int] $line)
    {
        if($excludeLineNumbers) { return }

        ## By default, write the line number in gray, and use
        ## a simple pipe as the separator
        $hColor = "DarkGray"
        $separator = "|"

        ## If we need to highlight the line, use the highlight
        ## color and highlight separator as the separator
        if($highlightRange -contains $line)
        {
            $hColor = $highlightColor
            $separator = $highlightCharacter
        }

        ## Write the formatted line
        $text = $formatString -f $line,$separator
        Write-Host -NoNewLine -Fore $hColor -Back White $text
    }

    ## Complete the current line with filler cells
    function CompleteLine($column)
    {
        ## Figure how much space is remaining
        $lineRemaining = $host.UI.RawUI.WindowSize.Width -
            $column - $highlightWidth + 1

        ## If we have less than 0 remaining, we've wrapped onto the
        ## next line. Add another buffer width worth of filler
        if($lineRemaining -lt 0)
        {
            $lineRemaining += $host.UI.RawUI.WindowSize.Width
        }

        Write-Host -NoNewLine -Back White (" " * $lineRemaining)
    }


    ## Write the first line of context information (line number,
    ## highlight character.)
    Write-Host
    WriteFormattedLine "{0:D3} {1} " 1

    ## Now, go through each of the tokens in the input
    ## script
    $column = 1
    foreach($token in $parsed)
    {
        $color = "Gray"

        ## Determine the highlighting colour for that token by looking
        ## in the hashtable that maps token types to their color
        $color = $replacementColours[[string]$token.Type]
        if(-not $color) { $color = "Gray" }

        ## If it's a newline token, write the next line of context
        ## information
        if(($token.Type -eq "NewLine") -or ($token.Type -eq "LineContinuation"))
        {
            CompleteLine $column
            WriteFormattedLine "{0:D3} {1} " ($token.StartLine + 1)
            $column = 1
        }
        else
        {
            ## Do any indenting
            if($column -lt $token.StartColumn)
            {
                $text = " " * ($token.StartColumn - $column)
                Write-Host -Back White -NoNewLine $text
                $column = $token.StartColumn
            }

            ## See where the token ends
            $tokenEnd = $token.Start + $token.Length - 1

            ## Handle the line numbering for multi-line strings and comments
            if(
                (($token.Type -eq "String") -or
                ($token.Type -eq "Comment")) -and
                ($token.EndLine -gt $token.StartLine))
            {
                ## Store which line we've started at
                $lineCounter = $token.StartLine

                ## Split the content of this token into its lines
                ## We use the start and end of the tokens to determine
                ## the position of the content, but use the content
                ## itself (rather than the token values) for manipulation.
                $stringLines = $(
                    -join $content[$token.Start..$tokenEnd] -split "`n")

                ## Go through each of the lines in the content
                foreach($stringLine in $stringLines)
                {
                    $stringLine = $stringLine.Trim()

                    ## If we're on a new line, fill the right hand
                    ## side of the line with spaces, and write the header
                    ## for the new line.
                    if($lineCounter -gt $token.StartLine)
                    {
                        CompleteLine $column
                        WriteFormattedLine "{0:D3} {1} " $lineCounter
                        $column = 1
                    }

                    ## Now write the text of the current line
                    Write-Host -NoNewLine -Fore $color -Back White $stringLine
                    $column += $stringLine.Length
                    $lineCounter++
                }
            }
            ## Write out a regular token
            else
            {
                ## We use the start and end of the tokens to determine
                ## the position of the content, but use the content
                ## itself (rather than the token values) for manipulation.
                $text = (-join $content[$token.Start..$tokenEnd])
                Write-Host -NoNewLine -Fore $color -Back White $text
            }

            ## Update our position in the column
            $column = $token.EndColumn
        }
    }

    CompleteLine $column
    Write-Host
}

function Show-Object
{
    #############################################################################
    ##
    ## Show-Object
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Provides a graphical interface to let you explore and navigate an object.


    .EXAMPLE

    PS > $ps = { Get-Process -ID $pid }.Ast
    PS > Show-Object $ps

    #>

    param(
        ## The object to examine
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    Set-StrictMode -Version 3

    Add-Type -Assembly System.Windows.Forms

    ## Figure out the variable name to use when displaying the
    ## object navigation syntax. To do this, we look through all
    ## of the variables for the one with the same object identifier.
    $rootVariableName = dir variable:\* -Exclude InputObject,Args |
        Where-Object {
            $_.Value -and
            ($_.Value.GetType() -eq $InputObject.GetType()) -and
            ($_.Value.GetHashCode() -eq $InputObject.GetHashCode())
    }

    ## If we got multiple, pick the first
    $rootVariableName = $rootVariableName| % Name | Select -First 1

    ## If we didn't find one, use a default name
    if(-not $rootVariableName)
    {
        $rootVariableName = "InputObject"
    }

    ## A function to add an object to the display tree
    function PopulateNode($node, $object)
    {
        ## If we've been asked to add a NULL object, just return
        if(-not $object) { return }

        ## If the object is a collection, then we need to add multiple
        ## children to the node
        if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object))
        {
            ## Some very rare collections don't support indexing (i.e.: $foo[0]).
            ## In this situation, PowerShell returns the parent object back when you
            ## try to access the [0] property.
            $isOnlyEnumerable = $object.GetHashCode() -eq $object[0].GetHashCode()

            ## Go through all the items
            $count = 0
            foreach($childObjectValue in $object)
            {
                ## Create the new node to add, with the node text of the item and
                ## value, along with its type
    	        $newChildNode = New-Object Windows.Forms.TreeNode
                $newChildNode.Text = "$($node.Name)[$count] = $childObjectValue"
                $newChildNode.ToolTipText = $childObjectValue.GetType()

                ## Use the node name to keep track of the actual property name
                ## and syntax to access that property.
                ## If we can't use the index operator to access children, add
                ## a special tag that we'll handle specially when displaying
                ## the node names.
                if($isOnlyEnumerable)
                {
                    $newChildNode.Name = "@"
                }

                $newChildNode.Name += "[$count]"
    	        $null = $node.Nodes.Add($newChildNode)

                ## If this node has children or properties, add a placeholder
                ## node underneath so that the node shows a '+' sign to be
                ## expanded.
                AddPlaceholderIfRequired $newChildNode $childObjectValue

                $count++
            }
        }
        else
        {
            ## If the item was not a collection, then go through its
            ## properties
            foreach($child in $object.PSObject.Properties)
            {
                ## Figure out the value of the property, along with
                ## its type.
    	        $childObject = $child.Value
                $childObjectType = $null
                if($childObject)
                {
                    $childObjectType = $childObject.GetType()
                }

                ## Create the new node to add, with the node text of the item and
                ## value, along with its type
    	        $childNode = New-Object Windows.Forms.TreeNode
                $childNode.Text = $child.Name + " = $childObject"
                $childNode.ToolTipText = $childObjectType
                if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($childObject))
                {
                    $childNode.ToolTipText += "[]"
                }

                $childNode.Name = $child.Name
    	        $null = $node.Nodes.Add($childNode)

                ## If this node has children or properties, add a placeholder
                ## node underneath so that the node shows a '+' sign to be
                ## expanded.
                AddPlaceholderIfRequired $childNode $childObject
            }
        }
    }

    ## A function to add a placeholder if required to a node.
    ## If there are any properties or children for this object, make a temporary
    ## node with the text "..." so that the node shows a '+' sign to be
    ## expanded.
    function AddPlaceholderIfRequired($node, $object)
    {
        if(-not $object) { return }

        if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object) -or
            @($object.PSObject.Properties))
        {
            $null = $node.Nodes.Add( (New-Object Windows.Forms.TreeNode "...") )
        }
    }

    ## A function invoked when a node is selected.
    function OnAfterSelect
    {
        param($Sender, $TreeViewEventArgs)

        ## Determine the selected node
        $nodeSelected = $Sender.SelectedNode

        ## Walk through its parents, creating the virtual
        ## PowerShell syntax to access this property.
        $nodePath = GetPathForNode $nodeSelected

        ## Now, invoke that PowerShell syntax to retrieve
        ## the value of the property.
        $resultObject = Invoke-Expression $nodePath
        $outputPane.Text = $nodePath

        ## If we got some output, put the object's member
        ## information in the text box.
        if($resultObject)
        {
            $members = Get-Member -InputObject $resultObject | Out-String
            $outputPane.Text += "`n" + $members
        }
    }

    ## A function invoked when the user is about to expand a node
    function OnBeforeExpand
    {
        param($Sender, $TreeViewCancelEventArgs)

        ## Determine the selected node
        $selectedNode = $TreeViewCancelEventArgs.Node

        ## If it has a child node that is the placeholder, clear
        ## the placehoder node.
        if($selectedNode.FirstNode -and
            ($selectedNode.FirstNode.Text -eq "..."))
        {
            $selectedNode.Nodes.Clear()
        }
        else
        {
            return
        }

        ## Walk through its parents, creating the virtual
        ## PowerShell syntax to access this property.
        $nodePath = GetPathForNode $selectedNode

        ## Now, invoke that PowerShell syntax to retrieve
        ## the value of the property.
        Invoke-Expression "`$resultObject = $nodePath"

        ## And populate the node with the result object.
        PopulateNode $selectedNode $resultObject
    }

    ## A function to handle key presses on the tree view.
    ## In this case, we capture ^C to copy the path of
    ## the object property that we're currently viewing.
    function OnTreeViewKeyPress
    {
        param($Sender, $KeyPressEventArgs)

        ## [Char] 3 = Control-C
        if($KeyPressEventArgs.KeyChar -eq 3)
        {
            $KeyPressEventArgs.Handled = $true

            ## Get the object path, and set it on the clipboard
            $node = $Sender.SelectedNode
            $nodePath = GetPathForNode $node
            [System.Windows.Forms.Clipboard]::SetText($nodePath)

            $form.Close()
        }
        elseif([System.Windows.Forms.Control]::ModifierKeys -eq "Control")
        {
            if($KeyPressEventArgs.KeyChar -eq '+')
            {
                $SCRIPT:currentFontSize++
                UpdateFonts $SCRIPT:currentFontSize

                $KeyPressEventArgs.Handled = $true
            }
            elseif($KeyPressEventArgs.KeyChar -eq '-')
            {
                $SCRIPT:currentFontSize--
                if($SCRIPT:currentFontSize -lt 1) { $SCRIPT:currentFontSize = 1 }
                UpdateFonts $SCRIPT:currentFontSize

                $KeyPressEventArgs.Handled = $true
            }
        }
    }

    ## A function to handle key presses on the form.
    ## In this case, we handle Ctrl-Plus and Ctrl-Minus
    ## to adjust font size.
    function OnKeyUp
    {
        param($Sender, $KeyUpEventArgs)

        if([System.Windows.Forms.Control]::ModifierKeys -eq "Control")
        {
            if($KeyUpEventArgs.KeyCode -in 'Add','OemPlus')
            {
                $SCRIPT:currentFontSize++
                UpdateFonts $SCRIPT:currentFontSize

                $KeyUpEventArgs.Handled = $true
            }
            elseif($KeyUpEventArgs.KeyCode -in 'Subtract','OemMinus')
            {
                $SCRIPT:currentFontSize--
                if($SCRIPT:currentFontSize -lt 1) { $SCRIPT:currentFontSize = 1 }
                UpdateFonts $SCRIPT:currentFontSize

                $KeyUpEventArgs.Handled = $true
            }
            elseif($KeyUpEventArgs.KeyCode -eq 'D0')
            {
                $SCRIPT:currentFontSize = 12
                UpdateFonts $SCRIPT:currentFontSize

                $KeyUpEventArgs.Handled = $true
            }
        }
    }

    ## A function to handle mouse wheel scrolling.
    ## In this case, we translate Ctrl-Wheel to zoom.
    function OnMouseWheel
    {
        param($Sender, $MouseEventArgs)

        if(
            ([System.Windows.Forms.Control]::ModifierKeys -eq "Control") -and
            ($MouseEventArgs.Delta -ne 0))
        {
            $SCRIPT:currentFontSize += ($MouseEventArgs.Delta / 120)
            if($SCRIPT:currentFontSize -lt 1) { $SCRIPT:currentFontSize = 1 }

            UpdateFonts $SCRIPT:currentFontSize
            $MouseEventArgs.Handled = $true
        }
    }

    ## A function to walk through the parents of a node,
    ## creating virtual PowerShell syntax to access this property.
    function GetPathForNode
    {
        param($Node)

        $nodeElements = @()

        ## Go through all the parents, adding them so that
        ## $nodeElements is in order.
        while($Node)
        {
            $nodeElements = ,$Node + $nodeElements
            $Node = $Node.Parent
        }

        ## Now go through the node elements
        $nodePath = ""
        foreach($Node in $nodeElements)
        {
            $nodeName = $Node.Name

            ## If it was a node that PowerShell is able to enumerate
            ## (but not index), wrap it in the array cast operator.
        if($nodeName.StartsWith('@'))
            {
                $nodeName = $nodeName.Substring(1)
            $nodePath = "@(" + $nodePath + ")"
            }
            elseif($nodeName.StartsWith('['))
            {
                ## If it's a child index, we don't need to
                ## add the dot for property access
            }
            elseif($nodePath)
            {
                ## Otherwise, we're accessing a property. Add a dot.
                $nodePath += "."
            }

            ## Append the node name to the path
            $nodePath += $nodeName
        }

        ## And return the result
        $nodePath
    }

    function UpdateFonts
    {
        param($fontSize)

        $treeView.Font = New-Object System.Drawing.Font "Consolas",$fontSize
        $outputPane.Font = New-Object System.Drawing.Font "Consolas",$fontSize
    }

    $SCRIPT:currentFontSize = 12

    ## Create the TreeView, which will hold our object navigation
    ## area.
    $treeView = New-Object Windows.Forms.TreeView
    $treeView.Dock = "Top"
    $treeView.Height = 500
    $treeView.PathSeparator = "."
    $treeView.ShowNodeToolTips = $true
    $treeView.Add_AfterSelect( { OnAfterSelect @args } )
    $treeView.Add_BeforeExpand( { OnBeforeExpand @args } )
    $treeView.Add_KeyPress( { OnTreeViewKeyPress @args } )

    ## Create the output pane, which will hold our object
    ## member information.
    $outputPane = New-Object System.Windows.Forms.TextBox
    $outputPane.Multiline = $true
    $outputPane.WordWrap = $false
    $outputPane.ScrollBars = "Both"
    $outputPane.Dock = "Fill"

    ## Create the root node, which represents the object
    ## we are trying to show.
    $root = New-Object Windows.Forms.TreeNode
    $root.ToolTipText = $InputObject.GetType()
    $root.Text = $InputObject
    $root.Name = '$' + $rootVariableName
    $root.Expand()
    $null = $treeView.Nodes.Add($root)

    UpdateFonts $currentFontSize

    ## And populate the initial information into the tree
    ## view.
    PopulateNode $root $InputObject

    ## Finally, create the main form and show it.
    $form = New-Object Windows.Forms.Form
    $form.Text = "Browsing " + $root.Text
    $form.Width = 1000
    $form.Height = 800
    $form.Controls.Add($outputPane)
    $form.Controls.Add($treeView)
    $form.Add_MouseWheel( { OnMouseWheel @args } )
    $treeView.Add_KeyUp( { OnKeyUp @args } )
    $treeView.Select()
    $null = $form.ShowDialog()
    $form.Dispose()
}

function Start-ProcessAsUser
{
    ##############################################################################
    ##
    ## Start-ProcessAsUser
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Launch a process under alternate credentials, providing functionality
    similar to runas.exe.

    .EXAMPLE

    PS > $file = Join-Path ([Environment]::GetFolderPath("System")) certmgr.msc
    PS > Start-ProcessAsUser Administrator mmc $file

    #>

    param(
        ## The credential to launch the process under
        $Credential = (Get-Credential),

        ## The process to start
        [Parameter(Mandatory = $true)]
        [string] $Process,

        ## Any arguments to pass to the process
        [string] $ArgumentList = ""
    )

    Set-StrictMode -Version 3

    ## Create a real credential if they supplied a username
    $credential = Get-Credential $credential

    ## Exit if they canceled out of the credential dialog
    if(-not ($credential -is "System.Management.Automation.PsCredential"))
    {
        return
    }

    ## Prepare the startup information (including username and password)
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.Filename = $process
    $startInfo.Arguments = $argumentList

    ## If we're launching as ourselves, set the "runas" verb
    if(($credential.Username -eq "$ENV:Username") -or
        ($credential.Username -eq "\$ENV:Username"))
    {
        $startInfo.Verb = "runas"
    }
    else
    {
        $startInfo.UserName = $credential.Username
        $startInfo.Password = $credential.Password
        $startInfo.UseShellExecute = $false
    }

    ## Start the process
    [Diagnostics.Process]::Start($startInfo)
}

function Test-Uri
{
    ##############################################################################
    ##
    ## Test-Uri
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Connects to a given URI and returns status about it: URI, response code,
    and time taken.

    .EXAMPLE

    PS > Test-Uri bing.com

    Uri               : bing.com
    StatusCode        : 200
    StatusDescription : OK
    ResponseLength    : 34001
    TimeTaken         : 459.0009

    #>

    param(
        ## The URI to test
        $Uri
    )

    $request = $null
    $time = try
    {
        ## Request the URI, and measure how long the response took.
    	$result = Measure-Command { $request = Invoke-WebRequest -Uri $uri }
        $result.TotalMilliseconds
    }
    catch
    {
        ## If the request generated an exception (i.e.: 500 server
        ## error or 404 not found), we can pull the status code from the
        ## Exception.Response property
        $request = $_.Exception.Response
        $time = -1
    }

    $result = [PSCustomObject] @{
        Time = Get-Date;
        Uri = $uri;
        StatusCode = [int] $request.StatusCode;
        StatusDescription = $request.StatusDescription;
        ResponseLength = $request.RawContentLength;
        TimeTaken = $time;
    }

    $result
}

function Use-Culture
{
    #############################################################################
    ##
    ## Use-Culture
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    #############################################################################

    <#

    .SYNOPSIS

    Invoke a scriptblock under the given culture

    .EXAMPLE

    PS > Use-Culture fr-FR { Get-Date -Date "25/12/2007" }
    mardi 25 decembre 2007 00:00:00

    #>

    param(
        ## The culture in which to evaluate the given script block
        [Parameter(Mandatory = $true)]
        [System.Globalization.CultureInfo] $Culture,

        ## The code to invoke in the context of the given culture
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock
    )

    Set-StrictMode -Version 3

    ## A helper function to set the current culture
    function Set-Culture([System.Globalization.CultureInfo] $culture)
    {
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    }

    ## Remember the original culture information
    $oldCulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture

    ## Restore the original culture information if
    ## the user's script encounters errors.
    trap { Set-Culture $oldCulture }

    ## Set the current culture to the user's provided
    ## culture.
    Set-Culture $culture

    ## Invoke the user's scriptblock
    & $ScriptBlock

    ## Restore the original culture information.
    Set-Culture $oldCulture
}

function Watch-Command
{
    ##############################################################################
    ##
    ## Watch-Command
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Watches the result of a command invocation, alerting you when the output
    either matches a specified string, lacks a specified string, or has simply
    changed.

    .EXAMPLE

    PS > Watch-Command { Get-Process -Name Notepad | Measure } -UntilChanged
    Monitors Notepad processes until you start or stop one.

    .EXAMPLE

    PS > Watch-Command { Get-Process -Name Notepad | Measure } -Until "Count    : 1"
    Monitors Notepad processes until there is exactly one open.

    .EXAMPLE

    PS > Watch-Command { Get-Process -Name Notepad | Measure } -While 'Count    : \d\s*\n'
    Monitors Notepad processes while there are between 0 and 9 open
    (once number after the colon).

    #>


    [CmdletBinding(DefaultParameterSetName = "Forever")]
    param(
        ## The scriptblock to invoke while monitoring
        [Parameter(Mandatory = $true, Position = 0)]
        [ScriptBlock] $ScriptBlock,

        ## The delay, in seconds, between monitoring attempts
        [Parameter()]
        [Double] $DelaySeconds = 1,

        ## Specifies that the alert sound should not be played
        [Parameter()]
        [Switch] $Quiet,

        ## Monitoring continues only until while the output of the
        ## command remains the same.
        [Parameter(ParameterSetName = "UntilChanged", Mandatory = $false)]
        [Switch] $UntilChanged,

        ## The regular expression to search for. Monitoring continues
        ## until this expression is found.
        [Parameter(ParameterSetName = "Until", Mandatory = $false)]
        [String] $Until,

        ## The regular expression to search for. Monitoring continues
        ## until this expression is not found.
        [Parameter(ParameterSetName = "While", Mandatory = $false)]
        [String] $While
    )

    Set-StrictMode -Version 3

    $initialOutput = ""

    ## Start a continuous loop
    while($true)
    {
        ## Run the provided script block
        $r = & $ScriptBlock

        ## Clear the screen and display the results
        Clear-Host
        $ScriptBlock.ToString().Trim()
        ""
        $textOutput = $r | Out-String
        $textOutput

        ## Remember the initial output, if we haven't
        ## stored it yet
        if(-not $initialOutput)
        {
            $initialOutput = $textOutput
        }

        ## If we are just looking for any change,
        ## see if the text has changed.
        if($UntilChanged)
        {
            if($initialOutput -ne $textOutput)
            {
                break
            }
        }

        ## If we need to ensure some text is found,
        ## break if we didn't find it.
        if($While)
        {
            if($textOutput -notmatch $While)
            {
                break
            }
        }

        ## If we need to wait for some text to be found,
        ## break if we find it.
        if($Until)
        {
            if($textOutput -match $Until)
            {
                break
            }
        }

        ## Delay
        Start-Sleep -Seconds $DelaySeconds
    }

    ## Notify the user
    if(-not $Quiet)
    {
        [Console]::Beep(1000, 1000)
    }
}

function Watch-DebugExpression
{
    #############################################################################
    ##
    ## Watch-DebugExpression
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    ##############################################################################

    <#

    .SYNOPSIS

    Updates your prompt to display the values of information you want to track.

    .EXAMPLE

    PS > Watch-DebugExpression { (Get-History).Count }

    Expression          Value
    ----------          -----
    (Get-History).Count     3

    PS > Watch-DebugExpression { $count }

    Expression          Value
    ----------          -----
    (Get-History).Count     4
    $count

    PS > $count = 100

    Expression          Value
    ----------          -----
    (Get-History).Count     5
    $count                100

    PS > Watch-DebugExpression -Reset
    PS >

    #>

    param(
        ## The expression to track
        [ScriptBlock] $ScriptBlock,

        ## Switch to no longer watch an expression
        [Switch] $Reset
    )

    Set-StrictMode -Version 3

    if($Reset)
    {
        Set-Item function:\prompt ([ScriptBlock]::Create($oldPrompt))

        Remove-Item variable:\expressionWatch
        Remove-Item variable:\oldPrompt

        return
    }

    ## Create the variableWatch variable if it doesn't yet exist
    if(-not (Test-Path variable:\expressionWatch))
    {
        $GLOBAL:expressionWatch = @()
    }

    ## Add the current variable name to the watch list
    $GLOBAL:expressionWatch += $scriptBlock

    ## Update the prompt to display the expression values,
    ## if needed.
    if(-not (Test-Path variable:\oldPrompt))
    {
        $GLOBAL:oldPrompt = Get-Content function:\prompt
    }

    if($oldPrompt -notlike '*$expressionWatch*')
    {
        $newPrompt = @'
            $results = foreach($expression in $expressionWatch)
            {
                New-Object PSObject -Property @{
                    Expression = $expression.ToString().Trim();
                    Value = & $expression
                } | Select Expression,Value
            }
            Write-Host "`n"
            Write-Host ($results | Format-Table -Auto | Out-String).Trim()
            Write-Host "`n"

'@

        $newPrompt += $oldPrompt

        Set-Item function:\prompt ([ScriptBlock]::Create($newPrompt))
    }
}