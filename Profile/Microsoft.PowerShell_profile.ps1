$PSDefaultParameterValues = @{

    'Export-Csv:NoTypeInformation' = $true
    'Install-Module:Scope' = CurrentUser

}

$modulesToImport = @(

    'PSUtil'

)

ForEach ($module in $modulesToImport) {

    $isValid = Get-Module $module -ListAvailable

    if ($null -eq $isValid) {Write-Warning "$module is not found in Modules path: $($env:PSModulePath)"}

}



$Functions = Get-ChildItem $PSScriptRoot\Functions\*.ps1 -ErrorAction SilentlyContinue

foreach ($import in $Functions) {

    Try {

        . $import.fullname

    }

    Catch {

        Write-Error "Failed to import function $($import.fullname): $_"

    }

}
