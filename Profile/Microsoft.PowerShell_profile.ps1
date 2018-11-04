$PSDefaultParameterValues = @{
    'Export-Csv:NoTypeInformation' = $true
    'Install-Module:Scope' = 'CurrentUser'
    'Out-Default:OutVariable' = 'LastOut'
    'Connect-Meraki:ApiKey' = (Get-PSFConfigValue merakips.apikey)
}


#region modules
$modulesToImport = @(
    'PSUtil'
    'HistoryPx'
    'PSConsoleTheme'
)

ForEach ($module in $modulesToImport) {
    $isValid = Get-Module $module -ListAvailable
    if ($null -eq $isValid) {Write-Warning "$module is not found in Modules path: $($env:PSModulePath)"}
    Import-Module $module
}
#endregion modules


#region functions
$Functions = Get-ChildItem $PSScriptRoot\Functions\*.ps1 -ErrorAction SilentlyContinue

foreach ($import in $Functions) {
    Try {
        . $import.fullname
    }
    Catch {
        Write-Error "Failed to import function $($import.fullname): $_"
    }
}
#endregion functions


#region one-off settings
Set-ConsoleTheme Flat
Set-PSReadlineOption -BellStyle None
#endregion one-off settings
