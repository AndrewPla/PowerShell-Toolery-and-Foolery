[cmdletbinding()]
param
(
	$baseUrl = 'https://psframework.org/documentation/commands/PSFramework/',
	
	$Path = 'C:\Repos\PowershellFrameworkCollective\psframework\PSFramework\functions'
)
$Files = Get-ChildItem $Path -Recurse -filter '*.ps1'




# Generate all the info for all functions
$Results = foreach ($File in $Files) {
	$CommandName = $File.basename
	$CommandUrl = $baseUrl + $CommandName
	$HelpUri = "[CmdletBinding(HelpUri = '$CommandUrl')]"
	$NewHelpUri = ", HelpUri = '$CommandUrl')"
	
	#region Grab the current Cmdletbinding
	$Content = Get-Content $File.FullName
	$cmdletbinding = ($Content | Select-String -Pattern 'cmdletbinding' | Out-String).trim()
	$CmdletBinding = $CmdletBinding.tostring()
	#endregion Grab the current Cmdletbinding
	
	#region Create the New CmdletBinding
	if ($cmdletbinding -like '*=*') {
		$NewCmdletBinding = $CmdletBinding.replace(')', "$NewHelpUri")
		
	}
	else {
		$NewCmdletBinding = $CmdletBinding.replace('[CmdletBinding()]', "$HelpUri")
	}
	#endregion Create the New CmdletBinding
	[pscustomobject]@{
		CommandName	     = $CommandName
		CmdletBinding    = $cmdletbinding
		NewCmdletBinding = $NewCmdletBinding
		FileName		 = $File.FullName
	}
}

# Display the output and manually select what to do and what not to do
$Changes = $Results | Out-GridView -PassThru | ForEach-Object -Process {
	$content = [System.IO.File]::ReadAllText($_.filename).Replace($_.CmdletBinding, $_.newcmdletbinding)
	
	$Result = [System.IO.File]::WriteAllText($_.filename, $content)
	
	[pscustomobject]@{
		CommandName = $_.commandName
		CmdletBinding = $_.cmdletbinding
		NewCmdletBinding = $_.newcmdletbinding
	}
}