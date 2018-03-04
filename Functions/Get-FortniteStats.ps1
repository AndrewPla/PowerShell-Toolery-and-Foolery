function Get-FortniteStats
{
<#
	.SYNOPSIS
		Get Fortnite stats for specified user on specified platform. i'
	
	.DESCRIPTION
		Get Fortnite stats for specified user on specified platform. This utilizes the Fortnite
		Tracker API at 'https://fortnitetracker.com/site-api' You will need to sign up and get an API
		key before proceeding.
	
	.PARAMETER EpicNickname
		This is the EpicNickname. This isn't necessarily the same as the name that you will see in-game.
		This is the name of the Epic account.
	
	.PARAMETER Platform
		Platform can either be xbl, psn, or pc.
	
	.PARAMETER APIKey
		Get your API Key from 'https://fortnitetracker.com/site-api'
	
	.EXAMPLE
		PS C:\> Get-FortniteStats -EpicNickname 'Nickname' -Platform 'PC' -apikey $APIKey
	
	
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$EpicNickname,
		[Parameter(Mandatory = $true)]
		[ValidateSet('pc', 'psn', 'xbl')]
		[string]$Platform,
		[Parameter(Mandatory = $true)]
		[String]$APIKey
	)
	
	$Url = "https://api.fortnitetracker.com/v1/profile/$Platform/$EpicNickname"
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	
	$headers.Add("TRN-Api-Key", $ApiKey)
	
	Invoke-RestMethod $url -Headers $headers
}