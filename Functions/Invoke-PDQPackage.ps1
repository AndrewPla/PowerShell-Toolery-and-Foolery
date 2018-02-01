

function Invoke-PDQPackage
{
<#
	.SYNOPSIS
		This is a function that allows you to send PDQ packages from PowerShell.
	
	.DESCRIPTION
		This sends a command to your pdq server that will tell it to deploy software to a computer or computers in your ogranization.
	
	.PARAMETER Package
		This is the name of the package that you want to deploy. 
	
	.PARAMETER NotificationName
		Specify the name of the Notification that you want this deployment to trigger. The default value is Full Details. This will send details of how
		the deployment ended up whatever email is specified for the Notification.
	
	.PARAMETER ComputerName
		This parameter contains the names of the computers that you want to deploy your package to. 
	
	.PARAMETER Credential
		You can use this parameter if you aren't running powershell as a console user for PDQ. 
	
	.PARAMETER PDQServer
		This is the name of your PDQ server. 
	
	.EXAMPLE
		PS C:\> Invoke-PDQPackage -Package '7-Zip' -computername 'Desktop-12345' -credential (get-credential) -pdqserver 'ServerName' -NotificationName 'Full Details'  
	
		This command uses (get-credential) to allow you to get a pop-up to enter your PDQ credentials in. 
		This will deploy the package '7-zip' to the computer 'Desktop-12345' by invoking a command to the 
		PDQ server 'ServerName' When the deployment is complete you will receive the notification named 'Full Details'
	
	
	.NOTES
		I should probably add better help soon.
#>
	
	[CmdletBinding()]
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true)]
		#[ValidateSet('Google Chrome','Mozilla Firefox','CustomPackageName')]   Uncomment if you like and add your own packages.
		[ValidateNotNull()]
		[string]$Package,
		[ValidateNotNull()]
		[string]$NotificationName = "Full Details",
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[string[]]$ComputerName,
		[ValidateNotNull()]
		#We set $credential to a value that will allow the user to run the command as the console user
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
		[string]$PDQServer #= 'ServerNameHere' 
	)
	
	#region 
	if ($Credential.UserName -like "")
	{
		Write-Verbose "No credential specified. Using Username:$Env:USERNAME"
	}
	
	
	if (!(Test-Connection $PDQServer -Count 1))
	{
		Write-Warning "Unable to ping $PDQServer"
	}
	#endregion Check Params	
	
	#region Invoke the Command on PDQ server
	
	Write-Verbose "Invoking command to $PDQServer"
	$ScriptBlock = { pdqdeploy deploy -package $using:package -targets $using:computername -notificationname $using:NotificationName }
	$Params = @{
		Computername    = $PDQServer
		Argumentlist    = $ComputerName, $Package, $NotificationName
		Credential	    = $Credential
		Scriptblock	    = $ScriptBlock
	} #Params
	
	$Result = Invoke-Command @Params
	#endregion
	
	#region Output an object
	
	Write-Verbose "Checking if the command was sent successfully."
	if ($Result -contains 'Deployment Started')
	{
		Write-Verbose "Deployment Started, Outputing successful object."
		
		$ID = $($Result | Select-String -Pattern "ID" | Out-String).split(":")[1]
		$Targets = $($Result | Select-String -Pattern "Targets:" | Out-String).split(":")[1]
		
		[pscustomobject] @{
			'ID'				   = $ID
			'Deployment Started'   = 'True'
			'Package'			   = "$Package"
			'Targets'			   = $targets
		}
	} #if
	else
	{
		Write-Warning "Deployment not started. "
		
		[pscustomobject] @{
			'ID'				   = 0
			'Deployment Started'   = 'False'
			'Package'			   = $Package
			'Targets'			   = $ComputerName.count
		} #pscustomobject
	} #Else	
	
	#endregion Output an object
	
} #function