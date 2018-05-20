function Start-PomodoroTimer
{
<#
	.SYNOPSIS
		Creates a Pomodoro Timer that displays a toast notification when complete.
	
	.DESCRIPTION
		Creates a Pomodoro Timer that displays a toast notification when complete. It creates a job
		This function requires the BurntToast module by Josh King @WindosNZ 
	
	.PARAMETER Minutes
		Length of timer
	
	.PARAMETER Sound
		
		Credit to Jeff Wouters for the Imperial March: http://jeffwouters.nl/index.php/2012/03/get-your-geek-on-with-powershell-and-some-music/
		
	
	.EXAMPLE
		PS C:\> Start-PomodoroTimer
	
	.NOTES
		You can download the BurntToast Module by running: Install-Module BurntToast -Scope CurrentUser
		This requires Windows 10 and PowerShell v5
#>
	
	[CmdletBinding()]
	param (
		[int]
		$Minutes = 25,
		
		# There are a lot more different sounds available, but that takes up too much space.
		[ValidateSet('Alarm',
					 'SMS',
					 'Imperial March'
					 )]
		[String]
		$Sound = 'Imperial March'
		
	)
	
	$Messages = @(
		'Go stretch a bit',
		'Call a loved one',
		'Rest your eyes',
		'Go refill your water',
		'Fix your posture',
		'Go for a short walk',
		'Clean up your workspace',
		'Relax, you earned it'
	)
	
	
	
	if ($Sound -match 'Imperial March')
	{
		Start-Job -Name 'Pomodoro Timer' -ArgumentList $Messages, $Minutes -ScriptBlock {
			Start-Sleep -Seconds (60 * $using:Minutes)
			New-BurntToastNotification -Text "Timer complete. Suggestion: $($using:Messages | Get-Random)." -SnoozeAndDismiss
			# Waiting for toast ding to finish, then IMPERIAL MARCH
			Start-Sleep -Seconds 1
			[console]::beep(440, 500)
			[console]::beep(440, 500)
			[console]::beep(440, 500)
			[console]::beep(349, 350)
			[console]::beep(523, 150)
			[console]::beep(440, 500)
			[console]::beep(349, 350)
			[console]::beep(523, 150)
			[console]::beep(440, 1000)
			[console]::beep(659, 500)
			[console]::beep(659, 500)
			[console]::beep(659, 500)
			[console]::beep(698, 350)
			[console]::beep(523, 150)
			[console]::beep(415, 500)
			[console]::beep(349, 350)
			[console]::beep(523, 150)
			[console]::beep(440, 1000)
			

		} 
	}
	else
	{
		Start-Job -Name 'Pomodoro Timer' -ArgumentList $Messages, $Minutes -ScriptBlock {
			Start-Sleep -Seconds (60 * $using:Minutes)
			New-BurntToastNotification -Text "Pomodoro Timer complete. Suggestion: $($Using:Messages | Get-Random)." -SnoozeAndDismiss -Sound $Sound
		} 
	}
}