function Install-MSIProgram
{
	[CmdletBinding()]
	Param (
		$File
	)
	$File = Get-ChildItem $File
	$DataStamp = get-date -Format yyyyMMddTHHmmss
	$logFile = '{0}-{1}.log' -f $file.fullname, $DataStamp
	$MSIArguments = @(
		"/i"
		('"{0}"' -f $file.fullname)
		"/qn"
		"/norestart"
		"/L*v"
		$logFile
	)
	Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow
}