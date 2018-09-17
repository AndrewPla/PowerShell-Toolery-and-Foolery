# This script installs RSAT on Windows 10 devices and installs Server Admin Tools on Windows Server

# Adapted from : https://blogs.technet.microsoft.com/drew/2016/12/23/installing-remote-server-admin-tools-rsat-via-powershell/
# By Drew Robinson
$web = Invoke-WebRequest 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=45520'
$MachineOS = (Get-CimInstance Win32_OperatingSystem).Name

#Check for Windows Server 2012 R2
if ($MachineOS -like '*Microsoft Windows Server*')
{
	Write-Verbose 'Windows Server detected. '
	Install-WindowsFeature RSAT-AD-PowerShell
}

if ($ENV:PROCESSOR_ARCHITECTURE -eq 'AMD64')
{
	Write-Verbose 'x64 Detected'
	$Link = (($web.AllElements | Where-Object class -eq 'multifile-failover-url').innerhtml[0].split(' ') |
		select-string href).tostring().replace('href=', '').trim('"')
}
else
{
	Write-Verbose 'x86 Detected'
	$Link = (($web.AllElements | Where-Object class -eq 'multifile-failover-url').innerhtml[1].split(' ') |
		select-string href).tostring().replace('href=', '').trim('"')
}

$DLPath = ($ENV:USERPROFILE) + '\Downloads\' + ($link.split('/')[8])

Write-Verbose 'Downloading RSAT MSU file' 
Start-BitsTransfer -Source $Link -Destination $DLPath

$Authenticatefile = Get-AuthenticodeSignature $DLPath

$WusaArguments = $DLPath + ' /quiet'
if ($Authenticatefile.status -ne 'valid') { write-error "Can't confirm download, exiting" }
Write-Verbose 'Installing RSAT for Windows 10 - please wait' 
Start-Process -FilePath 'C:\Windows\System32\wusa.exe' -ArgumentList $WusaArguments -Wait