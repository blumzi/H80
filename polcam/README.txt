To register the PowerShell module, run the following command in an elevated PowerShell session:

1. open an elevated 32bit PowerShell prompt (Run as Administrator).
	& "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
2. run .\register.bat
3. check by creating a new-oject of type H80

	$p = New-Object -ComObject H80
	$p | Get-Member
