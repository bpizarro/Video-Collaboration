# Testing the batch file
$read=Read-host "Please type something"
write-host -ForegroundColor $warningColor "You typed " -nonewline 
write-host -ForegroundColor Cyan "$($read)"

#.$execDir\config\ConnectToExO.ps1
#get-Pssession

IwasUsed $MyInvocation.MyCommand.Name
.$execDir\config\PressAnyKey.ps1