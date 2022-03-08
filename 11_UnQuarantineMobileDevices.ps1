# Manually allow mobile devices
# EXOnline only
Write-Host -ForegroundColor $warningColor "Enter the mailboxes you'd like to check"

.$execDir\config\ConnectToExO.ps1

$InputText=$mailbox=""
$mailbox=(Read-Host "Enter mailbox to check, type 'quit' to quit").trim()
Do{
    $InputText+="$($mailbox) "
    if(Get-mailbox $mailbox -ErrorAction SilentlyContinue){
        Write-Host -ForegroundColor $warningColor "Checking $($mailbox)"
        $MobileDevices=@(Get-MobileDevice -Mailbox $mailbox |sort DeviceAccessState)
        $Allowed=@($MobileDevices | where {$_.DeviceAccessState -eq "Allowed" })
        $Blocked=@($MobileDevices | where {$_.DeviceAccessState -eq "Blocked" })
        $Quarantined=@($MobileDevices | where {$_.DeviceAccessState -eq "Quarantined" })
        Write-Host -ForegroundColor $statusColor "$($MobileDevices.count) devices found. $($Allowed.count) allowed. $($Quarantined.count) Quarantined. $($Blocked.count) Denied."
        $i=0
        foreach($device in $mobileDevices){
            write-host -ForegroundColor Cyan "$($i)" 
            write-host -ForegroundColor Gray "FriendlyName: $($device.FriendlyName)"
            write-host -ForegroundColor Gray "DeviceUserAgent: $($device.DeviceUserAgent)"
            write-host -ForegroundColor Gray "DeviceId: $($device.DeviceId)"
            write-host -ForegroundColor Gray "DeviceAccessState: $($device.DeviceAccessState)"
            write-host -ForegroundColor Gray "DeviceAccessStateReason: $($device.DeviceAccessStateReason)"
            $i++
        }
        if($Quarantined -or $Blocked){
            Do{
                $pick=Read-Host "Which device do you want to unquarantine? [0-$($i-1)] (Q to quit)"
                if($mobileDevices[$pick]){
                    Set-CASMailbox $mailbox -ActiveSyncAllowedDeviceIDs @{add=$mobileDevices[$pick].deviceid}
                }
            }
            while($pick -ne "q")
        }
    }else{
        Write-Host -ForegroundColor $ErrorColor "$($mailbox) not found."
    }
    Write-Host ""
    $mailbox=(Read-Host "Enter mailbox to check, type 'quit' to quit").trim()
}while($mailbox -ne "quit")

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1