# Check and remove stale Skype user record
# Remove-StaleUserRecord 
#
# Reference: https://accidentalitguy.wordpress.com/2013/07/25/list-all-lync-front-end-servers-that-are-hosting-a-response-group-in-power-shell/
# https://www.jeffbrown.tech/single-post/2014/11/01/Cleaning-Up-Leftover-Lync-Accounts

function Remove-StaleUserRecord {
    $sip=Read-Host "Enter the sip address (without the 'sip:' prefix) of the problematic user"
    if($sip -like "*@starbucks.*"){
        Import-Module Lync
        Import-Module activedirectory

        if(Get-ADUser $sip.split("@")[0] -ErrorAction SilentlyContinue){
            if((Get-Module).Name -match "Lync"){
                $pool="ChdLyncPool01.Starbucks.net"
                $FrontEndServers=(Get-CSPool $pool).Computers

                if(Get-CSUser $sip -ErrorAction SilentlyContinue){
                    Write-Host "Skype service is still enabled for $($sip). Disabling now..."
                    Disable-CSUser $Sip
                    sleep 60
                }

                # Import SQL Module
                if(Get-Module|where {$_.name -eq "sqlserver"}){
                    Import-Module sqlserver 
                }else{
                    Install-Module sqlserver
                }    
                $data=@()
                foreach($Server in $FrontEndServers){
                    $Sql=Invoke-Sqlcmd -Query "SELECT [ResourceId],[UserAtHost] FROM [rtc].[dbo].[Resource] Where UserAtHost ='$sip'" -ServerInstance "$server\rtclocal"
                    $sql|Add-Member -Type NoteProperty -Name "Server" -Value $server
                    $data+=$sql
                }
                foreach($entry in $data){
                    Write-Host "Performing user record removal on $($entry.Server): resrouceID $($entry.ResourceID)"
                    Invoke-Sqlcmd -Query "execute rtc.dbo.RtcDeleteResource '$sip'" -ServerInstance "$($entry.Server)\rtclocal"
                }
                Enable-CsUser $sip -SipAddress "sip:$($sip)" -RegistrarPool $pool
                if($loginUser -match "m-00061-4" -or $loginUser -match "ashu"){
                    $cloudCred = Read-Credential -target https://outlook.office365.com/powershell-liveid/
                    Move-CsUser -Identity $sip -Target sipfed.online.lync.com -Credential $cloudCred -HostedMigrationOverrideUrl https://admin0b.online.lync.com/HostedMigration/hostedmigrationService.svc -confirm:$False
                }
                sleep 20
                Get-CSUser $sip
            }else{        
                write-host -ForegroundColor $WarningColor "Unable to import Lync module."
            }
        }else{
            write-host -ForegroundColor $WarningColor "$($sip) not found."
        }
    }else{
        write-host -ForegroundColor $WarningColor "$($sip) is not a valid sip address."
    }
}
Remove-StaleUserRecord 

IwasUsed $MyInvocation.MyCommand.Name
.$execDir\config\PressAnyKey.ps1