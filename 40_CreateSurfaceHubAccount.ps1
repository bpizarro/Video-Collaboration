# Create Surface Hub account
#
# 09.24.2020 added ExO Powershell v2 via ConnectToExO script

$InputText=$DisplayName=$UPN=""
$DoMore=$true
$DisplayName=(Read-Host "Enter Display Name of Surface Hub").trim()
$UPN=(Read-Host "Enter Username of Surface Hub").trim().replace(" ","")
Write-host -fore gray "Please enter the usage location for this device account (where the account is being used)."
$loc = Read-Host "This is a 2-character code that is used to assign licenses (e.g. US/GB)"

$error.clear()
$SurfaceHubPolicy="SurfaceHubMobileDevicePolicy"
Add-Type -AssemblyName System.web
clear 
$cloudCred = Read-Credential -target https://outlook.office365.com/powershell-liveid/
if(!$cloudCred -or (($loginUser -notlike "m-00061-*" -and $loginUser -notlike "*s-exchangemaint") -and ((Read-Host "Is your O365 cred in Credential Manager up to date? [y/n]") -match "n"))){
    $cloudCred=Get-Credential -Message "Enter your O365 admin credentials"
}

while($DoMore){
    $DoMore=$false
    Connect-MicrosoftTeams -Credential $cloudCred
    if($UPN -and $DisplayName -and (!($UPN -match "&"))){        
        .$execDir\config\ConnectToExO.ps1
        #Create Mailbox
        $UPN="$($UPN.split('@')[0])@retail.starbucks.com" 
        $InputText+="$($UPN) ;"
        if(Get-Recipient $UPN -ErrorAction SilentlyContinue){
            Write-Host -fore $ErrorColor "$($UPN) already exist."
            #.$execDir\config\ConnectToCSOnline.ps1
            if(!(Get-CsMeetingRoom $UPN -erroraction silentlycontinue)){
                write-host -fore darkgray "Enabling CS Meeting Room."
                Enable-CsMeetingRoom -Identity $UPN -RegistrarPool "sippoolBLU0B08.infra.lync.com" -SipAddressType EmailAddress
                if(!($Error[0] -match "Management object not found")){
                    Write-Host -fore $statusColor "CS Meeting Room enabled for $UPN"
                }
            }else{
                Write-Host -fore $statusColor "CS Meeting Room already enabled for $UPN"
            }
            Write-Host ""
        }else{
            Write-Host -fore darkgray "Creating Surface Hub mailbox $($UPN)..."
            $pass=[system.web.security.membership]::GeneratePassword(15,4)
            $password=(ConvertTo-SecureString -String $pass -AsPlainText -Force)
            $mailbox=(New-Mailbox -MicrosoftOnlineServicesID $UPN -room -Name $DisplayName -RoomMailboxPassword $Password -EnableRoomMailboxAccount $true)
            sleep 30
            if(!(Get-Mailbox $UPN -ErrorAction SilentlyContinue)){
                Write-Host -Fore $ErrorColor "Surface Hub mailbox not created."            
                break
            }
            # Convert mailbox to user type so we can apply the policy (necessary)
            # while loop as converting mailbox type may take a while
            Write-Host ""
            write-host -fore darkgray "Converting mailbox to UserMailbox to set mobile device policy."
            Set-Mailbox -Identity $UPN -Type Regular
            while((Get-Mailbox $UPN).RecipientTypeDetails -ne "UserMailbox"){
                sleep 15}
            Set-CASMailbox $UPN -ActiveSyncMailboxPolicy $SurfaceHubPolicy 

            # Convert back to room mailbox
            write-host -fore darkgray "Converting mailbox back to room mailbox";
            Set-Mailbox $UPN -Type Room
            while((Get-Mailbox $UPN).ResourceType -ne "Room"){
                sleep 15}
            Set-Mailbox $UPN -RoomMailboxPassword $Password -EnableRoomMailboxAccount $true

            #Set Calendar Processing
            write-host -fore darkgray "Setting calendar processing rules..."
            Set-CalendarProcessing -Identity $UPN -AutomateProcessing AutoAccept
            Set-CalendarProcessing -Identity $UPN -RemovePrivateProperty $false -AddOrganizerToSubject $false -AddAdditionalResponse $true -DeleteSubject $false -DeleteComments $false -AdditionalResponse "This is a Surface Hub room!"

            #Set Password never expire/assign license
            Connect-MsolService -Credential $CloudCred
            if(!(Get-MSOLDomain)){
                Write-Host -Fore $ErrorColor "Unable to Connect to MsolService."
                Write-Host -Fore $warningColor "Please manually set the account to 'Password never expire' and assign E3 license."
                Write-Host ""
            }else{
                if(Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue){  
                    Set-MsolUser -UserPrincipalName $UPN -PasswordNeverExpires $true
                    if(!$loc){$loc="US"}
                    Set-MsolUser -UserPrincipalName $UPN -UsageLocation $loc
                    Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses "RETAILSTARBUCKS1COM:ENTERPRISEPACK"

                    $MSOLUser=Get-MSOLUser -UserPrincipalName $UPN                
                    Write-Host ""
                    write-host -fore darkgray "Location is set to $($MSOLUser.UsageLocation) and $($MSOLUser.Licenses.AccountSkuId) assigned "
                }else{
                    Write-Host -Fore $ErrorColor "Unable to find MSOL User for the Surface mailbox. Failed to set the password to never expire."
                    Write-Host -Fore $warningColor "Please manually set the account to 'Password never expire' and assign E3 license.."
                    Write-Host ""
                }
            } 

            # Setup Skype for Business. 
            sleep 120
            Enable-CsMeetingRoom -Identity $UPN -RegistrarPool "sippoolBLU0B08.infra.lync.com" -SipAddressType EmailAddress
            while(!(Get-CsMeetingRoom $UPN -erroraction silentlycontinue)){
                clear
                write-host -fore $WarningColor "Waiting for object to propagate in AD to try again..."
                sleep 60
                Enable-CsMeetingRoom -Identity $UPN -RegistrarPool "sippoolBLU0B08.infra.lync.com" -SipAddressType EmailAddress
            }
            Write-Host -fore $statusColor "CS Meeting Room enabled."
            Write-Host ""
            Write-Host -Fore cyan "Surface mailbox $($UPN) created."
            Write-Host -Fore cyan "Password is: $($pass)"
            Write-Host ""
        }

    }else{
        write-host -fore red "Missing display name or UPN, or UPN contains invalid characters, such as '&'"
        write-host -fore red "Please try again..."
    }
    Get-PSSession | Remove-PSSession 
    if((Read-Host "Create another surface hub mailbox? Y/N") -match "y"){
        Clear
        $DoMore=$true
        $DisplayName=(Read-Host "Enter Display Name of Surface Hub").trim()
        $UPN=(Read-Host "Enter Username of Surface Hub").trim().replace(" ","")
        Write-host -fore gray "Please enter the usage location for this device account (where the account is being used)."
        $loc = Read-Host "This is a 2-character code that is used to assign licenses (e.g. US/GB)"
    }
}
IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1
