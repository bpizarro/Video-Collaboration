# New Store Mailbox Creation
# EXOnline only

#set env variables
$Date = "{0:yyyy_MM_dd}" -f (get-date)
$fileLocation="$($ReportDir)\StoreMailboxCreation"
$cloudCred = Read-Credential -target https://outlook.office365.com/powershell-liveid/ -ErrorAction SilentlyContinue
$From="StoreMailboxCreation@starbucks.com"
$to="meseng-starbucks@starbucks.com"
$to=@("ashu@starbucks.com","nehsani@starbucks.com")
$MsgSubject="New Store Mailbox Creation Issue(s):$($Date)"
$MsgBody=""

### MAIN ### 
$NumbersToValidate=Get-StoresFromFile
if(!$NumbersToValidate){
    Write-host -fore white "Enter store number. One per line."
    [array]$NumbersToValidate=@(Create-Array).trim() 
    #[array]$NumbersToValidate=Read-Host "Manually provide store number"
}
    $NewStores=(Get-StoreValidation $NumbersToValidate).NewStores

    #Connect to O365
    if(!$NewStores){
        [System.Windows.Forms.MessageBox]::Show("The number(s) you provided is not valid. Please check." , "Warning")
    }else{
        .$execDir\config\ConnectToExO.ps1
        Connect-MsolService -Credential $cloudCred
        clear
        if($error[0] -match "Authentication Error: Bad username or password."){
            [System.Windows.Forms.MessageBox]::Show("Unable to connect to O365. Check your credentials." , "Warning")
            break
        }
        $MSOLUsers=@(); $storeMailboxExists=@(); $notSynced=@()
        $LoopCount = 1
        foreach($Store in $NewStores){
            $PercentComplete = [Math]::Round(($LoopCount++ / $NewStores.Count * 100),1)
            Write-Progress -Activity ("Checking and assigning license temporarily...") -PercentComplete $PercentComplete `
            -Status "$PercentComplete% Complete" -CurrentOperation "Processing store $($LoopCount-1)/$($NewStores.Count): $($Store)" 
            
            if($user=Get-User $store -ErrorAction SilentlyContinue){
                if(Get-Mailbox -Identity $user.identity -ErrorAction SilentlyContinue){
                    $storeMailboxExists+=$store
                    $MSOLUsers+=Get-MSOLUser -UserPrincipalName $user.UserPrincipalName
                }else{
                    $MSOLUsers+=Set-StoreLicense $user
                }
            }else{
                $notSynced+=$store
            }
        }
        Write-Progress -Activity ("Checking and assigning license temporarily...") -Status "Ready" -Completed
    }
    if($MSOLUsers){
        if($MSOLUsers.count -gt $storeMailboxExists.count){
            write-host "Waiting for mailboxes to be provisioned..."
            Start-Sleep 120
        }
        if($storeMailboxExists){ write-host "Some store mailboxes already exist. Permission and Forwarding will be configured for them as well."}
        $LoopCount = 1
        foreach($MSOLUser in $MSOLUsers){
            $PercentComplete = [Math]::Round(($LoopCount++ / $MSOLUsers.Count * 100),1)
            Write-Progress -Activity ("Configuring store mailboxes") -PercentComplete $PercentComplete `
            -Status "$PercentComplete% Complete" -CurrentOperation "Processing store $($LoopCount-1)/$($MSOLUsers.Count): $($MSOLUser.UserPrincipalName)" 
            
            while (!($Mbx=Get-Mailbox $MSOLUser.UserPrincipalName -ErrorAction SilentlyContinue)){ 
                write-host -ForegroundColor $warningColor "Mailbox $($MSOLUser.UserPrincipalName) not yet provisioned, please wait..." ;
                Start-Sleep 30
            }
            #if($mbx.RecipientTypeDetails -ne "SharedMailbox"){
                Write-Host "updateing store calendar auto processing and converting to shared"
                Set-Mailbox -Identity $Mbx.UserPrincipalName -Type Equipment
                Start-Sleep 30
                Set-CalendarProcessing -Identity $Mbx.UserPrincipalName -AutomateProcessing AutoAccept -AllowConflicts:$true -ConflictPercentageAllowed 100 -MaximumConflictInstances 100 -DeleteNonCalendarItems:$false -DeleteComments:$False -DeleteSubject:$false -DeleteAttachments:$false -AddOrganizerToSubject:$false -MaximumDurationInMinutes 0
                Set-Mailbox -Identity $Mbx.UserPrincipalName -Type Shared;
                Start-Sleep 60
                while ((Get-Mailbox $Mbx.Identity).RecipientTypeDetails -ne "SharedMailbox"){write-host -ForegroundColor $statusColor "waiting for mailbox conversion"; Start-Sleep 60}
            #}
            #Grant MB group permission and forwarding
            if($MBGroup=Get-DistributionGroup "MB-D-$($Mbx.alias)-Managers" -ErrorAction SilentlyContinue){
                write-host -ForegroundColor $statusColor "Configuring MB group permissions and forwarding."
                Set-Mailbox -Identity $Mbx.UserPrincipalName -DeliverToMailboxAndForward $true -ForwardingAddress $MBGroup.Name
                Add-MailboxPermission -Identity $Mbx.UserPrincipalName -User $MBGroup.identity -AccessRights FullAccess
                Add-RecipientPermission -Identity $Mbx.UserPrincipalName -Trustee $MBGroup.identity -AccessRights SendAs -Confirm:$false
            }
            #remove license
            if((Get-MSOLUser -UserPrincipalName $MSOLUser.userprincipalname).Licenses.AccountSkuId -match "RETAILSTARBUCKS1COM:STANDARDPACK"){
                Write-Host "Removing E1 license from $($mbx.alias)"
                Set-MsolUserLicense -UserPrincipalName $MSOLUser.userprincipalname -RemoveLicenses "RETAILSTARBUCKS1COM:STANDARDPACK"        
            }
            Write-Progress -Activity ("Configuring store mailboxes") -Status "Ready" -Completed
        }
    }
    Get-PSSession | Remove-PSSession
    $Report=@()
    if($Validation.invalid){
        [void][System.Windows.MessageBox]::Show("Invalid store number(s) found. Please contact partners." , "Warning")
        $MsgBody+="<b>Invalid store numbers found: </b><br>"
        $MsgBody+=$Validation.invalid -join "<br>"
        $MsgBody+="<br><br>"
        foreach($number in $Validation.invalid){
            $Report+=[pscustomobject]@{StoreNumber=$number;issue="invalid store Number."}
        }
    }
    if($notSynced){
    #store user account not yet sync to the cloud
        $MsgBody+="<b>Store accounts not yet sync to O365: </b><br>"
        $MsgBody+=$notSynced -join "<br>"
        $MsgBody+="<br><br>"
        foreach($number in $notSynced){
            $Report+=[pscustomobject]@{StoreNumber=$number;issue="Not found in O365"}
        }
    }
    if($storeMailboxExists){
        #$MsgBody+="<b>Store mailboxes already exist: </b><br>"
        #$MsgBody+=$storeMailboxExists -join "<br>"
        #$MsgBody+="<br><br>"
        foreach($number in $storeMailboxExists){
            $Report+=[pscustomobject]@{StoreNumber=$number;issue="Store mailbox exists"}
        }
    }
    if($MsgBody){
        Send-MailMessage -Subject $MsgSubject -From $From -To $To -Body $MsgBody -SmtpServer SMTP-Prod.starbucks.net -BodyAsHtml
    }
    if($Report){
        $Report|Export-csv "$($fileLocation)\$($Date)_StoreMailboxCreationIssues.csv" -NoTypeInformation
    }
    
IwasUsed $MyInvocation.MyCommand.Name $NumbersToValidate
.$execDir\config\PressAnyKey.ps1
clear