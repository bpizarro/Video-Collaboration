# Create One-Time Use DL
#
# https://docs.starbucks.net/display/~ashu/Request+a+one-time+use+static+DL+for+communication
# 
# Purpose of the communication
# Preferred name for the DL (try to be not too generic)
# Name of the mailbox from which the communication will be sent
# attachment: email addresses of partners to be added in the requested DL in a xlsx or .csv file
# 

#input param
$DLName=$InputText=(Read-Host "Name/alias of the one-time use DL").Trim()

Write-Host "Email Address of the approved sender mailbox/group"
$ApprovedSenders=(Create-Array).Trim()
Write-Host -fore white "Select the CSV file that contains members for this DL"
Write-Host -fore cyan "The CSV file should only contain email addresses with 'Member' as header."
$Confirm=Read-Host "Are you ready to import the file? [Y/N]"
if($confirm -like "y*"){
    $Members=(Import-Csv (Get-FileName))|select -Unique Member
}

if(($DLName -like "DL-*") -and $Members){
    #establish connection
    if(!($PSSession.ConfigurationName -eq "Microsoft.Exchange" -AND $PSSession.State -eq "Opened")){
        Get-PSSession | Remove-PSSession
        .$execDir\config\ConnectToOnPremEx.ps1
    }
    #Create DL
    $OU="OU=Corp Distribution Lists,OU=Corp Groups,OU=Corp,DC=starbucks,DC=net"
    $alias=$DLName.replace(' ','')
    New-DistributionGroup -Name $DLName -DisplayName $DLName -Alias $alias -PrimarySMTPAddress "$($alias)@starbucks.com" -OrganizationalUnit $OU
    $DL=Get-DistributionGroup $alias
    sleep 300
    write-host -foregroundcolor yellow "Waiting for the DL to be provisioned..."
    #Set Allowed Sender
    if($ApprovedSenders){    
        write-host -foregroundcolor white "Adding allowed sender..."
        foreach($ApprovedSender in $ApprovedSenders){
            $Sender=Get-Recipient $ApprovedSender
            if($Sender.recipienttype -match "group"){
                Set-DistributionGroup -Identity $DL.Identity -AcceptMessagesOnlyFromDLMembers @{Add=$Sender.DistinguishedName}
            }else{
                Set-DistributionGroup -Identity $DL.Identity -AcceptMessagesOnlyFrom @{Add=$Sender.DistinguishedName}
            }
        }
    }

    #Add Members
    $notready=$true
    while($notready){
        $error.clear()
        Add-DistributionGroupMember -Identity $DL.Identity -Member $members[0].member.trim()
        if(!$error){
            $notready=$false
        }elseif($error[0] -match "The operation couldn't be performed because object"){
            sleep 300
        }else{
            $notready=$false
        }
    }
    $i=1
    foreach($member in $members){
        write-host "Adding member $($i)\$($members.count)"
        Add-DistributionGroupMember -Identity $DL.Identity -Member $member.member.trim()
        $i++
    }

    write-host -fore cyan "$($DLName) created."
}

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1
