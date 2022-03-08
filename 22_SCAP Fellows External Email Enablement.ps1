#SCAP Fellows External Email Enablement
.$execDir\config\ConnectToOnPremEx.ps1
Import-Module ActiveDirectory 

write-host "please provide Global User Number (including country code) of SCAP fellows: (e.g. US2896508)"
$PartnerNumbers=create-Array 

$Workitems=@()
$PartnerNumbers|%{$Workitems+=Get-ADUser $_ -Server ms20075.ext.starbucks.com -Properties sbuxLocalMarket,sbuxJobNumber,EmployeeNumber,employeeType,mail,sn,GivenName,DisplayName,Title,mail }

$contactOU="OU=Corp SCAP Fellows,OU=Corp Contacts,OU=Corp,DC=starbucks,DC=net"
$alreadyExist=@()
foreach($workitem in $workitems){
    if($workitem.mail){
        Write-host -fore cyan "Mail contact $($workitem.mail) for $($workitem.DisplayName) ($($workitem.EmployeeNumber)): $($workitem.Title) already exist."
        $alreadyExist+=$workitem
    }else{    
        $contact=Get-Contact $workitem.EmployeeNumber -ErrorAction silentlycontinue
        if(!$contact){ 
            New-ADObject -Type Contact -Path $contactOU -Name $workitem.EmployeeNumber `
            -OtherAttributes @{'EmployeeNumber'=$workitem.EmployeeNumber;'employeeType'=$workitem.employeeType; `
            'sn'=$workitem.sn;'GivenName'=$workitem.GivenName;'DisplayName'=$workitem.DisplayName}        
            $contact=Get-Contact $workitem.EmployeeNumber -ErrorAction silentlycontinue
             while(!$contact){
                write-host -fore yellow "$($workitem.EmployeeNumber) waiting for contact creation..."
                sleep 120
                $contact=Get-Contact $workitem.EmployeeNumber -ErrorAction silentlycontinue
            }
        }
    }
    $mailcontact=Get-MailContact $workitem.EmployeeNumber -ErrorAction silentlycontinue
    if(!$mailcontact){
        write-host "mail enabling $($workitem.EmployeeNumber)"
        Enable-MailContact -Identity $Contact.DistinguishedName -Alias $workitem.EmployeeNumber `
                        -ExternalEmailAddress "$($workitem.EmployeeNumber)@retailstarbucks1com.onmicrosoft.com" `
                        -PrimarySmtpAddress "$($workitem.EmployeeNumber)@starbucks.com" `
                        -ErrorAction Stop
    }else{
        $alreadyexist+=$workitem
    }
}
Get-PSSession | Remove-PSSession 

IwasUsed $MyInvocation.MyCommand.Name "$($PartnerNumbers -join ';')"
.$execDir\config\PressAnyKey.ps1