#Add Members to an Azure Group

$cloudcred=Read-Credential -Target "https://outlook.office365.com/powershell-liveid/"
Connect-AzureAD -Credential $cloudcred

$GroupName=(Read-Host "Enter the name of the Azure Group").trim()
Do{
    $AzADGroup=(Get-AzureADGroup -All:$true)|where {$_.Displayname -eq $GroupName}
    if($AzADGroup){
        Write-Host "Enter UPN of new members"
        $UPNs=(Create-Array).Trim()
        foreach($UPN in $UPNs){
            $member=Get-AzureADUser -ObjectId $UPN 
            if($member){
                Try{
                    Add-AzureADGroupMember -ObjectId $AzADGroup.ObjectId -RefObjectId $member.ObjectId
                    write-host "$($UPN) added to $($groupName)"
                }Catch{
                    Write-host -fore cyan $upn
                    write-host -fore red $_
                }   
            }
        }
    }else{
        write-host -fore Red "$($GroupName) not found."
    }
    $GroupName=(Read-Host "Enter the name of the Azure Group. Press 'Enter' to quit.").trim()

}while ($groupName)