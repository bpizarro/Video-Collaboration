# Export DL members (O365, static, dynamic groups)
# 2013 compatible
#
# O365 groups or AD groups
# dynamic or static (including FIM)
# export member DisplayName, Alias, PrimarySmtpAddress
# 

#$Location="E:\O365\Reports\GroupMembership"
$Location="$($ReportDir)\GroupMembership"

$groupName=(Read-Host "Please enter the name/email address of the group").trim()
write-host ""
$O365Group=Read-Host "Is it an O365 Group? Y/N"
Do{
    [string]$inputText+="$($groupName); "
    $PSSession=Get-PSSession
    if($O365Group -match "y"){
        #check session and reconnect if necessary
        if(!($PSSession.ComputerName -eq "outlook.office365.com" -AND $PSSession.State -eq "Opened")){
            Get-PSSession | Remove-PSSession
            .$execDir\config\ConnectToExO.ps1
        }
        #Retrieve and export Membership
        if($group=Get-UnifiedGroup $groupName){
            $members=Get-UnifiedGroupLinks $groupName -LinkType members -ResultSize unlimited
            $members|Select DisplayName, Alias, PrimarySmtpAddress | Export-csv "$($location)\O365_$($group.DisplayName)_Members.csv" -NoTypeInformation
            write-host -fore yellow "Export completed."
            ii $location
        }else{
            write-host -fore red "$($groupName) not found."
        }
    }else{
        #check session and reconnect if necessary
        if(!($PSSession.ConfigurationName -eq "Microsoft.Exchange" -AND $PSSession.State -eq "Opened")){
            Get-PSSession | Remove-PSSession
            .$execDir\config\ConnectToOnPremEx.ps1
        }

        #static/dynamic
        $group=Get-DistributionGroup $groupName -ErrorAction SilentlyContinue
        $dynamic=$false
        if(!$group){
            $group=Get-DynamicDistributionGroup $groupName -ErrorAction SilentlyContinue
            if($group){$dynamic=$true}
        }
        #Retrieve and export Membership
        if($group){
            if($dynamic){
                $date=(get-date -format yyyy-MM-dd)
                $members=Get-Recipient -RecipientPreviewFilter $group.RecipientFilter -ResultSize unlimited        
                $members|Select DisplayName, Alias, PrimarySmtpAddress | Export-csv "$($location)\Dynamic_$($group.DisplayName)_Members_$($date).csv" -NoTypeInformation
            }else{
                $data=Get-NestedMembership $group.alias -confirmed
                if($data.groups){
                    $data.recipients | Select DisplayName, Alias, PrimarySmtpAddress | Export-csv "$($location)\Nested_$($group.DisplayName)_Members.csv" -NoTypeInformation
                    $data.groups | Select DisplayName, Alias, PrimarySmtpAddress | Export-csv "$($location)\NestedGroups_of_$($group.DisplayName).csv" -NoTypeInformation
                }else{
                    $data.recipients | Select DisplayName, Alias, PrimarySmtpAddress | Export-csv "$($location)\$($group.DisplayName)_Members.csv" -NoTypeInformation
                }
            }    
            write-host -fore yellow "Export completed."
            ii $location
        }else{
            write-host -fore red "$($groupName) not found."
        }
    }
    
    write-host ""
    $groupName=(Read-Host "Please enter the name/email address of the group. (Type 'quit' to exit)").trim()
    if($groupName -ne "quit"){
        write-host ""
        $O365Group=Read-Host "Is it an O365 Group? Y/N"
    }
}while($groupName -ne "quit")

IwasUsed $MyInvocation.MyCommand.Name $inputText
.$execDir\config\PressAnyKey.ps1