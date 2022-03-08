#Deprovision Teams Phone Number
# 
#Need to remote number(LineURI) first and then remove user from TeamsPhoneSystem Azure group to remove license
#
$InputText=$UserPrincipalName=Read-Host "Please provide the UPN of the user to be de-provisioned"

if($UserPrincipalName){
# Start the session.
    $cloudcred = Read-Credential -target https://outlook.office365.com/powershell-liveid/
    $session = New-CsOnlineSession -Credential $cloudcred -OverrideAdminDomain $OverrideAdminDomain
    # Just import the commands we want as it saves several seconds from a full import.
    Import-PSSession -Session $session -AllowClobber -DisableNameChecking `
        -CommandName @(
            'Grant-CsOnlineVoiceRoutingPolicy',
            'Set-CsUserPstnSettings',
            'Set-CsUser',
            'Get-CsUser',
            'Get-CSOnlineUser',
            'Get-CsOnlineVoiceRoutingPolicy'
        ) `
    | Out-Null
    Set-CsUser -Identity $UserPrincipalName -OnPremLineURI $null -EnterpriseVoiceEnabled $false
    Connect-AzureAD -Credential $cloudcred
    
    $groupId="17370e68-220c-4671-8558-8569caab96c3"
    $userId=(Get-AzureADUser -SearchString $UPN).ObjectId 
    Remove-AzureADGroupMember -ObjectId $groupId -MemberId $userId

    Get-PSSession | Remove-PSSession
}


IwasUsed $MyInvocation.MyCommand.Name $NumbersToValidate
.$execDir\config\PressAnyKey.ps1
clear