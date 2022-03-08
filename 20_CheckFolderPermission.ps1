# Check folder permission in a mailbox (O365)
# on-perm not tested. currently disabled.
function Get-MailboxToCheck($mbxInput){
    Write-Host -ForegroundColor $StatusColor "Verifying mailbox input..."
    $ExMailbox=Get-Mailbox "$($mbxInput)*" -ErrorAction SilentlyContinue
    if(!$ExMailbox){
        Write-Host -ForegroundColor $ErrorColor "$($mbxInput) not found."
    }else{    
        $MailboxToCheck=$ExMailbox
        if($ExMailbox.count -gt 0 -and $ExMailbox.count -le 10){
            Write-Host -ForegroundColor $warningColor "More than one mailboxes found."
            Write-Host ""
            $i=1
            write-host -ForegroundColor $SeparatorColor "#: DisplayName`tUPN`tPrimarySMTPAddress"
            foreach($m in $ExMailbox){
                write-host -ForegroundColor $statusColor "$($i): " -NoNewline
                write-host -ForegroundColor cyan "$($m.DisplayName) " -NoNewline
                write-host -ForegroundColor Gray "`t$($m.UserPrincipalName) " -NoNewline
                write-host -ForegroundColor Gray "`t$($m.PrimarySMTPAddress) " 
                $i++
            }
            $pick=Read-Host "Please select the one you want to check [1 - $($ExMailbox.count)]"
            $MailboxToCheck=$ExMailbox[$pick-1]
            clear
        }elseif($ExMailbox.count -gt 10){
            Write-Host -ForegroundColor $WarningColor "$($ExMailbox.count) mailboxes found, please try using email address."
            break
        }
        Return $MailboxToCheck.UserPrincipalName
    }
}

function Get-AllFolderPermission{
    Param (
    [parameter(Mandatory = $True)][String]$mbxToProcess,
    [parameter(Mandatory = $false)][String]$userInput
    )
    if($userInput){
        $user=Get-User "$($userInput)" -ErrorAction SilentlyContinue
        if(!$user){
            Write-Host -ForegroundColor $WarningColor "$($userInput) not found. Will return permission for all users on this mailbox."
        }
    }
    Write-Host -ForegroundColor $statusColor "Retrieving folders for $($mbxToProcess)"
    $allFolders=Get-MailboxFolderStatistics $mbxToProcess | where {$_.FolderPath -notlike "/Sync Issues*" -and $_.FolderPath -notcontains "Top of Information Store" -and $_.FolderPath -notlike "/Archive*" -and $_.ContainerClass -like "IPF.*"}
    $FolderPerms=@()
    $LoopCount=1
    foreach($folder in $allFolders){
        $PercentComplete = [Math]::Round(($LoopCount++ / $allFolders.Count * 100),1)
        Write-Progress -Activity ("Checking folder permission") -PercentComplete $PercentComplete `
        -Status "$PercentComplete% Complete" -CurrentOperation "Processing folder $($LoopCount-1)/$($allFolders.Count): $($folder.folderpath)" 
        
        $folderPath="$($mbxToProcess):$($folder.folderpath.replace("/","\"))"
        $CommandString="Get-MailboxFolderPermission '$($FolderPath)' "
        if($user){$CommandString+=" -User '$($user)' -ErrorAction SilentlyContinue"}
        $perms=Invoke-Expression $CommandString | where {$_.user.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous"}
        if($perms){
            foreach($perm in $perms){
                $FolderPerms+=[pscustomobject]@{FolderPath=$folderPath.replace("$($mbxToProcess):","");User=$perm.User;AccessRights=($perm.AccessRights -join ";");SharingPermissionFlags=$perm.SharingPermissionFlags}
            }
        }
    }
    Write-Progress -Activity ("Checking folder permission") -Status "Ready" -Completed
    Return $FolderPerms
}

$Date = get-date -Format yyyy-MM-dd-HH-mm
$fileLocation="$($ReportDir)\FolderPermission"

Write-Host -ForegroundColor $WarningColor "Please note, folders with special characters cannot be processed.`n"
$mbxInput=(Read-Host "Enter the mailbox (display name or email address)").Trim()
#if((Read-Host "Is this a on-prem mailbox? [Y/N]") -match "y"){$OnPrem=$true}
$userInput=Read-Host "Enter the user (press 'eneter' to skip)"

if($OnPrem){
    #. 'E:\Program Files\Microsoft\Exchange Server\bin\RemoteExchange.ps1'; 
    #Connect-ExchangeServer -auto -AllowClobber
    .$execDir\config\ConnectToOnPremEx.ps1
}else{
    .$execDir\config\ConnectToExO.ps1
}

if($MailboxToCheck=Get-MailboxToCheck $mbxInput){
    $CommandStr="Get-AllFolderPermission -mbxToProcess '$($MailboxToCheck)'"
    if($mbxInput){$CommandStr += " -userInput '$($userInput)'"}
    $allFolderPermissions=Invoke-Expression $CommandStr
}
if($allFolderPermissions){
    $allFolderPermissions | Export-CSV "$($fileLocation)\$($date)_$($MailboxToCheck)_FolderPerms.csv" -NoTypeInformation
    ii $fileLocation
}else{
    Write-Host -ForegroundColor $warningColor "No result found."
}

IwasUsed $MyInvocation.MyCommand.Name
.$execDir\config\PressAnyKey.ps1
