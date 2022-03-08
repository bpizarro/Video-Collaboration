# Migrate mailbox (On-Prem -> O365)
# 2013 compatible: Y

#Connect to O365, set env variables
$more="yes"
.$execDir\config\ConnectToExO.ps1
$opCred = Read-Credential -target Starbucks\s-exchangemaint -ErrorAction SilentlyContinue 
if(!$OPCred){$opCred =Get-Credential -Message "Enter your On-Prem Exchange admin credentials"}
$BatchName="$($loginUser)_Move"
$InputText=""

DO{ 
    if($InputText){$InputText+=";"}
    Write-Host "alias or UPN of the mailbox to be moved (one per line)"
    $MbxsToMove=(Create-Array).Trim()
    [int]$BadItemLimit=Read-Host "Bad item limit (0-1000)"
    foreach($mbxToMove in $mbxsToMove){
        $BuildCmdStr=""
        $BuildCmdStr="New-MoveRequest `$mbxToMove -remote -RemoteHostName mrs1.starbucks.com -RemoteCredential `$OPCred -BatchName `$BatchName"
        $BuildCmdStr+=" -TargetDeliveryDomain retailstarbucks1com.mail.onmicrosoft.com -BadItemLimit `$BadItemLimit -AcceptLargeDataLoss"
        if($SuspendWhenReadyToComplete){$BuildCmdStr+=" -SuspendWhenReadyToComplete "}
        $BuildCmdStr+=" -Confirm:`$False"
        Invoke-Expression $BuildCmdStr
    }
    write-host ""
    write-host ""
    $InputText+=$MbxsToMove -join ";"
    $more=Read-Host "More mailboxes to move? [Y/N]"
}while($more -match "y")

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1

