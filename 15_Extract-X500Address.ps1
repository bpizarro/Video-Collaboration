# Extract X500 address from bounce back message
# 2013 compatible
$retry=$true
DO{
    $fromNDR=$x500=$null
    $title = 'Enter your the address you got from the NDR'
    $msg   = 'It should look something like this: IMCEAEX-_o=Starbucks_ou=Exchange+20Administrative+20Group+20+28FYDIBOHF23SPDLT+29_cn=Recipients_cn=cszasz@namprd05.prod.outlook.com'

    Add-Type -AssemblyName Microsoft.VisualBasic
    $fromNDR=[Microsoft.VisualBasic.Interaction]::InputBox($msg, $title, "$env:fromNDR")

    if($fromNDR -match "="){ 
        $x500=Get-X500 $fromNDR 
        Write-Host -ForegroundColor $StatusColor "Your X500 address is: "
        Write-Host -ForegroundColor Cyan $X500
    }else{
        Write-Host -ForegroundColor $WarningColor "No valid input received." 
        if((Read-Host "quit this task? [Y/N]") -match "y"){$retry=$false}
    }
    if($retry){
        if((Read-Host "quit this task? [Y/N]") -match "y"){$retry=$false}
    }
}
while($retry)

IwasUsed $MyInvocation.MyCommand.Name