# O365 Email Forwarding
# EXOnline only
function validateFowardAddress($FwdAddress){
    $consumerDomains=@("@gmail.com","@hotmail.com","@yahoo.com","@live.com","@icloud.com","@me.com","@mac.com","@outlook.com","@msn.com","@aol.com","@comcast.net")
    $Return=$null
    if($FwdAddress -match "@"){
        $FwdDomain="@$($FwdAddress.split('@')[1])"
        if($consumerDomains -match $fwdDomain){
            write-host -ForegroundColor $WarningColor "$($fwdDomain) is not a valid vendor domain."
            $return=$false
        }else{
            $return=$true
        }
    }else{
        write-host -ForegroundColor $WarningColor "$($FwdAddress) is not a valid email address."
        $return=$false
    }
    Return $Return
}

$InputText=$emailAddress=""
$emailAddress=(Read-Host "Enter email address of mailbox").trim()
if($emailAddress){
    .$execDir\config\ConnectToExO.ps1
    Do{
        $InputText+="$($emailAddress) ;"
        $mailbox=Get-mailbox $emailAddress -ErrorAction SilentlyContinue
        $confirm=$true
        if($mailbox){
            Write-Host ""
            if($mailbox.ForwardingAddress){
                $Confirm=$false
                write-host -ForegroundColor $WarningColor "$($mailbox.Name):$($mailbox.PrimarySMTPAddress) is currently forwarding to $($mailbox.ForwardingAddress)"
                Write-Host ""
                if((Read-Host "Please confirm if you would like to overwrite the forwarding address. [Y/N]") -match "y"){
                    $confirm=$true
                }
            }
            if($confirm){
                $ForwardingAddress=(Read-Host "Forwarding Email Address").trim()
                if(validateFowardAddress $ForwardingAddress){
                    $Contact=Get-MailContact $ForwardingAddress -ErrorAction SilentlyContinue
                    if($Contact){
                        Set-Mailbox $mailbox.identity -ForwardingAddress $Contact.identity -DeliverToMailboxAndForward:$true -ForwardingSmtpAddress $null
                        Write-Host ""
                        write-host -ForegroundColor Cyan "Setting forwarding for $($mailbox.Name):$($mailbox.PrimarySMTPAddress) to $($contact.name):$(($contact.ExternalEmailAddress).replace('SMTP:',''))"
                    }else{
                        write-host -ForegroundColor $WarningColor "Mail contact not yet created for $($ForwardingAddress)."
                    }
                }
            }
        }else{
            Write-Host -ForegroundColor $ErrorColor "$($emailAddress) not found."
        }
        Write-Host ""
        $emailAddress=(Read-Host "Enter email address of mailbox").trim()
    }while($emailAddress)
}

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1
