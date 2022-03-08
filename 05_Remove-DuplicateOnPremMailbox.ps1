# Remove duplicate mailbox on prem
# 2013 compatible: 

.$execDir\config\ConnectToOnPremEx.ps1 -E2013
$inputText=""
$filepath="\\ms58100\E$\MailboxPSTExports"
Do{
    $continue="y"
    $MailboxToRemove=(Read-Host "`nalias of the duplicate mailbox (Press 'q' to quit)").trim()
    $inputText+="$($MailboxToRemove);"
    if($MailboxToRemove -ne "q"){
        $Mailbox=Get-Mailbox $MailboxToRemove -ErrorAction silentlycontinue
        if($Mailbox){
            $Stats=Get-MailboxStatistics $Mailbox.identity -ErrorAction silentlycontinue 
            if($Stats){
                $Stats
                if($stats.ItemCount -gt 20){
                    $continue="n"
                    Write-Host -ForegroundColor $warningColor "`n$($stats.ItemCount) items found. `nExporting mailbox to PST now...`n"
                    New-MailboxExportRequest -Mailbox $Mailbox.alias -Name $Mailbox.alias -BatchName "ExportDuplicate" -FilePath "$($filepath)\$($mailbox.alias).pst"
                    
                    while((Get-MailboxExportRequest -Batchname ExportDuplicate).Status -ne "Completed"){
                        Get-MailboxExportRequest -Batchname ExportDuplicate
                        sleep 60
                    }
                    Get-MailboxExportRequest -Batchname ExportDuplicate
                    ii $filepath
                    Get-MailboxExportRequest -Batchname ExportDuplicate | Remove-MailboxExportRequest -Confirm:$false
                    $continue=Read-host "Are you sure you want to disable the mailbox now? [Y/N]"
                }
            }
            if($continue -match "y"){
                write-host -ForegroundColor $StatusColor "`nDisabling local mailbox for $($MailboxToRemove)"
                Disable-Mailbox $Mailbox.Identity -Confirm:$false
                $RemoteRouting=@(($mailbox.emailaddresses.AddressString |where {$_ -like "*@RETAILSTARBUCKS1COM.mail.onmicrosoft.com"}).replace("smtp:","").replace("SMTP:",""))[0]
                #make it into an array and take the first instance in case there's more than 1 @RETAILSTARBUCKS1COM.mail.onmicrosoft.com addresses
                if($RemoteRouting){
                    write-host -ForegroundColor $StatusColor "`nEnabling remote mailbox for $($MailboxToRemove)"
                    Enable-RemoteMailbox $Mailbox.alias -RemoteRoutingAddress $RemoteRouting                    
                    write-host -ForegroundColor $StatusColor "`nUpdating proxy addresses for $($MailboxToRemove)"
                    Set-RemoteMailbox $Mailbox.alias -EmailAddresses $mailbox.Emailaddresses
                }
            }
        }else{
            Write-Host -ForegroundColor $ErrorColor "$($MailboxToRemove) not found."
        }
    }
}while($MailboxToRemove -ne "q")

IwasUsed $MyInvocation.MyCommand.Name $inputText
.$execDir\config\PressAnyKey.ps1