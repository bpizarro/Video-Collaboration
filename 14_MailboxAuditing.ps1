# Enable/Retrive Mailbox Auditing Logs (On-Prem) 
# 
# not compatible with 2013 right now

import-module psworkbench
.$execDir\config\ConnectToOnPremEx.ps1
clear
write-host -fore cyan "Enable/Retrive Mailbox Auditing Logs (On-Prem)"
write-host ""
$InputText=""
$continue=$true
Do{
    #User Input
    $select=0
    $UserInput=(Read-Host "Mailbox UPN").trim()
    if(!(Get-Mailbox $UserInput -ErrorAction SilentlyContinue)){    
        write-host ""
        Write-Host -fore $WarningColor "$($UserInput) is not a mailbox on Prem."
        write-host ""
    }else{
        write-host ""
        Write-Host -fore cyan "Please select one of the options below:"
        Write-Host "(1) Enable mailbox auditing"
        Write-Host "(2) Disable mailbox auditing"
        Write-Host "(3) Export mailbox auditing log"
        write-host ""
        [int]$select=Read-Host "Please enter your selection [1/2/3]"
        $mailbox=Get-Mailbox $UserInput

        switch ($select){
            #enabling mailbox auditing
            1 {
                #Auditing Attributes
                $auditOwner=@("Create","HardDelete","Move","MoveToDeletedItems","SoftDelete","Update")
                $auditDelegate=@("Create","HardDelete","Move","MoveToDeletedItems","SoftDelete","Update","SendAs","SendOnBehalf")
                $auditAdmin=@("Copy","Create","HardDelete","Move","MoveToDeletedItems","SoftDelete","Update","SendAs","SendOnBehalf")

                Set-Mailbox $mailbox.identity -AuditEnabled:$true -AuditOwner $auditOwner -auditDelegate $auditDelegate    -auditAdmin $auditAdmin
                write-host ""
                write-host "Mailbox auditing enabled for $($UserInput)."
                write-host ""
            }
            #disabling mailbox auditing
            2 {
                Set-Mailbox $mailbox.identity -AuditEnabled:$false
                write-host ""
                write-host "Mailbox auditing disabled for $($UserInput)."
                write-host ""
            }
            #retrieving mailbox audit log
            3 {
                write-host ""
                write-host -fore $WarningColor "Exporting..."
                $Date=((Get-Date).toshortdatestring()).replace("/","-")
                $FileLocation="$($ReportDir)\MailboxAuditingLogs"
                $file="$($fileLocation)\$($date)-$($UserInput).csv"
                $log = Search-MailboxAuditLog -ShowDetails -Identity $mailbox.Identity -LogonTypes owner,delegate,admin -StartDate $Start -EndDate $end
                $log | select LastAccessed,LogonUserDisplayName,LogonType,Operation,OperationResult,FolderPathName,ItemSubject,SourceItemFolderPathNamesList,SourceItemSubjectsList,ClientProcessName |Export-csv $File -NoTypeInformation
                write-host ""
                write-host -fore $WarningColor "Mailbox audit log exported to $($File)"
                ii $FileLocation
                write-host ""
            }
        }
    }
    write-host ""
    if((Read-Host "Do you want to exit this task? [Y/N]") -match "Y"){$continue=$false}
    write-host ""
    $InputText+="$($UserInput): $($select) ;"
}while($continue)

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1