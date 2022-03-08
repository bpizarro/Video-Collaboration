# Mailbox Folder Statistics (O365)
#
.$execDir\config\ConnectToExO.ps1
$date=get-date -format yyyy-MM-dd
$reportLocation="$($ReportDir)\MailboxFolderStatistics"

$inputText=$mailbox=$reportGenerated=""
$mailbox=(Read-Host "Enter mailbox full email address, type 'quit' to quit").Trim()
Do{
    $InputText+="$($mailbox) ;"
    if(Get-mailbox $mailbox -ErrorAction SilentlyContinue){
        $reportGenerated+="1"
        Write-Host -ForegroundColor $warningColor "retrieving statistics for $($mailbox)..."
        $FolderStats=Get-MailboxFolderStatistics $mailbox | where { $_.FolderPath -ne "/Deletions" } 
        Write-Host -ForegroundColor $statusColor "Exporting statistics..."
        $FolderStats|select FolderPath, ItemsInFolder, @{N="FolderSize(MB)";E={[math]::round( ([decimal](($_.FolderSize -replace "[0-9\.]+ [A-Z]* \(([0-9,]+) bytes\)","`$1") -replace ",","") / 1MB),2)}}|Export-csv "$($reportLocation)\$($date)_$($mailbox)_folderStats.csv" -NotypeInformation 
    }else{
        Write-Host -ForegroundColor $ErrorColor "$($mailbox) not found."
    }
    Write-Host ""
    $mailbox=(Read-Host "Enter mailbox to check, type 'quit' to quit").Trim()
}while($mailbox -ne "quit")

if($reportGenerated){
    ii $reportLocation
}

IwasUsed $MyInvocation.MyCommand.Name $inputText
.$execDir\config\PressAnyKey.ps1