# Remove On-Prem search results (by Search Names)
# 2013 compatible: 

.$execDir\config\ConnectToOnPremEx.ps1
Write-Host ""

write-host -ForegroundColor $WarningColor "Enter search names"
$Searches=(Create-Array).Trim()
if($Searches){
    $targetMailbox="SearchNDelete"
    $SearchRequests=@()
    $allSearches=Get-MailboxSearch 
    clear
    foreach($searchName in $Searches ){ 
        $search=$allSearches|where {$_.name -eq $searchName}
        if( $search){
            $SearchRequests += $search
        }else{
            write-host -ForegroundColor red "$($searchName) not found."
        }
    }

    foreach($search in $searchRequests){
        $confirm=$null
        write-host -ForegroundColor $StatusColor "Original search estimate for $($Search.Name):"
        write-host -ForegroundColor Gray "Created By: $($search.CreatedBy)"
        write-host -ForegroundColor Gray "Source Mailboxes: $($search.SourceMailboxes.Name -join ";")"
        write-host -ForegroundColor Gray "Senders: $($search.Senders -join ";")"
        write-host -ForegroundColor Gray "SearchQuery: $($search.SearchQuery)"
        write-host -ForegroundColor Gray "Status: $($search.Status)"
        write-host -ForegroundColor Gray "Result Number Estimate: $($search.ResultNumberEstimate)"
        $BuildQuery=""
        if($Search.ResultNumberEstimate -gt 0){    
            write-host -ForegroundColor $StatusColor "building search..."
            $SourceMailboxes=($Search.SourceMailboxes | Get-Mailbox )
            if($SourceMailboxes){
                if($Search.SearchQuery){ 
                    $SearchQuery=$Search.SearchQuery.replace("`"","")
                    $BuildQuery="'$($SearchQuery)'" } 
                if($search.senders){
                    if($BuildQuery){$BuildQuery+=" AND " }
                    $i=0
                    while($i -lt $search.senders.count){
                        if($i -eq 0){$BuildQuery+="("
                        }else{ $BuildQuery+=" OR "}
                        $BuildQuery+="from:'$($search.senders[$i])'"
                        $i++
                    }
                    $BuildQuery+=")"
                }  
                if($search.StartDate){
                    if($BuildQuery){$BuildQuery+=" AND " }
                    $BuildQuery+="(Received: $(([datetime]$Search.StartDate).ToShortDateString()).."
                    if($search.EndDate){
                        $BuildQuery+="$(([datetime]$Search.EndDate).ToShortDateString()))"
                    }
                }
            }
            Write-Host -ForegroundColor gray "query string: $($buildQuery)"
            $SourceMailboxes | Search-Mailbox -SearchQuery $BuildQuery -TargetMailbox $TargetMailbox -TargetFolder $Search.Name 
            $confirm = Read-Host "Please confirm if the search result looks correct. [Y/N]"
            if($confirm -match "y"){
                $SourceMailboxes | Search-Mailbox -SearchQuery $BuildQuery -TargetMailbox $TargetMailbox -TargetFolder $Search.Name -DeleteContent -Confirm:$false 
            }
            Write-Host ""
        }else{
            Write-Host ""
            write-host -ForegroundColor $WarningColor "No result found for search $($Search.Name)"
            Write-Host ""
        }
    }
}else{
    write-host -ForegroundColor $warningColor "No search name provided."
}
IwasUsed $MyInvocation.MyCommand.Name ($Searches -join ";")
.$execDir\config\PressAnyKey.ps1