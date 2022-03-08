# Remove O365 search results (by Search Names)
# 2013 compatible: Y

#load supporting scripts/functions
#Add-Type -AssemblyName PresentationFramework
#Import-Module PSWorkBench
$cloudCred = Read-Credential -target https://outlook.office365.com/powershell-liveid/
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Cloudcred -Authentication Basic -AllowRedirection 
#Import-PSSession $Session -AllowClobber -DisableNameChecking | Out-Null
Connect-IPPSSession -Credential $cloudCred

function PurgeSearchResult ($ComplianceSearch){
    Do{
        New-ComplianceSearchAction -SearchName $ComplianceSearch.Name -Purge -PurgeType SoftDelete -Confirm:$false
        if($Error[0] -match $ComplianceSearch.Name -and $error[0] -match "aren't up to date. Please update the search results to"){
            Start-ComplianceSearch $ComplianceSearch.Name
        }
        Write-Host -ForegroundColor $StatusColor "Updating search $($ComplianceSearch.Name)..."
        sleep 60

    }while(!(Get-ComplianceSearchAction "$($ComplianceSearch.Name)_Purge" -ErrorAction silentlycontinue))
}

write-host -ForegroundColor yellow "Enter search names"
$Searches=(Create-Array).Trim()

Do{
    $SearchFound=@()
    foreach($searchName in $searches){
        $search=Get-ComplianceSearch $searchName -ErrorAction SilentlyContinue
        if($search){
            $SearchFound+=$search
        }else{
            write-host ""
            write-host -ForegroundColor red "Search $($searchName) not found."
        }
    }

    if($SearchFound){
        foreach($search in $searchfound){
            $search |select Name,SearchType,Items,SuccessResults,ExchangeLocation
            if($search.status -eq "Completed"){
                $Selection=Read-Host "Select the task you want to perform: (1) Purge content (2) Check Purging Status. [1/2]"
                if($Selection -eq 1){
                    PurgeSearchResult $search
                }elseif ($Selection -eq 2){
                    $Result=Get-ComplianceSearchAction "$($search.Name)_Purge"
                    write-host -ForegroundColor $StatusColor "Processing $($Search):"
                    write-host -ForegroundColor Gray "Action: $($Result.Action)"
                    write-host -ForegroundColor Gray "RunBy: $($Result.RunBy)"
                    write-host -ForegroundColor Gray "JobEndTime: $($Result.JobEndTime)"
                    write-host -ForegroundColor Gray "Status: $($Result.Status)"
                }
            }else{
                Start-ComplianceSearch $search.identity -ErrorAction silentlycontinue
                write-host -ForegroundColor $warningColor "$($search.Name) not yet completed. Please try again later."
            }
        
        }

    }else{
        [void][System.Windows.MessageBox]::Show("No existing search found matching the input." , "Warning")
    }
    write-host ""
    write-host -ForegroundColor yellow "Enter search names"
    $Searches=(Create-Array).Trim()
}While($Searches)

IwasUsed $MyInvocation.MyCommand.Name $Searches
.$execDir\config\PressAnyKey.ps1