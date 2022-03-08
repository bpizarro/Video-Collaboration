#Miro Report
#Using 'E-Mail' from source file as identifying attribute to get users in AD
#Attributes to add: Country, State, CostCenter, and sbuxEmployeeStatus, Manager
#
# v1
# manually select input file and provide recipient address for the report
#

#Adding this function in the script even though it's in SharedFunctions.ps1
#Since we're not sure where we're going to invoke thie script from yet. 

Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
$Date=Get-Date -format yyyy-MM-dd
if(Test-Path "E:\O365\Reports\MiroReport"){ $FileLocation="E:\O365\Reports\MiroReport"}else{ $FileLocation="E:\Reports\MiroReport"}
$Report="$($FileLocation)\MiroReport_$($Date).csv"

$To=Read-Host "Provide recipient email address for the report"
Write-Host "Select Input file..."
$InputText=Get-FileName

if($InputText){
    $Csv=Import-Csv $InputText
    $headers=($csv|Get-Member|where {$_.MemberType -eq "NoteProperty"}).Name
    Import-Module activeDirectory

    $Output=@()
    $i=1
    foreach($user in $csv){    
        write-host "Processing $($i)/$($csv.count) entry..."
        $ADuser=$null
        if($user.'E-mail' -like "*@starbucks.com"){
            $ADUser=Get-ADUser $user.'E-mail'.split("@")[0] -Properties c,State,sbuxCostCenter,sbuxEmployeeStatus,Manager,DisplayName
            clear
            $row=New-Object System.Object
            if($ADUser){
                write-host "AD User: $($ADUser.DisplayName) found."
                $row | Add-Member -Type NoteProperty -Name "DisplayName" -Value $AdUser.DisplayName
                $row | Add-Member -Type NoteProperty -Name "Country" -Value $AdUser.c
                $row | Add-Member -Type NoteProperty -Name "State" -Value $AdUser.State
                $row | Add-Member -Type NoteProperty -Name "CostCenter" -Value $AdUser.sbuxCostCenter
                $row | Add-Member -Type NoteProperty -Name "sbuxEmployeeStatus" -Value $AdUser.sbuxEmployeeStatus
                if($ADUser.Manager){
                    $manager=(Get-ADUser $AdUser.Manager -Properties DisplayName).DisplayName
                    $row | Add-Member -Type NoteProperty -Name "SbuxManager" -Value $Manager
                }else{        
                    $row | Add-Member -Type NoteProperty -Name "SbuxManager" -Value ""
                }
                if($ADUser.Enabled){
                    $row | Add-Member -Type NoteProperty -Name "ADAccount" -Value "Enabled" 
                }else{
                    $row | Add-Member -Type NoteProperty -Name "ADAccount" -Value "Disabled" 
                }
            }else{
                $row | Add-Member -Type NoteProperty -Name "DisplayName" -Value ""
                if($user.'E-mail'.split("@")[0] -like 'ca[0-9][0-9][0-9][0-9][0-9][0-9][0-9]'){
                    $Country="CA"
                }elseif($user.'E-mail'.split("@")[0] -like 'us[0-9][0-9][0-9][0-9][0-9][0-9][0-9]'){
                    $Country="US"
                }else{
                    $Country=""
                }
                $row | Add-Member -Type NoteProperty -Name "Country" -Value $Country
                $row | Add-Member -Type NoteProperty -Name "State" -Value ""
                $row | Add-Member -Type NoteProperty -Name "CostCenter" -Value ""
                $row | Add-Member -Type NoteProperty -Name "sbuxEmployeeStatus" -Value ""
                $row | Add-Member -Type NoteProperty -Name "SbuxManager" -Value ""
                $row | Add-Member -Type NoteProperty -Name "ADAccount" -Value "Does Not Exist in AD"
            }        
            foreach($header in $headers){
                $row | Add-Member -Type NoteProperty -Name $header -Value $user.$header
            }
            $Output+=$row
        }else{
            $row=New-Object System.Object
            $row | Add-Member -Type NoteProperty -Name "DisplayName" -Value ""
            $row | Add-Member -Type NoteProperty -Name "Country" -Value ""
            $row | Add-Member -Type NoteProperty -Name "State" -Value ""
            $row | Add-Member -Type NoteProperty -Name "CostCenter" -Value ""
            $row | Add-Member -Type NoteProperty -Name "sbuxEmployeeStatus" -Value ""
            $row | Add-Member -Type NoteProperty -Name "SbuxManager" -Value ""
            $row | Add-Member -Type NoteProperty -Name "ADAccount" -Value "Not Starbucks"
            foreach($header in $headers){
                $row | Add-Member -Type NoteProperty -Name $header -Value $user.$header
            }
            $Output+=$row
        }
        $i++
    }

    $output|export-csv $Report -NoTypeInformation
    if($to){
        Send-MailMessage -Subject "MiroReport_$($Date)" -Attachments $Report -From "MiroReport@starbucks.com" -To $To -SmtpServer SMTP-Prod.starbucks.net -BodyAsHtml
    }else{
        ii $FileLocation
    }
}else{
    Write-Host -foregroundcolor $WarningColor "No input file selected"
}

IwasUsed $MyInvocation.MyCommand.Name $InputText
.$execDir\config\PressAnyKey.ps1
