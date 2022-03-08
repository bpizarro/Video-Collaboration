# On-Prem Message Track
#
# Connect to On-Prem Exchange
# Remove EMEA servers 4/20/2020
.$execDir\config\ConnectToOnPremEx.ps1
clear
try{Add-Type -Path "$execDir\config\MessageTracking.cs" -ReferencedAssemblies ("System.Windows.Forms","System.Drawing","System.Configuration.Install") -WarningAction SilentlyContinue -ErrorAction SilentlyContinue}catch{}
[System.Windows.Forms.Application]::EnableVisualStyles();
#create new window and show it
$window = New-Object MessageTracking.FormWindow

$debug = $false
$debugcolor = "Magenta"
$again = "`$true"
do{
	$result = $window.ShowDialog()
	if($result -eq "Yes"){
		#get data after window closed
		$textboxes = "sender","recipient","messageid","messagesubject","server","source"
		foreach($textbox in $textboxes){ Set-Variable -Name $textbox -Value $window.Controls.Find($textbox,$true)[0].Text.Trim().Replace("`$","```$").Replace("`"","```"") }

		$dropdowns = "eventid","Coast"
		foreach($dropdown in $dropdowns){ Set-Variable -Name $dropdown -Value $window.Controls.Find($dropdown,$true)[0].SelectedItem }

		if($window.Controls.Find("startdate",$true)[0].Checked -eq "True"){$startdate = $window.Controls.Find("startdate",$true)[0].Value.tostring().split(" ",2)[0]}
		else{$startdate = ""}

		if($window.Controls.Find("enddate",$true)[0].Checked -eq "True"){$enddate = $window.Controls.Find("enddate",$true)[0].Value.tostring().split(" ",2)[0]}
		else{$enddate = ""}

		if($window.Controls.Find("starttime",$true)[0].Checked -eq "True"){$starttime = $window.Controls.Find("starttime",$true)[0].Value.tostring().split(" ",2)[1]}
		else{$starttime = ""}

		if($window.Controls.Find("endtime",$true)[0].Checked -eq "True"){$endtime = $window.Controls.Find("endtime",$true)[0].Value.tostring().split(" ",2)[1]}
		else{$endtime = ""}
		
		$checkboxes = "ClientHostname","ClientIP","ConnectorID","EventData","EventID","InternalMessageID","MessageID","MessageInfo","MessageLatency",`
					  "MessageLatencyType","MessageSubject","RecipientCount","Recipients","RecipientStatus","Reference","RelatedRecipientAddress",`
					  "ReturnPath","Sender","ServerHostname","ServerIP","Source","SourceContext","Timestamp","TotalBytes"
		foreach($checkbox in $checkboxes){ Set-Variable -Name "chk$checkbox" -Value $window.Controls.Find("chk$checkbox",$true)[0].CheckState }

		if($window.Controls.Find("ExportToCSV",$true)[0].checked -eq "True"){
			$exportToCSV = $true
			$savefile = $window.Controls.Find("savefile",$true)[0].Text.Trim()
			$path=($savefile.replace($($savefile.split("\")[-1]),""))
			if($debug){ Write-Host -ForegroundColor $debugcolor "CSV export: $savefile" }
		}else{
			$exportToCSV = $false
			if($debug){ Write-Host -ForegroundColor $debugcolor "--no export to CSV" }
		}

		if($window.Controls.Find("FormatTable",$true)[0].checked -eq "True"){
			$formatting = "ft -a"
			if($debug){ Write-Host -ForegroundColor $debugcolor "--using format-table" }
		}elseif($window.Controls.Find("FormatList",$true)[0].checked -eq "True"){
			$formatting = "fl"
			if($debug){ Write-Host -ForegroundColor $debugcolor "--using format-list" }
		}
		
		if($window.Controls.Find("chkExtraCommands",$true)[0].checked -eq "True"){
			$chkExtraCommands = $true
			$ExtraCommands = $window.Controls.Find("ExtraCommands",$true)[0].Text.Trim()
			if($debug){ Write-Host -ForegroundColor $debugcolor "Extra commands: $ExtraCommands" }
		}else{
			$chkExtraCommands = $false
			if($debug){ Write-Host -ForegroundColor $debugcolor "--no extra commands found" }
		}

		$ExecStr = ""
		$MTResults=$TransportServers=@()

		if($server){
			$ExecStr = "`$server | %{ `$MTResults += Get-MessageTrackingLog -Server `$_ -resultsize unlimited "
			if($debug){ Write-Host -ForegroundColor $debugcolor "Server: $server" }
		}else{
			$CHD=@('chdms071','ms01176','ms01177','ms01178','ms05112','ms05113','ms13752','ms13753','ms13863','ms13864','ms13865')
			$IAD=@('ms02948','ms02949','ms02950','ms02951','MS05097','ms05403','ms05404','MS05096','ms13876','ms13877','ms13878','ms13879')
			#$EMEA=@('ms02604','ms02605','ms02606')
			$China=@('shams460','shams461','shams462','SHAMS360','SHAMS361','SHAMS362','SHAMS363','SHAMS364','SHAMS365')
			Write-Host "Coast = $($Coast)"
			switch($coast){
				"All"{
					#Write-Host -ForegroundColor $statusColor "All"
					$TransportServers=$CHD+$IAD+$EMEA+$China
				}
				"North America"{
					#Write-Host -ForegroundColor $statusColor "NA"
					$TransportServers=$CHD+$IAD
				}
				#"EMEA"{
				#	#Write-Host -ForegroundColor $statusColor "EMEA"
				#	$TransportServers=$EMEA
				#}
				"China"{
					#Write-Host -ForegroundColor $statusColor "China"
					$TransportServers=$China
				}
				default{
					#Write-Host -ForegroundColor $statusColor "no input"
					$TransportServers=$CHD+$IAD+$EMEA+$China
				}
			}
			Write-Host -ForegroundColor $debugcolor "Server: $TransportServers" 
			$ExecStr  = "`$TransportServers | %{ `$MTResults += Get-MessageTrackingLog -Server `$_ -resultsize unlimited "
		}
		$variables = "sender","recipient","messageid","messagesubject","source","eventid"
		$wildcards = $false
		foreach($variable in $variables){
			$var = $variable
			$val = (Get-Variable -Name $variable).Value
			if(($var -eq "sender" -or $var -eq "recipient") -and ($val -match "\*")){
				$wildcards = $true
				if($debug){ Write-Host -ForegroundColor $debugcolor "--found wildcard" }
			}else{
				if($val){ $ExecStr += " -$var `"$val`"" }
			}
		}

		if($startdate -or $starttime -or $enddate -or $endtime){
			if($startdate -and $starttime){ $ExecStr += " -start `"$startdate $starttime`""}
			elseif($startdate -and !$starttime){ $ExecStr += " -start `"$startdate`""}
			elseif(!$starttime -and $starttime){ $ExecStr += " -start `"$starttime`""}
			if($enddate -and $endtime){ $ExecStr += " -end `"$enddate $endtime`""}
			elseif($enddate -and !$endtime){ $ExecStr += " -end `"$enddate`""}
			elseif(!$endtime -and $endtime){$ExecStr += " -end `"$endtime`""}
		}

		if($wildcards){
			if(($sender -match "\*") -and ($recipient -match "\*")){
				$ExecStr += " | where {`$_.sender -like `"$sender`" -and `$_.recipients -like `"$recipient`"}"
				if($debug){ Write-Host -ForegroundColor $debugcolor "--determined recipient and sender wildcards" }
			}elseif(($sender -match "\*") -and !($recipient -match "\*")){
				$ExecStr += " | where {`$_.sender -like `"$sender`"}"
				if($debug){ Write-Host -ForegroundColor $debugcolor "--determined sender wildcard" }
			}elseif(!($sender -match "\*") -and ($recipient -match "\*")){
				$ExecStr += " | where {`$_.recipients -like `"$recipient`"}"
				if($debug){ Write-Host -ForegroundColor $debugcolor "--determined recipient wildcard" }
			}
		}
		$ExecStr += " } | sort-object timestamp"

		Write-Host -ForegroundColor DarkGray "Running message track...`n"
		if($debug){ Write-Host -ForegroundColor $debugcolor "Invoked string: $ExecStr" }else{ Write-Host -nonewline "Command Used: "; write-host -foregroundcolor green "$ExecStr" }
		Invoke-Expression -Command "$ExecStr"
		if($MTResults){
			if($chkExtraCommands){ #screen output
				$checkboxlist = ""
				foreach($checkbox in $checkboxes){
					$var = $checkbox
					$val = (Get-Variable -Name "chk$checkbox").Value
					if($val -eq "Checked"){ $checkboxlist += "$var," }
				}
				$checkboxlist = $checkboxlist.TrimEnd(",")
				if($debug){ Write-Host -ForegroundColor $debugcolor "Checkboxes: $checkboxlist" }
				Invoke-Expression -Command "`$MTResults | $ExtraCommands | $formatting $checkboxlist"
			}else{
				$checkboxlist = ""
				foreach($checkbox in $checkboxes){
					$var = $checkbox
					$val = (Get-Variable -Name "chk$checkbox").Value
					if($val -eq "Checked"){ $checkboxlist += "$var," }
				}
				$checkboxlist = $checkboxlist.TrimEnd(",")
				if($debug){ Write-Host -ForegroundColor $debugcolor "Checkboxes: $checkboxlist" }
				Invoke-Expression -Command "`$MTResults | $formatting $checkboxlist"
			}
			if($exportToCSV){ #csv output
				$checkboxlist = ""
				foreach($checkbox in $checkboxes){
					$var = $checkbox
					$val = (Get-Variable -Name "chk$checkbox").Value
					if($val -eq "Checked"){
						switch($var ){
							"Recipients"		{ $checkboxlist += "@{Name=`"Recipients`";Expression={`$_.Recipients}}," }
							"RecipientStatus"	{ $checkboxlist += "@{Name=`"RecipientStatus`";Expression={`$_.RecipientStatus}}," }
							"EventData"			{ $checkboxlist += "@{Name=`"EventData`";Expression={`$_.EventData}}," }
							"Reference"			{ $checkboxlist += "@{Name=`"Reference`";Expression={`$_.Reference}}," }
							default				{ $checkboxlist += "$var," }
						}
					}
				}
				$checkboxlist = $checkboxlist.TrimEnd(",")
				if($debug){ Write-Host -ForegroundColor $debugcolor "Checkboxes: $checkboxlist" }
				Invoke-Expression -Command "`$MTResults | select $checkboxlist | export-csv -NoTypeInformation -force -path `"$savefile`""
			}
		}else{
			Write-Host -ForegroundColor Yellow "`nNo results found.`n"
		}
		
		Write-Host -ForegroundColor DarkGray "`nMessage track complete."
		#if($exportToCSV -and $MTResults){ Invoke-Expression "notepad `"$savefile`"" }
		if($exportToCSV -and $MTResults){ Invoke-Item $path }
	}else{
		Write-Host -ForegroundColor red "Window closed. Message track canceled."
	}
	Remove-Variable mtresults,execstr -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
	$again = PromptBool @("&Yes","&No") "Do you want to run another message track?"
}while($again -eq "`$true")

IwasUsed $MyInvocation.MyCommand.Name $ExecStr