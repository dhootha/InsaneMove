<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
	Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: InsaneMove.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.26
	Last Modified	: 11-07-2016
.LINK
	Source Code
	http://www.github.com/spjeff/insanemove
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='CSV list of source and destination SharePoint site URLs to copy to Office 365.')]
	[string]$fileCSV,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Verify all Office 365 site collections.  Prep step before real migration.')]
	[Alias("v")]
	[switch]$verifyCloudSites = $false
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

# Config
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$settings = Get-Content "$root\InsaneMove.xml"
$maxWorker = $settings.settings.maxWorker

Function VerifyPSRemoting() {
	$ssp = Get-WSManCredSSP
	if ($ssp[0] -match "not configured to allow delegating") {
		# Enable remote PowerShell over CredSSP authentication
		Enable-WSManCredSSP -DelegateComputer * -Role Client -Force
		Restart-Service WinRM
	}
}

Function ReadIISPW {
	# Read IIS password for current logged in user
	Write-Host "===== Read IIS PW ===== $(Get-Date)" -Fore Yellow

	# Current user (ex: Farm Account)
	$domain = $env:userdomain
	$user = $env:username
	Write-Host "Logged in as $domain\$user"
	
	# Start IISAdmin if needed
	$iisadmin = Get-Service IISADMIN
	if ($iisadmin.Status -ne "Running") {
		#set Automatic and Start
		Set-Service -Name IISADMIN -StartupType Automatic -ErrorAction SilentlyContinue
		Start-Service IISADMIN -ErrorAction SilentlyContinue
	}
	
	# Attempt to detect password from IIS Pool (if current user is local admin and farm account)
	Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null
	$m = Get-Module WebAdministration
	if ($m) {
		#PowerShell ver 2.0+ IIS technique
		$appPools = Get-ChildItem "IIS:\AppPools\"
		foreach ($pool in $appPools) {	
			if ($pool.processModel.userName -like "*$user") {
				Write-Host "Found - "$pool.processModel.userName
				$pass = $pool.processModel.password
				if ($pass) {
					break
				}
			}
		}
	} else {
		#PowerShell ver 3.0+ WMI technique
		$appPools = Get-CimInstance -Namespace "root/MicrosoftIISv2" -ClassName "IIsApplicationPoolSetting" -Property Name, WAMUserName, WAMUserPass | select WAMUserName, WAMUserPass
		foreach ($pool in $appPools) {	
			if ($pool.WAMUserName -like "*$user") {
				Write-Host "Found - "$pool.WAMUserName
				$pass = $pool.WAMUserPass
				if ($pass) {
					break
				}
			}
		}
	}

	# Prompt for password
	if (!$pass) {
		$sec = Read-Host "Enter password for $domain\$user" -AsSecureString
	} else {
		$sec = $pass | ConvertTo-SecureString -AsPlainText -Force
	}
	$global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList "$domain\$user", $sec
}

Function DetectVendor() {
	# SharePoint Servers in local farm
	$spservers = Get-SPServer |? {$_.Role -ne "Invalid"} | sort Address

	# Detect if Vendor software installed
	$coll = @()
	foreach ($s in $spservers) {
		$found = Get-ChildItem "\\$($s.Address)\C$\Program Files (x86)\Sharegate\Sharegate.exe" -ErrorAction SilentlyContinue
		if ($found) {
            #if ($s.Address -ne "PWSYS-APSP-03") {
			    $coll += $s.Address
            #}
		}
	}
	
	# Display and return
	$coll |% {Write-Host $_ -Fore Green}
	$global:servers = $coll
}

Function ReadCloudPW() {
	# Prompt for admin password
    return (Read-Host "Enter O365 Cloud Password for $($settings.settings.tenant.adminUser)")
}

Function CloseSession() {
	# Close remote PS sessions
	Get-PSSession | Remove-PSSession
}

Function CreateWorkers() {
	# Open worker sessions per server.  Runspace to create local SCHTASK on remote PC
    # Template command
    $cmd = @'
mkdir "d:\InsaneMove" -ErrorAction SilentlyContinue | Out-Null

Function VerifySchtask($name, $file) {
	$found = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
	if ($found) {
		$found | Unregister-ScheduledTask -Confirm:$false
	}

	$user = "domain\user"
	$pw = "password"
	
	$folder = Split-Path $file
	$a = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $file -WorkingDirectory $folder
	$p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType Password
	$task = New-ScheduledTask -Action $a -Principal $p
	return Register-ScheduledTask -TaskName $name -InputObject $task -Password $pw -User $user
}

VerifySchtask "worker1" "d:\InsaneMove\worker1.ps1"
'@

	# Loop available servers
	$global:workers = @()
	$i = 0
	foreach ($pc in $global:servers) {
		# Loop maximum worker
		$s = New-PSSession -ComputerName $pc -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
        $s
        1..$maxWorker |% {
            # create worker
            $curr = $cmd -replace "worker1","worker$i"
            Write-Host "CREATE Worker$i on $pc ..." -Fore Yellow
            $sb = [Scriptblock]::Create($curr)
            $result = Invoke-Command -Session $s -ScriptBlock $sb
            $result | ft
			
			# purge old worker XML output
			$resultfile = "\\$pc\d$\insanemove\worker$i.xml"
            Remove-Item $resultfile -confirm:$false -ErrorAction SilentlyContinue
			
            # track worker
			$obj = New-Object -TypeName PSObject -Prop (@{"Id"=$i;"PC"=$pc})
			$global:workers += $obj			
			$i++
		}
	}
	$global:workers | ft
}

Function CreateTracker() {
	# CSV migration source/destination URL
	Write-Host "===== Populate Tracking table ===== $(Get-Date)" -Fore Yellow
	$i = 0	
	$j = 0
	$global:track = @()
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		# Assign each row to a Worker
		$pc = $global:workers[$j].PC
		
		# Get SharePoint total storage
		$site = Get-SPSite $row.SourceURL
		if ($site) {
			$SPStorage = [Math]::Round($site.Usage.Storage/1MB,2)
		}

		# Add row
		$obj = New-Object -TypeName PSObject -Prop (@{
			"SourceURL"=$row.SourceURL;
			"DestinationURL"=$row.DestinationURL;
			"CsvID"=$i;
			"WorkerID"=$j;
			"PC"=$pc;
			"Status"="New";
			"SGResult"="";
			"SGServer"="";
			"SGSessionId"="";
			"SGSiteObjectsCopied"="";
			"SGItemsCopied"="";
			"SGWarnings"="";
			"SGErrors"="";
			"Error"="";
			"ErrorCount"="";
			"TaskXML"="";
			"SPStorage"=$SPStorage;
			"TimeCopyStart"="";
			"TimeCopyEnd"=""
		})
		$global:track += $obj

		# Increment ID
		$i++
		$j++
		if ($j -ge $global:workers.count) {
			# Reset, back to first Session
			$j = 0
		}
	}
	
	# Display
	Get-PSSession | ft -a
}

Function UpdateTracker () {
	# Update tracker with latest SCHTASK status
	$active = $global:track |? {$_.Status -eq "InProgress"}
	foreach ($row in $active) {
		# Monitor remote SCHTASK
		$wid = $row.WorkerID
        $pc = $row.PC
		
		# Reconnect Broken
		$b = Get-PSSession |? {$_.State -eq "Broken"}
		if ($b) {
			# Make new session
			New-PSSession -ComputerName $b.ComputerName -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
			
			# Close old session
			$b | Remove-PSSession
		}
		
		# Check SCHTASK State=Ready
		$s = Get-PSSession |? {$_.ComputerName -eq $pc}
		$cmd = "Get-Scheduledtask -TaskName 'worker$wid'"
		$sb = [Scriptblock]::Create($cmd)
		$schtask = Invoke-Command -Session $s -Command $sb
		if ($schtask) {
			$schtask | select {$pc},TaskName,State | ft -a
			if ($schtask.State -eq 3) {
				$row.Status = "Completed"
				$row.TimeCopyEnd = (Get-Date).ToString()
				
				# Do we have ShareGate XML?
				$resultfile = "\\$pc\d$\insanemove\worker$wid.xml"
				if (Test-Path $resultfile) {
					# Read XML
					$x = $null
					[xml]$x = Get-Content $resultfile
					if ($x) {
						# Parse XML nodes
						$row.SGServer = $pc
						$row.SGResult = ($x.Objs.Obj.Props.S |? {$_.N -eq "Result"})."#text"
						$row.SGSessionId = ($x.Objs.Obj.Props.S |? {$_.N -eq "SessionId"})."#text"
						$row.SGSiteObjectsCopied = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "SiteObjectsCopied"})."#text"
						$row.SGItemsCopied = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "ItemsCopied"})."#text"
						$row.SGWarnings = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "Warnings"})."#text"
						$row.SGErrors = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "Errors"})."#text"
						
						# TaskXML
						$row.TaskXML = $x.OuterXml
						
						# Delete XML
						Remove-Item $resultfile -confirm:$false -ErrorAction SilentlyContinue
					}

					# Error
					$err = ""
					$errcount = 0
					$task.Error |% {
						$err += ($_|ConvertTo-Xml).OuterXml
						$errcount++
					}
					$row.ErrorCount = $errCount
				}
			}
		}
	}
}

Function ExecuteSiteCopy($row, $worker) {
	# Parse fields
	$name = $row.Name
	$srcUrl = $row.SourceURL
	$destUrl = FormatCloudMP $row.DestinationURL
	
	# Make NEW Session - remote PowerShell
    $wid = $worker.Id	
    $pc = $worker.PC
	$s = Get-PSSession |? {$_.ComputerName -eq $pc}
	
	# Generate PS1 worker script
	$ps = "Start-Transcript ""d:\insanemove\log\worker$wid.log"";`n`$secpw=""$global:cloudPW"" | ConvertTo-SecureString -AsPlainText -Force;`n`$cred = New-Object System.Management.Automation.PSCredential (""$($settings.settings.tenant.adminUser)"", `$secpw);`nImport-Module ShareGate;`n`$src=`$null;`n`$dest=`$null;`n`$src = Connect-Site ""$srcUrl"";`n`$dest = Connect-Site ""$destUrl"" -Credential `$cred;`n`$result=Copy-Site -Site `$src -DestinationSite `$dest -Merge -InsaneMode -VersionLimit 100;`n`$result | Export-Clixml ""d:\insanemove\worker$wid.xml"" -Force;`nStop-Transcript"
    $ps | Out-File "\\$pc\d$\insanemove\worker$wid.ps1" -Force
    Write-Host $ps -Fore Yellow

    # Invoke SCHTASK
    $cmd = "Get-ScheduledTask -TaskName ""worker$wid"" | Start-ScheduledTask"
	
	# Display
    Write-Host "START worker $wid on $pc" -Fore Green
	Write-Host "$srcUrl,$destUrl" -Fore yellow

	# Execute
	$sb = [Scriptblock]::Create($cmd) 
	return Invoke-Command $sb -Session $s
}

Function WriteCSV() {
    # Write new CSV output with detailed results
    $file = $fileCSV.Replace(".csv", "-results.csv")
    $global:track | select SourceURL,DestinationURL,CsvID,WorkerID,PC,Status,SGResult,SGServer,SGSessionId,SGSiteObjectsCopied,SGItemsCopied,SGWarnings,SGErrors,Error,ErrorCount,TaskXML,SPStorage | Export-Csv $file -NoTypeInformation -Force
}

Function CopySites() {
	# Monitor and Run loop
	Write-Host "===== Start Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	CreateTracker
	
	$i = 0
	do {
		$i++
		# Get latest Job status
		UpdateTracker
		Write-Host "." -NoNewline
		
		# Ensure all sessions are active
		foreach ($worker in $global:workers) {
			# Count active sessions per server
			$wid = $worker.Id
			$active = $global:track |? {$_.Status -eq "InProgress" -and $_.WorkerID -eq $wid}

            # Available session.  Assign new work
			if (!$active) {
				# Next row
                $row = $global:track |? {$_.Status -eq "New" -and $_.WorkerID -eq $wid}
			
                if ($row) {
                    if ($row -is [Array]) {
                        $row = $row[0]
                    }

                    # Kick off copy
					Sleep 5
				    $result = ExecuteSiteCopy $row $worker

				    # Update DB tracking
				    $row.Status = "InProgress"
					$row.TimeCopyStart = (Get-Date).ToString()
                }
			}
				
			# Progress bar %
			$complete = ($global:track |? {$_.Status -eq "Completed"}).Count
			$total = $global:track.Count
			$prct = [Math]::Round(($complete/$total)*100)
			
			# ETA
			if ($prct) {
				$elapsed = (Get-Date) - $start
				$remain = ($elapsed.TotalSeconds) / ($prct / 100.0)
				$eta = (Get-Date).AddSeconds($remain - $elapsed.TotalSeconds)
			}
			
			# Display
			Write-Progress -Activity "Copy site - ETA $eta" -Status "$name ($prct %)" -PercentComplete $prct

			# Detail table
			$global:track |? {$_.Status -eq "InProgress"} | select CsvID,WorkerID,PC,SourceURL,DestinationURL | ft -a
			$grp = $global:track | group Status
			$grp | select Count,Name | sort Name | ft -a
		}
		
		# Write Csv
		if ($i -gt 5) {
			WriteCSV
			$i = 0
		}

		# Latest counter
		$remain = $global:track |? {$_.status -ne "Completed" -and $_.status -ne "Failed"}
		sleep 5
	} while ($remain)
	
	# Complete
	Write-Host "===== Finish Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	$global:track | group status | ft -a
	$global:track | select CsvID,JobID,SessionID,SGSessionId,PC,SourceURL,DestinationURL | ft -a
}

Function VerifyCloudSites() {
	# Read CSV and ensure cloud sites exists for each row
	Write-Host "===== Verify Site Collections exist in O365 ===== $(Get-Date)" -Fore Yellow
	
	# Connect-SPO
	$secpw = ConvertTo-SecureString -String $global:cloudPW -AsPlainText -Force
	$c = New-Object System.Management.Automation.PSCredential ($settings.settings.tenant.adminUser, $secpw)
	Connect-SPOService -URL $settings.settings.tenant.adminURL -Credential $c
	
	# Loop CSV
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$row | ft
		EnsureCloudSite $row.SourceURL $row.DestinationURL $true
	}
}

Function EnsureCloudSite($srcUrl, $destUrl, $noWait) {
	# Create site in O365 if does not exist
	$destUrl = FormatCloudMP $destUrl
	Write-Host $destUrl -Fore Yellow
	$web = (Get-SPSite $srcUrl).RootWeb
	if ($web.RequestAccessEmail) {
		$rae = $web.RequestAccessEmail.Split(",;")[0].Split("@")[0] + "@" + $settings.settings.tenant.suffix;
	}
	if (!$rae) {
		$rae = $settings.settings.tenant.adminUser
	}
	
	# Verify SPOUser
     try {
	    $u = Get-SPOUser -Site $settings.settings.tenant.adminURL -LoginName $rae -ErrorAction SilentlyContinue
    } catch {}
	if (!$u) {
		$rae = $settings.settings.tenant.adminUser
	}
	
	# Verisy SPOSite
	try {
		$cloud = Get-SPOSite $destUrl -ErrorAction SilentlyContinue
	} catch {}
	if (!$cloud) {
		Write-Host "- CREATING $destUrl"
		if ($noWait) {
			New-SPOSite -Owner $rae -Url $destUrl -NoWait -StorageQuota (1024*50)
		} else {
			New-SPOSite -Owner $rae -Url $destUrl -StorageQuota (1024*50)
		}
	} else {
		Write-Host "- FOUND $destUrl"
	}
}

Function FormatCloudMP($url) {
	# Replace Managed Path with O365 /sites/ only
	$managedPath = "sites"
	$i = $url.indexOf("://")+3
	$split = $url.substring($i, $url.length-$i).Split("/")
	$split[1] = $managedPath
	$final = ($url.substring(0,$i) + ($split -join "/")).replace("http:","https:")
	return $final
}

Function Main() {
	# Start LOG
	$start = Get-Date
	$when = $start.ToString("yyyy-MM-dd-hh-mm-ss")
	$logFile = "$root\log\InsaneMove-$when.txt"
	mkdir "$root\log" -ErrorAction SilentlyContinue | Out-Null
	if (!$psISE) {Start-Transcript $logFile}
	Write-Host "fileCSV = $fileCSV"

	# Core
	if ($verifyCloudSites) {
		$global:cloudPW = ReadCloudPW
		VerifyCloudSites
	} else {
		VerifyPSRemoting
		ReadIISPW
		$global:cloudPW = ReadCloudPW
		DetectVendor
		CloseSession
		CreateWorkers
		CopySites
		CloseSession
		WriteCSV
	}
	
	# Finish LOG
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th				= [Math]::Round(((Get-Date) - $start).TotalHours, 2)
	$attemptMb		= ($global:track |measure SPStorage -Sum).Sum
	$actualMb		= ($global:track |? {$_.SGSessionId -ne ""} |measure SPStorage -Sum).Sum
	$actualSites	= ($global:track |? {$_.SGSessionId -ne ""}).Count
	Write-Host ("Duration Hours              : {0:N2}" -f $th) -Fore Yellow
	Write-Host ("Total Sites Attempted       : {0}" -f $($global:track.count)) -Fore Green
	Write-Host ("Total Sites Copied          : {0}" -f $actualSites) -Fore Green
	Write-Host ("Total Storage Attempted (MB): {0:N0}" -f $attemptMb) -Fore Green
	Write-Host ("Total Storage Copied (MB)   : {0:N0}" -f $actualMb) -Fore Green
	Write-Host ("Total Objects               : {0:N0}" -f $(($global:track |measure SGItemsCopied -Sum).Sum)) -Fore Green
	Write-Host ("Total Worker Threads        : {0}" -f $maxWorker) -Fore Green
	Write-Host "====="  -Fore Yellow
	Write-Host ("GB per Hour                 : {0:N2}" -f (($actualMb/1KB)/$th)) -Fore Green
	Write-Host $fileCSV
	if (!$psISE) {Stop-Transcript}
}

Main