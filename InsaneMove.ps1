<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
	Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: InsaneMove.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.12
	Last Modified	: 08-31-2016
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
[xml]$settings = Get-Content "$root\settings.xml"
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
			$coll += $s.Address
		}
	}
	
	# Display and return
	$coll | ft
	$global:servers = $coll
}

Function ReadAdminPW() {
	# Prompt for admin password
    return (Read-Host "Enter Password for $($settings.settings.tenant.adminUser)")

}

Function CloseSession() {
	# Close remote PS sessions
	Get-PSSession | Remove-PSSession
	Get-Job | Remove-Job
}

Function OpenSession() {
	# Open worker sessions per server.  Runspace to execute jobs
	# Loop available servers
	foreach ($server in $global:servers) {
		# Loop maximum worker
		1..$maxWorker |% {
			New-PSSession -ComputerName $server -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
		}
	}
}

Function CreateTracker() {
	# CSV migration source/destination URL
	Write-Host "===== Populate Tracking table ===== $(Get-Date)" -Fore Yellow
	$i = 0	
	$j = 0
	$global:track = @()
	$csv = Import-Csv $fileCSV
	$sessions = Get-PSSession
	foreach ($row in $csv) {
		# Assign each row to a Session
		$sid = (Get-PSSession)[$j].Id
		$pc = (Get-PSSession)[$j].ComputerName
		
		# Get SharePoint total storage
		$SPStorage = [Math]::Round((Get-SPSite $row.SourceURL).Usage.Storage/1MB,2)

		# Add row
		$obj = New-Object -TypeName PSObject -Prop (@{"SourceURL"=$row.SourceURL;"DestinationURL"=$row.DestinationURL;"CsvID"=$i;"SessionID"=$sid;"JobID"=0;"PC"=$pc;"Status"="New";"Log"="";;"SGResult"="";"SGServer"="";"SGSessionId"="";"SGSiteObjectsCopied"="";"SGItemsCopied"="";"SGWarnings"="";"SGErrors"="";"Error"="";"ErrorCount"="";"SPStorage"=$SPStorage})
		$global:track += $obj

		# Increment ID
		$i++
		$j++
		if ($j -ge $sessions.count) {
			# Reset, back to first Session
			$j = 0
		}
	}
	
	# Display
	Get-PSSession |ft -a
}

Function UpdateTracker () {
	# Update tracker with latest Job ID status
	$active = $global:track |? {$_.Status -eq "InProgress"}
	foreach ($row in $active) {
		# Monitor remote server job
		if ($row.JobID) {
			$job = Get-Job $row.JobID
            
            $status = ""
			if ($job.State -eq "Completed") {
				# Update DB tracking
				$status = "Completed"
            } elseif ($job.State -ne "Running" -and $job.State -ne "NotStarted") {              
				# Update DB tracking
				$status = "Failed"
           
			}

            # Update to save
            if ($status) {
                # Status
                $row.Status = $status


                 # Details from ShareGate
                $out = $job[0].ChildJobs[0].output
                if ($out) {
			        $row.SGServer = $out.PSComputerName
			        $row.SGResult = $out.Result
			        $row.SGSessionId = $out.SessionId
			        $row.SGSiteObjectsCopied = $out.SiteObjectsCopied
			        $row.SGItemsCopied = $out.ItemsCopied
			        $row.SGWarnings = $out.Warnings
			        $row.SGErrors = $out.Errors
                }

                # Error
                $err = ""
                $errcount = 0
                $job[0].ChildJobs[0].Error |% {
                    $err += ($_|ConvertTo-Xml).OuterXml
                    $errcount++
                }
                $job[0].Error |% {
                    $err += ($_|ConvertTo-Xml).OuterXml
                    $errcount++
                }
                $row.Error = $err
                $row.ErrorCount = $errCount

            }
			
		}
	}
}

Function ExecuteSiteCopy($row, $session) {
	# Parse fields
	$id = $row.Id
	$name = $row.Name
	$srcUrl = $row.SourceURL
	$destUrl = FormatCloudMP $row.DestinationURL
	
	# Core command
	$id = $session.Id
	$str = "`$secpw=""$global:adminPass"" | ConvertTo-SecureString -AsPlainText -Force`n`$cred = New-Object System.Management.Automation.PSCredential (""$($settings.settings.tenant.adminUser)"", `$secpw)`n`nImport-Module ShareGate`n`$src=`$null`n`$dest=`$null`n`$src = Connect-Site ""$srcUrl""`n`$dest = Connect-Site ""$destUrl"" -Credential `$cred`nCopy-Site -Site `$src -DestinationSite `$dest -Merge -InsaneMode -VersionLimit 100"
	
    Write-Host $session.ComputerName -Fore Green
	Write-Host $str -Fore yellow

	# Execute
	$cmd = [Scriptblock]::Create($str) 
	return Invoke-Command $cmd -Session $session -AsJob
}

Function WriteCSV() {
    # Write new CSV output with detailed results
    $file = $fileCSV.Replace(".csv", "-results.csv")
    $global:track | select SourceURL,DestinationURL,CsvID,SessionID,JobID,PC,Status,Log,SGResult,SGServer,SGSessionId,SGSiteObjectsCopied,SGItemsCopied,SGWarnings,SGErrors,Error,ErrorCount | Export-Csv $file -NoTypeInformation -Force
}

Function CopySites() {
	# Monitor and Run loop
	Write-Host "===== Start Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	CreateTracker
	
	do {
		# Get latest Job status
		UpdateTracker
		Write-Host "." -NoNewline
		
		# Ensure all sessions are active
		foreach ($session in Get-PSSession) {
			# Count active sessions per server
			$sid = $session.Id
			$active = $global:track |? {$_.Status -eq "InProgress" -and $_.SessionID -eq $sid}

            # Available session.  Assign new work
			if (!$active) {
				# Next row
                $row = $global:track |? {$_.Status -eq "New" -and $_.SessionID -eq $sid}
                if ($row) {
                    if ($row -is [Array]) {
                        $row = $row[0]
                    }

                    # Kick off copy
				    $result = ExecuteSiteCopy $row $session

				    # Update DB tracking
				    $row.JobID = $result.Id
				    $row.Status = "InProgress"
				    $row.Log = $global:log
                }
			}
				
			# Progress bar %
			$complete = ($global:track |? {$_.Status -eq "Completed"}).Count
			$total = $global:track.Count
			$prct = [Math]::Round(($complete/$total)*100)
			
			# ETA
			$elapsed = (Get-Date) - $start
			$remain = ($elapsed.TotalSeconds) / ($prct / 100.0)
			$eta = (Get-Date).AddSeconds($remain - $elapsed.TotalSeconds)
			
			# Display
			Write-Progress -Activity "Copy site - ETA $eta" -Status "$name ($prct %)" -PercentComplete $prct

			# Detail table
			$global:track |? {$_.Status -eq "InProgress"} | select CsvID,JobID,SessionID,PC,SourceURL,DestinationURL | ft -a
			$grp = $global:track | group Status
			$grp | select Count,Name | sort Name | ft -a
			
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
	$secpw = ConvertTo-SecureString -String $global:adminPass -AsPlainText -Force
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
	$web = (Get-SPSite $srcUrl).RootWeb;
	$rae = $web.RequestAccessEmail.Split(",;")[0].Split("@")[0] + "@" + $settings.settings.tenant.suffix;
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
		$global:adminPass = ReadAdminPW
		VerifyCloudSites
	} else {
		VerifyPSRemoting
		ReadIISPW
		$global:adminPass = ReadAdminPW
		DetectVendor
		CloseSession
		OpenSession
		CopySites
		CloseSession
		WriteCSV
	}
	
	# Finish LOG
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th = [Math]::Round(((Get-Date) - $start).TotalHours, 2)
	Write-Host "Duration Hours: $th`n" -Fore Yellow
	Write-Host "Total Sites: "$($global:track.count) -Fore Green
	Write-Host "Total Objects: "$(($global:track |measure SGItemsCopied -Sum).Sum) -Fore Green
	if (!$psISE) {Stop-Transcript}
}

$fileCSV = "custom.csv"
Main