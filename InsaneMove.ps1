<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
	Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: InsaneMove.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.10
	Last Modified	: 08-16-2016
.LINK
	Source Code
	http://www.github.com/spjeff/insanemove
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='CSV list of source and destination SharePoint site URLs to copy to Office 365.')]
	[string]$fileCSV,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Clear saved passwords in HKCU registry.')]
	[Alias("c")]
	[switch]$clearSavedPW = $false
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
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
		$found = Get-ChildItem "\\$($s.Address)\C$\Program Files (x86)\Sharegate\Sharegate.exe"
		if ($found) {
			$coll += $s.Address
		}
	}
	
	# Display and return
	$coll | ft
	$global:servers = $coll
}

Function ReadAdminPW() {
	# Registry HKCU folder
	$path = "HKCU:\Software\InsaneMove"
	if (!(Test-Path $path)) {md $path | Out-Null}
	$name = $settings.settings.tenant.adminUser
	
	# Do we need to clear old paswords?
	if ($clearSavedPW) {
		Remove-ItemProperty -Path $path -Name $name -Confirm:$false
		Write-Host "Deleted password OK for $name" -Fore Yellow
		Exit
	}
	
	# Do we have registry HKCU saved password?
	$hash = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue)."$name"
	
	# Prompt for input
	if (!$hash) {
		$sec = Read-Host "Enter Password for $($settings.settings.tenant.adminUser)" -AsSecureString
		if (!$sec) {
			Write-Error "Exit - No password given"
			Exit
		}
		$hash = $sec | ConvertFrom-SecureString
		
		# Prompt to save to HKCU
		$save = Read-Host "Save to HKCU registry (secure hash) [Y/N]?"
		if ($save -like "Y*") {
			Set-ItemProperty -Path $path -Name $name -Value $hash -Force
			Write-Host "Saved password OK for $name" -Fore Yellow
		}
	}
	
	# Return
	return $hash
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
		$obj = New-Object -TypeName PSObject -Prop (@{"SourceURL"=$row.SourceURL;"DestinationURL"=$row.DestinationURL;"CsvID"=$i;"SessionID"=$sid;"JobID"=0;"Status"="New";"SGResult"="";"SGServer"="";"SGSessionId"="";"SGSiteObjectsCopied"="";"SGItemsCopied"="";"SGWarnings"="";"SGErrors"=""})
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
	$global:track | select JobID,SessionID,CsvId,SourceURL,DestinationURL |ft -a
	Get-PSSession |ft -a
}

Function UpdateTracker () {
	# Update tracker with latest Job ID status
	$active = $global:track |? {$_.Status -eq "InProgress"}
	foreach ($row in $active) {
		# Monitor remote server job
		if ($row.JobID) {
			$job = Get-Job $row.JobID
			if ($job.State -eq "Completed") {
				# Update DB tracking
				$row.Status = "Completed"

                # Detailed output from ShareGate "Copy-Site" cmdlet
                $out = $job[0].ChildJobs[0].output
				$row.SGServer = $out.PSComputerName
				$row.SGResult = $out.Result
				$row.SGSessionId = $out.SessionId
				$row.SGSiteObjectsCopied = $out.SiteObjectsCopied
				$row.SGItemsCopied = $out.ItemsCopied
				$row.SGWarnings = $out.Warnings
				$row.SGErrors = $out.Errors
			} elseif ($job.State -ne "Running" -and $job.State -ne "NotStarted") {
				# Update DB tracking
				$row.Status = "Failed"
			}
		}
	}
}

Function ExecuteSiteCopy($row, $session) {
	# Parse fields
	$id = $row.Id
	$name = $row.Name
	$srcUrl = $row.SourceURL
	$destUrl = $row.DestinationURL
	
	# Core command
	$when = (Get-Date).ToString("yyyy-MM-dd-hh-mm-ss")
	$str = "`$logFile=""InsaneMove-Worker-$when.txt"";Start-Transcript `$logFile;`$hash = ""$global:adminPass"";`$secpw = ConvertTo-SecureString -String `$hash;`n`$cred = New-Object System.Management.Automation.PSCredential (""$($settings.settings.tenant.adminUser)"", `$secpw);`nImport-Module ShareGate;`n`$src = Connect-Site $srcUrl;`n`$dest = Connect-Site $destUrl -Credential `$cred;`nCopy-Site -Site `$src -DestinationSite `$dest -Merge -InsaneMode -VersionLimit 100;Stop-Transcript"
	Write-Host $str -Fore yellow

	# Execute
	$cmd = [Scriptblock]::Create($str) 
	return Invoke-Command $cmd -Session $session -AsJob
}

Function WriteCSV() {
    # Write new CSV output with detailed results
    $file = $fileCSV.Replace(".csv", "-results.csv")
    $global:track | SourceURL,DestinationURL,CsvID,SessionID,JobID,Status,SGResult,SGServer,SGSessionId,SGSiteObjectsCopied,SGItemsCopied,SGWarnings,SGErrors | Export-Csv $file -NoTypeInformation -Force
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
			if ($active.count -lt $maxWorker) {
			
				# Available session.  Assign new work
				$avail = $global:track |? {$_.Status -eq "New" -and $_.SessionID -eq $sid}
				if ($avail) {
					# Assign first available row
					if ($avail -is [array]) {
						$row = $avail[0]
					} else {
						$row = $avail
					}
					
					# Kick off site copy on remote session
					$result = ExecuteSiteCopy $row $session
					
					# Update DB tracking
					$row.JobID = $result.Id
					$row.Status = "InProgress"
				}
				
				# Progress bar %
				$counter = ($global:track |? {$_.Status -eq "Completed"}).Count
				$prct = [Math]::Round(($counter/$global:track.Count)*100)
				Write-Progress -Activity "Copy site" -Status "$name ($prct %)" -PercentComplete $prct

				# Detail table
				$global:track |? {$_.Status -ne "Completed"} | ft -a
				$grp = $global:track | group Status
				$grp | select Count,Name | sort Name | ft -a
			}
		}

		# Latest counter
		$remain = $global:track |? {$_.status -ne "Completed" -and $_.status -ne "Failed"}
		sleep 5
	} while ($remain)
	
	# Complete
	Write-Host "===== Finish Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	$global:track | group status | ft -a
	$global:track | ft -a
}

Function Main() {
	# Start LOG
	$start = Get-Date
	$when = $start.ToString("yyyy-MM-dd-hh-mm-ss")
	$logFile = "$root\log\InsaneMove-$when.txt"
	mkdir "$root\log" -ErrorAction SilentlyContinue | Out-Null
	Start-Transcript $logFile
	Write-Host "$fileCSV = $fileCSV"

	# Core 	
	VerifyPSRemoting
	ReadIISPW
	$global:adminPass = ReadAdminPW
	DetectVendor
	CloseSession
	OpenSession
	CopySites
	CloseSession
    WriteCSV
	
	# Finish LOG
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th = [Math]::Round(((Get-Date) - $start).TotalHours, 2)
	Write-Host "Duration Hours: $th" -Fore Yellow
	Stop-Transcript
}
Main