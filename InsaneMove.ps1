<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
	Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Name		: InsaneMove.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.01
	Last Modified	: 08-10-2016
.LINK
	Source Code
	http://www.github.com/spjeff/insanemove
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$False, ValueFromPipeline=$false, HelpMessage='CSV list of source SharePoint site URLs to copy to Office 365.')]
	[string]$csvSites
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

Function EnablePSRemoting() {
	$ssp = Get-WSManCredSSP
	if ($ssp[0] -match "not configured to allow delegating") {
		# Enable remote PowerShell over CredSSP authentication
		Enable-WSManCredSSP -DelegateComputer * -Role Client -Force
		Restart-Service WinRM
	}
}

Function ReadIISPW {
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
		$sec = Read-Host "Enter password: " -AsSecureString
	} else {
		$sec = $pass | ConvertTo-SecureString -AsPlainText -Force
	}
	$global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList "$domain\$user",$sec
}

Function DetectVendor() {
	# Servers
	$servers = Get-SPServer |? {$_.Role -ne "Invalid"} | sort Address

	# Detect if Vendor software installed
	$coll = @()
	foreach ($s in $servers) {
		$found = Get-ChildItem "\\$($s.Address)\c$\program files\ShareGate\ShareGate.exe"
		if ($found) {
			$coll += $s.Address
		}
	}
	$global:servers = $coll
}
Function UpgradeContent() {
	Write-Host "===== Copy Sites to O365 ===== $(Get-Date)" -Fore Yellow
	
	# Tracking table - assign DB to server
	$rows = Import-CSV $csvSites
	$maxWorkers = 4
	$i = 0
	$track = @()
	foreach ($row in $rows) {
		# Assign to SPServer
		$mod = $i % $global:servers.count
		$pc = $global:servers[$mod].Address
		
		# Collect
		$obj = New-Object -TypeName PSObject -Prop (@{"Name"=$row.URL;"Id"=$row.Id;"UpgradePC"=$pc;"JID"=0;"Status"="New"})
		$track += $obj
		$i++
	}
	$track | Format-Table -Auto

	# Cleanup
	Get-PSSession | Remove-PSSession
	Get-Job | Remove-Job
	
	# Open sessions
	foreach ($server in $global:servers) {
		$remote = New-PSSession -ComputerName $server -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
		if (!$remote) {
			$remote = New-PSSession -ComputerName $server -Credential $global:cred -Authentication Negotiate -ErrorAction SilentlyContinue 
		}
	}

	# Monitor and Run loop
	do {
		# Get latest PID status
		$active = $track |? {$_.Status -eq "InProgress"}
		foreach ($row in $active) {
			# Monitor remote server job
			if ($row.JID) {
				$job = Get-Job $row.JID
				if ($job.State -eq "Completed") {
					# Update DB tracking
					$row.Status = "Completed"
				} elseif ($job.State -eq "Failed") {
					# Update DB tracking
					$row.Status = "Failed"
				} else {
					Write-host "-" -NoNewline
				}
			}
		}
		
		# Ensure workers are active
		foreach ($server in $global:servers) {
			# Count active workers per server
			$active = $track |? {$_.Status -eq "InProgress" -and $_.UpgradePC -eq $server}
			if ($active.count -lt $maxWorkers) {
			
				# Choose next available DB
				$avail = $track |? {$_.Status -eq "New" -and $_.UpgradePC -eq $server}
				if ($avail) {
					if ($avail -is [array]) {
						$row = $avail[0]
					} else {
						$row = $avail
					}
				
					# Kick off new worker
					$id = $row.Id
					$name = $row.Name
					$remoteStr = "Add-PSSnapIn ShareGate; $src = Connect-SGSite $srcUrl; $dest = Connect-SGSite $destUrl; Copy-SGSite -Source $src -Destinatoin $dest -InsaneMode"
					
					# Run on remote server
					$remoteCmd = [Scriptblock]::Create($remoteStr) 
					$pc = $server
					Write-Host $pc -fore green
					Get-PSSession | Format-Table -AutoSize
					$session = Get-PSSession |? {$_.ComputerName -like "$pc*"}
					$result = Invoke-Command $remoteCmd -Session $session -AsJob
					
					# Update DB tracking
					$row.JID = $result.Id
					$row.Status = "InProgress"
				}
				
				# Progress
				$counter = ($track |? {$_.Status -eq "Completed"}).Count
				$prct = [Math]::Round(($counter/$track.Count)*100)
				Write-Progress -Activity "Upgrade database" -Status "$name ($prct %)" -PercentComplete $prct
				$track | Format-Table -AutoSize
			}
		}

		# Latest counter
		$remain = $track |? {$_.status -ne "Completed" -and $_.status -ne "Failed"}
	} while ($remain)
	Write-Host "===== DONE ====="
	$track | group status | Format-Table -AutoSize
	$track | Format-Table -AutoSize
	
	# Cleanup
	Get-PSSession | Remove-PSSession
	Get-Job | Remove-Job
}

function Main() {
	# Start
	$start = Get-Date
	$when = $start.ToString("yyyy-MM-dd-hh-mm-ss")
	$logFile = "$root\log\InsaneMove-$when.txt"
	mkdir "$root\log" -ErrorAction SilentlyContinue | Out-Null
	Start-Transcript $logFile

	# Core 	
	EnablePSRemoting
	ReadIISPW
	DetectVendor
	CopySites
	
	# Finish
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th = [Math]::Round(((Get-Date) - $start).TotalHours, 2)
	Write-Host "Duration Hours: $th" -Fore Yellow
	
	# Cleanup
	Stop-Transcript
}
Main