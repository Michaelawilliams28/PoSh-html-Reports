Function Test-htmlModule {
	$script:Mod = Get-Module -ListAvailable -Name "ReportHTML"
	If (!$Mod) {
		# Write-Host "ReportHTML Module is not present, attempting to install it"
		Install-Module -Name ReportHTML -Force
		Import-Module ReportHTML -ErrorAction SilentlyContinue
	}
}
Function get-reportSettings {
	
	#                               Log
	$script:log = "$currentdir\$logdate.log"
	#                               report dir
	$script:reportdir = "$currentdir\report"

}
Function Write-ProgressHelper {
	param(
		[int]$StepNumber,
		[string]$Message
	)
	Write-Progress -Activity 'Building Report...' -Status $Message -PercentComplete (($StepNumber / $steps) * 100)
}
# Write progess to log file
Function New-LogWrite {
	Param ([string]$logstring)
	Add-content $log -value $logstring
}
# Wait Timer
Function New-WriteTime {
	Param ([Int]$time)
	Start-Sleep -Milliseconds $time
}
# Check and create for report folder
Function New-reportFolder {
	Write-ProgressHelper 1 "New-reportFolder"
	New-LogWrite "[$loggingDate]  Function New-reportFolder"
	If ($reportdir ) {
		If (Test-Path $reportdir) { 
			New-LogWrite  "[$loggingDate] report folder exists! "
		} 
		ElseIf (!(Test-Path $reportdir)) {  
			New-LogWrite  "[$loggingDate] Creating report folder "
			New-Item  $currentdir -Name "report" -ItemType "directory" | Out-File -Append -Encoding Default  $log
			If (!$?) {
				New-LogWrite  "Failed Creating report folder "
                
			}
		}
	}
	ElseIf (!$reportdir) {
                    
		New-LogWrite "[$loggingDate]  Failed New-reportFolder"
        
	}
	New-WriteTime 300
}

# Log starting Header
Function Set-Header {

	New-LogWrite  "[$loggingdate] Starting report....."
	New-LogWrite  "[$loggingdate] Report Settings....."
	New-LogWrite  "[$loggingdate] `$CompanyLogo = $CompanyLogo"
	New-LogWrite  "[$loggingdate] `$RightLogo = $RightLogo"
	New-LogWrite  "[$loggingdate] `$ReportTitle $ReportTitle"
	New-LogWrite  "[$loggingdate] `$ReportSavePath = $ReportSavePath"
	New-LogWrite  "[$loggingdate] `$Days = $Days"
	New-LogWrite  "[$loggingdate]  `$UserCreatedDays = $UserCreatedDays"
	New-LogWrite  "[$loggingdate] `$DaysUntilPWExpireINT = $DaysUntilPWExpireINT"
	New-LogWrite  "[$loggingdate] `$ADModNumber = $ADModNumber"
    
	New-WriteTime 300
}
# Convert object time
Function Set-FileTime ($FileTime) {
	New-LogWrite  "[$loggingdate] Function Set-FileTime"
	New-LogWrite  "[$loggingdate] `$FileTime $FileTime"
	$Date = [DateTime]::FromFileTime($FileTime)
	if ((!$date ) -or $Date -lt (Get-Date '1/1/1900') -or $date -eq 0) {
		'Never'
	}
	else {
		$Date
	}
}
# Default Security Groups
Function Get-DefaultSecurityGroups {
	Write-ProgressHelper 2 "Get-DefaultSecurityGroups"
	New-LogWrite  "[$loggingdate] Function Get-DefaultSecurityGroups " 
	$script:DefaultSGs = @(
		'Access Control Assistance Operators'
		'Account Operators'
		'Administrators'
		'Allowed RODC Password Replication Group'
		'Backup Operators'
		'Certificate Service DCOM Access'
		'Cert Publishers'
		'Cloneable Domain Controllers'
		'Cryptographic Operators'
		'Denied RODC Password Replication Group'
		'Distributed COM Users'
		'DnsUpdateProxy'
		'DnsAdmins'
		'Domain Admins'
		'Domain Computers'
		'Domain Controllers'
		'Domain Guests'
		'Domain Users'
		'Enterprise Admins'
		'Enterprise Key Admins'
		'Enterprise Read-only Domain Controllers'
		'Event Log Readers'
		'Group Policy Creator Owners'
		'Guests'
		'Hyper-V Administrators'
		'IIS_IUSRS'
		'Incoming Forest Trust Builders'
		'Key Admins'
		'Network Configuration Operators'
		'Performance Log Users'
		'Performance Monitor Users'
		'Print Operators'
		'Pre-Windows 2000 Compatible Access'
		'Protected Users'
		'RAS and IAS Servers'
		'RDS Endpoint Servers'
		'RDS Management Servers'
		'RDS Remote Access Servers'
		'Read-only Domain Controllers'
		'Remote Desktop Users'
		'Remote Management Users'
		'Replicator'
		'Schema Admins'
		'Server Operators'
		'Storage Replica Administrators'
		'System Managed Accounts Group'
		'Terminal Server License Servers'
		'Users'
		'Windows Authorization Access Group'
		'WinRMRemoteWMIUsers'
	)
	New-WriteTime 300
}
# Create Empty Tables
Function Get-CreateObjects {
	Write-ProgressHelper 3 "Get-CreateObjects"
	New-LogWrite  "[$loggingdate] Function Get-CreateObjects " 
	$script:Table = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupTypetable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultGrouptable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:EnabledDisabledUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DomainAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ExpiringAccountsTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:securityeventtable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUGPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupMembershipTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:PasswordExpirationTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:PasswordExpireSoonTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:EnterpriseAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:NewCreatedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ADObjectTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ProtectedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputerProtectedTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputersEnabledTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultComputersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultUsersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPGroupsTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GraphComputerOS = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:Type = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:LinkedGPOs = New-Object 'System.Collections.Generic.List[System.Object]'
	# $script:userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:recentgpostable = New-Object 'System.Collections.Generic.List[System.Object]'

	New-WriteTime 300
}
# Get All AD Users
Function Get-AllUsers {
	Write-ProgressHelper 4 "Get-AllUsers"
	New-LogWrite  "[$loggingdate] Function Get-AllUsers"
	$script:AllUsers = Get-ADUser -Filter * -Properties *
	If (!$AllUsers) {
		New-LogWrite  "[$loggingdate] (!`$AllUsers) $true "
	}
	New-WriteTime 300
}
# Get all GPOs
Function Get-AllGPOs {
	Write-ProgressHelper 5 "Get-AllGPOs"
	New-LogWrite  "[$loggingdate] Function Get-AllGPOs"
	$script:GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, ModificationTime, @{ Label = 'ComputerVersion'; Expression = { $_.computer.dsversion } }, @{ Label = 'UserVersion'; Expression = { $_.user.dsversion } }
	If (!$GPOs) {
		New-LogWrite  "[$loggingdate] (!`$GPOs) $true"
	}
	New-WriteTime 300
}
# Get AD Objects
Function Get-ADOBJ {
	Write-ProgressHelper 6 "Get-ADOBJ"
	New-LogWrite  "[$loggingdate] Function Get-ADOBJ "
	$setdate = (Get-Date).AddDays(- $ADModNumber)
	If ($setdate) {
		$script:ADObjs = Get-ADObject -Filter { whenchanged -gt $setdate -and ObjectClass -ne 'domainDNS' -and ObjectClass -ne 'rIDManager' -and ObjectClass -ne 'rIDSet' } -Properties *
		If (!$ADObjs) {
			New-LogWrite  "[$loggingdate] (!`$ADObjs) $true"
		}
	}
	New-WriteTime 300
}
# AD Objects to tables
Function Set-ADOBJ {
	Write-ProgressHelper 7 "Set-ADOBJ"
	New-LogWrite  "[$loggingdate] Function Set-ADOBJ "
	If ($ADObjs) {
		foreach ($ADObj in $ADObjs) {
			if ($ADObj.ObjectClass -eq 'GroupPolicyContainer') {
				$Name = $ADObj.DisplayName
			}
			else {
				$Name = $ADObj.Name
			}
			$adobjecttype = [PSCustomObject]@{
				'Name'         = $Name
				'Object Type'  = $ADObj.ObjectClass
				'When Changed' = $ADObj.WhenChanged
			}
			$script:ADObjectTable.Add($adobjecttype)
		}
	}
	ElseIf (!$ADObjs) {
		New-LogWrite  "[$loggingdate] (!`$ADObjs)  $true "
	}
	New-WriteTime 300
}
#  AD Trash Bin
Function Get-ADRecBin {
	Write-ProgressHelper 8 "Get-ADRecBin"
	New-LogWrite  "[$loggingdate] Function Get-ADRecBin"
	$ADRecycleBinStatus = (Get-ADOptionalFeature -Identity "Recycle Bin Feature").EnabledScopes
	If ($ADRecycleBinStatus) {
		if ($ADRecycleBinStatus.Count -lt 1) {
			$script:ADRecycleBin = 'Disabled'
		}
		else {	
			$script:ADRecycleBin = 'Enabled'
		}
	}
	ElseIf (!$ADRecycleBinStatus) {
		New-LogWrite  "[$loggingdate] (!`$ADRecycleBinStatus)  $true "
	}
	New-WriteTime 300
}
#  Get AD Domain, Forest info
Function Get-ADInfo {
	Write-ProgressHelper 9 "Get-ADInfo"
	New-LogWrite  "[$loggingdate] Function Get-ADInfo"
	$script:ADInfo = Get-ADDomain
	$script:ForestObj = Get-ADForest
	$script:DomainControllerobj = Get-ADDomain
    
	$script:Forest = $ADInfo.Forest
	$script:InfrastructureMaster = $DomainControllerobj.InfrastructureMaster
	$script:RIDMaster = $DomainControllerobj.RIDMaster
	$script:PDCEmulator = $DomainControllerobj.PDCEmulator
	$script:DomainNamingMaster = $ForestObj.DomainNamingMaster
	$script:SchemaMaster = $ForestObj.SchemaMaster


	$addomains = [PSCustomObject]@{
		'Domain'                = $Forest
		'AD Recycle Bin'        = $ADRecycleBin
		'Infrastructure Master' = $InfrastructureMaster
		'RID Master'            = $RIDMaster
		'PDC Emulator'          = $PDCEmulator
		'Domain Naming Master'  = $DomainNamingMaster
		'Schema Master'         = $SchemaMaster
	}
	$script:CompanyInfoTable.Add($addomains)
	New-WriteTime 300
}
# Get Created Users
Function Get-CreatedUsers {
	Write-ProgressHelper 10 "Get-CreatedUsers"
	New-LogWrite  "[$loggingdate] Function Get-CreatedUsers"
	# Get newly created users
	$When = ((Get-Date).AddDays(- $UserCreatedDays)).Date
	If ($When) {
		$NewUsers = $AllUsers | Where-Object { $_.whenCreated -ge $When }
		If ($NewUsers) {
			foreach ($Newuser in $Newusers) {
				$createduserss = [PSCustomObject]@{
					'Name'          = $Newuser.Name
					'Enabled'       = $Newuser.Enabled
					'Creation Date' = $Newuser.whenCreated
				}
				$script:NewCreatedUsersTable.Add($createduserss)
			}
		}
	}
	New-WriteTime 300
}
Function Get-DomainAdmins {
	Write-ProgressHelper 11 "Get-DomainAdmins"
	New-LogWrite  "[$loggingdate] Function Get-DomainAdmins"
	$script:DomainAdminMembers = Get-ADGroupMember 'Domain Admins'
	New-WriteTime 300
}
# Get Domain Admins
Function Set-DomainAdminsTable {
	Write-ProgressHelper 12 "Set-DomainAdminsTable"
	New-LogWrite  "[$loggingdate] Function Set-DomainAdminsTable"
	# Get Domain Admins
    
	If ($DomainAdminMembers) { 
		foreach ($DomainAdminMember in $DomainAdminMembers) {
			$Name = $DomainAdminMember.Name
			$Type = $DomainAdminMember.ObjectClass
			$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
			$adgroupmemebers = [PSCustomObject]@{
				'Name'    = $Name
				'Enabled' = $Enabled
				'Type'    = $Type
			}
			$script:DomainAdminTable.Add($adgroupmemebers)
			New-LogWrite  "[$loggingdate] `$adgroupmemebers $adgroupmemebers"
		}
	}
	New-WriteTime 300
}
# Get Enterprise Admins
Function Get-EntAdmins {
	Write-ProgressHelper 13 "Get-EntAdmins"
	New-LogWrite  "[$loggingdate] Function Get-EntAdmins "
	# Get Enterprise Admins
	$EnterpriseAdminsMembers = Get-ADGroupMember 'Enterprise Admins' -Server $SchemaMaster
	If ($EnterpriseAdminsMembers) {
		foreach ($EnterpriseAdminsMember in $EnterpriseAdminsMembers) {
			$Name = $EnterpriseAdminsMember.Name
			$Type = $EnterpriseAdminsMember.ObjectClass
			$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
			$entadmin = [PSCustomObject]@{
				'Name'    = $Name
				'Enabled' = $Enabled
				'Type'    = $Type
			}
			$script:EnterpriseAdminTable.Add($entadmin)
		}
	}
	New-WriteTime 300
}
# Get Ad Computers
Function Get-Computers {
	Write-ProgressHelper 14 "Get-Computers"
	New-LogWrite  "[$loggingdate] Function Get-Computers"
	$DefaultComputersOU = (Get-ADDomain).computerscontainer
	If ($DefaultComputersOU) {
		$DefaultComputers = Get-ADComputer -Filter * -Properties * -SearchBase "$DefaultComputersOU"
		If ($DefaultComputers) {
			foreach ($DefaultComputer in $DefaultComputers) {
				$computersenabled = [PSCustomObject]@{
					'Name'                  = $DefaultComputer.Name
					'Enabled'               = $DefaultComputer.Enabled
					'Operating System'      = $DefaultComputer.OperatingSystem
					'Modified Date'         = $DefaultComputer.Modified
					'Password Last Set'     = $DefaultComputer.PasswordLastSet
					'Protect from Deletion' = $DefaultComputer.ProtectedFromAccidentalDeletion
				}
				$script:DefaultComputersinDefaultOUTable.Add($computersenabled)
			}
		}
	}
	New-WriteTime 300
}
# AD Users
Function Get-ADUsers {
	Write-ProgressHelper 15 "Get-ADUsers"
	New-LogWrite "[$loggingDate]  Function Get-ADUsers "
	$DefaultUsersOU = (Get-ADDomain).UsersContainer
	If ($DefaultUsersOU ) {
		# $DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, @{ Name = 'lastlogon'; Expression = { $_.lastlogon } }, DistinguishedName
		$DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, DistinguishedName, LastLogon # @{ Name = 'lastlogon'; Expression = { Set-FileTime $_.lastlogon } }, 
		#$DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, DistinguishedName, lastlogon
		If ($DefaultUsers) {
			foreach ($DefaultUser in $DefaultUsers) {
				# $DefaultUserlogon = $DefaultUser | Set-FileTime $_.lastlogon
				New-LogWrite "[$loggingDate]  $DefaultUser "
				$DefaultUserlogon = $DefaultUser.Lastlogon 
				$DefaultUserlogon = Set-FileTime $DefaultUserlogon
				New-LogWrite "[$loggingDate]  $DefaultUserlogon "
				$aduserobject = [PSCustomObject]@{
					'Name'                    = $DefaultUser.Name
					'UserPrincipalName'       = $DefaultUser.UserPrincipalName
					'Enabled'                 = $DefaultUser.Enabled
					'Protected from Deletion' = $DefaultUser.ProtectedFromAccidentalDeletion
					'Last Logon'              = $DefaultUserlogon
					'Email Address'           = $DefaultUser.EmailAddress
				}
				$script:DefaultUsersinDefaultOUTable.Add($aduserobject)
			}
		}
	}
	ElseIf (!$DefaultUsersOU) {
		New-LogWrite "[$loggingDate] Failed: `$DefaultUsersOU $DefaultUsersOU"
	}
	New-WriteTime 300
}
# Expired Accounts
Function Get-ExpiredAccounts {
	Write-ProgressHelper 16 "Get-ExpiredAccounts"
	New-LogWrite "[$loggingDate]  Function Get-ExpiredAccounts "
	# Expiring Accounts
	$ExpiringAccounts = Search-ADAccount -AccountExpiring -UsersOnly
	If ($ExpiringAccounts) {
		foreach ($ExpiringAccount in $ExpiringAccounts) {
			$NameExpiringAccounts = $ExpiringAccount.Name
			$UPNExpiringAccounts = $ExpiringAccount.UserPrincipalName
			$ExpirationDate = $ExpiringAccount.AccountExpirationDate
			$enabled = $ExpiringAccount.Enabled
			$expaccount = [PSCustomObject]@{
		
				'Name'              = $NameExpiringAccounts
				'UserPrincipalName' = $UPNExpiringAccounts
				'Expiration Date'   = $ExpirationDate
				'Enabled'           = $enabled
			}
			$script:ExpiringAccountsTable.Add($expaccount)
		}
	}
	New-WriteTime 300
}
# Security Logs
Function Get-Seclogs {
	Write-ProgressHelper 17 "Get-Seclogs"
	New-LogWrite "[$loggingDate]  Function Get-Seclogs "
	# Security Logs
	$SecurityLogs = Get-EventLog -Newest 7 -LogName 'Security' | Where-Object { $_.Message -like '*An account*' }
	If ($SecurityLogs) {
		foreach ($SecurityLog in $SecurityLogs) {
			$TimeGenerated = $SecurityLog.TimeGenerated
			$EntryType = $SecurityLog.EntryType
			$Recipient = $SecurityLog.Message
			$securitylogss = [PSCustomObject]@{
				'Time'    = $TimeGenerated
				'Type'    = $EntryType
				'Message' = $Recipient
			}
			$script:SecurityEventTable.Add($securitylogss)
		}
	}
	New-WriteTime 300
}
# Domains
Function Get-Domains {
	Write-ProgressHelper 18 "Get-Domains"
	New-LogWrite "[$loggingDate]  Function Get-Domains "
	# Tenant Domain
	$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes
	If ($Domains) {
		ForEach ($Domain  in $Domains) {
			$domainforest = [PSCustomObject]@{
				'UPN Suffixes' = $Domain
				'True'    = $true  
			}
			$script:DomainTable.Add($domainforest)
			New-LogWrite "[$loggingDate]  `DomainTable $DomainTable"
		}
	}
	New-WriteTime 300
}
# Buld Group tables
Function Get-Groups {
	Write-ProgressHelper 19 "Get-Groups"
	New-LogWrite "[$loggingDate]  Function Get-Groups "

	$SecurityCount = 0
	$MailSecurityCount = 0
	$CustomGroup = 0
	$DefaultGroup = 0
	$Groupswithmemebrship = 0
	$Groupswithnomembership = 0
	$GroupsProtected = 0
	$GroupsNotProtected = 0

	# Get groups and sort in alphabetical order
	$Groups = Get-ADGroup -Filter * -Properties *
	$DistroCount = $Groups | Where-Object { $_.GroupCategory -eq 'Distribution' }
	$DistroCounts = $DistroCount.DistinguishedName.Count
	foreach ($Group in $Groups) {
		$DefaultADGroup = 'False'
		$Gemail = (Get-ADGroup $Group -Properties mail).mail
		If ($Gemail) {
			if ($group.GroupCategory -eq 'Security') {
				$MailSecurityCount++
			}
		}
		If ($Gemail) {
			if ($group.GroupCategory -eq 'Security') {
				$SecurityCount++
			}
		}
		if ($Group.ProtectedFromAccidentalDeletion -eq $True) {
			$GroupsProtected++
		}
		else {
			$GroupsNotProtected++
		}
		if ($DefaultSGs -contains $Group.Name) {
			$DefaultADGroup = 'True'
			$DefaultGroup++
		}
		else {
			$CustomGroup++
		}
		if ($group.GroupCategory -eq 'Distribution') {
			$Type = 'Distribution Group'
		}
		If ($Gemail) {
			if ($group.GroupCategory -eq 'Security') {
				$Type = 'Security Group'
			}
		}
		If ($Gemail) {
			if ($group.GroupCategory -eq 'Security') {
				$Type = 'Mail-Enabled Security Group'
			}
		}
		if ($Group.Name -ne 'Domain Users') {
			$Users = (Get-ADGroupMember -Identity $Group | Sort-Object DisplayName | Select-Object -ExpandProperty Name) -join ', '
			if (!$Users) {
				$Groupswithnomembership++
			}
			else {
				$Groupswithmemebrship++
			}
		}
		else {
			$Users = 'Skipped Domain Users Membership'
		}
		$OwnerDN = Get-ADGroup -Filter { name -eq $Group.Name } -Properties managedBy | Select-Object -ExpandProperty ManagedBy
		Try {
			$Manager = Get-ADUser -Filter { distinguishedname -like $OwnerDN } | Select-Object -ExpandProperty Name
		}
		Catch {
			$groupname = $group.Name
			New-LogWrite "[$loggingDate]  Manager attribute:$Manager  on the group  $groupname  missing"
		}
		# $Manager = $AllUsers | Where-Object { $_.distinguishedname -eq $OwnerDN } | Select-Object -ExpandProperty Name
		$adgroupobject = [PSCustomObject]@{
		
			'Name'                    = $Group.name
			'Type'                    = $Type
			'Members'                 = $users
			'Managed By'              = $Manager
			'E-mail Address'          = $GEmail
			'Protected from Deletion' = $Group.ProtectedFromAccidentalDeletion
			'Default AD Group'        = $DefaultADGroup
		}
		$script:table.Add($adgroupobject)
	}
	# TOP groups table
	$objectmailgroups = [PSCustomObject]@{
		'Total Groups'                 = $Groups.Count
		'Mail-Enabled Security Groups' = $MailSecurityCount.Count
		'Security Groups'              = $SecurityCount.Count
		'Distribution Groups'          = $DistroCounts
	}
	$script:TOPGroupsTable.Add($objectmailgroups)
    
	# Default Group Type Pie Chart
	$objectmailgroupssec = [PSCustomObject]@{
		'Name'  = 'Mail-Enabled Security Groups'
		'Count' = $MailSecurityCount.Count
	}
	$script:GroupTypetable.Add($objectmailgroupssec)

	$secgroups = [PSCustomObject]@{
		'Name'  = 'Security Groups'
		'Count' = $SecurityCount.Count
	}
	$script:GroupTypetable.Add($secgroups)

	$distgroups = [PSCustomObject]@{
		'Name'  = 'Distribution Groups'
		'Count' = $DistroCounts
	}
	$script:GroupTypetable.Add($distgroups)

	# Default Group Pie Chart
	$defaultgroups = [PSCustomObject]@{
		'Name'  = 'Default Groups'
		'Count' = $DefaultGroup
	}
	$script:DefaultGrouptable.Add($defaultgroups)

	$customgroups = [PSCustomObject]@{
		'Name'  = 'Custom Groups'
		'Count' = $CustomGroup
	}
	$script:DefaultGrouptable.Add($customgroups)
	# Group Protection Pie Chart
	$protectedgroups = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $GroupsProtected
	}
	$script:GroupProtectionTable.Add($protectedgroups)

	$notprotectedgroups = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $GroupsNotProtected
	}
	$script:GroupProtectionTable.Add($notprotectedgroups)

	# Groups with membership vs no membership pie chart
	$groupwithmembers = [PSCustomObject]@{
		'Name'  = 'With Members'
		'Count' = $Groupswithmemebrship
	}
	$script:GroupMembershipTable.Add($groupwithmembers)

	$nogroupwithmembers = [PSCustomObject]@{
		'Name'  = 'No Members'
		'Count' = $Groupswithnomembership
	}
	$script:GroupMembershipTable.Add($nogroupwithmembers)
	New-WriteTime 300
}
# Build OU tables
Function Get-OU {
	Write-ProgressHelper 20 "Get-OU"
	New-LogWrite "[$loggingDate]  Function Get-OU  "

	$OUwithLinked = 0
	$OUwithnoLink = 0
	$OUProtected = 0
	$OUNotProtected = 0
	$OUs = Get-ADOrganizationalUnit -Filter * -Properties *
	foreach ($OU in $OUs) {
		if (($OU.linkedgrouppolicyobjects).length -lt 1) {
			$LinkedGPOs = 'None'
			$OUwithnoLink++
		}
		else {
			$OUwithLinked++
			$GPOslinks = $OU.linkedgrouppolicyobjects
			foreach ($GPOlink in $GPOslinks) {
				$Split1 = $GPOlink -split '{' | Select-Object -Last 1
				$Split2 = $Split1 -split '}' | Select-Object -First 1
				# $LinkedGPOs.Add((Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName)
				$LinkedGPOs += (Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName
				# $LinkedGPOs.Add($OUs.linkedgrouppolicyobjects.count)

			}
		}
		if ($OU.ProtectedFromAccidentalDeletion -eq $True) {
			$OUProtected++
		}
		else {
			$OUNotProtected++
		}
		$LinkedGPOs = $LinkedGPOs -join ', '
		$linkedgpoobjects = [PSCustomObject]@{
			'Name'                    = $OU.Name
			'Linked GPOs'             = $LinkedGPOs
			'Modified Date'           = $OU.WhenChanged
			'Protected from Deletion' = $OU.ProtectedFromAccidentalDeletion
		}
		$script:OUTable.Add($linkedgpoobjects)
	}
	# OUs with no GPO Linked
	$nolinkedous = [PSCustomObject]@{
		'Name'  = "OUs with no GPO's linked"
		'Count' = $OUwithnoLink
	}
	$script:OUGPOTable.Add($nolinkedous)
	$linkedous = [PSCustomObject]@{
		'Name'  = "OUs with GPO's linked"
		'Count' = $OUwithLinked
	}
	$script:OUGPOTable.Add($linkedous)
	# OUs Protected Pie Chart
	$protectedous = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $OUProtected
	}
	$script:OUProtectionTable.Add($protectedous)
	$notprotectedous = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $OUNotProtected
	}
	$script:OUProtectionTable.Add($notprotectedous)
	New-WriteTime 300

}
# Build User Tables - Protected and pwd expiring
Function Get-Users {
	Write-ProgressHelper 21 "Get-Users"
	New-LogWrite "[$loggingDate]  Function Get-Users "

	$UserEnabled = 0
	$UserDisabled = 0
	$UserPasswordExpires = 0
	$UserPasswordNeverExpires = 0
	$ProtectedUsers = 0
	$NonProtectedUsers = 0
	## $UsersWIthPasswordsExpiringInUnderAWeek = 0
	## $UsersNotLoggedInOver30Days = 0
	## $AccountsExpiringSoon = 0
	## Get users that haven't logged on in X amount of days, var is set at start of script
	foreach ($User in $AllUsers) {
		# $AttVar = $User | Select-Object Enabled, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, Name, SamAccountName, EmailAddress, AccountExpirationDate, @{ Name = 'lastlogon'; Expression = { $_.lastlogon } }, DistinguishedName
		$AttVar = $User | Select-Object Enabled, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, Name, SamAccountName, EmailAddress, AccountExpirationDate, DistinguishedName, LastLogon
		$defaultuserlastlogon = $User.lastlogon
		$defaultuserlastlogon = Set-FileTime $defaultuserlastlogon 
		New-LogWrite "[$loggingDate] `$defaultuserlastlogon $defaultuserlastlogon "
		$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
		if ((($AttVar.PasswordNeverExpires) -eq $False) -and (($AttVar.Enabled) -ne $false)) {
			# Get Password last set date
			$passwordSetDate = ($User | ForEach-Object { $_.PasswordLastSet })
			if (!$passwordSetDate) {
				$daystoexpire = 'User has never logged on'
			}
			else {
				# Check for Fine Grained Passwords
				$PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
				if ($PasswordPol) {
					$maxPasswordAge = ($PasswordPol).MaxPasswordAge
				}
				$expireson = $passwordsetdate.AddDays($maxPasswordAge)
				$today = (Get-Date)
				# Gets the count on how many days until the password expires and stores it in the $daystoexpire var
				$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
			}
		}
		else {
			$daystoexpire = 'N/A'
		}
		New-LogWrite "((`"$defaultuserlastlogon`" -like `"Never`") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))"
		If ($defaultuserlastlogon) {
			New-LogWrite "((`"$defaultuserlastlogon`" -like `"Never`") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))"
			If (("$defaultuserlastlogon" -like "Never") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))) {
				New-LogWrite "[$loggingDate]  If ($User -and $AttVar -and $Days )"
				New-LogWrite "[$loggingDate] psobject lastlogon =  $defaultuserlastlogon"
				if (($User.Enabled -eq $True)) {
					$objlastlogonobject = [PSCustomObject]@{
						'Name'                        = $User.Name
						'UserPrincipalName'           = $User.UserPrincipalName
						'Enabled'                     = $AttVar.Enabled
						'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
						'Last Logon'                  = $defaultuserlastlogon
						'Password Never Expires'      = $AttVar.PasswordNeverExpires
						'Days Until Password Expires' = $daystoexpire
					}
					$script:userphaventloggedonrecentlytable.Add($objlastlogonobject)
					New-LogWrite "`$userphaventloggedonrecentlytable $userphaventloggedonrecentlytable"
				}
			}
		}
		# Items for protected vs non protected users
		if ($User.ProtectedFromAccidentalDeletion -eq $False) {
			$NonProtectedUsers++
		}
		else {
			$ProtectedUsers++
		}
		# Items for the enabled vs disabled users pie chart
		if (($AttVar.PasswordNeverExpires) -ne $false) {
			$UserPasswordNeverExpires++
		}
		else {
			$UserPasswordExpires++
		}
		# Items for password expiration pie chart
		if (($AttVar.Enabled) -ne $false) {
			$UserEnabled++
		}
		else {
			$UserDisabled++
		}
		$Name = $User.Name
		$UPN = $User.UserPrincipalName
		$Enabled = $AttVar.Enabled
		$EmailAddress = $AttVar.EmailAddress
		$AccountExpiration = $AttVar.AccountExpirationDate
		$PasswordExpired = $AttVar.PasswordExpired
		$PasswordLastSet = $AttVar.PasswordLastSet
		$PasswordNeverExpires = $AttVar.PasswordNeverExpires
		# $daysUntilPWExpire = $daystoexpire
		$adobjectouser = [PSCustomObject]@{
			'Name'                        = $Name
			'UserPrincipalName'           = $UPN
			'Enabled'                     = $Enabled
			'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
			'Last Logon'                  = $defaultuserlastlogon
			# 'Last Logon'                  = $LastLogon
			'Email Address'               = $EmailAddress
			'Account Expiration'          = $AccountExpiration
			'Change Password Next Logon'  = $PasswordExpired
			'Password Last Set'           = $PasswordLastSet
			'Password Never Expires'      = $PasswordNeverExpires
			'Days Until Password Expires' = $daystoexpire
		}
		$script:usertable.Add($adobjectouser)
		if ($daystoexpire -lt $DaysUntilPWExpireINT) {
			$noadobjectouser = [PSCustomObject]@{
				'Name'                        = $Name
				'Days Until Password Expires' = $daystoexpire
			}
			$script:PasswordExpireSoonTable.Add($noadobjectouser)

		}
	}

	# Data for users enabled vs disabled pie graph
	$userenabled = [PSCustomObject]@{
		'Name'  = 'Enabled'
		'Count' = $UserEnabled
	}
	$script:EnabledDisabledUsersTable.Add($userenabled)

	$userdisabled = [PSCustomObject]@{
		'Name'  = 'Disabled'
		'Count' = $UserDisabled
	}
	$script:EnabledDisabledUsersTable.Add($userdisabled)

	# Data for users password expires pie graph
	$userpasswordexp = [PSCustomObject]@{
		'Name'  = 'Password Expires'
		'Count' = $UserPasswordExpires
	}
	$script:PasswordExpirationTable.Add($userpasswordexp)

	$userpasswordneverexp = [PSCustomObject]@{
		'Name'  = 'Password Never Expires'
		'Count' = $UserPasswordNeverExpires
	}
	$script:PasswordExpirationTable.Add($userpasswordneverexp)

	#Data for protected users pie graph
	$protecteduser = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $ProtectedUsers
	}
	$script:ProtectedUsersTable.Add($protecteduser)

	$notprotecteduser = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $NonProtectedUsers
	}
	$script:ProtectedUsersTable.Add($notprotecteduser)

	if ($userphaventloggedonrecentlytable.count -eq 0) {
		# if (($userphaventloggedonrecentlytable).Information) {
		$script:UHLONXD = '0'
	}
	Else {
		$script:UHLONXD = $userphaventloggedonrecentlytable.Count
	}
	# TOP User table
	If ($ExpiringAccounts) {
		# If (($ExpiringAccounts).Information) {
		$expaccounts = [PSCustomObject]@{
			'Total Users'                                                           = $AllUsers.Count
			"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
			'Expiring Accounts'                                                     = $ExpiringAccounts.Count
			"Users Haven't Logged on in $Days Days or more"                         = $UHLONXD
		}
		$script:TOPUserTable.Add($expaccounts)
	}
	Else {
		$expaccounts = [PSCustomObject]@{
			'Total Users'                                                           = $AllUsers.Count
			"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
			'Expiring Accounts'                                                     = '0'
			"Users Haven't Logged on in $Days Days or more"                         = $UHLONXD
		}
		$script:TOPUserTable.Add($expaccounts)
	}
	New-WriteTime 300
}
# GPO Tables
Function Get-GPOs {
	Write-ProgressHelper 22 "Get-GPOs"
	New-LogWrite "[$loggingDate]  Get-GPOs "
	foreach ($GPO in $GPOs) {
		$gpoobject = [PSCustomObject]@{
			'Name'             = $GPO.DisplayName
			'Status'           = $GPO.GpoStatus
			'Modified Date'    = $GPO.ModificationTime
			'User Version'     = $GPO.UserVersion
			'Computer Version' = $GPO.ComputerVersion
		}
		$script:GPOTable.Add($gpoobject)
	}
	New-WriteTime 300
}
# Recent GPOs Table
Function Get-RecentGPOs {
	Write-ProgressHelper 22 "Get-RecentGPOs"
	New-LogWrite "[$loggingDate]  Function Get-RecentGPOs"
	$createdinthelast = ((Get-Date).AddDays( - 30)).Date
	$grouppolicyrecent = Get-GPO -all | Select-Object DisplayName, GPOStatus, CreationTime , @{ Label = 'ComputerVersion'; Expression = { $_.computer.dsversion } }, @{ Label = 'UserVersion'; Expression = { $_.user.dsversion } } 
	foreach ($grouppolicyrecent in $grouppolicyrecent) {
		$grouppolicyrecenttest5 = $grouppolicyrecent.CreationTime
		If (($grouppolicyrecenttest5 -gt $createdinthelast)) {
			$gpocreatedobject = [PSCustomObject]@{
				'Name'             = $grouppolicyrecent.DisplayName
				'Status'           = $grouppolicyrecent.GpoStatus
				'Created Date'     = $grouppolicyrecenttest5
				'User Version'     = $grouppolicyrecent.UserVersion
				'Computer Version' = $grouppolicyrecent.ComputerVersion
			}
			$script:recentgpostable.Add($gpocreatedobject)
		}
	}
	New-WriteTime 300
}
# AD Computers tables - Protected
Function Get-Adcomputers {
	Write-ProgressHelper 23 "Get-Adcomputers"
	New-LogWrite "[$loggingDate]  Function Get-Adcomputers "
	$Computers = Get-ADComputer -Filter * -Properties *

	$ComputersProtected = 0
	$ComputersNotProtected = 0
	$ComputerEnabled = 0
	$ComputerDisabled = 0
	# Only search for versions of windows that exist in the Environment
	$WindowsRegex = '(Windows (Server )?(\d+|XP)?( R2)?).*'
	$OsVersions = $Computers | Select-Object OperatingSystem -unique | ForEach-Object {
		if ($_.OperatingSystem -match $WindowsRegex ) { 
			return $matches[1]
		}
		elseif ($_.OperatingSystem) {
			return $_.OperatingSystem
		}
	} | Select-Object -unique | Sort-Object
	$OsObj = [PSCustomObject]@{ }
	$OsVersions | ForEach-Object {
		$OsObj | Add-Member -Name $_ -Value 0 -Type NoteProperty
	}
	foreach ($Computer in $Computers) {
		if ($Computer.ProtectedFromAccidentalDeletion -eq $True) {
			$ComputersProtected++
		}
		else {
			$ComputersNotProtected++
		}
		if ($Computer.Enabled -eq $True) {
			$ComputerEnabled++
		}
		else {
			$ComputerDisabled++
		}
		$computerlastlogin = $Computer.lastLogonTimestamp
		$computerlastlogin = Set-FileTime $computerlastlogin
		$computerobjects = [PSCustomObject]@{
			'Name'                  = $Computer.Name
			'Enabled'               = $Computer.Enabled
			'Operating System'      = $Computer.OperatingSystem
			'Modified Date'         = $Computer.Modified
			'Last Login'            = $computerlastlogin
			'Password Last Set'     = $Computer.PasswordLastSet
			'Protect from Deletion' = $Computer.ProtectedFromAccidentalDeletion
		}
		$script:ComputersTable.Add($computerobjects)
		if ($Computer.OperatingSystem -match $WindowsRegex) {
			$OsObj."$($matches[1])"++
		}
	}
	#Pie chart breaking down OS for computer obj
	$OsObj.PSObject.Properties | ForEach-Object {
		$script:GraphComputerOS.Add([PSCustomObject]@{'Name' = $_.Name; 'Count' = $_.Value })
	}
	#Data for TOP Computers data table
	$OsObj | Add-Member -Name 'Total Computers' -Value $Computers.Count -Type NoteProperty
	$script:TOPComputersTable.Add($OsObj)
	#Data for protected Computers pie graph
	$protectedcomputers = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $ComputersProtected
	}
	$script:ComputerProtectedTable.Add($protectedcomputers)

	$notprotectedcomputers = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $ComputersNotProtected
	}
	$script:ComputerProtectedTable.Add($notprotectedcomputers)

	#Data for enabled/vs Computers pie graph
	$computersenabled = [PSCustomObject]@{
		'Name'  = 'Enabled'
		'Count' = $ComputerEnabled
	}
	$script:ComputersEnabledTable.Add($computersenabled)

	$computersdisabled = [PSCustomObject]@{
		'Name'  = 'Disabled'
		'Count' = $ComputerDisabled
	}
	$script:ComputersEnabledTable.Add($computersdisabled)
	New-WriteTime 300
}
# Checking for Empty Tables
Function compare-tables {
	Write-ProgressHelper 24 "compare-tables"
	New-LogWrite "[$loggingDate]  Function compare-tables"
	if (!$PasswordExpireSoonTable) {
		$pwdexp = [PSCustomObject]@{
			Information = ' No users were found to have passwords expiring soon'
		}
		$script:PasswordExpireSoonTable.Add($pwdexp)
	}
	if (!$userphaventloggedonrecentlytable) {
		$noobjlastlogonobject = [PSCustomObject]@{
			Information = " No Users were found to have not logged on in $Days days or more"
		}
		$script:userphaventloggedonrecentlytable.Add($noobjlastlogonobject)
	}
	if (!$ADObjectTable ) {
		$noadobjecttype = [PSCustomObject]@{
			Information = ' No AD Objects have been modified recently'
		}
		$script:ADObjectTable.Add($noadobjecttype)
	}
	if (!$CompanyInfoTable) {
		$noaddomains = [PSCustomObject]@{
			Information = ' Could not get items for table'
		}
		$script:CompanyInfoTable.Add($noaddomains)
	}
	if (!$NewCreatedUsersTable ) {
		$nocreateduserss = [PSCustomObject]@{
			Information = ' No new users have been recently created'
		}
		$script:NewCreatedUsersTable.Add($nocreateduserss)
	}
	if (!$DomainAdminTable) {
		$noadgroupmemebers = [PSCustomObject]@{
			Information = ' No Domain Admin Members were found'
		}
		$script:DomainAdminTable.Add($noadgroupmemebers)
	}
	if (!$EnterpriseAdminTable) {
		$noentadmin = [PSCustomObject]@{
			Information = ' Enterprise Admin members were found'
		}
		$script:EnterpriseAdminTable.Add($noentadmin)
	}
	if (!$DefaultComputersinDefaultOUTable ) {
		$computersdisabled = [PSCustomObject]@{
			Information = ' No computers were found in the Default OU'
		}
		$script:DefaultComputersinDefaultOUTable.Add($computersdisabled)
	}
	if (!$DefaultUsersinDefaultOUTable ) {
		$noaduserobject = [PSCustomObject]@{
			Information = ' No Users were found in the default OU'
		}
		$script:DefaultUsersinDefaultOUTable.Add($noaduserobject)
	}
	if (!$SecurityEventTable) {
		$nosecuritylogss = [PSCustomObject]@{
			Information = ' No logon security events were found'
		}
		$script:SecurityEventTable.Add($nosecuritylogss)
	}
	if (!$DomainTable ) {
		$nodomainforest = [PSCustomObject]@{
			Information = ' No UPN Suffixes were found'
		}
		$script:DomainTable.Add($nodomainforest)
	}
	if (!$table ) {
		$noadgroupobject = [PSCustomObject]@{
			Information = ' No Groups were found'
		}
		$script:table.Add($noadgroupobject)
	}
	if (!$usertable ) {
		$adobjectousers = [PSCustomObject]@{
			Information = ' No users were found'
		}
		$script:usertable.Add($adobjectousers)
	}
	if (!$OUTable ) {
		$nolinkedgpoobjects = [PSCustomObject]@{
			Information = ' No OUs were found'
		}
		$script:OUTable.Add($nolinkedgpoobjects)
	}
	if (!$GPOTable ) {
		$gpoobject = [PSCustomObject]@{
			Information = ' No Group Policy Obejects were found'
		}
		$script:GPOTable.Add($gpoobject)
	}
	if (!$ComputersTable ) {
		$compuerobjects = [PSCustomObject]@{
			Information = ' No computers were found'
		}
		$script:ComputersTable.Add($compuerobjects)
	}
	If (!$ExpiringAccountsTable ) {
		$noexpaccount = [PSCustomObject]@{
			Information = ' No Users were found to expire soon' 
		}
		$script:ExpiringAccountsTable.Add($noexpaccount)
	}
	If (!$recentgpostable ) {
		$norecentgpostable = [PSCustomObject]@{
			Information = ' No GPOs were created recently' 
		}
		$script:recentgpostable.Add($norecentgpostable)
	}
	New-WriteTime 300
}