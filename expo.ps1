
$script:command = $($args[0])
# $script:switch = $($args[1])
$script:currentdir = Get-Location
Get-ChildItem -Path $currentdir\Fn -Filter *.ps1 | ForEach-Object { . $_.FullName }
Get-ChildItem -Path $currentdir\conf -Filter *.ps1 | ForEach-Object { . $_.FullName }

Get-Config

Test-htmlModule
get-reportSettings
New-reportFolder
Set-Header
Get-DefaultSecurityGroups
Get-CreateObjects
Get-AllUsers
Get-AllGPOs
Get-ADOBJ
Set-ADOBJ
Get-ADRecBin
Get-ADInfo
Get-CreatedUsers
Get-DomainAdmins
Set-DomainAdminsTable
Get-EntAdmins
Get-Computers
Get-ADUsers
Get-ExpiredAccounts
Get-Seclogs
Get-Domains
Get-Groups
Get-OU
Get-Users
Get-GPOs
Get-RecentGPOs
Get-Adcomputers
compare-tables 
Get-Report