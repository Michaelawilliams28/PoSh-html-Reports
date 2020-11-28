# Build Report
Function Get-Report {
	Write-ProgressHelper 25 "Get-Report"
	New-LogWrite "[$loggingDate]  Function Get-Report"

	$tabarray = @('Dashboard', 'Groups', 'Organizational Units', 'Users', 'Group Policy', 'Computers')


	#--------------------------------OU Protection PIE CHART--##
	#  Basic Properties 
	$PO12 = Get-HTMLPieChartObject
	$PO12.Title = 'Organizational Units Protected from Deletion'
	$PO12.Size.Height = 250
	$PO12.Size.width = 250
	$PO12.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PO12.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PO12.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PO12.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PO12.DataDefinition.DataNameColumnName = 'Name'
	$PO12.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Computer OS Breakdown PIE CHART--##
	$PieObjectComputerObjOS = Get-HTMLPieChartObject
	$PieObjectComputerObjOS.Title = 'Computer Operating Systems'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'Random'

	#--------------------------------Computers Protection PIE CHART--##
	#  Basic Properties 
	$PieObjectComputersProtected = Get-HTMLPieChartObject
	$PieObjectComputersProtected.Title = 'Computers Protected from Deletion'
	$PieObjectComputersProtected.Size.Height = 250
	$PieObjectComputersProtected.Size.width = 250
	$PieObjectComputersProtected.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectComputersProtected.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectComputersProtected.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Computers Enabled PIE CHART--##
	#  Basic Properties 
	$PieObjectComputersEnabled = Get-HTMLPieChartObject
	$PieObjectComputersEnabled.Title = 'Computers Enabled vs Disabled'
	$PieObjectComputersEnabled.Size.Height = 250
	$PieObjectComputersEnabled.Size.width = 250
	$PieObjectComputersEnabled.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectComputersEnabled.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectComputersEnabled.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------USERS Protection PIE CHART--##
	#  Basic Properties 
	$PieObjectProtectedUsers = Get-HTMLPieChartObject
	$PieObjectProtectedUsers.Title = 'Users Protected from Deletion'
	$PieObjectProtectedUsers.Size.Height = 250
	$PieObjectProtectedUsers.Size.width = 250
	$PieObjectProtectedUsers.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectProtectedUsers.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectProtectedUsers.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Basic Properties 
	$PieObjectOUGPOLinks = Get-HTMLPieChartObject
	$PieObjectOUGPOLinks.Title = 'OU GPO Links'
	$PieObjectOUGPOLinks.Size.Height = 250
	$PieObjectOUGPOLinks.Size.width = 250
	$PieObjectOUGPOLinks.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectOUGPOLinks.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectOUGPOLinks.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Basic Properties 
	$PieObject4 = Get-HTMLPieChartObject
	$PieObject4.Title = 'Office 365 Unassigned Licenses'
	$PieObject4.Size.Height = 250
	$PieObject4.Size.width = 250
	$PieObject4.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObject4.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObject4.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObject4.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObject4.DataDefinition.DataNameColumnName = 'Name'
	$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

	#--------------------------------Basic Properties 
	$PieObjectGroupType = Get-HTMLPieChartObject
	$PieObjectGroupType.Title = 'Group Types'
	$PieObjectGroupType.Size.Height = 250
	$PieObjectGroupType.Size.width = 250
	$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'

	#--------------------------------Pie Chart Groups with members vs no members
	$PieObjectGroupMembersType = Get-HTMLPieChartObject
	$PieObjectGroupMembersType.Title = 'Group Membership'
	$PieObjectGroupMembersType.Size.Height = 250
	$PieObjectGroupMembersType.Size.width = 250
	$PieObjectGroupMembersType.ChartStyle.ChartType = 'doughnut'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'ColorScheme1'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'Generated1'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'Random'
	$PieObjectGroupMembersType.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupMembersType.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Basic Properties 
	$PieObjectGroupType2 = Get-HTMLPieChartObject
	$PieObjectGroupType2.Title = 'Custom vs Default Groups'
	$PieObjectGroupType2.Size.Height = 250
	$PieObjectGroupType2.Size.width = 250
	$PieObjectGroupType2.ChartStyle.ChartType = 'doughnut'

	#--------------------------------These file exist in the module directoy, There are 4 schemes by default
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'ColorScheme1'

	# There are 8 generated schemes, randomly generated at runtime 
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Generated1'

	# you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectGroupType.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Enabled users vs Disabled Users PIE CHART--##
	#  Basic Properties 
	$EnabledDisabledUsersPieObject = Get-HTMLPieChartObject
	$EnabledDisabledUsersPieObject.Title = 'Enabled vs Disabled Users'
	$EnabledDisabledUsersPieObject.Size.Height = 250
	$EnabledDisabledUsersPieObject.Size.width = 250
	$EnabledDisabledUsersPieObject.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$EnabledDisabledUsersPieObject.DataDefinition.DataNameColumnName = 'Name'
	$EnabledDisabledUsersPieObject.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------PasswordNeverExpires PIE CHART--##
	#  Basic Properties 
	$PWExpiresUsersTable = Get-HTMLPieChartObject
	$PWExpiresUsersTable.Title = 'Password Expiration'
	$PWExpiresUsersTable.Size.Height = 250
	$PWExpiresUsersTable.Size.Width = 250
	$PWExpiresUsersTable.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PWExpiresUsersTable.DataDefinition.DataNameColumnName = 'Name'
	$PWExpiresUsersTable.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Group Protection PIE CHART--##
	#   Basic Properties 
	$PieObjectGroupProtection = Get-HTMLPieChartObject
	$PieObjectGroupProtection.Title = 'Groups Protected from Deletion'
	$PieObjectGroupProtection.Size.Height = 250
	$PieObjectGroupProtection.Size.width = 250
	$PieObjectGroupProtection.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectGroupProtection.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupProtection.DataDefinition.DataValueColumnName = 'Count'

	#--------------------------------Dashboard Report
	$FinalReport = New-Object 'System.Collections.Generic.List[System.Object]'
	$FinalReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo))
	$FinalReport.Add($(Get-HTMLTabHeader -TabNames $tabarray))
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[0] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Company Information'))
	$FinalReport.Add($(Get-HTMLContentTable $CompanyInfoTable))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Groups'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DomainAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Objects in Default OUs'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Computers'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DefaultComputersinDefaultOUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DefaultUsersinDefaultOUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "AD Objects Modified in Last $ADModNumber Days "))
	$FinalReport.Add($(Get-HTMLContentDataTable $ADObjectTable))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Expiring Items'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText " Users with Passwords Expiring in less than $DaysUntilPWExpireINT days "))
	$FinalReport.Add($(Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days or more"))
	$FinalReport.Add($(Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Security Logs'))
	$FinalReport.Add($(Get-HTMLContentDataTable $securityeventtable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'UPN Suffixes'))
	$FinalReport.Add($(Get-HTMLContentTable $DomainTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))
    #--------------------------------
    #--------------------------------Groups Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[1] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Groups Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPGroupsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Groups'))
	$FinalReport.Add($(Get-HTMLContentDataTable $Table -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumn1of2))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DomainAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Groups Chart'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupType2 -DataSet $DefaultGrouptable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupMembersType -DataSet $GroupMembershipTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 4 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupProtection -DataSet $GroupProtectionTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#-------------------------------- Organizational Unit Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[2] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Organizational Units'))
	$FinalReport.Add($(Get-HTMLContentDataTable $OUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Organizational Units Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectOUGPOLinks -DataSet $OUGPOTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PO12 -DataSet $OUProtectionTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentclose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#--------------------------------Users Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[3] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPUserTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Users'))
	$FinalReport.Add($(Get-HTMLContentDataTable $UserTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Expiring Items'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days"))
	$FinalReport.Add($(Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days or more"))
	$FinalReport.Add($(Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $EnabledDisabledUsersPieObject -DataSet $EnabledDisabledUsersTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PWExpiresUsersTable -DataSet $PasswordExpirationTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectProtectedUsers -DataSet $ProtectedUsersTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

    #--------------------------------GPO Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[4] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Group Policies'))
	$FinalReport.Add($(Get-HTMLContentDataTable $GPOTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	#$FinalReport.Add($(Get-HTMLTabContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Recently Created GPOs'))
	$FinalReport.Add($(Get-HTMLContentDataTable $recentgpostable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

    #--------------------------------Computers Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[5] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPComputersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ComputersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputersProtected -DataSet $ComputerProtectedTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputersEnabled -DataSet $ComputersEnabledTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentclose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Operating System Breakdown'))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputerObjOS -DataSet $GraphComputerOS))
	$FinalReport.Add($(Get-HTMLContentclose))
    #--------------------------------
	$FinalReport.Add($(Get-HTMLTabContentClose))
	$FinalReport.Add($(Get-HTMLClosePage))
    #--------------------------------
	$Day = (Get-Date).Day
	$Month = (Get-Date).Month
	$Year = (Get-Date).Year
	$ReportName = ("$Day - $Month - $Year - AD Report")
	Write-ProgressHelper 25 "Save-HTMLReport"
	Save-HTMLReport -ReportContent $FinalReport -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath
}