#===============================================
# 	Date: 			26/03/2014
# 	Author: 		Stephen Furniss
#	Blog: 			myblog.furnissathome.co.uk
# 	Twitter: 		twitter.com/furnissathome
#
#	Description:	This script automates user account creation
#					from .CSV files created by SIMS. Optionally
#					mailboxes for Exchange and Office 365 can 
#					also be created.
#
#	Requirements:	PowerShell V3.0
#					.Net 4.5.1
#					Microsoft Online Services Sign-in Assistant
#					Windows Azure Active Directory Module for Windows PowerShell
#
#==============================================
$Version = "1.0"
#----------------------------------------------
#region Change Log
#----------------------------------------------
#	26/03/2014	New script created.
#	01/05/2014	Updated after testing.
#	07/05/2014	Added Office 365 integration.
#----------------------------------------------
#endregion Change Log
#----------------------------------------------

#----------------------------------------------
#region Application Functions
#----------------------------------------------
#	Clear console output screen.
cls

#	Check & Load Active Directory PowerShell Module
if ((Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue | foreach { $_.Name }) -ne "ActiveDirectory") 
{
#    # write-host "Loading ActiveDirectory PowerShell Module..."
    Import-Module ActiveDirectory
}

#	Get folder path for script location... (All required files should be in this folder)
$strScriptPath = Split-Path (Resolve-Path $myInvocation.MyCommand.Path)
if ($strScriptPath -match 'Quest Software') {$strScriptPath = [System.AppDomain]::CurrentDomain.BaseDirectory}

#	Get todays date.
$LogDate = (Get-Date -Format yyyy-MM-dd)
$CurrentDate = (Get-Date -Format dd/MM/yyyy)
$CurrentDate = Get-Date $CurrentDate


#	Get the site that this computer is  a member of.
$siteName = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
##  Get all Servers in a specified Active Directory site.
$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
$siteContainerDN = ("CN=Sites," + $configNCDN)
$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
$DCs = (Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel -filter { objectClass -eq "Server" } -Properties "DNSHostName", "Description" | Select DNSHostName)

function TestPupilFileExists {
#	Test if Pupil file exists and create one if it does not.
if(!(Test-Path -Path $PupilOutput))
	{
		#	Create new log file.
		New-Item $PupilOutput -type file
		"UserID,FirstName,LastName,Year,InitialPassword" | out-file -filepath $PupilOutput -append
	}	
}

function TestStaffFileExists	{
		#	Test if Staff file exists and create one if it does not.
		if(!(Test-Path -Path $StaffOutput))
			{
				#	Create new log file.
				New-Item $StaffOutput -type file
				"UserID,FirstName,LastName,InitialPassword" | out-file -filepath $StaffOutput -append
			}
}
#----------------------------------------------
#endregion Application Functions
#----------------------------------------------

#----------------------------------------------
#region Create Log file
#----------------------------------------------
#	Build name for log file based on todays date.
$LogFile = ($strScriptPath + "\Logs\" + $LogDate + "-Daily.log")

#	Test if log file exists and create one if it does not.
if(!(Test-Path -Path $LogFile))
	{
		#	Create new log file.
		New-Item $LogFile -type file
	}
	
"-------------------------------------" | out-file -filepath $LogFile -append
(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
"ADUA - User Automation task started." | out-file -filepath $LogFile -append
"-------------------------------------" | out-file -filepath $LogFile -append
#----------------------------------------------
#endregion Create Log file
#----------------------------------------------

#----------------------------------------------
#region Read Configuration file
#----------------------------------------------
$XMLOptions = "ADUA.Options.xml"
$XMLFile = Join-Path $strScriptPath $XMLOptions

if(!(Test-Path $XMLFile))
	{
#		# write-host $XMLFile " does not exist"
	}
else
	{
		#	Read XML file - this is the main configuration file.
		[XML]$Script:XML = Get-Content $XMLFile
 		$DNS_Name = $XML.Options.Global.Domain_DNS_Name
		$Domain_LDAP_Path = $XML.Options.Global.Domain_LDAP_Path
		$RetentionPeriod = $XML.Options.Global.AccountRetentionPeriod
		$CreateHomeFolders = $XML.Options.Global.CreateHomeFolders
		$RecreateHomeFolders = $XML.Options.Global.RecreateHomeFolders
		$Group_Prefix = $XML.Options.Global.Group_Prefix
		$DisabledUsersOU = $XML.Options.Global.DisabledUsersOU
		$CSVExtractLocation = $XML.Options.Global.CSVExtractLocation
		
		$365MailDomainName = $XML.Options.Mail.Office365.MailDomainName
		$365StaffLicenseAssignment = $XML.Options.Mail.Office365.StaffLicenseAssignment
		$365StudentLicenseAssignment = $XML.Options.Mail.Office365.StudentLicenseAssignment

		$Default_Staff_Group = $XML.Options.Staff.Default_Staff_Group
		$Staff_OU_LDAP = $XML.Options.Staff.Staff_OU_LDAP		
		$Staff_Naming_Convention = $XML.Options.Staff.Staff_Naming_Convention
		$StaffHomeServer = $XML.Options.Staff.StaffHomeServer
		$StaffHomeFolderRoot = $XML.Options.Staff.StaffHomeFolderRoot
		$InitialStaffPwd = ConvertTo-SecureString –AsPlainText –Force –String $XML.Options.Staff.InitialStaffPwd
		$Office365StaffPwd = $XML.Options.Staff.InitialStaffPwd
		$StaffPwdNeverExpires = $XML.Options.Staff.StaffPwdNeverExpires
		$StaffOutput = $XML.Options.Staff.StaffOutPath + "\" + $LogDate + "-Staff.csv"
		$StaffMailSolution = $XML.Options.Staff.StaffMailSolution
		
		$Default_Student_Group = $XML.Options.Students.Default_Student_Group
		$Student_OU_LDAP = $XML.Options.Students.Student_OU_LDAP
		$Student_Naming_Convention = $XML.Options.Students.Student_Naming_Convention
		$StudentHomeServer = $XML.Options.Students.StudentHomeServer
		$StudentHomeFolderRoot = $XML.Options.Students.StudentHomeFolderRoot
		$InitialStudentPwd = ConvertTo-SecureString –AsPlainText –Force –String $XML.Options.Students.InitialStudentPwd
		$Office365StudentPwd = $XML.Options.Students.InitialStudentPwd
		$StudentPwdNeverExpires = $XML.Options.Students.StudentPwdNeverExpires
		$PupilOutput = $XML.Options.Students.StudentOutPath + "\" + $LogDate + "-Pupil.csv"
		$StudentMailSolution = $XML.Options.Students.StudentMailSolution
		
	}
#----------------------------------------------
#endregion Read Configuration file
#----------------------------------------------

#----------------------------------------------
#region Connect to Office 365
#----------------------------------------------
#	Check if we need to create an Office 365
#	mailbox for the user accounts and connect
#	to Office 365 online.
#	-------------------------------------------
if($StaffMailSolution -eq "Office365")
	{
		#	Check & Load Windows Azure Active Directory Module for Windows PowerShell
		if ((Get-Module -name MSOnline -ErrorAction SilentlyContinue | foreach { $_.Name }) -ne "MSOnline") 
		{
		    # write-host "Loading Windows Azure Active Directory Module for Windows PowerShell..."
		    Import-Module MSOnline
		}
		
		#	Check for credentials file. If one
		#	exists connect to Office 365. If 
		#	not create a new file and request
		#	credentials.
		if(!(Test-Path -Path ($strScriptPath + "\Office365\Credential.xml")))
			{
				#	The credentials file does not exist.... creating new one.
				Get-Credential | Export-Clixml ($strScriptPath + "\Office365\Credential.xml")
			}
		else
			{
				#	Credentials file exists so we can try and make a connection.
				$Cred = Import-Clixml ($strScriptPath + "\Office365\Credential.xml")
				
				Try
					{
						Connect-MsolService -Credential $Cred -ErrorAction Stop
					}
				Catch 
					{
						$ErrMsg = "Error. Cannot Connect-MsolService. $_"					
						"***************************************************************************" | out-file -filepath $LogFile -append
						(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
						$ErrMsg | out-file -filepath $LogFile -append 
						"***************************************************************************" | out-file -filepath $LogFile -append
						Write-Error -Message "Error. Cannot Connect-MsolService because $_" -ErrorAction Stop 
					}
			}		
	}
else
	{
		#	Do not make a connection to Office 365.
	}
#----------------------------------------------
#endregion Connect to Office 365
#----------------------------------------------

#----------------------------------------------
#region Read Staff Current CSV file
#----------------------------------------------
#	Test if CSV file exists.
if(!(Test-Path -Path ($CSVExtractLocation + "\AD Staff Current.csv")))
	{
		#	CSV file does not exist...
		"***************************************************************************" | out-file -filepath $LogFile -append
		(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
		"Staff CSV file does not exist..." | out-file -filepath $LogFile -append 
		"***************************************************************************" | out-file -filepath $LogFile -append
	}
else
	{
		#	Read 'Staff_Extract.csv' - this file contains all current staff users for the school.
		"***************************************************************************" | out-file -filepath $LogFile -append
		(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
		"Reading Staff CSV file..." | out-file -filepath $LogFile -append 
		"***************************************************************************" | out-file -filepath $LogFile -append
		$StaffUsers = Import-Csv ($CSVExtractLocation + "\AD Staff Current.csv")
		## write-host "Creating Staff"
		foreach($UD in $StaffUsers) 
			{
				$Exit = 0
				$Count = 1
				$External_ID = $UD.External_ID	#bullTCLextension1
				if ($External_ID -eq $Null)
					{
						#	This is blank so skip user creation.
					}
				elseif($External_ID -eq "")
					{
						#	"empty string"
					}
				else
					{
						$FirstName = $UD.Forename		#givenName
						$LastName = $UD.Surname		#sn
				
						# Strip out invalid characters.
						$FirstName = [System.Text.RegularExpressions.Regex]::Replace($FirstName,"[^1-9a-zA-Z_]","");
						$LastName = [System.Text.RegularExpressions.Regex]::Replace($LastName,"[^1-9a-zA-Z_]","");
				
						$OrigFirstName = $FirstName
						$OrigLastName = $LastName
				
						# Check length of FirstName & LastName is not longer than 20 characters.
						if ($FirstName.Length + $LastName.Length -gt 18) 
							{
								$FirstName = $FirstName[0]
								if ($FirstName.Length + $LastName.Length -gt 18)
									{
										$LastName = $LastName.Substring(0,17)
									}
							}
						# Select Naming convention to use.
						$Name = switch ($Staff_Naming_Convention) 
					    		{ 
					        		"FirstName.LastName"		{"{0}.{1}" -f $FirstName,$LastName}
					        		"FirstInitial.LastName"  	{"{0}.{1}" -f ($OrigFirstName)[0],$LastName}
									"LastName.FirstInitial"  	{"{0}.{1}" -f $LastName,($OrigFirstName)[0]}
									"FirstNameLastName"			{"{0}{1}" -f $FirstName,$LastName}
									"FirstInitialLastName"  	{"{0}{1}" -f ($OrigFirstName)[0],$LastName}
					        		"LastNameFirstInitial"  	{"{0}{1}" -f $LastName,($OrigFirstName)[0]}
					        		Default                 	{"{0}.{1}" -f $FirstName,$LastName}  
					    		}	
						$defaultname = $Name
				
						$FirstName = $OrigFirstName
						$LastName = $OrigLastName
				
						$External_ID_Match = (Get-ADUser -Filter {bullTCLextension1 -eq $External_ID} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
#						write-host $External_ID_Match
						if ($External_ID_Match -eq $Null)
							{
								"***************************************************************************" | out-file -filepath $LogFile -append
								(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
								"No existing user found... creating new one with the following properties..." | out-file -filepath $LogFile -append
								Do	
									{
							    	Try
							    		{
											#	Checks for all matching sAMAccountName.
											$User = (Get-ADUser -Filter {SamAccountName -eq $Name} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
											if ($User -eq $null) {$Exit = 1}
							        		else
							         			{
							        				#	The user exists, add +1 to count
							            			$Name  = $defaultname + $Count++
												}
							        	}
							    	Catch
								    	{
											#	User does not exist. got to exit.
								        	$Exit = 1       
								    	}
									}
								While ($Exit -eq 0)
							
									#write-host $Name
								$HomeFolderPath = ("\\" + $StaffHomeServer + "\" + $StaffHomeFolderRoot + "\" + $Name)	#full path to user home folder.
									## write-host $HomeFolderPath
							
								$UserProperties = @{
									"Name" = $Name
									"SamAccountName" = $Name
									"GivenName" = $FirstName
									"SurName" = $LastName
									"UserPrincipalName" = ($Name + "@" + $DNS_Name)
									"DisplayName" = ($FirstName[0] + " " + $LastName)
									"Path" = $Staff_OU_LDAP
									"OtherAttributes" = @{'bullTCLextension1' = $External_ID;'bullTCLextension2' = $HomeFolderPath}
									"AccountPassword" = $InitialStaffPwd
									"ChangePasswordAtLogon" = $true
									"Enabled" = $true
									}	

								#	Create AD User & add to default group.
				 				New-ADUser @UserProperties
								Add-ADGroupMember -Identity ($Group_Prefix + " " + $Default_Staff_Group) -Members $Name -Confirm:$false
								TestStaffFileExists
								
								#	Write to log file.
								"SamAccountName: " + $Name | out-file -filepath $LogFile -append
								"GivenName: " + $FirstName | out-file -filepath $LogFile -append
								"SurName: " + $LastName | out-file -filepath $LogFile -append
								"UserPrincipalName: " + ($Name + "@" + $DNS_Name) | out-file -filepath $LogFile -append
								"DisplayName: " + ($FirstName[0] + " " + $LastName) | out-file -filepath $LogFile -append
								"OU Location: " + $Staff_OU_LDAP | out-file -filepath $LogFile -append
								"bullTCLextension1: " + $External_ID | out-file -filepath $LogFile -append
								"bullTCLextension2: " + $HomeFolderPath | out-file -filepath $LogFile -append
								"MemberOf: " + ($Group_Prefix + " " + $Default_Staff_Group) | out-file -filepath $LogFile -append
								$Name + "," + $FirstName + "," + $LastName + "," + $XML.Options.Staff.InitialStaffPwd | out-file -filepath $StaffOutput -append
							
								#	Check that user account has been created on all Domain Controllers in the AD Site before creating home folder.
								foreach($DC in $DCs)
									{
										$DomainController = $DC.DNSHostName
										$Exit = 0
										Do	
											{
												#	Checks for a matching sAMAccountName.
												$User = (Get-ADUser -Server $DomainController -Filter {SamAccountName -eq $Name} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
												if ($User -eq $null) 
													{
														$Exit = 0
													}
												else
													{
														#	The user exists so exit this loop.
														$Exit = 1
													}
											}
										While ($Exit -eq 0)
									}
								
								if($CreateHomeFolders -eq "No")
									{
										#	Don't create home folders for new users.
									}
								else
									{
										#	Test if home folder exists and create one if it does not.
										if(!(Test-Path -Path $HomeFolderPath))
											{
												#	Create folder.
												"Attempting to create new homefolder... " + $HomeFolderPath | out-file -filepath $LogFile -append
												
												Try
													{
														md $HomeFolderPath -ErrorAction Stop
														
														"Home Folder Created: " + $HomeFolderPath | out-file -filepath $LogFile -append
														#	Read folder ACL.
														$acl = Get-Acl $HomeFolderPath
														$acl.SetAccessRuleProtection($false, $true)
														#	Append User account to ACL permissions with "Modify" .
														$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($Name, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
														#	Write new ACL to folder.
														Set-Acl $HomeFolderPath $acl
														"Permissions set on Home Folder for: " + $Name | out-file -filepath $LogFile -append
													}
												Catch
													{
														$ErrMsg = "Error. Unable to create folder: $_"	
														$ErrMsg | out-file -filepath $LogFile -append
													}
											}
										else
											{
												#	Read existing folder ACL.
												$acl = Get-Acl $HomeFolderPath
												$acl.SetAccessRuleProtection($false, $true)
												#List Users/Groups in ACL with permissions
												$acl.Access | Select IdentityReference, FileSystemRights
												#Remove All non-inherited Permissions
												$acl.Access | ForEach-Object {if ($_.IsInherited -eq $False) {$acl.RemoveAccessRule($_)}}
												#	Append User account to ACL permissions with "Modify" .
												$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($Name, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
												#	Write new ACL to folder.
												Set-Acl $HomeFolderPath $acl
												"Permissions set on Home Folder for: " + $Name | out-file -filepath $LogFile -append
											}
									}	
								#-----------------	
								#region New Staff Office 365 Mailbox
								#	-------------------------------------------------------------------------------------------
								#	Check if we need to create a mailbox for the user account on either Office 365 or Exchange.
								#	-------------------------------------------------------------------------------------------
								if($StaffMailSolution -eq "Office365")
									{
										$365UPN = ($Name + "@" + $365MailDomainName)
										
										$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

										if($Office365UPN_Match -eq $null)
											{
												#	This is null so create new user mailbox.
		#										Write-Host "Creating new mailbox... please wait"
												
												$365UserProperties = @{
													"UserPrincipalName" = $365UPN
													"DisplayName" = ($FirstName[0] + " " + $LastName)
													"FirstName" = $FirstName
													"LastName" = $LastName
													"UsageLocation" = "GB"
													"LicenseAssignment" = $365StaffLicenseAssignment
													"StrongPasswordRequired" = $false
													"Password" = $Office365StaffPwd
													"ForceChangePassword" = $true
													}
												
												New-MsolUser @365UserProperties
												
												#	Check that mailbox has been created and get unique ObjectId.
												$365Exit = 0
												Do	
													{
														#	Checks mailbox is created.
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})
														if ($Office365UPN_Match -eq $null) 
															{
																$365Exit = 0
															}
														else
															{
																#	The mailbox exists so exit this loop.
																$365Exit = 1
															}
													}
												While ($365Exit -eq 0)
		#										Write-Host $Office365UPN_Match.ObjectId	#	Unique Id of Office 365 user.
												Set-ADUser -Identity $Name -Replace @{bullTCLextension3=$Office365UPN_Match.ObjectId}
												
												#	Write to log file
												"Creating Office 365 mailbox..." | out-file -filepath $LogFile -append
												"365UserPrincipalName: " + $365UPN | out-file -filepath $LogFile -append
												"365DisplayName: " + ($FirstName[0] + " " + $LastName) | out-file -filepath $LogFile -append
												"365FirstName: " + $FirstName | out-file -filepath $LogFile -append
												"365LastName: " + $LastName | out-file -filepath $LogFile -append 
												"365ObjectId: " + $Office365UPN_Match.ObjectId | out-file -filepath $LogFile -append
											}
										else	
											{
												#	Match found... check properties.
												ForEach($365 in $Office365UPN_Match)
													{												
														if($365.DisplayName -ne ($FirstName[0] + " " + $LastName))
															{
																"365DisplayName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -DisplayName ($FirstName[0] + " " + $LastName)
															}
														
														if($365.FirstName -ne $FirstName)
															{
																"365FirstName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -FirstName $FirstName
															}
														
														if($365.LastName -ne $LastName)
															{
																"365LastName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -LastName $LastName
															}

													}
											}
									
									}
								else
									{
										#	Do not create Office 365 mailboxes.
									}
										#---------------------
										#endregion New Staff Office 365 Mailbox
										#---------------------
							}
#	------Update the user-------
						else 
							{	
									# write-host "Existing user account found... checking properties match..."
								#	Update existing AD User.
								#	# write-host $External_ID_Match
								$ADUserProps = (Get-ADUser -Identity $External_ID_Match -Properties GivenName, SurName, DisplayName, Memberof)
								$DN = $(Get-ADUser -Identity $External_ID_Match).distinguishedName
								if ($ADUserProps.Get_Item("GivenName") -ne $FirstName)
									{
											# write-host "FirstName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -GivenName $FirstName
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"FirstName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
								if ($ADUserProps.Get_Item("SurName") -ne $LastName)
									{
											# write-host "LastName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -SurName $LastName
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"LastName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
								if ($ADUserProps.Get_Item("DisplayName") -ne ($FirstName[0] + " " + $LastName))
									{
											# write-host "DisplayName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -DisplayName ($FirstName[0] + " " + $LastName)
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"DisplayName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
							
																#-----------------	
								#region Update Satff Office 365 Mailbox
								#	-------------------------------------------------------------------------------------------
								#	Check if we need to create a mailbox for the user account on either Office 365 or Exchange.
								#	-------------------------------------------------------------------------------------------
								if($StaffMailSolution -eq "Office365")
									{
										$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
										
										$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

										if($Office365UPN_Match -eq $null)
											{
												#	This is null so create new user mailbox.
		#										Write-Host "Creating new mailbox... please wait"
												
												$365UserProperties = @{
													"UserPrincipalName" = $365UPN
													"DisplayName" = ($FirstName[0] + " " + $LastName)
													"FirstName" = $FirstName
													"LastName" = $LastName
													"UsageLocation" = "GB"
													"LicenseAssignment" = $365StaffLicenseAssignment
													"StrongPasswordRequired" = $false
													"Password" = $Office365StaffPwd
													"ForceChangePassword" = $true
													}
												
												New-MsolUser @365UserProperties
												
												#	Check that mailbox has been created and get unique ObjectId.
												$365Exit = 0
												Do	
													{
														#	Checks mailbox is created.
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})
														if ($Office365UPN_Match -eq $null) 
															{
																$365Exit = 0
															}
														else
															{
																#	The mailbox exists so exit this loop.
																$365Exit = 1
															}
													}
												While ($365Exit -eq 0)
		#										Write-Host $Office365UPN_Match.ObjectId	#	Unique Id of Office 365 user.
												Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension3=$Office365UPN_Match.ObjectId}
												
												#	Write to log file
												"Creating Office 365 mailbox..." | out-file -filepath $LogFile -append
												"365UserPrincipalName: " + $365UPN | out-file -filepath $LogFile -append
												"365DisplayName: " + ($FirstName[0] + " " + $LastName) | out-file -filepath $LogFile -append
												"365FirstName: " + $FirstName | out-file -filepath $LogFile -append
												"365LastName: " + $LastName | out-file -filepath $LogFile -append 
												"365ObjectId: " + $Office365UPN_Match.ObjectId | out-file -filepath $LogFile -append
											}
										else	
											{
												#	Match found... check properties.
												ForEach($365 in $Office365UPN_Match)
													{
														if($365.DisplayName -ne ($FirstName[0] + " " + $LastName))
															{
																"365DisplayName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -DisplayName ($FirstName[0] + " " + $LastName)
															}
														
														if($365.FirstName -ne $FirstName)
															{
																"365FirstName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -FirstName $FirstName
															}
														
														if($365.LastName -ne $LastName)
															{
																"365LastName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -LastName $LastName
															}

													}
											}
									
									}
								else
									{
										#	Do not create Office 365 mailboxes.
									}
										#---------------------
										#endregion Update Staff Office 365 Mailbox
										#---------------------
							
								$Groups = (Get-ADPrincipalGroupMembership $External_ID_Match | select name)
								$DefGroupMatch = 0
						
								foreach($Group in $Groups)
									{
										if ($Group.name -eq ($Group_Prefix + " " + $Default_Staff_Group))
											{
												$DefGroupMatch = 1
											}
									}
								
								if ($DefGroupMatch -ne 1)
									{
										Add-ADGroupMember -Identity ($Group_Prefix + " " + $Default_Staff_Group) -Members $External_ID_Match -Confirm:$false
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"Adding " + $External_ID_Match + " to " + ($Group_Prefix + " " + $Default_Staff_Group) + " group" | out-file -filepath $LogFile -append
									}
									
								if($RecreateHomeFolders -eq "No")
									{
										#	Don't recreate user home folders.
									}
								else
									{
										$HomeFolderPath = ("\\" + $StaffHomeServer + "\" + $StaffHomeFolderRoot + "\" + $External_ID_Match)	#full path to user home folder.
										#	Test if home folder exists and create one if it does not.
										if(!(Test-Path -Path $HomeFolderPath))
											{
												#	Create folder.
												Try
													{
														md $HomeFolderPath -ErrorAction Stop
														
														"Home Folder Created: " + $HomeFolderPath | out-file -filepath $LogFile -append
														#	Read folder ACL.
														$acl = Get-Acl $HomeFolderPath
														$acl.SetAccessRuleProtection($false, $true)
														#	Append User account to ACL permissions with "Modify" .
														$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($External_ID_Match, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
														#	Write new ACL to folder.
														Set-Acl $HomeFolderPath $acl
														"Permissions set on Home Folder for: " + $External_ID_Match | out-file -filepath $LogFile -append
													}
												Catch
													{
														$ErrMsg = "Error. Unable to create folder: $_"	
														$ErrMsg | out-file -filepath $LogFile -append
													}
											}
										else
											{
												#	Read existing folder ACL.
												$acl = Get-Acl $HomeFolderPath
												$acl.SetAccessRuleProtection($false, $true)
												#List Users/Groups in ACL with permissions
												$acl.Access | Select IdentityReference, FileSystemRights
												#Remove All non-inherited Permissions
												$acl.Access | ForEach-Object {if ($_.IsInherited -eq $False) {$acl.RemoveAccessRule($_)}}
												#	Append User account to ACL permissions with "Modify" .
												$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($External_ID_Match, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
												#	Write new ACL to folder.
												Set-Acl $HomeFolderPath $acl
												"Permissions set on Home Folder for: " + $External_ID_Match | out-file -filepath $LogFile -append
											}
									}
									
								if ($DN -replace ("CN=" + $External_ID_Match + ",")  -ne $Staff_OU_LDAP)
									{
											# write-host "Account in wrong OU... moving to correct one..."
										Move-ADObject $DN -TargetPath $Staff_OU_LDAP
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"User Account in wrong OU (" + $DN + ")... moving " + $External_ID_Match + " to: " + $Staff_OU_LDAP | out-file -filepath $LogFile -append
									}
							}
					}	
			}
	}
#----------------------------------------------
#endregion Read Staff Current CSV file
#----------------------------------------------

#----------------------------------------------
#region Read Pupil Current CSV file
#----------------------------------------------
#	Test if CSV file exists.
if(!(Test-Path -Path ($CSVExtractLocation + "\AD Pupils Current.csv")))
	{
		#	CSV file does not exist...
		"***************************************************************************" | out-file -filepath $LogFile -append
		(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
		"Student CSV file does not exist..." | out-file -filepath $LogFile -append 
		"***************************************************************************" | out-file -filepath $LogFile -append
	}
else
	{
		#	Read 'Student_Extract.csv' - this file contains all current students users for the school.
		"***************************************************************************" | out-file -filepath $LogFile -append
		(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
		"Reading Student CSV file..." | out-file -filepath $LogFile -append 
		"***************************************************************************" | out-file -filepath $LogFile -append
		$StudentUsers=Import-Csv ($CSVExtractLocation + "\AD Pupils Current.csv")
		# write-host "Creating Pupils"
		foreach($UD in $StudentUsers) 
			{
				$Exit = 0
				$Count = 1
				$External_ID = $UD.External_ID	#bullTCLextension1
				if ($External_ID -eq $Null)
					{
						#	This is blank so skip user creation.
					}
				elseif($External_ID -eq "")
					{
						# "empty string"
					}
				else
					{
						$FirstName = $UD.Forename	#givenName
						$LastName = $UD.Surname		#sn
						$Year = $UD.Year	#Current student year e.g. Year1 etc...
				
						# Strip out invalid characters.
						$FirstName = [System.Text.RegularExpressions.Regex]::Replace($FirstName,"[^1-9a-zA-Z_]","");
						$LastName = [System.Text.RegularExpressions.Regex]::Replace($LastName,"[^1-9a-zA-Z_]","");
				
						$OrigFirstName = $FirstName
						$OrigLastName = $LastName
				
						# Check length of FirstName & LastName is not longer than 20 characters.
						if ($FirstName.Length + $LastName.Length -gt 18) 
							{
								$FirstName = $FirstName[0]
								if ($FirstName.Length + $LastName.Length -gt 18)
									{
										$LastName = $LastName.Substring(0,17)
									}
							}
						# Select Naming convention to use.
						$Name = switch ($Student_Naming_Convention) 
							    { 
							        "FirstName.LastName"		{"{0}.{1}" -f $FirstName,$LastName}
					    		    "FirstInitial.LastName"  	{"{0}.{1}" -f ($OrigFirstName)[0],$LastName}
									"LastName.FirstInitial"  	{"{0}.{1}" -f $LastName,($OrigFirstName)[0]}
									"FirstNameLastName"			{"{0}{1}" -f $FirstName,$LastName}
									"FirstInitialLastName"  	{"{0}{1}" -f ($OrigFirstName)[0],$LastName}
					        		"LastNameFirstInitial"  	{"{0}{1}" -f $LastName,($OrigFirstName)[0]}
					        		Default                 	{"{0}.{1}" -f $FirstName,$LastName}  
					    		}	
						$defaultname = $Name
				
						$FirstName = $OrigFirstName
						$LastName = $OrigLastName
				
						$External_ID_Match = (Get-ADUser -Filter {bullTCLextension1 -eq $External_ID} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
						#	# write-host $External_ID_Match
						if ($External_ID_Match -eq $Null)
							{
								## write-host "Creating new pupil account"
								"***************************************************************************" | out-file -filepath $LogFile -append
								(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
								"No existing user found... creating new one with the following properties..." | out-file -filepath $LogFile -append
								Do	
									{
							    		Try
							    			{
												#	Checks for all matching sAMAccountName.
												$User = (Get-ADUser -Filter {SamAccountName -eq $Name} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
												if ($User -eq $null) {$Exit = 1}
							        			else
							         			{
							        				#	The user exists, add +1 to count
							            			$Name  = $defaultname + $Count++
												}
							        		}
							    		Catch
								    		{
												#	User does not exist. got to exit.
								        		$Exit = 1       
								    		}
									}
								While ($Exit -eq 0)
							
#									 write-host $Name
								$HomeFolderPath = ("\\" + $StudentHomeServer + "\" + $StudentHomeFolderRoot + "\" + $Year + "\" + $Name)	#full path to user home folder.
									## write-host $HomeFolderPath
							
								$UserProperties = @{
								"Name" = $Name
								"SamAccountName" = $Name
								"GivenName" = $FirstName
								"SurName" = $LastName
								"UserPrincipalName" = ($Name + "@" + $DNS_Name)
								"DisplayName" = ($FirstName + " " + $LastName)
								"Path" = ("OU=" + $Year + "," +$Student_OU_LDAP)
								"OtherAttributes" = @{'bullTCLextension1' = $External_ID;'bullTCLextension2' = $HomeFolderPath}
								"AccountPassword" = $InitialStudentPwd
								"ChangePasswordAtLogon" = $true
								"Enabled" = $true
								}
								
								#	Create AD User.
			 					New-ADUser @UserProperties
								Add-ADGroupMember -Identity ($Group_Prefix + "-" + $Year) -Members $Name -Confirm:$false
								Add-ADGroupMember -Identity ($Group_Prefix + " " + $Default_Student_Group) -Members $Name -Confirm:$false
								TestPupilFileExists
							
								#	Write to log file.
								"SamAccountName: " + $Name | out-file -filepath $LogFile -append
								"GivenName: " + $FirstName | out-file -filepath $LogFile -append
								"SurName: " + $LastName | out-file -filepath $LogFile -append
								"UserPrincipalName: " + ($Name + "@" + $DNS_Name) | out-file -filepath $LogFile -append
								"DisplayName: " + ($FirstName + " " + $LastName) | out-file -filepath $LogFile -append
								"OU Location: " + ("OU=" + $Year + "," +$Student_OU_LDAP) | out-file -filepath $LogFile -append
								"bullTCLextension1: " + $External_ID | out-file -filepath $LogFile -append
								"bullTCLextension2: " + $HomeFolderPath | out-file -filepath $LogFile -append
								"MemberOf: " + ($Group_Prefix + "-" + $Year) | out-file -filepath $LogFile -append
								"MemberOf: " + ($Group_Prefix + " " + $Default_Student_Group) | out-file -filepath $LogFile -append
								$Name + "," + $FirstName + "," + $LastName + "," + $Year + "," + $XML.Options.Students.InitialStudentPwd | out-file -filepath $PupilOutput -append
											
								#	Check that user account has been created on all Domain Controllers in the AD Site.
								foreach($DC in $DCs)
									{
										$DomainController = $DC.DNSHostName
										## write-host $DomainController
										$Exit = 0
										Do	
											{
												#	Checks for a matching sAMAccountName.
												$User = (Get-ADUser -Server $DomainController -Filter {SamAccountName -eq $Name} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
												if ($User -eq $null) 
													{
														$Exit = 0
													}
												else
													{
														#	The user exists so exit this loop.
														$Exit = 1
													}
											}
										While ($Exit -eq 0)
									}
								## write-host $CreateHomeFolders	
								if($CreateHomeFolders -eq "No")
									{
										#	Don't create home folders for new users.
									}
								else	
									{
										#	Test if home folder exists and create one if it does not.
										if(!(Test-Path -Path $HomeFolderPath))
											{
												#	Create folder.
												Try
													{
														md $HomeFolderPath -ErrorAction Stop
														
														"Home Folder Created: " + $HomeFolderPath | out-file -filepath $LogFile -append
														#	Read folder ACL.
														$acl = Get-Acl $HomeFolderPath
														$acl.SetAccessRuleProtection($false, $true)
														#	Append User account to ACL permissions with "Modify" .
														$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($Name, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
														#	Write new ACL to folder.
														Set-Acl $HomeFolderPath $acl
														"Permissions set on Home Folder for: " + $Name | out-file -filepath $LogFile -append
													}
												Catch
													{
														$ErrMsg = "Error. Unable to create folder: $_"	
														$ErrMsg | out-file -filepath $LogFile -append
													}
											}
										else
											{
												#	Read folder ACL.
												$acl = Get-Acl $HomeFolderPath
												$acl.SetAccessRuleProtection($false, $true)
												#List Users/Groups in ACL with permissions
												$acl.Access | Select IdentityReference, FileSystemRights
												#Remove All non-inherited Permissions
												$acl.Access | ForEach-Object {if ($_.IsInherited -eq $False) {$acl.RemoveAccessRule($_)}}
												#	Append User account to ACL permissions with "Modify".
												$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($Name, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
												#	Write new ACL to folder.
												Set-Acl $HomeFolderPath $acl
												"Permissions set on Home Folder for: " + $Name | out-file -filepath $LogFile -append
											}
									}
								#-----------------	
								#region New Pupil Office 365 Mailbox
								#	-------------------------------------------------------------------------------------------
								#	Check if we need to create a mailbox for the user account on either Office 365 or Exchange.
								#	-------------------------------------------------------------------------------------------
								if($StudentMailSolution -eq "Office365")
									{
										$365UPN = ($Name + "@" + $365MailDomainName)
										
										$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

										if($Office365UPN_Match -eq $null)
											{
												#	This is null so create new user mailbox.
		#										Write-Host "Creating new mailbox... please wait"
												
												$365UserProperties = @{
													"UserPrincipalName" = $365UPN
													"DisplayName" = ($FirstName + " " + $LastName)
													"FirstName" = $FirstName
													"LastName" = $LastName
													"UsageLocation" = "GB"
													"LicenseAssignment" = $365StudentLicenseAssignment
													"StrongPasswordRequired" = $false
													"Password" = $Office365StudentPwd
													"ForceChangePassword" = $true
													}
												
												New-MsolUser @365UserProperties
												
												#	Check that mailbox has been created and get unique ObjectId.
												$365Exit = 0
												Do	
													{
														#	Checks mailbox is created.
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})
														if ($Office365UPN_Match -eq $null) 
															{
																$365Exit = 0
															}
														else
															{
																#	The mailbox exists so exit this loop.
																$365Exit = 1
															}
													}
												While ($365Exit -eq 0)
		#										Write-Host $Office365UPN_Match.ObjectId	#	Unique Id of Office 365 user.
												Set-ADUser -Identity $Name -Replace @{bullTCLextension3=$Office365UPN_Match.ObjectId}
												
												#	Write to log file
												"Creating Office 365 mailbox..." | out-file -filepath $LogFile -append
												"365UserPrincipalName: " + $365UPN | out-file -filepath $LogFile -append
												"365DisplayName: " + ($FirstName[0] + " " + $LastName) | out-file -filepath $LogFile -append
												"365FirstName: " + $FirstName | out-file -filepath $LogFile -append
												"365LastName: " + $LastName | out-file -filepath $LogFile -append 
												"365ObjectId: " + $Office365UPN_Match.ObjectId | out-file -filepath $LogFile -append
											}
										else	
											{
												#	Match found... check properties.
												ForEach($365 in $Office365UPN_Match)
													{
														if($365.DisplayName -ne ($FirstName + " " + $LastName))
															{
																"365DisplayName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -DisplayName ($FirstName[0] + " " + $LastName)
															}
														
														if($365.FirstName -ne $FirstName)
															{
																"365FirstName is wrong..." | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -FirstName $FirstName
															}
														
														if($365.LastName -ne $LastName)
															{
																"365LastName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -LastName $LastName
															}

													}
											}
									
									}
								else
									{
										#	Do not create Office 365 mailboxes.
									}
										#---------------------
										#endregion New Pupil Office 365 Mailbox
										#---------------------	
							}
#	------Update the user-------							
						else 
							{	
									# write-host "Matching user account found... checking properties match..."
								#	Check and update existing AD User.

								$ADUserProps = (Get-ADUser -Identity $External_ID_Match -Properties GivenName, SurName, DisplayName, bullTCLextension2)
								$DN = $(Get-ADUser -Identity $External_ID_Match).distinguishedName
								if ($ADUserProps.Get_Item("GivenName") -ne $FirstName)
									{
											# write-host "FirstName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -GivenName $FirstName
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"FirstName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
								if ($ADUserProps.Get_Item("SurName") -ne $LastName)
									{
											# write-host "LastName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -SurName $LastName
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"LastName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
								if ($ADUserProps.Get_Item("DisplayName") -ne ($FirstName + " " + $LastName))
									{
											# write-host "DisplayName is wrong... changing to defined standard..."
										Set-ADUser -Identity $External_ID_Match -DisplayName ($FirstName + " " + $LastName)
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"DisplayName is wrong... changing to defined standard on " + $External_ID_Match | out-file -filepath $LogFile -append
									}
					
								#region Update Pupil Office 365 Mailbox
								#	-------------------------------------------------------------------------------------------
								#	Check if we need to create a mailbox for the user account on either Office 365 or Exchange.
								#	-------------------------------------------------------------------------------------------
								if($StudentMailSolution -eq "Office365")
									{
										$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
										
										$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

										if($Office365UPN_Match -eq $null)
											{
												#	This is null so create new user mailbox.
		#										Write-Host "Creating new mailbox... please wait"
												
												$365UserProperties = @{
													"UserPrincipalName" = $365UPN
													"DisplayName" = ($FirstName + " " + $LastName)
													"FirstName" = $FirstName
													"LastName" = $LastName
													"UsageLocation" = "GB"
													"LicenseAssignment" = $365StudentLicenseAssignment
													"StrongPasswordRequired" = $false
													"Password" = $Office365StudentPwd
													"ForceChangePassword" = $true
													}
												
												New-MsolUser @365UserProperties
												
												#	Check that mailbox has been created and get unique ObjectId.
												$365Exit = 0
												Do	
													{
														#	Checks mailbox is created.
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})
														if ($Office365UPN_Match -eq $null) 
															{
																$365Exit = 0
															}
														else
															{
																#	The mailbox exists so exit this loop.
																$365Exit = 1
															}
													}
												While ($365Exit -eq 0)
		#										Write-Host $Office365UPN_Match.ObjectId	#	Unique Id of Office 365 user.
												Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension3=$Office365UPN_Match.ObjectId}
												
												#	Write to log file
												"Creating Office 365 mailbox..." | out-file -filepath $LogFile -append
												"365UserPrincipalName: " + $365UPN | out-file -filepath $LogFile -append
												"365DisplayName: " + ($FirstName[0] + " " + $LastName) | out-file -filepath $LogFile -append
												"365FirstName: " + $FirstName | out-file -filepath $LogFile -append
												"365LastName: " + $LastName | out-file -filepath $LogFile -append 
												"365ObjectId: " + $Office365UPN_Match.ObjectId | out-file -filepath $LogFile -append
											}
										else	
											{
												#	Match found... check properties.
												ForEach($365 in $Office365UPN_Match)
													{												
														if($365.DisplayName -ne ($FirstName + " " + $LastName))
															{
																"365DisplayName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -DisplayName ($FirstName[0] + " " + $LastName)
															}
														
														if($365.FirstName -ne $FirstName)
															{
																"365FirstName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -FirstName $FirstName
															}
														
														if($365.LastName -ne $LastName)
															{
																"365LastName is wrong... updating" | out-file -filepath $LogFile -append
																Set-MsolUser -UserPrincipalName $365UPN -LastName $LastName
															}

													}
											}
									
									}
								else
									{
										#	Do not create Office 365 mailboxes.
									}
										#---------------------
										#endregion Update Pupil Office 365 Mailbox
										#---------------------
					
									$Groups = (Get-ADPrincipalGroupMembership $External_ID_Match | select name)
									$GroupMatch = 0
									$DefGroupMatch = 0
						
								foreach($Group in $Groups)
									{
											# write-host $Group.name
										if ($Group.name -eq ($Group_Prefix + " " + $Default_Staff_Group))
											{
												Remove-ADGroupMember -Identity $Group.name -Member $External_ID_Match -Confirm:$false
												"***************************************************************************" | out-file -filepath $LogFile -append
												(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
												"Removing " + $External_ID_Match + " from " + $Group.name + " group" | out-file -filepath $LogFile -append
											}
										if ($Group.name -eq ($Group_Prefix + " " + $Default_Student_Group))
											{
												$DefGroupMatch = 1
											}
										if ($Group.name -like ($Group_Prefix + "-" + "Year*"))	#list groups like 'Year' that user is a MemberOf
											{
												if ($Group.name -ne ($Group_Prefix + "-" + $Year))
													{
														Remove-ADGroupMember -Identity $Group.name -Member $External_ID_Match -Confirm:$false
														"***************************************************************************" | out-file -filepath $LogFile -append
														(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
														"Removing " + $External_ID_Match + " from " + $Group.name + " group" | out-file -filepath $LogFile -append
													}
												else
													{
														$GroupMatch = 1
													}
											}
									}
								if ($GroupMatch -ne 1)
									{
										Add-ADGroupMember -Identity ($Group_Prefix + "-" + $Year) -Members $External_ID_Match -Confirm:$false
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"Adding " + $External_ID_Match + " to " + ($Group_Prefix + "-" + $Year) + " group" | out-file -filepath $LogFile -append
									}
								if ($DefGroupMatch -ne 1)
									{
										Add-ADGroupMember -Identity ($Group_Prefix + " " + $Default_Student_Group) -Members $External_ID_Match -Confirm:$false
										"***************************************************************************" | out-file -filepath $LogFile -append
										(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
										"Adding " + $External_ID_Match + " to " + ($Group_Prefix + " " + $Default_Student_Group) + " group" | out-file -filepath $LogFile -append
									}
									
								## write-host $RecreateHomeFolders
								$HomeFolderPath = ("\\" + $StudentHomeServer + "\" + $StudentHomeFolderRoot + "\" + $Year + "\" + $External_ID_Match)	#calculated path to user home folder.
								$CurrentHomeFolderPath = $ADUserProps.Get_Item("bullTCLextension2")	#full path to user home folder from bullTCLextension2.
								
								if($RecreateHomeFolders -eq "No")
									{
										#	Don't recreate user home folders but check path for extensionAttribute2 is correct.
										if($HomeFolderPath -ne $CurrentHomeFolderPath)
											{
												#	Update value in bullTCLextension2 with $HomeFolderPath
												Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension2=$HomeFolderPath}
												"Updated bullTCLextension2: " + $HomeFolderPath | out-file -filepath $LogFile -append
											}
									}
								else
									{
										if($HomeFolderPath -ne $CurrentHomeFolderPath)
											{
												#	Test if home folder exists and create one if it does not.
												if(!(Test-Path -Path $CurrentHomeFolderPath))
													{
														#	Create folder.
														Try
															{
																md $HomeFolderPath -ErrorAction Stop
														
																"***************************************************************************" | out-file -filepath $LogFile -append
																(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
																"Home Folder does not exist in expected location...New one Created: " + $HomeFolderPath | out-file -filepath $LogFile -append
																#	Read folder ACL.
																$acl = Get-Acl $HomeFolderPath
																$acl.SetAccessRuleProtection($false, $true)
																#	Append User account to ACL permissions with "Modify".
																$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($External_ID_Match, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
																#	Write new ACL to folder.
																Set-Acl $HomeFolderPath $acl
																"Permissions set on Home Folder for: " + $External_ID_Match | out-file -filepath $LogFile -append
																#	Update value in bullTCLextension2 with $HomeFolderPath
																Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension2=$HomeFolderPath}
																"Updated bullTCLextension2: " + $HomeFolderPath | out-file -filepath $LogFile -append
															}
														Catch
															{
																$ErrMsg = "Error. Unable to create folder: $_"	
																$ErrMsg | out-file -filepath $LogFile -append
															}
													}
												else
													{
														#	Folder exists in old path location...moving to new location.
														#	# write-host "home folder found at " $CurrenthomeFolderPath " ...moving to new location " $HomeFolderPath
														Try
															{
																Move-Item $CurrenthomeFolderPath $HomeFolderPath -ErrorAction Stop
																
																"***************************************************************************" | out-file -filepath $LogFile -append
																(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
																"Home Folder moved from " + $CurrenthomeFolderPath + " to " + $HomeFolderPath | out-file -filepath $LogFile -append
																#	Update value in bullTCLextension2 with $HomeFolderPath
																Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension2=$HomeFolderPath}
																"Updated bullTCLextension2: " + $HomeFolderPath | out-file -filepath $LogFile -append
															}
														Catch
															{
																$ErrMsg = "Error. Unable to move folder: $_"	
																$ErrMsg | out-file -filepath $LogFile -append
															}
													}

											}
										else
											{
												#	Read existing folder ACL.
												$acl = Get-Acl $HomeFolderPath
												$acl.SetAccessRuleProtection($false, $true)
												#List Users/Groups in ACL with permissions
												$acl.Access | Select IdentityReference, FileSystemRights
												#Remove All non-inherited Permissions
												$acl.Access | ForEach-Object {if ($_.IsInherited -eq $False) {$acl.RemoveAccessRule($_)}}
												#	Append User account to ACL permissions with "Modify" .
												$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($External_ID_Match, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")))
												#	Write new ACL to folder.
												Set-Acl $HomeFolderPath $acl
												"Permissions set on Home Folder for: " + $External_ID_Match | out-file -filepath $LogFile -append
											}
									}		
							
							if ($DN -replace ("CN=" + $External_ID_Match + ",")  -ne ("OU=" + $Year + "," + $Student_OU_LDAP))
								{
									# write-host ("Account in wrong OU... (" + $DN + ") moving to correct one...")
									
									
									Move-ADObject $DN -TargetPath ("OU=" + $Year + "," + $Student_OU_LDAP)
									"***************************************************************************" | out-file -filepath $LogFile -append
									(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
									"User Account in wrong OU (" + $DN + ")... moving " + $External_ID_Match + " to: " + ("OU=" + $Year + "," + $Student_OU_LDAP) | out-file -filepath $LogFile -append
								}
							}
					}	
			}
	}
#----------------------------------------------
#endregion Read Pupil Current CSV file
#----------------------------------------------

#----------------------------------------------
#region Read Staff Leavers CSV file
#----------------------------------------------
#	Test if CSV file exists.
if(!(Test-Path -Path ($CSVExtractLocation + "\AD Staff Leavers.csv")))
	{
		#	CSV file does not exist...
	}
else
	{
		#	Read 'AD Staff Leavers.csv' - this file contains all staff that have left the school.
		$StaffLeavers = Import-Csv ($CSVExtractLocation + "\AD Staff Leavers.csv")

		foreach($UD in $StaffLeavers) 
			{	
				$External_ID = $UD.External_ID	#bullTCLextension1
				
				if($External_ID -eq $null)
					{
						#	CSV file is empty so do nothing.
					}
				elseif($External_ID -eq"")
					{
						#	"empty string"
					}
				else
					{		
						$DateOfLeaving = Get-Date $UD.DateOfLeaving
						$DateOfLeaving = Get-Date $DateOfLeaving
				
						if($DateOfLeaving -gt $CurrentDate)
							{
								#	# write-host "not yet time to do anything 1"
							}
						elseif($DateOfLeaving -match $CurrentDate)
							{
								#	# write-host "not yet time to do anything 2"
							}
						else
							{
								$External_ID_Match = (Get-ADUser -Filter {bullTCLextension1 -eq $External_ID} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
								#	# write-host $External_ID_Match
								
								if ($External_ID_Match -eq $Null)
									{
										#	User account has already been removed from Active Directory.
									}
								else
									{
										$ADUserProps = (Get-ADUser -Identity $External_ID_Match -Properties bullTCLextension2)
										$DN = $(Get-ADUser -Identity $External_ID_Match).distinguishedName
										$CurrentHomeFolderPath = $ADUserProps.Get_Item("bullTCLextension2")	#full path to user home folder from bullTCLextension2.
										#	# write-host $CurrentHomeFolderPath
										$DisabledFolderPath = ("\\" + $StaffHomeServer + "\" + $StaffHomeFolderRoot + "\DisabledUsers\" + $External_ID_Match)
		
										if(!(Test-Path -Path $CurrentHomeFolderPath))
											{
												#	# write-host "home folder does not exist in expected location...."
												"***************************************************************************" | out-file -filepath $LogFile -append
												(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
												"Home Folder does not exist in expected location...nothig to move." | out-file -filepath $LogFile -append	
											}
										else
											{
												if($CurrentHomeFolderPath -eq $DisabledFolderPath)
													{
														#	Home folder has already been moved.
													}
												else
													{
														#	Folder exists in expected path location...moving to DisabledUsers folder.
														#	# write-host "home folder found at " $CurrenthomeFolderPath " ...moving to DisabledUsers folder location. "
														Try
															{
																Move-Item $CurrenthomeFolderPath $DisabledFolderPath -ErrorAction Stop
																
																"***************************************************************************" | out-file -filepath $LogFile -append
																(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
																"Home Folder moved from " + $CurrenthomeFolderPath + " to " + $DisabledFolderPath | out-file -filepath $LogFile -append
																#	Update value in bullTCLextension2 with $HomeFolderPath
																Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension2=$DisabledFolderPath}
																"Updated bullTCLextension2: " + $DisabledFolderPath | out-file -filepath $LogFile -append
															}
														Catch
															{
																$ErrMsg = "Error. Unable to move folder: $_"	
																$ErrMsg | out-file -filepath $LogFile -append
															}
													}
											}
					
										#------------------------------------------
										#region Disable Staff Office 365 Mailbox
										#	---------------------------------------
										#	Check if we need to disable a mailbox.
										#	---------------------------------------
										if($StaffMailSolution -eq "Office365")
											{
												$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
												
												$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

												if($Office365UPN_Match -eq $null)
													{
														#	This is null so user does not have a mailbox.
													}
												else	
													{
														#	Match found... disbaling mailbox.
														Set-MsolUser -UserPrincipalName $365UPN -BlockCredential $true	#	Disables mailbox access.
														"Office 365 mailbox disabled." | out-file -filepath $LogFile -append
													}
											
											}

										#------------------------------------------
										#endregion Disable Staff Office 365 Mailbox
										#------------------------------------------
										
										if ($ADUserProps.Enabled -eq $true)
											{
												#	Disable Active Directory Account.
												Disable-ADAccount -Identity $External_ID_Match	
											}
										if($DN -eq $DisabledUsersOU)
											{
												#	User is already in the DisabledUsers OU.
											}
										else
											{
												#	# write-host ("Account in wrong OU... (" + $DN + ") moving to correct one...")
												Move-ADObject $DN -TargetPath $DisabledUsersOU
												"***************************************************************************" | out-file -filepath $LogFile -append
												(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
												"Moving User Account to DisabledUsers OU (" + $DisabledUsersOU + ")... for " + $External_ID_Match | out-file -filepath $LogFile -append
											}	
									
										if(($CurrentDate - $DateOfLeaving).Days -gt $RetentionPeriod)
											{
												#	# write-host "delete user account"
												Remove-ADUser -Identity $External_ID_Match -Confirm:$false
												#	# write-host "delete home folder"
												Remove-Item $CurrentHomeFolderPath -Recurse
												
												#------------------------------------------
												#region Delete Staff Office 365 Mailbox
												#	---------------------------------------
												#	Check if we need to delete a mailbox.
												#	---------------------------------------
												if($StaffMailSolution -eq "Office365")
													{
														$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
														
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

														if($Office365UPN_Match -eq $null)
															{
																#	This is null so user does not have a mailbox.	
															}
														else	
															{
																#	Match found... delete mailbox.
																Remove-MsolUser -UserPrincipalName $365UPN -Force
																"Office 365 mailbox deleted." | out-file -filepath $LogFile -append
															}
													}

												#------------------------------------------
												#endregion Delete Staff Office 365 Mailbox
												#------------------------------------------
											}
									}
							}
					}
			}
	}
#----------------------------------------------
#endregion Read Staff Leavers CSV file
#----------------------------------------------	

#----------------------------------------------
#region Read Pupil Leavers CSV file
#----------------------------------------------
#	Test if CSV file exists.
if(!(Test-Path -Path ($CSVExtractLocation + "\AD Pupils Leavers.csv")))
	{
		#	CSV file does not exist...
	}
else
	{	
		#	Read 'AD Pupils Leavers.csv' - this file contains all students that have left the school.
		$StudentLeavers=Import-Csv ($CSVExtractLocation + "\AD Pupils Leavers.csv")
		## write-host "Checking Pupil Leavers"
		foreach($UD in $StudentLeavers) 
			{
				$External_ID = $UD.External_ID	#bullTCLextension1
				## write-host "External ID = " $External_ID
				if($External_ID -eq $null)
					{
						#	CSV file is empty so do nothing.
					}
				elseif($External_ID -eq"")
					{
						#	"empty string"
					}
				else
					{		
						$DateOfLeaving = Get-Date $UD.DateOfLeaving
						$DateOfLeaving = Get-Date $DateOfLeaving
				
						if($DateOfLeaving -gt $CurrentDate)
							{
									## write-host "not yet time to do anything 1"
							}
						elseif($DateOfLeaving -match $CurrentDate)
							{
									## write-host "not yet time to do anything 2"
							}
						else
							{
								#	Disable user and move account & home folder to 'DisabledUser' OU/folder location.
								$External_ID_Match = (Get-ADUser -Filter {bullTCLextension1 -eq $External_ID} -SearchBase $Domain_LDAP_Path -Properties SamAccountName).SamAccountName
								
								if ($External_ID_Match -eq $Null)
									{
										#	User account has already been removed from Active Directory.
									}
								else
									{
										$ADUserProps = (Get-ADUser -Identity $External_ID_Match -Properties bullTCLextension2)
										$DN = $(Get-ADUser -Identity $External_ID_Match).distinguishedName
										$CurrentHomeFolderPath = $ADUserProps.Get_Item("bullTCLextension2")	#full path to user home folder from bullTCLextension2.
										$DisabledFolderPath = ("\\" + $StudentHomeServer + "\" + $StudentHomeFolderRoot + "\DisabledUsers\" + $External_ID_Match)
		
										if(!(Test-Path -Path $CurrentHomeFolderPath))
											{
												#	# write-host "home folder does not exist in expected location...."
												"***************************************************************************" | out-file -filepath $LogFile -append
												(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
												"Home Folder does not exist in expected location...nothig to move." | out-file -filepath $LogFile -append	
											}
										else
											{
												if($CurrentHomeFolderPath -eq $DisabledFolderPath)
													{
														#	Home folder has already been moved.
													}
												else
													{
														#	Folder exists in expected path location...moving to DisabledUsers folder.
														#	# write-host "home folder found at " $CurrenthomeFolderPath " ...moving to DisabledUsers folder location. "
														Try
															{
																Move-Item $CurrenthomeFolderPath $DisabledFolderPath -ErrorAction Stop
																
																"***************************************************************************" | out-file -filepath $LogFile -append
																(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
																"Home Folder moved from " + $CurrenthomeFolderPath + " to " + $DisabledFolderPath | out-file -filepath $LogFile -append
																#	Update value in bullTCLextension2 with $HomeFolderPath
																Set-ADUser -Identity $External_ID_Match -Replace @{bullTCLextension2=$DisabledFolderPath}
																"Updated bullTCLextension2: " + $DisabledFolderPath | out-file -filepath $LogFile -append
															}
														Catch
															{
																$ErrMsg = "Error. Unable to move folder: $_"	
																$ErrMsg | out-file -filepath $LogFile -append
															}
													}
											}
											
										#------------------------------------------
										#region Disable Pupil Office 365 Mailbox
										#	---------------------------------------
										#	Check if we need to disable a mailbox.
										#	---------------------------------------
										if($StudentMailSolution -eq "Office365")
											{
												$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
												
												$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

												if($Office365UPN_Match -eq $null)
													{
														#	This is null so user does not have a mailbox.
													}
												else	
													{
														#	Match found... disbaling mailbox.
														Set-MsolUser -UserPrincipalName $365UPN -BlockCredential $true	#	Disables mailbox access.
														"Office 365 mailbox disabled." | out-file -filepath $LogFile -append
													}
											
											}

										#------------------------------------------
										#endregion Disable Pupil Office 365 Mailbox
										#------------------------------------------
										
										
										if ($ADUserProps.Enabled -eq $true)
											{
												#	Disable Active Directory Account.
												Disable-ADAccount -Identity $External_ID_Match	
											}
										if($DN -eq $DisabledUsersOU)
											{
												#	User is already in the DisabledUsers OU.
											}
										else
											{
												#	# write-host ("Account in wrong OU... (" + $DN + ") moving to correct one...")
												Move-ADObject $DN -TargetPath $DisabledUsersOU
												"***************************************************************************" | out-file -filepath $LogFile -append
												(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
												"Moving User Account to DisabledUsers OU (" + $DisabledUsersOU + ")... for " + $External_ID_Match | out-file -filepath $LogFile -append
											}
											
										if(($CurrentDate - $DateOfLeaving).Days -gt $RetentionPeriod)
											{
												#	# write-host "delete user account"
												Remove-ADUser -Identity $External_ID_Match -Confirm:$false
												#	# write-host "delete home folder"
												Remove-Item $CurrentHomeFolderPath -Recurse
												
												#------------------------------------------
												#region Delete Pupil Office 365 Mailbox
												#	---------------------------------------
												#	Check if we need to delete a mailbox.
												#	---------------------------------------
												if($StudentMailSolution -eq "Office365")
													{
														$365UPN = ($External_ID_Match + "@" + $365MailDomainName)
														
														$Office365UPN_Match = (Get-MsolUser | where {$_.UserPrincipalName -eq $365UPN})

														if($Office365UPN_Match -eq $null)
															{
																#	This is null so user does not have a mailbox.	
															}
														else	
															{
																#	Match found... delete mailbox.
																Remove-MsolUser -UserPrincipalName $365UPN -Force
																"Office 365 mailbox deleted." | out-file -filepath $LogFile -append
															}
													}

												#------------------------------------------
												#endregion Delete Pupil Office 365 Mailbox
												#------------------------------------------
											}
									}
							}	
					}
			}
	}
#----------------------------------------------
#endregion Read Pupil Leavers CSV file
#----------------------------------------------
"---------------------------------------" | out-file -filepath $LogFile -append
(Get-Date -Format HH:mm:ss) | out-file -filepath $LogFile -append
"ADUA - User Automation task completed." | out-file -filepath $LogFile -append
"---------------------------------------" | out-file -filepath $LogFile -append