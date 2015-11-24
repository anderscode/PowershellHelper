# Script manifest
#----------------------------------------------------------------------------
# Version number of Script
# ScriptVersion = '0.6'

# Author of this module
# Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
# Description = 'IT GUI event and function code'
#----------------------------------------------------------------------------
<#
The MIT License (MIT)

Copyright (c)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
#>
#############################################################
#############################################################
# Start of GUI function code
#------------------------------------------------------------
function IT_displayErrorText($outputText)
{
	$currentSelectionColor = $IT_richtextboxOutput.SelectionColor;
	$IT_richtextboxOutput.SelectionStart = $IT_richtextboxOutput.TextLength
	$IT_richtextboxOutput.SelectionLength = 0
	$IT_richtextboxOutput.SelectionColor = "Red";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$IT_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$IT_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function IT_displayWarningText($outputText)
{
	$currentSelectionColor = $IT_richtextboxOutput.SelectionColor;
	$IT_richtextboxOutput.SelectionStart = $IT_richtextboxOutput.TextLength
	$IT_richtextboxOutput.SelectionLength = 0
	$IT_richtextboxOutput.SelectionColor = "MediumVioletRed";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$IT_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$IT_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function IT_displayOutputText($outputText)
{
	$currentSelectionColor = $IT_richtextboxOutput.SelectionColor;
	if (($outputText -like "*failed*") -or ($outputText -like "*error*"))
	{
		$IT_richtextboxOutput.SelectionStart = $IT_richtextboxOutput.TextLength
		$IT_richtextboxOutput.SelectionLength = 0
		$IT_richtextboxOutput.SelectionColor = "MediumVioletRed";
	}
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$IT_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	if (($outputText -like "*failed*") -or ($outputText -like "*error*"))
	{
		# Restore selection color
		$IT_richtextboxOutput.SelectionColor = $currentSelectionColor;
	}
}

#------------------------------------------------------------
function IT_runCommand($dynamicFunctionCall)
{
	if ($IT_textboxComputer.Text -ne "")
	{
		if ($dynamicFunctionCall -ne "")
		{
			if ($dynamicFunctionCall -like "*User")
			{
				if ($IT_listboxUsers.SelectedIndex -ne -1)
				{
					Invoke-WmiMethod -Path win32_process -Name create -ArgumentList "cmd /c ipconfig /flushdns"
					IT_displayOutputText "---------------------------------------------------------------"
					# Call selected function with ComputerName and User arguments
					$computerNameInput = $IT_textboxComputer.Text
					$userNameInput = $IT_listboxUsers.SelectedItem
					$currentDomain = $script:currentDomainByLoggedInUser
					if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
					{
						if (isWMIWorkingOnPC $computerNameInput $currentDomain)
						{
							if (isComputerNameInAD $computerNameInput $currentDomain)
							{
								if(isComputerNameInADEnabled $computerNameInput $currentDomain)
								{
									# Running simple profile check to make sure it works ok, no temp profile etc
									$profileErrorCheck = listErrorInUserProfile $computerNameInput $UserNameInput $currentDomain
									if ($profileErrorCheck -eq $False)
									{
										
										IT_displayOutputText "Running $dynamicFunctionCall on $computerNameInput, $UserNameInput"
										$return = &"$dynamicFunctionCall" $computerNameInput $UserNameInput $currentDomain
										foreach ($textOut in $return)
										{
											IT_displayOutputText $textOut
										}
										IT_displayOutputText "Completed $dynamicFunctionCall on $computerNameInput, $UserNameInput"
									}
									else
									{
										foreach ($textOut in $profileErrorCheck)
										{
											IT_displayErrorText $textOut
										}
									}
								}
								else
								{
									IT_displayErrorText "$computerNameInput is NOT Enabled in AD please fix this first"
								}
							}
							else
							{
								IT_displayErrorText "$computerNameInput is NOT in AD please fix this first"
							}
						}
						else
						{
							IT_displayErrorText "WMI ERROR on $computerNameInput"
						}
					}
					else
					{
						IT_displayWarningText "$computerNameInput is Offline"
					}
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Please select a user before running this function")
				}
			}
			elseif (($dynamicFunctionCall -like "*Computer") -or ($dynamicFunctionCall -like "*CreateButton"))
			{
				IT_displayOutputText "---------------------------------------------------------------"
				# Call selected function with ComputerName argument only
				$computerNameInput = $IT_textboxComputer.Text
				$currentDomain = $script:currentDomainByLoggedInUser
				if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
				{
					if (isWMIWorkingOnPC $computerNameInput $currentDomain)
					{
						if (isComputerNameInAD $computerNameInput $currentDomain)
						{
							if(isComputerNameInADEnabled $computerNameInput $currentDomain)
							{
								IT_displayOutputText "Running $dynamicFunctionCall on $computerNameInput"
								$return = &"$dynamicFunctionCall" $computerNameInput $currentDomain
								foreach ($textOut in $return)
								{
									IT_displayOutputText $textOut
								}
								IT_displayOutputText "Completed $dynamicFunctionCall on $computerNameInput"
							}
							else
							{
								IT_displayErrorText "$computerNameInput is NOT Enabled in AD please fix this first"
							}
						}
						else
						{
							IT_displayErrorText "$computerNameInput is NOT in AD please fix this first"
						}
					}
					else
					{
						IT_displayErrorText "WMI ERROR on $computerNameInput"
					}
				}
				else
				{
					IT_displayWarningText "$computerNameInput is Offline"
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("The function you have selected is not OK for this tool please have someone correct it")
			}
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show("Please select a function to run")
		}
	}
	else 
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a computername before trying to run a function")
	}
}

#------------------------------------------------------------
function IT_buttonRun_Click
{
	IT_runCommand $IT_comboboxFunction.Text
}

#------------------------------------------------------------
function IT_computerStatusDisplay
{
	if ($IT_textboxComputer.Text -ne "")
	{
		IT_displayOutputText "---------------------------------------------------------------"
		$computerNameInput = $IT_textboxComputer.Text
		$currentDomain = $script:currentDomainByLoggedInUser
		IT_displayOutputText "Gathering information on $computerNameInput please wait..."
		if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
		{
			if (isWMIWorkingOnPC $computerNameInput $currentDomain)
			{
				IT_displayOutputText "$computerNameInput is Online"
				$status = returnComputerInfo $computerNameInput $currentDomain
				foreach ($textOut in $status)
				{
					IT_displayOutputText $textOut
				}
			}
			else
			{
				IT_displayErrorText "WMI ERROR on $computerNameInput"
			}
			if ((isComputerNameInADEnabled $computerNameInput $currentDomain) -eq $False)
			{
				IT_displayErrorText "$computerNameInput is Disabled in AD"
			}
		}
		else
		{
			IT_displayWarningText "$computerNameInput is Offline"
			if ((isComputerNameInADEnabled $computerNameInput $currentDomain) -eq $False)
			{
				IT_displayErrorText "$computerNameInput is Disabled in AD"
			}
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a computername before trying to get computer status")
	}
}

#------------------------------------------------------------
function IT_buttonSearchStatus_Click
{
	$clickTime = Get-Date
	$secondsSinceLastClick = (New-TimeSpan -Start $script:StartTime1 -End $clickTime).TotalSeconds
	if ($secondsSinceLastClick -ge 3)
	{
		$IT_textboxComputer.Text = $IT_textboxComputer.Text.Trim()
		$searchText = $IT_textboxComputer.Text
		$searchDomain = $script:currentDomainByLoggedInUser
		if ($searchText.Length -le 1)
		{
			[System.Windows.Forms.MessageBox]::Show("Please enter more then 2 charaters to search for")
		}
		else
		{
			if ($searchText -ne "")
			{
				if (([regex]::matches($searchText,"\*")).count -lt $searchText.length)
				{
					$samMatch = Get-ADComputer -Server $searchDomain -Identity $searchText -Properties Name,DisplayName,DNSHostName
					if ($samMatch -eq $null)
					{
						$computerMatch = Get-ADComputer -Server $searchDomain -Properties Name,DisplayName,DNSHostName -Filter {(Name -like $searchText) -or (DisplayName -like $searchText) -or (DNSHostName -like $searchText)}
						if ($computerMatch -ne $null)
						{
							$returnComputer = SearchForADComputerAndReturn $computerMatch
							if ($returnComputer -ne $null)
							{
								$IT_textboxComputer.Text = $returnComputer
								IT_computerStatusDisplay
							}
						}
						else
						{
							IT_displayErrorText "No computer named $searchText has been found in $searchDomain"
						}
					}
					else
					{
						IT_computerStatusDisplay
					}
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Please enter atleast one normal charater to search for. (all * not allowed)")
				}
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter name before trying to search.")
			}
		}
	}
	else
	{
		IT_displayErrorText "Please wait 3 seconds or more between searches."
	}
}

#------------------------------------------------------------
function IT_buttonAllUsers_Click
{
	$IT_listboxUsers.Items.Clear()
	if ($IT_textboxComputer.Text -ne "")
	{
		$computerNameInput = $IT_textboxComputer.Text
		$currentDomain = $script:currentDomainByLoggedInUser
		if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
		{
			if (isWMIWorkingOnPC $computerNameInput $currentDomain)
			{
				$arrayUserList = ListUserNamesOnPC $computerNameInput $currentDomain
				foreach ($userName in $arrayUserList)
				{
					$IT_listboxUsers.Items.add("$userName")
				}
				
				$arrayOrphanedProfileList = ListOrphanedUserProfilesOnPC $computerNameInput $currentDomain
				if ($arrayOrphanedProfileList)
				{
					IT_displayOutputText "---------------------------------------------------------------"
					IT_displayOutputText "Listing profiles on $computerNameInput where no AD account exists"
					foreach ($profilePath in $arrayOrphanedProfileList)
					{
						IT_displayWarningText "$profilePath doesnt have a matching AD account"
					}
				}
			}
			else
			{
				IT_displayErrorText "WMI ERROR on $computerNameInput"
			}
		}
		else
		{
			IT_displayWarningText "$computerNameInput is Offline"
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a computername before trying to list users")
	}
}

#------------------------------------------------------------
function IT_buttonCurrentUser_Click
{
	$IT_listboxUsers.Items.Clear()
	if ($IT_textboxComputer.Text -ne "")
	{
		IT_displayOutputText "---------------------------------------------------------------"
		$computerNameInput = $IT_textboxComputer.Text
		$currentDomain = $script:currentDomainByLoggedInUser
		if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
		{
			if (isWMIWorkingOnPC $computerNameInput $currentDomain)
			{
				$userName = ReturnLoggedinUserOnPC $computerNameInput $currentDomain
				if ($userName -eq $False)
				{
					[System.Windows.Forms.MessageBox]::Show("No user logged in to $computerNameInput")
				}
				else
				{
					$IT_listboxUsers.Items.add("$userName")
					# Do error check for user
					$profileErrorCheck = listErrorInUserProfile $computerNameInput $userName $currentDomain
					if ($profileErrorCheck -eq $False)
					{
						# No error found, not displayed but leaving space here for it if needed
					}
					else
					{
						foreach ($textOut in $profileErrorCheck)
						{
							IT_displayErrorText $textOut
						}
					}
					# Ordinary user info displayed here, maybe move up to "# No error found" place if its to much
					$localUserPathCheck = GetLocalUserPathForPC $computerNameInput $userName $currentDomain
					$remoteUserPathCheck = GetRemoteUserPathForPC $computerNameInput $userName $currentDomain
					IT_displayOutputText "$userName on $computerNameInput has remote path:"
					foreach ($userPathCheck in $remoteUserPathCheck)
					{
						if(Test-Path $userPathCheck)
						{
							IT_displayOutputText $userPathCheck
						}
						else
						{
							IT_displayOutputText "$userPathCheck (Path doesnt exist)"
						}
					}
					IT_displayOutputText "$userName on $computerNameInput has local path:"
					foreach ($userPathCheck in $localUserPathCheck)
					{
						IT_displayOutputText $userPathCheck
					}
				}
			}
			else
			{
				IT_displayErrorText "WMI ERROR on $computerNameInput"
			}
		}
		else
		{
			IT_displayWarningText "$computerNameInput is Offline"
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a computername before trying to show current user")
	}
}

#------------------------------------------------------------
function IT_buttonUserStatus_Click
{
	if ($IT_textboxComputer.Text -ne "")
	{
		$computerNameInput = $IT_textboxComputer.Text
		$currentDomain = $script:currentDomainByLoggedInUser
		IT_displayOutputText "---------------------------------------------------------------"
		IT_displayOutputText "Gathering user information on $computerNameInput please wait..."
		if (isComputerOnlineAndAccessible $computerNameInput $currentDomain)
		{
			if (isWMIWorkingOnPC $computerNameInput $currentDomain)
			{
				if ($IT_listboxUsers.SelectedIndex -ne -1)
				{
					$userName = $IT_listboxUsers.SelectedItem
					if ($userName -eq $False)
					{
						[System.Windows.Forms.MessageBox]::Show("No user logged in to $computerNameInput")
					}
					else
					{
						# Do error check for user profile
						$profileErrorCheck = listErrorInUserProfile $computerNameInput $userName $currentDomain
						if ($profileErrorCheck -eq $False)
						{
							# No error found, not displayed but leaving space here for it if needed
						}
						else
						{
							foreach ($textOut in $profileErrorCheck)
							{
								IT_displayErrorText $textOut
							}
						}
						# AD check
						if(isUserNameInAD $userName $currentDomain)
						{
							if((isUserNameInADEnabled $userName $currentDomain) -ne $true)
							{
								IT_displayErrorText "$userName found in AD but is Disabled"
							}
						}
						else
						{
							IT_displayWarningText "$userName not found in AD"
						}
						# Ordinary user info displayed here, maybe move up to "# No error found" place if its to much
						$localUserPathCheck = GetLocalUserPathForPC $computerNameInput $userName $currentDomain
						$remoteUserPathCheck = GetRemoteUserPathForPC $computerNameInput $userName $currentDomain
						IT_displayOutputText "$userName on $computerNameInput remote profilepath:"
						foreach ($userPathCheck in $remoteUserPathCheck)
						{
							if(Test-Path $userPathCheck) 
							{
								IT_displayOutputText $userPathCheck
							}
							else
							{
								IT_displayOutputText "$userPathCheck (Path doesnt exist)"
							}
						}
						IT_displayOutputText "$userName on $computerNameInput local profilepath:"
						foreach ($userPathCheck in $localUserPathCheck)
						{
							IT_displayOutputText $userPathCheck
						}
					}
				}
				else
				{
					[System.Windows.Forms.MessageBox]::Show("Please select a username before trying to show status for it")
				}
			}
			else
			{
				IT_displayErrorText "WMI ERROR on $computerNameInput"
			}
		}
		else
		{
			IT_displayWarningText "$computerNameInput is Offline"
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a computername and select a user before trying to show status for it")
	}
}


#############################################################
# Start of GUI event code
#------------------------------------------------------------
function IT_buttonDynamic_Click
{
	$dynamicFunctionCall = $this.Text + "CreateButton"
	IT_RunCommand $dynamicFunctionCall
}

#------------------------------------------------------------
function IT_textboxComputer_TextChanged
{
	$IT_listboxUsers.Items.Clear()
	$IT_textboxComputer.Text = $IT_textboxComputer.Text.Trim()
	$IT_textboxComputer.Select($IT_textboxComputer.Text.Length, 0) # Prevent trim from jumping selection back to start
}

#------------------------------------------------------------
# Force scroll to end of Richtext box text otherwise user need to do it.
function IT_richtextboxOutput_TextChanged
{
  	$IT_richtextboxOutput.SuspendLayout()
	#$richtextboxOutput.Text = $richtextboxOutput.Text.Replace(' ', [char]0x0020); # trying to prevent spaces in URI links to break them but it didnt work
	$IT_richtextboxOutput.Select($IT_richtextboxOutput.Text.Length - 1, 0)
	$IT_richtextboxOutput.ScrollToCaret()
  	$IT_richtextboxOutput.ResumeLayout()
}

#------------------------------------------------------------
# link clicked
function IT_richtextboxOutput_LinkClicked
{
	$textLink = $_.LinkText
	IT_displayOutputText "Trying to open link: $textLink"
	Invoke-Item $textLink
}

