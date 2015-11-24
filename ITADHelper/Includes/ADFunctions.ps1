# Script manifest
#----------------------------------------------------------------------------
# Version number of Script
# ScriptVersion = '0.6'

# Author of this module
# Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
# Description = 'AD GUI event and function code'
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
function AD_displayErrorText($outputText)
{
	$currentSelectionColor = $AD_richtextboxOutput.SelectionColor;
	$AD_richtextboxOutput.SelectionStart = $AD_richtextboxOutput.TextLength
	$AD_richtextboxOutput.SelectionLength = 0
	$AD_richtextboxOutput.SelectionColor = "Red";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$AD_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$AD_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function AD_displayWarningText($outputText)
{
	$currentSelectionColor = $AD_richtextboxOutput.SelectionColor;
	$AD_richtextboxOutput.SelectionStart = $AD_richtextboxOutput.TextLength
	$AD_richtextboxOutput.SelectionLength = 0
	$AD_richtextboxOutput.SelectionColor = "MediumVioletRed";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$AD_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$AD_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function AD_displayOutputText($outputText)
{
	$currentSelectionColor = $rAD_ichtextboxOutput.SelectionColor;
	if (($OutputText -like "*failed*") -or ($OutputText -like "*error*"))
	{
		$AD_richtextboxOutput.SelectionStart = $AD_richtextboxOutput.TextLength
		$AD_richtextboxOutput.SelectionLength = 0
		$AD_richtextboxOutput.SelectionColor = "MediumVioletRed";
	}
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$AD_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	if (($outputText -like "*failed*") -or ($outputText -like "*error*"))
	{
		# Restore selection color
		$AD_richtextboxOutput.SelectionColor = $currentSelectionColor;
	}
}

#------------------------------------------------------------
function AD_runCommand($dynamicFunctionCall)
{
	if ($dynamicFunctionCall -like "*ForUser")
	{
		if ($AD_textboxUser.Text -ne "")
		{
			AD_displayOutputText "---------------------------------------------------------------"
			$userNameInput = $AD_textboxUser.Text
			$domainChoise = $AD_listboxDomains.SelectedItem
			if (isUserNameInAD $userNameInput $domainChoise)
			{
				AD_displayOutputText "Starting $dynamicFunctionCall on $userNameInput"
				$returnText = &"$dynamicFunctionCall" $userNameInput $domainChoise
				foreach ($textOut in $returnText)
				{
					AD_displayOutputText $textOut
				}
				AD_displayOutputText "Completed $dynamicFunctionCall on $userNameInput"
			}
			else
			{
				AD_displayErrorText "$userNameInput does not exist in domain: $domainChoise"
			}
		}
		else 
		{
			[System.Windows.Forms.MessageBox]::Show("Please enter a username before trying to run a function")
		}
	}
	elseif ($dynamicFunctionCall -like "*CreateButton")
	{
		if ($AD_textboxUser.Text -ne "")
		{
			AD_displayOutputText "---------------------------------------------------------------"
			$userNameInput = $AD_textboxUser.Text
			$domainChoise = $AD_listboxDomains.SelectedItem
			if (isUserNameInAD $userNameInput $domainChoise)
			{
				AD_displayOutputText "Starting $dynamicFunctionCall on $userNameInput"
				$returnText = &"$dynamicFunctionCall" $userNameInput $domainChoise
				foreach ($textOut in $returnText)
				{
					AD_displayOutputText $textOut
				}
				AD_displayOutputText "Completed $dynamicFunctionCall on $userNameInput"
			}
			else
			{
				AD_displayErrorText "$userNameInput does not exist in domain: $domainChoise"
			}
		}
		else 
		{
			[System.Windows.Forms.MessageBox]::Show("Please enter a username before trying to run a function")
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("The function you have selected is not OK for this tool please have someone correct it")
	}
}

#------------------------------------------------------------
function AD_buttonRun_Click
{
	AD_runCommand $AD_comboboxFunction.Text
}

#------------------------------------------------------------
function AD_userStatusDisplay
{
	$AD_textboxUser.Text = $AD_textboxUser.Text.Trim()
	if ($AD_textboxUser.Text -ne "")
	{
		$userNameInput = $AD_textboxUser.Text
		$domainText = $AD_listboxDomains.SelectedItem
		AD_displayOutputText "---------------------------------------------------------------"
		AD_displayOutputText "Gathering information on $userNameInput in $domainText please wait..."
		$currentUser = Get-ADUser -Server $domainText -Identity $userNameInput -Properties *
		$displayName = ($currentUser).DisplayName
		$accountExpirationDate = ($currentUser).AccountExpirationDate
		$cannotChangePassword = ($currentUser).CannotChangePassword
		$description = ($currentUser).Description
		$distinguishedName = ($currentUser).DistinguishedName
		$emailAddress = ($currentUser).EmailAddress
		$homeDirectory = ($currentUser).HomeDirectory
		$lastBadPasswordAttempt = ($currentUser).LastBadPasswordAttempt
		$passwordExpired = ($currentUser).PasswordExpired
		$passwordLastSet = ($currentUser).PasswordLastSet
		$passwordNeverExpires = ($currentUser).PasswordNeverExpires
		$modified = ($currentUser).Modified
		$created = ($currentUser).Created
		$title = ($currentUser).Title
		$office = ($currentUser).physicalDeliveryOfficeName
		AD_displayOutputText "Title: $title"
		AD_displayOutputText "DisplayName: $displayName"
		AD_displayOutputText "Description: $description"
		AD_displayOutputText "Title: $title"
		AD_displayOutputText "Office: $office"
		AD_displayOutputText "DistinguishedName: $distinguishedName"
		AD_displayOutputText "EmailAddress: $emailAddress"
		AD_displayOutputText "CannotChangePassword: $cannotChangePassword" 
		if ($passwordLastSet) { AD_displayOutputText "PasswordLastSet: $passwordLastSet" }
		else {AD_displayOutputText "PasswordLastSet: Temp password set with change on logon"}
		AD_displayOutputText "LastBadPasswordAttempt: $lastBadPasswordAttempt"
		$PasswordExpires = ""
		if (($passwordExpired -ne $True) -and ($passwordNeverExpires -ne $True))
		{
			$TimeUntilPassExpire = (Get-ADUser -Server $domainText -Identity $userNameInput -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
			$DaysUntilPassExpire = (([datetime]::FromFileTime($TimeUntilPassExpire))-(Get-Date)).Days # Converts from "Special" Microsoft time to days left
			$PasswordExpires = "PasswordExpires: $DaysUntilPassExpire Days"
		}
		if ($passwordNeverExpires -eq $False)
		{
			AD_displayOutputText "PasswordExpired: $passwordExpired || $PasswordExpires"
		}
		AD_displayOutputText "PasswordNeverExpires: $passwordNeverExpires"
		if ($accountExpirationDate) 
		{
			AD_displayOutputText "ExpireDate: $accountExpirationDate" 
		}
		AD_displayOutputText "HomeDirectory: $HomeDirectory"
		AD_displayOutputText "Created: $created || Modified: $modified"
		if ((isUserNameInADLocked $userNameInput $domainText) -eq $true)
		{
			AD_displayErrorText "$userNameInput is Locked in AD: $domainText" 
		}
		if ((isUserNameInADEnabled $userNameInput $domainText) -eq $false)
		{
			AD_displayErrorText "$userNameInput is Disabled in AD: $domainText"
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a userame before status")
	}
}

#------------------------------------------------------------
function AD_buttonSearchStatus_Click
{
	$clickTime = Get-Date
	$secondsSinceLastClick = (New-TimeSpan -Start $script:StartTime1 -End $clickTime).TotalSeconds
	if ($secondsSinceLastClick -ge 3)
	{
		$script:StartTime1 = Get-Date
		$AD_textboxUser.Text = $AD_textboxUser.Text.Trim()
		$AD_textboxDomain.Text = "";
		$searchText = $AD_textboxUser.Text
		$searchDomain = $AD_listboxDomains.SelectedItem
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
					$samMatch = Get-ADUser -Server $searchDomain -Identity $searchText -Properties Name,DisplayName,SamAccountName,Description,physicalDeliveryOfficeName,mail
					if ($samMatch -eq $null)
					{
						$usersMatch = Get-ADUser -Server $searchDomain -Properties Name,DisplayName,SamAccountName,Description,physicalDeliveryOfficeName,mail -Filter {(DisplayName -like $searchText) -or (SamAccountName -like $searchText) -or (Description -like $searchText) -or (physicalDeliveryOfficeName -like $searchText) -or (mail -like $searchText)}
						if ($usersMatch -ne $null)
						{
							$returnUser = SearchForADUserAndReturn $usersMatch
							if ($returnUser -ne $null)
							{
								$AD_textboxUser.Text = $returnUser
								$AD_textboxDomain.Text = $searchDomain
								AD_userStatusDisplay
							}
						}
						else
						{
							AD_displayErrorText "No user named $searchText has been found in $searchDomain"
						}
					}
					else
					{
						$AD_textboxDomain.Text = $searchDomain
						AD_userStatusDisplay
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
		AD_displayErrorText "Please wait 3 seconds or more between searches."
	}
}


#############################################################
# Start of GUI event code
#------------------------------------------------------------
function AD_buttonDynamic_Click
{
	$dynamicFunctionCall = $this.Text + "CreateButton"
	AD_RunCommand $dynamicFunctionCall
}

#------------------------------------------------------------
function AD_textboxUser_TextChanged
{
	$AD_textboxUser.Text = $AD_textboxUser.Text.TrimStart() # Prevent beginning with space
	$AD_textboxDomain.Text = ""
}

#------------------------------------------------------------
# Force scroll to end of Richtext box text otherwise user need to do it.
function AD_richtextboxOutput_TextChanged
{
  	$AD_richtextboxOutput.SuspendLayout()
	#$richtextboxOutput.Text = $richtextboxOutput.Text.Replace(' ', [char]0x0020); # trying to prevent spaces in URI links breaking them but it didnt work
	$AD_richtextboxOutput.Select($AD_richtextboxOutput.Text.Length - 1, 0)
	$AD_richtextboxOutput.ScrollToCaret()
  	$AD_richtextboxOutput.ResumeLayout()
}

#------------------------------------------------------------
# link clicked
function AD_richtextboxOutput_LinkClicked
{
	$textLink = $_.LinkText
	AD_displayOutputText "Trying to open link: $textLink"
	Invoke-Item $textLink
}

