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
function 365_displayErrorText($outputText)
{
	$currentSelectionColor = $365_richtextboxOutput.SelectionColor;
	$365_richtextboxOutput.SelectionStart = $365_richtextboxOutput.TextLength
	$365_richtextboxOutput.SelectionLength = 0
	$365_richtextboxOutput.SelectionColor = "Red";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$365_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$365_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function 365_displayWarningText($outputText)
{
	$currentSelectionColor = $365_richtextboxOutput.SelectionColor;
	$365_richtextboxOutput.SelectionStart = $365_richtextboxOutput.TextLength
	$365_richtextboxOutput.SelectionLength = 0
	$365_richtextboxOutput.SelectionColor = "MediumVioletRed";
	
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$365_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	# Restore selection color
	$365_richtextboxOutput.SelectionColor = $currentSelectionColor;
}

#------------------------------------------------------------
function 365_displayOutputText($outputText)
{
	$currentSelectionColor = $r365_ichtextboxOutput.SelectionColor;
	if (($OutputText -like "*failed*") -or ($OutputText -like "*error*"))
	{
		$365_richtextboxOutput.SelectionStart = $365_richtextboxOutput.TextLength
		$365_richtextboxOutput.SelectionLength = 0
		$365_richtextboxOutput.SelectionColor = "MediumVioletRed";
	}
	$dateTimeString = [System.DateTime]::Now.ToString("yyyy.MM.dd hh:mm:ss")
	$365_richtextboxOutput.AppendText($dateTimeString + " - " + $outputText + "`n")
	
	if (($outputText -like "*failed*") -or ($outputText -like "*error*"))
	{
		# Restore selection color
		$365_richtextboxOutput.SelectionColor = $currentSelectionColor;
	}
}

#------------------------------------------------------------
function 365_runCommand($dynamicFunctionCall)
{
	if ($dynamicFunctionCall -like "*ForUser")
	{
		if ($365_textboxUser.Text -ne "")
		{
			365_displayOutputText "---------------------------------------------------------------"
			$userNameInput = $365_textboxUser.Text
			$domainChoise = $365_listboxDomains.SelectedItem
			if (isUserNameInAD $userNameInput $domainChoise)
			{
				365_displayOutputText "Starting $dynamicFunctionCall on $userNameInput"
				$returnText = &"$dynamicFunctionCall" $userNameInput $domainChoise
				foreach ($textOut in $returnText)
				{
					365_displayOutputText $textOut
				}
				365_displayOutputText "Completed $dynamicFunctionCall on $userNameInput"
			}
			else
			{
				365_displayErrorText "$userNameInput does not exist in domain: $domainChoise"
			}
		}
		else 
		{
			[System.Windows.Forms.MessageBox]::Show("Please enter a username before trying to run a function")
		}
	}
	elseif ($dynamicFunctionCall -like "*CreateButton")
	{
		if ($365_textboxUser.Text -ne "")
		{
			365_displayOutputText "---------------------------------------------------------------"
			$userNameInput = $365_textboxUser.Text
			$domainChoise = $365_listboxDomains.SelectedItem
			if (isUserNameInAD $userNameInput $domainChoise)
			{
				365_displayOutputText "Starting $dynamicFunctionCall on $userNameInput"
				$returnText = &"$dynamicFunctionCall" $userNameInput $domainChoise
				foreach ($textOut in $returnText)
				{
					365_displayOutputText $textOut
				}
				365_displayOutputText "Completed $dynamicFunctionCall on $userNameInput"
			}
			else
			{
				365_displayErrorText "$userNameInput does not exist in domain: $domainChoise"
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
function 365_buttonRun_Click
{
	365_runCommand $365_comboboxFunction.Text
}

#------------------------------------------------------------
function 365_userStatusDisplay
{
	$365_textboxUser.Text = $365_textboxUser.Text.Trim()
	if ($365_textboxUser.Text -ne "")
	{
		$userNameInput = $365_textboxUser.Text
		$domainText = $365_listboxDomains.SelectedItem
		365_displayOutputText "---------------------------------------------------------------"
		365_displayOutputText "Gathering information on $userNameInput in $domainText please wait..."
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
		365_displayOutputText "Title: $title"
		365_displayOutputText "DisplayName: $displayName"
		365_displayOutputText "Description: $description"
		365_displayOutputText "Title: $title"
		365_displayOutputText "Office: $office"
		365_displayOutputText "DistinguishedName: $distinguishedName"
		365_displayOutputText "EmailAddress: $emailAddress"
		365_displayOutputText "CannotChangePassword: $cannotChangePassword" 
		if ($passwordLastSet) { 365_displayOutputText "PasswordLastSet: $passwordLastSet" }
		else {365_displayOutputText "PasswordLastSet: Temp password set with change on logon"}
		365_displayOutputText "LastBadPasswordAttempt: $lastBadPasswordAttempt"
		$PasswordExpires = ""
		if (($passwordExpired -ne $True) -and ($passwordNeverExpires -ne $True))
		{
			$TimeUntilPassExpire = (Get-ADUser -Server $domainText -Identity $userNameInput -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
			$DaysUntilPassExpire = (([datetime]::FromFileTime($TimeUntilPassExpire))-(Get-Date)).Days # Converts from "Special" Microsoft time to days left
			$PasswordExpires = "PasswordExpires: $DaysUntilPassExpire Days"
		}
		if ($passwordNeverExpires -eq $False)
		{
			365_displayOutputText "PasswordExpired: $passwordExpired || $PasswordExpires"
		}
		365_displayOutputText "PasswordNeverExpires: $passwordNeverExpires"
		if ($accountExpirationDate) 
		{
			365_displayOutputText "ExpireDate: $accountExpirationDate" 
		}
		365_displayOutputText "HomeDirectory: $HomeDirectory"
		365_displayOutputText "Created: $created || Modified: $modified"
		if ((isUserNameInADLocked $userNameInput $domainText) -eq $true)
		{
			365_displayErrorText "$userNameInput is Locked in AD: $domainText" 
		}
		if ((isUserNameInADEnabled $userNameInput $domainText) -eq $false)
		{
			365_displayErrorText "$userNameInput is Disabled in AD: $domainText"
		}
	}
	else
	{
		[System.Windows.Forms.MessageBox]::Show("Please enter a userame before status")
	}
}

#------------------------------------------------------------
function 365_buttonSearchStatus_Click
{
	$clickTime = Get-Date
	$secondsSinceLastClick = (New-TimeSpan -Start $script:StartTime1 -End $clickTime).TotalSeconds
	if ($secondsSinceLastClick -ge 3)
	{
		$script:StartTime1 = Get-Date
		$365_textboxUser.Text = $365_textboxUser.Text.Trim()
		$365_textboxDomain.Text = "";
		$searchText = $365_textboxUser.Text
		$searchDomain = $365_listboxDomains.SelectedItem
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
								$365_textboxUser.Text = $returnUser
								$365_textboxDomain.Text = $searchDomain
								365_userStatusDisplay
							}
						}
						else
						{
							365_displayErrorText "No user named $searchText has been found in $searchDomain"
						}
					}
					else
					{
						$365_textboxDomain.Text = $searchDomain
						365_userStatusDisplay
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
		365_displayErrorText "Please wait 3 seconds or more between searches."
	}
}


#############################################################
# Start of GUI event code
#------------------------------------------------------------
function 365_buttonDynamic_Click
{
	$dynamicFunctionCall = $this.Text + "CreateButton"
	365_RunCommand $dynamicFunctionCall
}

#------------------------------------------------------------
function 365_textboxUser_TextChanged
{
	$365_textboxUser.Text = $365_textboxUser.Text.TrimStart() # Prevent beginning with space
	$365_textboxDomain.Text = ""
}

#------------------------------------------------------------
# Force scroll to end of Richtext box text otherwise user need to do it.
function 365_richtextboxOutput_TextChanged
{
  	$365_richtextboxOutput.SuspendLayout()
	#$richtextboxOutput.Text = $richtextboxOutput.Text.Replace(' ', [char]0x0020); # trying to prevent spaces in URI links breaking them but it didnt work
	$365_richtextboxOutput.Select($365_richtextboxOutput.Text.Length - 1, 0)
	$365_richtextboxOutput.ScrollToCaret()
  	$365_richtextboxOutput.ResumeLayout()
}

#------------------------------------------------------------
# link clicked
function 365_richtextboxOutput_LinkClicked
{
	$textLink = $_.LinkText
	365_displayOutputText "Trying to open link: $textLink"
	Invoke-Item $textLink
}

