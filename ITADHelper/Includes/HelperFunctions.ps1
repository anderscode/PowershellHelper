# Script manifest
#----------------------------------------------------------------------------
# Version number of Script
# ScriptVersion = '0.6'

# Author of this module
# Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
# Description = 'IT Helper script helper functions'
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
#----------------------------------------------------------------------------

function killProcesses($computerName, $currentDomain, $processName)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$Processes = Get-WmiObject -Class Win32_Process -ComputerName $computerNameFull -Filter "name LIKE '$processName'"
	foreach ($process in $processes) 
	{
		if ($process)
		{
			$returnValue = $process.terminate()
			$processID = $process.handle
			$procName = $process.Name
			if($returnValue.returnvalue -eq 0) 
			{
			  "$computerName : $procName `($processID`) successfully terminated"
			}
			else 
			{
				if ($procName) { "$computerName : $procName `($processID`) termination issue, unknown if it closed or not" }
			}
		}
	}
}
#----------------------------------------------------------------------------

function isComputerOnlineAndAccessible($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$isUp = Test-Connection -ComputerName $computerNameFull -Quiet -Count 1
	if($isUp)
	{
		if(Test-Path \\$computerNameFull\c$)
		{
			return $True
		}
		else
		{
			return $False
		}
	}
	else
	{
		return $False
	}
}
#----------------------------------------------------------------------------

function isGivenComputerNameEqualToConnectedPC($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$connectedPCName = (get-wmiobject Win32_ComputerSystem -computername $computerNameFull).Name
	if ($connectedPCName -eq $computerName)
	{
		return $True
	}
	else
	{
		return $connectedPCName
	}
}
#----------------------------------------------------------------------------

function getADUserData($userName, $currentDomain)
{
	try
	{
		if ($userName -ne "") # Here to prevent error passing $userName from causing it to try listing whole AD
		{
			$userData = Get-ADUser -Server $currentDomain -Identity $userName -Properties *,msDS-UserPasswordExpiryTimeComputed -EA Stop
			return $userData
		}
		else
		{
			throw "Error occured in isUserNameInAD"
		}
	}
	catch
	{
		return $false
	}
}
#----------------------------------------------------------------------------

function isUserNameInAD($userName, $currentDomain)
{
	try
	{
		if ($userName -ne "") # Here to prevent error passing $userName from causing it to try listing whole AD
		{
			Get-ADUser -Server $currentDomain -Identity $userName -Properties Name -EA Stop
			return $true
		}
		else
		{
			throw "Error occured in isUserNameInAD"
		}
	}
	catch
	{
		return $false
	}
}
#----------------------------------------------------------------------------

function isUserNameInADEnabled($userName, $currentDomain)
{
	if ($userName -ne "") # Here to prevent error passing $userName from causing it to try listing whole AD
	{
		$userObject = Get-ADUser -Server $currentDomain -Identity $userName -Properties Enabled
		If ($userObject.Enabled)
		{
			return $true
		}	
		else
		{
			return $false
		}
	}
	else
	{
		throw "Error occured in isUserNameInADEnabled"
	}
}
#----------------------------------------------------------------------------

function isUserNameInADLocked($userName, $currentDomain)
{
	if ($userName -ne "") # Here to prevent error passing $userName from causing it to try listing whole AD
	{
		$userObject = Get-ADUser -Server $currentDomain -Identity $userName -Properties LockedOut
		If ($userObject.LockedOut)
		{
			return $true
		}	
		else
		{
			return $false
		}
	}
	else
	{
		throw "Error occured in isUserNameInADLocked"
	}
}
#----------------------------------------------------------------------------

function isComputerNameInAD($computerName, $currentDomain)
{
	try
	{
		if ($computerName -ne "") # Here to prevent error passing $computerName from causing it to try listing whole AD
		{
			Get-ADComputer -Server $currentDomain $computerName -Properties Name -EA Stop
			return $true
		}
		else
		{
			throw "Error occured in isComputerNameInAD"
		}
	}
	catch
	{
		return $false		
	}
}
#----------------------------------------------------------------------------

function isComputerNameInADEnabled($computerName, $currentDomain)
{
	if ($computerName -ne "") # Here to prevent error passing $computerName from causing it to try listing whole AD
	{
		$computerObject = Get-ADComputer -Server $currentDomain $computerName -Properties Enabled
		If ($computerObject.Enabled)
		{
			return $true
		}	
		else
		{
			return $false
		}
	}
	else
	{
		throw "Error occured in isComputerNameInADEnabled"
	}
}
#----------------------------------------------------------------------------

function isWMIWorkingOnPC($computerName, $currentDomain)
{
	try
	{
		$computerNameFull = $computerName + "." + $currentDomain
		$wmi = Get-WmiObject -class "Win32_Process" -namespace "root\cimv2" -computername $computerNameFull -EA Stop
		return $true
	}
	catch
	{
		return $false
	}
}
#----------------------------------------------------------------------------

function returnComputerInfo($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$infoComputerSystem = get-wmiobject Win32_ComputerSystem -computername $computerNameFull
	if ($infoComputerSystem["UserName"])
	{
		$user = $infoComputerSystem["UserName"].SubString($infoComputerSystem["UserName"].IndexOf("\")+1)
	}
	else
	{
		$user = "No user logged in locally"
	}
	$model = $infoComputerSystem["Model"]
	$serial = (Get-WmiObject Win32_BIOS -ComputerName $computerNameFull).SerialNumber.ToString();
	$wmiOS = Get-WmiObject -ComputerName $computerNameFull -Class Win32_OperatingSystem;
	$OS = $wmiOS.Name;
	$make = $infoComputerSystem["Manufacturer"]
	
	# Return info text directly to GUI
	"Make: $make"
	"Model: $model"
	"Serial: $serial"
	"OS: $OS"
	"Logged-on user: $user"
	
	#Check if connected to correct pc.
	$nameMatch = isGivenComputerNameEqualToConnectedPC $computerName  $currentDomain
	if ($nameMatch -ne $true)
	{
		"Error: Connected pc name doesnt match entered name, you are connected to $nameMatch not $computerName"
	}
}
#----------------------------------------------------------------------------

function listErrorInUserProfile($computerName, $userName, $currentDomain)
{
	$objUser = New-Object System.Security.Principal.NTAccount($env:userdnsdomain, $userName)
	$objSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
	$strSID = $objSID.Value
	$computerNameFull = $computerName + "." + $currentDomain
	$localPath = (gwmi -class Win32_UserProfile -ComputerName $computerNameFull -filter "SID='$strSID'" | select localpath).localpath
	if($localPath -isnot [system.array]) # if it return an array the profile most likely has temp profile error
	{
		if ($localPath)
		{
			If ($localPath.IndexOf("TEMP") -ne -1)
			{
				return $localPath
			}
			else 
			{
				# Skipp this check for now as it didnt work repliably, might be issue with older powershell look into later
				return $false
				<#
				$remotepath = $localPath.SubString($localPath.IndexOf("\")+1)
				$remotepath = "\\$computerNameFull\C$\$remotepath"
				if(Test-Path $remotepath)
				{
					$folder = Get-ChildItem -Path $remotepath
					$AccessStatus = (Get-Acl $folder.FullName).AccessToString | findstr "$userName"
					if ($AccessStatus)
					{
						return $false
					}
					else
					{
						return "$userName might have issues with right to his/her profile folder"
					}
				}
				else
				{
					return "$userName Unable to remotely access users profile folder"
				}
				#>
			}
		}
		else
		{
			"Error: Unable to get any profilepath for $userName on $computerName"
		}
	}
	else
	{
		$tempProfile = $false
		foreach ($tempPathCheck in $localPath)
		{
			If ($tempPathCheck.IndexOf("TEMP") -ne -1)
			{
				$tempProfile = $true;
			}
		}
		if ($tempProfile)
		{
			return "$userName might have TEMP profile issues, please check on this"
		}
		else
		{
			"$userName : Multiple user folders reported please check on this"
			foreach ($profilePath in $localPath)
			{
				"$profilePath"
			}
		}
	}
}
#----------------------------------------------------------------------------

function GetLocalUserPathForPC($computerName, $userName, $currentDomain)
{
	$arrayAccessList = @()
	$objUser = New-Object System.Security.Principal.NTAccount($env:userdnsdomain, $userName)
	$objSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
	$strSID = $objSID.Value
	$computerNameFull = $computerName + "." + $currentDomain
	$localpath = (gwmi -class Win32_UserProfile -ComputerName $computerNameFull -filter "SID='$strSID'" | select localpath).localpath
	if($localpath -isnot [system.array]) # Only return if there is only one path specified
	{
		return $localpath
	}
	else
	{
		# listErrorInProfileForUser handles multiple profilepaths so only return false here
		return $false
	}
}
#----------------------------------------------------------------------------

function GetRemoteUserPathForPC($computerName, $userName, $currentDomain)
{
	$arrayAccessList = @()
	$objUser = New-Object System.Security.Principal.NTAccount($env:userdnsdomain, $userName)
	$objSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
	$strSID = $objSID.Value
	$computerNameFull = $computerName + "." + $currentDomain
	$localPath = (gwmi -class Win32_UserProfile -ComputerName $computerNameFull -filter "SID='$strSID'" | select localpath).localpath
	if ($localPath)
	{
		if($localPath -isnot [system.array])
		{
			$localPath = $localPath.SubString($localPath.IndexOf("\")+1)
		}
		else
		{
			# listErrorInProfileForUser handles multiple profilepaths so only return false here
			return $false
		}

		$remotePath = "\\$computerNameFull\C$\$localPath"
		if(Test-Path $remotePath)
		{
			return $remotePath
		}
		else
		{
			return $false
		}
	}
	else
	{
		return $false
	}
}
#----------------------------------------------------------------------------

function listOrphanedUserProfilesOnPC($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$arrayOrphanedProfileList = @()
	$userList = gwmi win32_userprofile -ComputerName $computerNameFull
	foreach ($user in $userList)
	{
		if ($user)
		{
			$uSID = $user.SID
			if ($uSID)
			{
				if ($uSID -like "S-1-5-21-*") # Only list Domain accounts
				{
					if (($uSID -like "S-1-5-21-*-50?") -ne $true) # Administrator,Guest... SIDs are ignored when checking
					{
						if (isUserNameInAD $uSID $currentDomain)
						{
							# profile ok ignore
						}
						else
						{
							$profilePath = $user.Localpath
							$arrayOrphanedProfileList = $arrayOrphanedProfileList + $profilePath
						}
					}
				}
			}
		}
	}
	return $arrayOrphanedProfileList
}
#----------------------------------------------------------------------------

function listUserNamesOnPC($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$arrayUserList = @()
	$userList = gwmi win32_userprofile -ComputerName $computerNameFull
	foreach ($user in $userList)
	{
		if ($user)
		{
			$uSID = $user.SID
			if ($uSID)
			{
				if ($uSID -like "S-1-5-21-*") # Only list Domain accounts
				{
					if (($uSID -like "S-1-5-21-*-50?") -ne $true) # Administrator,Guest... SIDs are ignored when checking
					{
						if (isUserNameInAD $uSID $currentDomain)
						{
							$objSID = New-Object System.Security.Principal.SecurityIdentifier("$uSID")
							$objUser = $objSID.Translate([System.Security.Principal.NTAccount])
							$userRAW = $objUser.Value
							$user = $userRAW.SubString($userRAW.IndexOf("\")+1)
							$arrayUserList = $arrayUserList + $user
						}
						else
						{
							# Listing orphaned profiles is handled separately by ListOrphanedUserProfilesOnPC function
						}
					}
				}
			}
		}
	}
	return $arrayUserList
}
#----------------------------------------------------------------------------

function returnLoggedinUserOnPC($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$infoComputerSystem = @()
	$infoComputerSystem = get-wmiobject Win32_ComputerSystem -computername $computerNameFull -property username
	if ($infoComputerSystem["UserName"] -ne $null)
	{
		$usernCurrent = $infoComputerSystem["UserName"].SubString($infoComputerSystem["UserName"].IndexOf("\")+1)
		return $usernCurrent
	}
	else
	{
		return $false
	}
}
#----------------------------------------------------------------------------
