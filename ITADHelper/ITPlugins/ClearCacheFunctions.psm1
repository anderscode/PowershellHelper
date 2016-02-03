# Module manifest
#----------------------------------------------------------------------------
@{
# Version number of this module.
ModuleVersion = '0.6'

# ID used to uniquely identify this module
GUID = '9efe47b5-e002-476e-9ba9-81601aa2001b'

# Author of this module
Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
Description = 'Plugin functions for clearing temp/cache files remotely'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Minimum version of the Windows PowerShell host required by this module
PowerShellHostVersion = '2.0'

# Minimum version of the .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''
}
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

#----------------------------------------------------------------------------
if(!$PSScriptRoot) 
{
	set-variable -name PSScriptRoot -Scope Script
	$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent # Needed incase of Powershell version 2
}
. "$PSScriptRoot\..\Includes\HelperFunctions.ps1" # Inlcude help function in plugin
#----------------------------------------------------------------------------

function out-default
{ 
	$optionalOutput = $_
	out-host
}
#----------------------------------------------------------------------------

function ClearChromeCacheForUser($computerName, $userName, $currentDomain)
{
	$userPath = GetRemoteUserPathForPC $computerName $userName $currentDomain
	$fullPathA = @("$userPath\AppData\Local\Google\Chrome\User Data\Default\Cache\")
	foreach ($fullPath in $fullPathA)
	{
		if ((Test-Path $fullPath))
		{
			$filePathA = Get-ChildItem -Path $fullPath
			foreach ($filePath in $filePathA)
			{
				$FileToRemove = $filePath.FullName
				try
				{
					Remove-Item $FileToRemove -recurse -force -EA Stop | Out-null
				}
				catch
				{
					$ErrorMessage = $_.Exception.Message
					"ERROR: $ErrorMessage"
				}
			}
		}
		else
		{
			"No Chrome cache folder found for user: $userName"
		}
	}
}
#----------------------------------------------------------------------------

function ClearIECacheForUser($computerName, $userName, $currentDomain)
{
	$userPath = GetRemoteUserPathForPC $computerName $userName $currentDomain
	$fullPathA = @("$userPath\AppData\Local\Microsoft\Windows\Temporary Internet Files\")
	foreach ($fullPath in $fullPathA)
	{
		if ((Test-Path $fullPath))
		{
			$filePathA = Get-ChildItem -Path $fullPath
			foreach ($filePath in $filePathA)
			{
				$FileToRemove = $filePath.FullName
				try
				{
					Remove-Item $FileToRemove -recurse -force -EA Stop | Out-null
				}
				catch
				{
					$ErrorMessage = $_.Exception.Message
					"ERROR: $ErrorMessage"
				}
			}
		}
		else
		{
			"No IE Cache found at: $fullPath"
		}
	}
}
#----------------------------------------------------------------------------
	
function ClearTempForUser($computerName, $userName, $currentDomain)
{
	$userPath = GetRemoteUserPathForPC $computerName $userName $currentDomain
	$fullPathA = @("$userPath\AppData\Local\Temp\")
	foreach ($fullPath in $fullPathA)
	{
		if ((Test-Path $fullPath))
		{
			$filePathA = Get-ChildItem -Path $fullPath
			foreach ($filePath in $filePathA)
			{
				$FileToRemove = $filePath.FullName
				try
				{
					Remove-Item $FileToRemove -recurse -force -EA Stop | Out-null
				}
				catch
				{
					$ErrorMessage = $_.Exception.Message
					"ERROR: $ErrorMessage"
				}
			}
		}
		else
		{
			"No temp file folder found at: $fullPath"
		}
	}
}
#----------------------------------------------------------------------------

function ClearFlashJavaCacheForUser($computerName, $userName, $currentDomain)
{
	$userPath = GetRemoteUserPathForPC $computerName $userName $currentDomain
	$fullPathA = @("$userPath\AppData\LocalLow\Sun\Java\Deployment\Cache\")
	$fullPathA += "$userPath\AppData\Roaming\Macromedia\Flash Player\"
	foreach ($fullPath in $fullPathA)
	{
		if ((Test-Path $fullPath))
		{
			$filePathA = Get-ChildItem -Path $fullPath
			foreach ($filePath in $filePathA)
			{
				$FileToRemove = $filePath.FullName
				try
				{
					Remove-Item $FileToRemove -recurse -force -EA Stop | Out-null
				}
				catch
				{
					$ErrorMessage = $_.Exception.Message
					"ERROR: $ErrorMessage"
				}
			}
		}
		else
		{
			"No cache found for path: $fullPath"
		}
	}
}
#----------------------------------------------------------------------------

Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer