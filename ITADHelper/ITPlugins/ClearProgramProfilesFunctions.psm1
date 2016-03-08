# Module manifest
#----------------------------------------------------------------------------
@{
# Version number of this module.
ModuleVersion = '0.6'

# ID used to uniquely identify this module
GUID = '39b17891-1b9f-440f-b9fb-c74ec7f4db12'

# Author of this module
Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
Description = 'Plugin functions for clearing profile/files for specific programs remotely'

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
if(!$PSScriptRoot) 
{
	set-variable -name PSScriptRoot -Scope Script
	$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent # Needed incase of Powershell version 2
}
. "$PSScriptRoot\..\Includes\HelperFunctions.ps1" # Inlcude help function in plugin
#----------------------------------------------------------------------------

function out-default
{ 
	$OptionalOutput = $_
	out-host
}

#----------------------------------------------------------------------------

function KillLotusForComputer($ComputerName, $CurrentDomain)
{
	killProcesses $ComputerName $CurrentDomain "ntaskldr.exe"
	killProcesses $ComputerName $CurrentDomain "ntmulti.exe"
	killProcesses $ComputerName $CurrentDomain "nsd.exe"
	killProcesses $ComputerName $CurrentDomain "notes.exe"
	killProcesses $ComputerName $CurrentDomain "notes2.exe"
	killProcesses $ComputerName $CurrentDomain "nlnotes.exe"
	killProcesses $ComputerName $CurrentDomain "nntspreld.exe"
	killProcesses $ComputerName $CurrentDomain "nevent.exe"
}
#----------------------------------------------------------------------------

function ClearLotusCacheForUser($ComputerName, $UserName, $CurrentDomain)
{
	KillLotusForComputer $ComputerName $CurrentDomain
	start-sleep 1
	$userPath = GetRemoteUserPathForPC $ComputerName $UserName $CurrentDomain
	$fullpathA = @("$userPath\AppData\Local\Lotus\Notes\Data\cache.ndk")
	$fullpathA += "$userPath\AppData\Local\IBM\Notes\Data\cache.ndk"
	$fullpathA += "$userPath\AppData\Local\IBM\Lotus\Notes\Data\cache.ndk"
	foreach ($fullpath in $fullpathA)
	{
		if ((Test-Path $fullpath))
		{
			try
			{
				Remove-Item $fullpath -force -EA Stop | Out-null
				"Removed: $fullpath"
			}
			catch
			{
				"Failed to remove: $fullpath"
				$ErrorMessage = $_.Exception.Message
				"ERROR: $ErrorMessage"
			}
		}
		else
		{
			"Path didnt exist: $fullpath"
		}
	}
}
#----------------------------------------------------------------------------

function ClearLotusProfileForUser($ComputerName, $UserName, $CurrentDomain)
{
	KillLotusForComputer $ComputerName $CurrentDomain
	start-sleep 1
	$userPath = GetRemoteUserPathForPC $ComputerName $UserName $CurrentDomain
	$fullpathA = @("$userPath\AppData\Local\Lotus\Notes\Data")
	$fullpathA += "$userPath\AppData\Local\IBM\Notes\Data"
	$fullpathA += "$userPath\AppData\Local\IBM\Lotus\Notes\Data"
	foreach ($fullpath in $fullpathA)
	{
		if ((Test-Path $fullpath))
		{
			try
			{
				Rename-Item -path $fullpath -newname "Data.OLD"
				"Renamed to .OLD: $fullpath"
			}
			catch
			{
				"Failed to rename: $fullpath"
				$ErrorMessage = $_.Exception.Message
				"ERROR: $ErrorMessage"
			}
		}
		else
		{
			"Path didnt exist: $fullpath"
		}
	}
}
#----------------------------------------------------------------------------

function KillOnedriveAndRelated($ComputerName, $CurrentDomain)
{
	killProcesses $ComputerName $CurrentDomain "groove.exe"
	killProcesses $ComputerName $CurrentDomain "msosync.exe"
	killProcesses $ComputerName $CurrentDomain "msouc.exe "
	killProcesses $ComputerName $CurrentDomain "csisyncclinet.exe"
	killProcesses $ComputerName $CurrentDomain "winword.exe"
	killProcesses $ComputerName $CurrentDomain "excel.exe"
	killProcesses $ComputerName $CurrentDomain "powerpnt.exe"
}
#----------------------------------------------------------------------------

function ClearOnedriveForUser($ComputerName, $UserName, $CurrentDomain)
{
	KillOnedriveAndRelated $ComputerName $CurrentDomain
	start-sleep 1
	$userPath = GetRemoteUserPathForPC $ComputerName $UserName $CurrentDomain
	$fullpathA = @("$userPath\AppData\Local\Microsoft\Office\16.0\OfficeFileCache")
	$fullpathA += "$userPath\AppData\Local\Microsoft\Office\Spw"
	foreach ($fullpath in $fullpathA)
	{
		if ((Test-Path $fullpath))
		{
			try
			{
				Remove-Item $fullpath -Recurse -force -EA Stop | Out-null
				"Removed: $fullpath"
			}
			catch
			{
				"Failed to remove: $fullpath"
				$ErrorMessage = $_.Exception.Message
				"ERROR: $ErrorMessage"
			}
		}
		else
		{
			"Path didnt exist: $fullpath"
		}
	}
}
#----------------------------------------------------------------------------


Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer