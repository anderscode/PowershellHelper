# Module manifest
#----------------------------------------------------------------------------
@{
# Version number of this module.
ModuleVersion = '0.6'

# ID used to uniquely identify this module
GUID = 'd06812b4-84e1-4232-aeb7-34fbef9b7f16'

# Author of this module
Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
Description = 'Misc plugin functions (GPUpdate/FlushDNS...)'

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
	$optionalOutput = $_
	out-host
}
#----------------------------------------------------------------------------

function StartRemoteRegistryForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
	if ($remoteRegStatus.Status -ne "Running")
	{
		"Starting RemoteRegistry..."
		Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name RemoteRegistry) | Out-Null
		start-sleep 2
	}
	else
	{
		"RemoteRegistry is aldready running"
	}
	$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
	if ($remoteRegStatus.Status -ne "Running")
	{
		"Failed to start RemoteRegistry"
	}
}
#----------------------------------------------------------------------------

function StopRemoteRegistryForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
	if ($remoteRegStatus.Status -ne "Running")
	{
		"RemoteRegistry is not running on $computerName"
	}
	else
	{
		"Stopping the RemoteRegistry service"
        Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name RemoteRegistry) | Out-Null
	}
	$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
	if ($remoteRegStatus.Status -eq "Running")
	{
		"Failed to stop RemoteRegistry"
	}
}

#----------------------------------------------------------------------------
function ForceRebootForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "cmd /c shutdown /r /t 30 /f /c RebootByIT"
	if ($return.ReturnValue -eq 0)
	{
		"$computerName will reboot in 30 seconds"
	}
	else
	{
		"$computerName reboot command failed"
	}
}

#----------------------------------------------------------------------------
function AbortRebootForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "cmd /c shutdown /a"
	if ($return.ReturnValue -eq 0)
	{
		"$computerName reboot aborted"
	}
	else
	{
		"$computerName abort reboot command failed"
	}
}

#----------------------------------------------------------------------------
function FlushDNSForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "cmd /c ipconfig /flushdns"
	if ($return.ReturnValue -eq 0)
	{
		"$computerName flushdns completed"
	}
	else
	{
		"$computerName flushdns failed"
	}
}

#----------------------------------------------------------------------------
function FlushDNSCreateButton($computerName, $currentDomain)
{
	FlushDNSForComputer $computerName $currentDomain
}

#----------------------------------------------------------------------------
# GPUUpdate function
function GPUpdateForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
    $targetOSInfo = Get-WmiObject -ComputerName $computerNameFull -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($targetOSInfo -eq $null)
    {
        "$computerName;GPUPDATE_FAILED"
    }
    else
    {
        If ($targetOSInfo.version -ge 5.1)
        {
            $return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "gpupdate /target:Computer /force /wait:0"
			if ($return.ReturnValue -eq 0)
			{
				"$computerName GPUpdate completed for computer"
			}
			else
			{
				"$computerName GPUpdate failed for computer"
			}
			$return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "gpupdate /target:User /force /wait:0"
			if ($return.ReturnValue -eq 0)
			{
				"$computerName GPUpdate completed for users"
			}
			else
			{
				"$computerName GPUpdate failed for users"
			}
        }
        else
        {
            $return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create –ArgumentList “secedit /refreshpolicy machine_policy /enforce“
			if ($return.ReturnValue -eq 0)
			{
				"$computerName GPUpdate completed"
			}
			else
			{
				"$computerName GPUpdate failed"
			}
        }
    }
}

#----------------------------------------------------------------------------
# Clear WSUS cache function
function ForceClearWsusCacheForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name wuauserv) | Out-Null
	Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name cryptSvc) | Out-Null
	Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name bits) | Out-Null
	Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name msiserver) | Out-Null
	start-sleep 2
	$fullpathA = @("C:\\Windows\\SoftwareDistribution")
	$fullpathA += "C:\\Windows\\System32\\catroot2"
	foreach ($fullpath in $fullpathA)
	{
		if ((Test-Path $fullpath))
		{
			try
			{
				$leafPart = Split-Path $fullpath -leaf
				Rename-Item -path $fullpath -newname "$leafPart.OLD"
				"Renamed: $fullpath"
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
	Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name wuauserv) | Out-Null
	Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name cryptSvc) | Out-Null
	Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name bits) | Out-Null
	Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name msiserver) | Out-Null
}

#----------------------------------------------------------------------------
function GPUpdateCreateButton($computerName, $currentDomain)
{
	GPUpdateForComputer $computerName $currentDomain
}

#----------------------------------------------------------------------------
function KillMsiExecForComputer($computerName, $currentDomain)
{
	killProcesses $computerName $currentDomain "msiexec.exe"
}

#----------------------------------------------------------------------------
function KillMSICreateButton($computerName, $currentDomain)
{
	KillMsiExecForComputer $computerName $currentDomain
}

Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer
Export-ModuleMember -Function *CreateButton