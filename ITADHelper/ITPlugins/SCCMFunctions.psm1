# Module manifest
#----------------------------------------------------------------------------
@{
# Version number of this module.
ModuleVersion = '0.6'

# ID used to uniquely identify this module
GUID = '44d72704-c64e-40e9-9590-6ec7ec394059'

# Author of this module
Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
Description = 'Plugin functions for handling SCCM Clients remotely'

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


function ClearSCCMClientCacheForComputer($computerName, $userName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	$CCMObjects = Get-WmiObject -ComputerName $computerNameFull -Class CacheInfoEx -Namespace root\ccm\softmgmtagent
	foreach ($CCMObject in $CCMObjects)
	{
		$CachePath = $CCMObject.Location
		$CachePath = $CachePath.replace(":","$")
		$CacheRemotePath = "\\$computerNameFull\$CachePath"
		Remove-Item $CacheRemotePath -Recurse -Force
		$CCMObject | Remove-WmiObject
		"Removed: $CachePath"
	}
	Stop-Service -InputObject (get-Service -ComputerName $computerNameFull -Name CcmExec) | Out-Null
	start-sleep 2
	Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name CcmExec) | Out-Null
}
#----------------------------------------------------------------------------

Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer