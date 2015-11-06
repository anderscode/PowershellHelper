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
			$return = Invoke-WmiMethod -ComputerName $computerNameFull -Path win32_process -Name create -ArgumentList "gpupdate /target:User /force /wait:0" # Should update every local users gp policy, websearch inconclusive testing needed
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

#----------------------------------------------------------------------------
function listInstalledPrinterDriversGUI($printers)
{
	#----------------------------------------------
	#region Form Objects
	#----------------------------------------------
	$label = New-Object 'System.Windows.Forms.Label'
	$listboxPrinters = New-Object 'System.Windows.Forms.ListBox'
    $buttonOk = New-Object 'System.Windows.Forms.Button'
	$buttonCancel = New-Object 'System.Windows.Forms.Button'
    $form = New-Object 'System.Windows.Forms.Form'
	
	
	#----------------------------------------------
	#region Form Code
	#----------------------------------------------
    # label
    $label.Location = '10, 10'
    $label.Size = '280, 20'
    $label.AutoSize = $true
    $label.Text = "Select Printer:"

    # listboxPrinters
	$listboxPrinters.Location = '10, 30'
	$listboxPrinters.Size = '345, 280'
	$listboxPrinters.Name = "listboxPrinters"
	$listboxPrinters.FormattingEnabled = $True
	
    # buttonOk
    $buttonOk.Location = '70,310'
    $buttonOk.Size = '75,25'
    $buttonOk.Text = "OK"
    $buttonOk.Add_Click({ 
		if ($listboxPrinters.SelectedIndex -ne -1)
		{
			if ($printers -is [system.array])
			{
				$form.Tag = $printers[$listboxPrinters.SelectedIndex]
			}
			else
			{
				$form.Tag = $printers
			}
		}
		else
		{
			$form.Tag = $null
		}
		$form.Close() 
	})
     
    # buttonCancel
    $buttonCancel.Location = '220,310'
    $buttonCancel.Size = '75,25'
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Add_Click({
		$form.Tag = $null
		$form.Close()
	})
     
    # form
    $form.Text = "Select printer to remove"
    $form.Size = '370,370'
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $True
    $form.AcceptButton = $buttonOk
    $form.CancelButton = $buttonCancel
    $form.ShowInTaskbar = $true
     
    # Add controls to form.
    $form.Controls.Add($label)
    $form.Controls.Add($listboxPrinters)
    $form.Controls.Add($buttonOk)
    $form.Controls.Add($buttonCancel)

	$form.Add_Shown({
		if ($printers -is [system.array])
		{
			Foreach ($printer in $printers)
			{
				$listboxPrinters.Items.Add($printer)
			}
		}
		else
		{
			$listboxPrinters.Items.Add($printers)
		}
	})
	
    # Initialize and show form.
    $form.Add_Shown({
		$form.Activate()
	})
    $form.ShowDialog() > $null
     
    # Return selected user
    return $form.Tag
}

#----------------------------------------------------------------------------
function RemovePrinterDriverForComputer($computerName, $currentDomain)
{
	$computerNameFull = $computerName + "." + $currentDomain
	if (((Test-Path "\\$computerNameFull\c$\Users") -eq $true))
	{
		#Define the printer driver key location 64Bit
		$printerDriverKey="SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\Windows x64\\Drivers\\Version-3"
		
		if (((Test-Path "\\$computerNameFull\c$\Program Files (x86)") -eq $true))
		{
			#Define the printer driver key location 64Bit
			$printerDriverKey="SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\Windows x64\\Drivers\\Version-3"
		}
		else
		{
			#Define the printer driver key location 32Bit
			$printerDriverKey="SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\Windows NT x86\\Drivers\\Version-3" 
		}

		# Check if RemoteRegistry is running if not start it
		$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
		$remoteRegStatusBefore = $remoteRegStatus;
		if ($remoteRegStatus.Status -ne "Running")
		{
			# "Starting remote registry on remote computer"
			"RemoteRegistry is not running, starting it temporarily"
			Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name RemoteRegistry) | Out-Null
			start-sleep 2
		}
		$remoteRegStatus = get-Service -ComputerName $computerNameFull -Name RemoteRegistry 
		if ($remoteRegStatus.Status -eq "Running")
		{
			"Listing printer drivers from : $PrinterDriverKey"
			#Create an instance of the Registry Object and open the HKLM base key
			$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computerNameFull)

			#Open printer driver key
			$regKey=$reg.OpenSubKey($printerDriverKey,$true)
			if ($regKey -ne $null)
			{
				#Retrieve an array of string that contain all the subkey names
				$subKeys=$regKey.GetSubKeyNames() 

				#List all subkeys/printers and return user choise
				$returnPrinter = listInstalledPrinterDriversGUI $subKeys
				if ($returnPrinter -ne $null)
				{
					Stop-Service -force -InputObject (get-Service -ComputerName $computerName -Name Spooler) | Out-Null
					$regKey.DeleteSubKey($returnPrinter)
					$FullKeypath = $printerDriverKey + "\\" + $returnPrinter
					$regKey=$reg.OpenSubKey($FullKeypath)
					if (($regKey -ne $false) -or ($regKey -ne $null))
					{
						"Printer key removed for : $returnPrinter"
					}
					else
					{
						"Printer key removal failed for : $returnPrinter"
					}
					Start-Service -InputObject (get-Service -ComputerName $computerNameFull -Name Spooler) | Out-Null
				}
				else
				{
					"User has canceld the removal"
				}				
			}
			else
			{
				"Unable to open registry key: $printerDriverKey"
			}
		}
		else
		{
			"Unable to start RemoteRegistry stopping RemovePrinterDriverForComputer"
		}

		# If remotereg whasnt running when we started stop it again.
		if ($remoteRegStatusBefore.Status -ne "Running")
		{
			# "Stopping remote registry on remote computer"
			"Stopping the RemoteRegistry service after use"
			Stop-Service -InputObject (get-Service -ComputerName $computerName -Name RemoteRegistry) | Out-Null
		}
    }
    else
    {
        "$computerNameFull doesnt seam to use Windows so it is not supported by this script"
    }
}
#----------------------------------------------------------------------------

Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer
Export-ModuleMember -Function *CreateButton