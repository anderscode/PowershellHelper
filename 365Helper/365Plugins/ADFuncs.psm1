﻿# Module manifest
#----------------------------------------------------------------------------
@{
# Version number of this module.
ModuleVersion = '0.6'

# ID used to uniquely identify this module
GUID = '0d64294f-608c-4526-8952-1bd2f611a884'

# Author of this module
Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
Description = 'AD functions'

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

#############################################################
# Start password Choise GUI for SetTemporaryPasswordForUser
$form_Shown=
{
	$prefix = "vinter"
	if (((Get-Date).Month -ge 10) -or ((Get-Date).Month -le 2))
	{
		$prefix = "vinter"
	}
	else
	{
		$prefix = "sommar"
	}
	$passText = $prefix + (Get-Date).Year
	$vinterText = "vinter" + (Get-Date).Year
	$sommarText = "sommar" + (Get-Date).Year
	$comboboxSelectedPassword.Items.add($vinterText)
	$comboboxSelectedPassword.Items.add($sommarText)
	$comboboxSelectedPassword.SelectedItem=$passText
}

#----------------------------------------------------------------------------
function read-comboboxDialog()
{

	$label = New-Object 'System.Windows.Forms.Label'
	$comboboxSelectedPassword = New-Object 'System.Windows.Forms.ComboBox'
    $buttonOk = New-Object 'System.Windows.Forms.Button'
	$buttonCancel = New-Object 'System.Windows.Forms.Button'
    $form = New-Object 'System.Windows.Forms.Form'

    # label
    $label.Location = '10, 10'
    $label.Size = '280, 20'
    $label.AutoSize = $true
    $label.Text = "Please select a password to set"
     
    # comboboxSelectedPassword
	$comboboxSelectedPassword.FormattingEnabled = $True
	$comboboxSelectedPassword.Location = '10, 30'
	$comboboxSelectedPassword.MaxDropDownItems = 12
	$comboboxSelectedPassword.Name = "comboboxSelectedPassword"
	$comboboxSelectedPassword.Size = '280, 21'
	$comboboxSelectedPassword.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
     
    # buttonOk
    $buttonOk.Location = '70,60'
    $buttonOk.Size = '75,25'
    $buttonOk.Text = "OK"
    $buttonOk.Add_Click({ 
		$form.Tag = $comboboxSelectedPassword.Text
		$form.Close() 
	})
     
    # buttonCancel
    $buttonCancel.Location = '160,60'
    $buttonCancel.Size = '75,25'
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Add_Click({
		$form.Tag = $null
		$form.Close() 
	})
     
    # form
    $form.Text = "Temporary password choise"
    $form.Size = '310,120'
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $buttonOk
    $form.CancelButton = $buttonCancel
    $form.ShowInTaskbar = $True
     
    # Add controls to form.
    $form.Controls.Add($label)
    $form.Controls.Add($comboboxSelectedPassword)
    $form.Controls.Add($buttonOk)
    $form.Controls.Add($buttonCancel)
     
    # Initialize and show form.
	$form.Add_Shown($form_Shown)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null
     
    # Return selcted password
    return $form.Tag
}

#----------------------------------------------------------------------------
function SetTemporaryPasswordForUser($userName, $server)
{
	$currentUser = Get-ADUser -Server $server -Identity $userName -Properties *
	$passwordNeverExpires = ($currentUser).PasswordNeverExpires
	$cannotChangePassword = ($currentUser).CannotChangePassword
	if (($cannotChangePassword -ne $true) -and ($passwordNeverExpires -ne $True))
	{
		$comboboxPassword = Read-comboboxDialog
		if ($comboboxPassword -eq $null) 
		{ 
			"You clicked Cancel no change performed" 
		}
		else 
		{
			try
			{
				Set-ADAccountPassword -Server $server -Identity $userName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $comboboxPassword -Force)
				Set-ADUser -Server $server -Identity $userName -changepasswordatlogon 1
				Unlock-ADAccount -Server $server -Identity $userName
				"$userName password set to: $comboboxPassword and forcing change of password at logon"
			}
			catch 
			{
				"$userName password reset failed"
			}
		}
	}
	else
	{
		if ($passwordNeverExpires) { "Failed: Password Never expires set for this account so no change made." }
		if ($cannotChangePassword) { "Failed: Cannot change password set for this account so no change made." }
	}
}

#----------------------------------------------------------------------------
function SetNewPasswordForUser($userName, $server)
{
	$currentUser = Get-ADUser -Server $server -Identity $userName -Properties *
	$newPassword = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new password for user", "", "")
	if ($newPassword -eq "") 
	{ 
		"Passwordchange canceled" 
	}
	else 
	{
		try
		{
			Set-ADAccountPassword -Server $server -Identity $userName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force)
			Unlock-ADAccount -Server $server -Identity $userName
			"$userName password set to: $newPassword"
		}
		catch 
		{
			"$userName password reset failed"
		}
	}
}

#----------------------------------------------------------------------------
function UnlockADAccountForUser($userName, $server)
{
	"---------------------------------------------------------------"
	if ((isUserNameInAD $userName $server) -eq $True)
	{
		"Unlocking user..."
		Unlock-ADAccount -Server $server -Identity $userName
		if ((isUserNameInADLocked $userName $server) -eq $True)
		{
			"$userName is still locked"
		}
		else
		{
			"$userName is now unlocked"
		}
	}
}

#----------------------------------------------------------------------------
function UnlockCreateButton($userName, $server)
{
	UnlockADAccountForUser $userName $server
}

#----------------------------------------------------------------------------
function EnableADAccountForUser($userName, $server)
{
	"---------------------------------------------------------------"
	if ((isUserNameInAD $userName $server) -eq $True)
	{
		"Enabling user..."
		Enable-ADAccount -Server $server -Identity $userName
		if ((isUserNameInADEnabled $userName $server) -eq $True)
		{
			"$userName is now enabled"
		}
		else
		{
			"$userName is still disabled"
		}
	}
}

#----------------------------------------------------------------------------
function EnableCreateButton($userName, $server)
{
	EnableADAccountForUser $userName $server
}

#----------------------------------------------------------------------------
function DisableADAccountForUser($userName, $server)
{
	"---------------------------------------------------------------"
	if ((isUserNameInAD $userName $server) -eq $True)
	{
		"Disabling user..."
		Disable-ADAccount -Server $server -Identity $userName
		if ((isUserNameInADEnabled $userName $server) -eq $True)
		{
			"$userName is still enabled"
		}
		else
		{
			"$userName is now disabled"
		}
	}
}

#----------------------------------------------------------------------------
function DisableCreateButton($userName, $server)
{
	DisableADAccountForUser $userName $server
}

#----------------------------------------------------------------------------
Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *CreateButton