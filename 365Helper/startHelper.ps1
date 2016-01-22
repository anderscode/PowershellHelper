#Requires -Version 2.0

# Script manifest
#----------------------------------------------------------------------------
# Version number of Script
$scriptVersion = '0.7'
$scriptDate = 'November 2015'

# ID used to uniquely identify this script
# GUID = '57cb658a-a4c7-46e4-a91c-04a27d886699'

# Author of this module
$authorName = 'Anderscode'
$authorEmail = 'git@c-solutions.se'
$authorWebsite = 'https://github.com/anderscode/PowershellHelper'

# Description of the functionality provided by this module
# Description = 'IT Helper script using GUI and plugin support'
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
# Start of initializing code

# Include file with functions used for the program
#------------------------------------------------------------
if(!$PSScriptRoot) 
{
	set-variable -name PSScriptRoot -Scope Script
	$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent # Needed incase of Powershell version 2
}
. "$PSScriptRoot\Includes\HelperFunctions.ps1" # Inlcude general functions
. "$PSScriptRoot\Includes\SearchGUI.ps1" # Include search result GUIs
. "$PSScriptRoot\Includes\365Functions.ps1" # Inlcude AD GUI eventcode and functions

#------------------------------------------------------------
#String Array containing all functions able to be called
$365_arrayFunctions = @()
$365_arrayFunctionList = @()

#------------------------------------------------------------
# Checking for ActiveDirectory module and loading if needed/possible
if(-not(Get-Module -name "ActiveDirectory")) 
{ 
    if(Get-Module -ListAvailable | Where-Object { $_.name -eq "ActiveDirectory" }) 
    { 
        Import-Module -Name "ActiveDirectory"
        "Imported ActiveDirectory module as its needed"
    }
    else 
    {
        "ActiveDirectory module doesnt exist exiting script"
        exit(1)
    }
}  

#Domain for user running this tool.
$script:currentDomainByLoggedInUser = (Get-ADDomain -Current LoggedOnUser | Select DNSRoot).DNSRoot
$script:startTime1 = Get-Date

#------------------------------------------------------------
# Unload all modules from itplugin folder incase they remained for unknown reason so we can reload them properly
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\365Plugins")
{
	if ($moduleToUnLoad.Attributes -ne "Directory")
	{
		if ($moduleToUnLoad -like "*.psm1")
		{
			Remove-Module -EA SilentlyContinue $moduleToUnLoad.BaseName | out-null
		}
	}
}

#------------------------------------------------------------
# Load all modules from adplugin folder and build array listing the functions included in them. These can then be used against any pc/user specified in GUI
foreach ($moduleToLoad in Get-ChildItem -Path "$PSScriptRoot\365Plugins")
{
	if ($moduleToLoad.Attributes -ne "Directory")
	{
		if ($moduleToLoad -like "*.psm1")
		{
			if ($PSVersionTable.PSVersion.Major -ge 3)
			{
				Import-Module -Name "$PSScriptRoot\365Plugins\$moduleToLoad" -Scope local
			}
			elseif ($PSVersionTable.PSVersion.Major -ge 2)
			{
				Import-Module -Name "$PSScriptRoot\365Plugins\$moduleToLoad"
				"Importing module on Powershell version 2.0 please upgrade powershell to be safe."
			}
			start-sleep -m 100 # Possibel fix for functions not listed sometimes.
			
			$365_arrayFunctions = Get-Command -CommandType function -Module $moduleToLoad.BaseName
			foreach ($functionName in $365_arrayFunctions)
			{
				$365_arrayFunctionList += $functionName
			}
		}
	}
}


#############################################################
# General GUI eventcode

#------------------------------------------------------------
$ITADHelper_Load=
{
	# Place any custom load script here
}

#------------------------------------------------------------
$ITADHelper_Shown=
{
	$IT_richtextboxOutput.font = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
	$365_richtextboxOutput.font = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
	
	#Fill from 365_arrayFunctionList in comboboxRun and so user can choose what to do
	foreach ($365_functionName in $365_arrayFunctionList)
	{
		# Add only ForComputer and ForUser functions to comboboxRun
		if (($365_functionName -like "*ForUser") -or ($365_functionName -like "*ForComputer"))
		{
			$365_comboboxFunction.Items.add($365_functionName)
		}
		else # Any other function is ignored unless someone wants to add something...
		{
			# Ignoring function
		}
	}
	
	#------------------------------------------------------------
	# Add AD Helper buttons from plugins
	foreach ($365_functionName in $365_functionName)
	{
		if ($365_functionName -like "*CreateButton")
		{
			[String]$buttonText = $365_functionName
			[String]$buttonText = $buttonText.Substring(0,$buttonText.Length-12)
			$object = New-Object System.Windows.Forms.Button
			$object.Size = '79, 20'
			$object.Margin = '3,3,3,2'
			$object.MaximumSize  = '150, 20'
			$object.AutoSize = $True
			$object.Text = $buttonText
			$object.TabStop = $False
			$object.UseVisualStyleBackColor = $True
			$object.add_Click({365_buttonDynamic_Click})
			$365_flowlayoutpanelQuickButtons.Controls.Add($object) 
		}
	}
	
	#------------------------------------------------------------
	# Add domain choises to AD Helper tab
	$domains = (Get-ADForest).Domains
	if ($domains -ne $Null)
	{
		foreach($domain in $domains)
		{
			$365_listboxDomains.Items.add($domain)
			$domainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"}
		    if ($domainTrusts -ne $Null)
			{
				if ($domainTrusts -is [array])
				{
					foreach($trust in $domainTrusts) 
					{
						if (!$365_listboxDomains.Items.Contains($trust.Name))
						{
							$365_listboxDomains.Items.add($trust.Name)
						}
					}
				}
				else
				{
					if (!$365_listboxDomains.Items.Contains($domainTrusts.Name))
					{
						$365_listboxDomains.Items.add($domainTrusts.Name)
					}
				}
			}
		}
	}
	$365_listboxDomains.SelectedItem=$script:currentDomainByLoggedInUser
}

#------------------------------------------------------------
function buttonAbout_Click
{
$authorLicense = "The MIT License (MIT)

Copyright (c)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
[void][System.Windows.Forms.MessageBox]::Show("Version:`t$ScriptVersion ($ScriptDate)`nName:`t$authorName`nEmail:`t$authorEmail`nWebSite:`t$authorWebsite`n`n`n$authorLicense","About")
}

#############################################################
# GUI Form setup
function run-ITADHelperForm 
{
	#------------------------------------------------------------
	#Import assembly
	[void][reflection.assembly]::Load('mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

	#------------------------------------------------------------
	# Create general objects
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$ITADHelper = New-Object 'System.Windows.Forms.Form'
	$tabcontrolMain = New-Object 'System.Windows.Forms.TabControl'
	$buttonAbout = New-Object 'System.Windows.Forms.Button'
	
	#------------------------------------------------------------
	# Create AD Helper objects
	$tabpage365Helper = New-Object 'System.Windows.Forms.TabPage'
	$365_buttonRun = New-Object 'System.Windows.Forms.Button'
	$365_buttonSearchStatus = New-Object 'System.Windows.Forms.Button'
	$365_listboxDomains = New-Object 'System.Windows.Forms.ListBox'
	$365_comboboxFunction = New-Object 'System.Windows.Forms.ComboBox'
	$365_richtextboxOutput = New-Object 'System.Windows.Forms.RichTextBox'
	$365_flowlayoutpanelQuickButtons = New-Object 'System.Windows.Forms.FlowLayoutPanel'
	$365_textboxUser = New-Object 'System.Windows.Forms.TextBox'
	$365_textboxDomain = New-Object 'System.Windows.Forms.TextBox'
	$365_labelPleaseEnterUsernameO = New-Object 'System.Windows.Forms.Label'
	
	#############################################################
	# Object setup
	$ITADHelper.SuspendLayout()
	$tabcontrolMain.SuspendLayout()
	$tabpage365Helper.SuspendLayout()
	
	#------------------------------------------------------------
	# ITADHelper
	$ITADHelper.Controls.Add($tabcontrolMain)
	$ITADHelper.Controls.Add($buttonAbout)
	$ITADHelper.ClientSize = '779, 474'
	$ITADHelper.FormBorderStyle = 'FixedSingle'
	$ITADHelper.MaximizeBox = $False
	$ITADHelper.MinimizeBox = $False
	$ITADHelper.StartPosition = 'CenterScreen'
	$ITADHelper.Name = 'ITADHelper'
	$ITADHelper.Text = "ITADHelper : " + $script:currentDomainByLoggedInUser + " - version: $ScriptVersion"
	$ITADHelper.add_Load($ITADHelper_Load)
	$ITADHelper.add_Shown($ITADHelper_Shown)
	
	#------------------------------------------------------------
	# tabcontrolMain
	$tabcontrolMain.Controls.Add($tabpage365Helper)
	$tabcontrolMain.Anchor = 'Top, Bottom, Left, Right'
	$tabcontrolMain.Location = '3, 1'
	$tabcontrolMain.Name = 'tabcontrolMain'
	$tabcontrolMain.SelectedIndex = 0
	$tabcontrolMain.Size = '772, 468'
	$tabcontrolMain.TabIndex = 0
	
	#------------------------------------------------------------
	# buttonAbout
	$buttonAbout.Location = '698, 1'
	$buttonAbout.Name = 'buttonAbout'
	$buttonAbout.Size = '75, 20'
	$buttonAbout.TabIndex = 6
	$buttonAbout.Text = 'About'
	$buttonAbout.UseVisualStyleBackColor = $True
	$buttonAbout.add_Click({buttonAbout_Click})
	$buttonAbout.BringToFront()
	
	
	#############################################################
	# AD Helper object setup
	$tabpage365Helper.Controls.Add($365_buttonRun)
	$tabpage365Helper.Controls.Add($365_comboboxFunction)
	$tabpage365Helper.Controls.Add($365_listboxDomains)
	$tabpage365Helper.Controls.Add($365_buttonSearchStatus)
	$tabpage365Helper.Controls.Add($365_richtextboxOutput)
	$tabpage365Helper.Controls.Add($365_flowlayoutpanelQuickButtons)
	$tabpage365Helper.Controls.Add($365_labelPleaseEnterUsernameO)
	$tabpage365Helper.Controls.Add($365_textboxUser)
	$tabpage365Helper.Controls.Add($365_textboxDomain)
	$tabpage365Helper.BackColor = 'Control'
	$tabpage365Helper.Location = '4, 22'
	$tabpage365Helper.Name = 'tabpage365Helper'
	$tabpage365Helper.Padding = '3, 3, 3, 3'
	$tabpage365Helper.Size = '764, 442'
	$tabpage365Helper.TabIndex = 1
	$tabpage365Helper.Text = '365Helper'
	
	#------------------------------------------------------------
	# buttonRun
	$365_buttonRun.Location = '683, 29'
	$365_buttonRun.Name = '365_buttonRun'
	$365_buttonRun.Size = '75, 20'
	$365_buttonRun.TabIndex = 6
	$365_buttonRun.Text = 'Run'
	$365_buttonRun.UseVisualStyleBackColor = $True
	$365_buttonRun.add_Click({365_buttonRun_Click})

	#------------------------------------------------------------
	# comboboxFunction
	$365_comboboxFunction.DropDownStyle = 'DropDownList'
	$365_comboboxFunction.FormattingEnabled = $True
	$365_comboboxFunction.Location = '426, 29'
	$365_comboboxFunction.Name = '365_comboboxFunction'
	$365_comboboxFunction.Size = '251, 21'
	$365_comboboxFunction.TabIndex = 5

	#------------------------------------------------------------
	# 365_listboxDomains
	$365_listboxDomains.FormattingEnabled = $True
	$365_listboxDomains.Location = '6, 55'
	$365_listboxDomains.Name = '365_listboxDomains'
	$365_listboxDomains.Size = '294, 95'
	$365_listboxDomains.TabIndex = 5

	#------------------------------------------------------------
	# buttonSearchStatus
	$365_buttonSearchStatus.Location = '306, 55'
	$365_buttonSearchStatus.Name = '365_buttonSearchStatus'
	$365_buttonSearchStatus.Size = '99, 20'
	$365_buttonSearchStatus.TabIndex = 2
	$365_buttonSearchStatus.Text = 'Search/Status'
	$365_buttonSearchStatus.UseVisualStyleBackColor = $True
	$365_buttonSearchStatus.add_Click({365_buttonSearchStatus_Click})
	
	#------------------------------------------------------------
	# richtextboxOutput
	$365_richtextboxOutput.Location = '6, 154'
	$365_richtextboxOutput.Name = '365_richtextboxOutput'
	$365_richtextboxOutput.ReadOnly = $True
	$365_richtextboxOutput.Size = '752, 282'
	$365_richtextboxOutput.TabIndex = 4
	$365_richtextboxOutput.Text = ''
	$365_richtextboxOutput.add_TextChanged({365_richtextboxOutput_TextChanged})
	$365_richtextboxOutput.add_LinkClicked({365_richtextboxOutput_LinkClicked})
	
	#------------------------------------------------------------
	# flowlayoutpanel1
	$365_flowlayoutpanelQuickButtons.FlowDirection = 'TopDown'
	$365_flowlayoutpanelQuickButtons.Location = '422, 52'
	$365_flowlayoutpanelQuickButtons.Name = '365_flowlayoutpanelQuickButtons'
	$365_flowlayoutpanelQuickButtons.Size = '350, 93'
	$365_flowlayoutpanelQuickButtons.TabIndex = 3

	#------------------------------------------------------------
	# labelPleaseEnterUsernameO
	$365_labelPleaseEnterUsernameO.AutoSize = $True
	$365_labelPleaseEnterUsernameO.Location = '6, 10'
	$365_labelPleaseEnterUsernameO.Name = '365_labelPleaseEnterUsernameO'
	$365_labelPleaseEnterUsernameO.Size = '586, 13'
	$365_labelPleaseEnterUsernameO.TabIndex = 1
	$365_labelPleaseEnterUsernameO.Text = 'Please enter username or searchphrase. (* is wildcard charater, Search in displayName/SAMName/Office/Email and Description)'
	
	#------------------------------------------------------------
	# textbox1
	$365_textboxUser.Location = '6, 29'
	$365_textboxUser.Name = '365_textboxUser'
	$365_textboxUser.Size = '245, 20'
	$365_textboxUser.TabIndex = 0
	$365_textboxUser.add_TextChanged({365_textboxUser_TextChanged})
	$365_textboxUser.Add_KeyDown({
		if ($_.KeyCode -eq "Enter") 
		{
			365_buttonSearchStatus_Click
		}
	}) 
	
	#------------------------------------------------------------
	# textbox1
	$365_textboxDomain.Location = '255, 29'
	$365_textboxDomain.Name = '365_textboxDomain'
	$365_textboxDomain.Size = '150, 20'
	$365_textboxDomain.Text = ""
	$365_textboxDomain.ReadOnly = $True
	$365_textboxDomain.TabStop = $False

	#############################################################
	#
	$tabpage365Helper.ResumeLayout()
	$tabcontrolMain.ResumeLayout()
	$ITADHelper.ResumeLayout()
	return $ITADHelper.ShowDialog()
}

#Run the GUI form
run-ITADHelperForm | Out-Null


#############################################################
# Cleanup,Unloading code
#------------------------------------------------------------
# Unload all modules from adplugin folder
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\365Plugins")
{
	if ($moduleToUnLoad.Attributes -ne "Directory")
	{
		if ($moduleToUnLoad -like "*.psm1")
		{
			Remove-Module -EA SilentlyContinue $moduleToUnLoad.BaseName | out-null
		}
	}
}

