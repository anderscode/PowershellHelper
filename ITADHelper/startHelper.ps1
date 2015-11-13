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
. "$PSScriptRoot\Includes\ADFunctions.ps1" # Inlcude AD GUI eventcode and functions
. "$PSScriptRoot\Includes\ITFunctions.ps1" # Inlcude IT GUI eventcode and functions

#------------------------------------------------------------
#String Array containing all functions able to be called
$IT_arrayFunctions = @()
$IT_arrayFunctionList = @()
$AD_arrayFunctions = @()
$AD_arrayFunctionList = @()

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
$currentDomainByLoggedInUser = (Get-ADDomain -Current LoggedOnUser | Select DNSRoot).DNSRoot
$script:startTime1 = Get-Date

#------------------------------------------------------------
# Unload all modules from adplugin folder incase they remained for unknown reason so we can reload them properly
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\ADPlugins")
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
# Unload all modules from itplugin folder incase they remained for unknown reason so we can reload them properly
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\ITPlugins")
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
foreach ($moduleToLoad in Get-ChildItem -Path "$PSScriptRoot\ADPlugins")
{
	if ($moduleToLoad.Attributes -ne "Directory")
	{
		if ($moduleToLoad -like "*.psm1")
		{
			if ($PSVersionTable.PSVersion.Major -ge 3)
			{
				Import-Module -Name "$PSScriptRoot\ADPlugins\$moduleToLoad" -Scope local
			}
			elseif ($PSVersionTable.PSVersion.Major -ge 2)
			{
				Import-Module -Name "$PSScriptRoot\ADPlugins\$moduleToLoad"
				"Importing module on Powershell version 2.0 please upgrade powershell to be safe."
			}
			start-sleep -m 100 # Possibel fix for functions not listed sometimes.
			
			$AD_arrayFunctions = Get-Command -CommandType function -Module $moduleToLoad.BaseName
			foreach ($functionName in $AD_arrayFunctions)
			{
				$AD_arrayFunctionList += $functionName
			}
		}
	}
}

#------------------------------------------------------------
# Load all modules from itplugin folder and build array listing the functions included in them. These can then be used against any pc/user specified in GUI
foreach ($moduleToLoad in Get-ChildItem -Path "$PSScriptRoot\ITPlugins")
{
	if ($moduleToLoad.Attributes -ne "Directory")
	{
		if ($moduleToLoad -like "*.psm1")
		{
			if ($PSVersionTable.PSVersion.Major -ge 3)
			{
				Import-Module -Name "$PSScriptRoot\ITPlugins\$moduleToLoad" -Scope local
			}
			elseif ($PSVersionTable.PSVersion.Major -ge 2)
			{
				Import-Module -Name "$PSScriptRoot\ITPlugins\$moduleToLoad"
				"Importing module on Powershell version 2.0 please upgrade powershell to be safe."
			}
			start-sleep -m 100 # Possibel fix for functions not listed sometimes.
			
			$IT_arrayFunctions = Get-Command -CommandType function -Module $moduleToLoad.BaseName
			foreach ($functionName in $IT_arrayFunctions)
			{
				$IT_arrayFunctionList += $functionName
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
	$AD_richtextboxOutput.font = New-Object System.Drawing.Font("Verdana",10,[System.Drawing.FontStyle]::Regular)
	
	#Fill from AD_arrayFunctionList in comboboxRun and so user can choose what to do
	foreach ($AD_functionName in $AD_arrayFunctionList)
	{
		# Add only ForComputer and ForUser functions to comboboxRun
		if (($AD_functionName -like "*ForUser") -or ($AD_functionName -like "*ForComputer"))
		{
			$AD_comboboxFunction.Items.add($AD_functionName)
		}
		else # Any other function is ignored unless someone wants to add something...
		{
			# Ignoring function
		}
	}
	
	#Fill from IT_arrayFunctionList in comboboxRun and so user can choose what to do
	foreach ($IT_functionName in $IT_arrayFunctionList)
	{
		# Add only ForComputer and ForUser functions to comboboxRun
		if (($IT_functionName -like "*ForUser") -or ($IT_functionName -like "*ForComputer"))
		{
			$IT_comboboxFunction.Items.add($IT_functionName)
		}
		else # Any other function is ignored unless someone wants to add something...
		{
			# Ignoring function
		}
	}
	
	#------------------------------------------------------------
	# Add IT Helper quick buttons from plugins
	foreach ($IT_functionName in $IT_arrayFunctionList)
	{
		if ($IT_functionName -like "*CreateButton")
		{
			[String]$buttonText = $IT_functionName
			[String]$buttonText = $buttonText.Substring(0,$buttonText.Length-12)
			$object = New-Object System.Windows.Forms.Button
			$object.Size = '79, 20'
			$object.Margin = '3,3,3,2'
			$object.MaximumSize  = '150, 20'
			$object.AutoSize = $True
			$object.Text = $buttonText
			$object.TabStop = $False
			$object.UseVisualStyleBackColor = $True
			$object.add_Click({IT_buttonDynamic_Click})
			$IT_flowlayoutpanelQuickButtons.Controls.Add($object) 
		}
	}
	
	#------------------------------------------------------------
	# Add AD Helper buttons from plugins
	foreach ($AD_functionName in $AD_arrayFunctionList)
	{
		if ($AD_functionName -like "*CreateButton")
		{
			[String]$buttonText = $AD_functionName
			[String]$buttonText = $buttonText.Substring(0,$buttonText.Length-12)
			$object = New-Object System.Windows.Forms.Button
			$object.Size = '79, 20'
			$object.Margin = '3,3,3,2'
			$object.MaximumSize  = '150, 20'
			$object.AutoSize = $True
			$object.Text = $buttonText
			$object.TabStop = $False
			$object.UseVisualStyleBackColor = $True
			$object.add_Click({AD_buttonDynamic_Click})
			$AD_flowlayoutpanelQuickButtons.Controls.Add($object) 
		}
	}
	
	#------------------------------------------------------------
	# Add domain choises to AD Helper tab
	$domains = (Get-ADForest).Domains
	if ($domains -ne $Null)
	{
		foreach($domain in $domains)
		{
			$AD_listboxDomains.Items.add($domain)
			$domainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"}
		    if ($domainTrusts -ne $Null)
			{
				if ($domainTrusts -is [array])
				{
					foreach($trust in $domainTrusts) 
					{
						if (!$AD_listboxDomains.Items.Contains($trust.Name))
						{
							$AD_listboxDomains.Items.add($trust.Name)
						}
					}
				}
				else
				{
					if (!$AD_listboxDomains.Items.Contains($domainTrusts.Name))
					{
						$AD_listboxDomains.Items.add($domainTrusts.Name)
					}
				}
			}
		}
	}
	$AD_listboxDomains.SelectedItem=$currentDomainByLoggedInUser
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
	$tabpageADHelper = New-Object 'System.Windows.Forms.TabPage'
	$AD_buttonRun = New-Object 'System.Windows.Forms.Button'
	$AD_buttonSearchStatus = New-Object 'System.Windows.Forms.Button'
	$AD_listboxDomains = New-Object 'System.Windows.Forms.ListBox'
	$AD_comboboxFunction = New-Object 'System.Windows.Forms.ComboBox'
	$AD_richtextboxOutput = New-Object 'System.Windows.Forms.RichTextBox'
	$AD_flowlayoutpanelQuickButtons = New-Object 'System.Windows.Forms.FlowLayoutPanel'
	$AD_textboxUser = New-Object 'System.Windows.Forms.TextBox'
	$AD_textboxDomain = New-Object 'System.Windows.Forms.TextBox'
	$AD_labelPleaseEnterUsernameO = New-Object 'System.Windows.Forms.Label'
	
	#------------------------------------------------------------
	# Create IT Helper objects
	$tabpageITHelper = New-Object 'System.Windows.Forms.TabPage'
	$IT_buttonRun = New-Object 'System.Windows.Forms.Button'
	$IT_buttonSearchStatus = New-Object 'System.Windows.Forms.Button'
	$IT_buttonAllUsers = New-Object 'System.Windows.Forms.Button'
	$IT_buttonCurrentUser = New-Object 'System.Windows.Forms.Button'
	$IT_buttonUserStatus = New-Object 'System.Windows.Forms.Button'
	$IT_buttonUserADStatus = New-Object 'System.Windows.Forms.Button'
	$IT_listboxUsers = New-Object 'System.Windows.Forms.ListBox'
	$IT_comboboxFunction = New-Object 'System.Windows.Forms.ComboBox'
	$IT_richtextboxOutput = New-Object 'System.Windows.Forms.RichTextBox'
	$IT_flowlayoutpanelQuickButtons = New-Object 'System.Windows.Forms.FlowLayoutPanel'
	$IT_textboxComputer = New-Object 'System.Windows.Forms.TextBox'
	$IT_labelPleaseEnterComputername = New-Object 'System.Windows.Forms.Label'
	
	#############################################################
	# Object setup
	$ITADHelper.SuspendLayout()
	$tabcontrolMain.SuspendLayout()
	$tabpageADHelper.SuspendLayout()
	
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
	$ITADHelper.Text = "ITADHelper : " + (Get-ADDomain -Current LoggedOnUser | Select DNSRoot).DNSRoot + " - version: $ScriptVersion"
	$ITADHelper.add_Load($ITADHelper_Load)
	$ITADHelper.add_Shown($ITADHelper_Shown)
	
	#------------------------------------------------------------
	# tabcontrolMain
	$tabcontrolMain.Controls.Add($tabpageITHelper)
	$tabcontrolMain.Controls.Add($tabpageADHelper)
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
	# IT Helper object setup
	$tabpageITHelper.Controls.Add($IT_buttonRun)
	$tabpageITHelper.Controls.Add($IT_comboboxFunction)
	$tabpageITHelper.Controls.Add($IT_listboxUsers)
	$tabpageITHelper.Controls.Add($IT_buttonSearchStatus)
	$tabpageITHelper.Controls.Add($IT_buttonAllUsers)
	$tabpageITHelper.Controls.Add($IT_buttonCurrentUser)
	$tabpageITHelper.Controls.Add($IT_buttonUserStatus)
	$tabpageITHelper.Controls.Add($IT_buttonUserADStatus)
	$tabpageITHelper.Controls.Add($IT_richtextboxOutput)
	$tabpageITHelper.Controls.Add($IT_flowlayoutpanelQuickButtons)
	$tabpageITHelper.Controls.Add($IT_labelPleaseEnterComputername)
	$tabpageITHelper.Controls.Add($IT_textboxComputer)
	$tabpageITHelper.BackColor = 'Control'
	$tabpageITHelper.ForeColor = 'ControlText'
	$tabpageITHelper.Location = '4, 22'
	$tabpageITHelper.Name = 'tabpageITHelper'
	$tabpageITHelper.Padding = '3, 3, 3, 3'
	$tabpageITHelper.Size = '764, 442'
	$tabpageITHelper.TabIndex = 0
	$tabpageITHelper.Text = 'ITHelper'
	
	#------------------------------------------------------------
	# buttonRun
	$IT_buttonRun.Location = '683, 29'
	$IT_buttonRun.Name = 'IT_buttonRun'
	$IT_buttonRun.Size = '75, 20'
	$IT_buttonRun.TabIndex = 6
	$IT_buttonRun.Text = 'Run'
	$IT_buttonRun.UseVisualStyleBackColor = $True
	$IT_buttonRun.add_Click({IT_buttonRun_Click})
	
	#------------------------------------------------------------
	# buttonSearchStatus
	$IT_buttonSearchStatus.Location = '306, 29'
	$IT_buttonSearchStatus.Name = 'IT_buttonSearchStatus'
	$IT_buttonSearchStatus.Size = '99, 20'
	$IT_buttonSearchStatus.TabIndex = 2
	$IT_buttonSearchStatus.Text = 'Search/Status'
	$IT_buttonSearchStatus.UseVisualStyleBackColor = $True
	$IT_buttonSearchStatus.add_Click({IT_buttonSearchStatus_Click})
	
	#------------------------------------------------------------
	# buttonAllUsers
	$IT_buttonAllUsers.Location = '306, 55'
	$IT_buttonAllUsers.Name = 'IT_buttonAllUsers'
	$IT_buttonAllUsers.Size = '99, 20'
	$IT_buttonAllUsers.TabIndex = 6
	$IT_buttonAllUsers.TabStop = $False
	$IT_buttonAllUsers.Text = 'All users'
	$IT_buttonAllUsers.UseVisualStyleBackColor = $True
	$IT_buttonAllUsers.add_Click({IT_buttonAllUsers_Click})

	#------------------------------------------------------------
	# buttonCurrentUser
	$IT_buttonCurrentUser.Location = '306, 80'
	$IT_buttonCurrentUser.Name = 'IT_buttonCurrentUser'
	$IT_buttonCurrentUser.Size = '99, 20'
	$IT_buttonCurrentUser.TabIndex = 7
	$IT_buttonCurrentUser.TabStop = $False
	$IT_buttonCurrentUser.Text = 'Current user'
	$IT_buttonCurrentUser.UseVisualStyleBackColor = $True
	$IT_buttonCurrentUser.add_Click({IT_buttonCurrentUser_Click})
	
	#------------------------------------------------------------
	# buttonUserStatus
	$IT_buttonUserStatus.Location = '306, 105'
	$IT_buttonUserStatus.Name = 'IT_buttonUserStatus'
	$IT_buttonUserStatus.Size = '99, 20'
	$IT_buttonUserStatus.TabIndex = 8
	$IT_buttonUserStatus.TabStop = $False
	$IT_buttonUserStatus.Text = 'User info'
	$IT_buttonUserStatus.UseVisualStyleBackColor = $True
	$IT_buttonUserStatus.add_Click({IT_buttonUserStatus_Click})
	
	#------------------------------------------------------------
	# IT_buttonUserADStatus
	$IT_buttonUserADStatus.Location = '306, 130'
	$IT_buttonUserADStatus.Name = 'IT_buttonUserADStatus'
	$IT_buttonUserADStatus.Size = '99, 20'
	$IT_buttonUserADStatus.TabIndex = 8
	$IT_buttonUserADStatus.TabStop = $False
	$IT_buttonUserADStatus.Text = 'AD Status'
	$IT_buttonUserADStatus.UseVisualStyleBackColor = $True
	$IT_buttonUserADStatus.add_Click({
		if ($IT_listboxUsers.SelectedIndex -ne -1)
		{
			$userName = $IT_listboxUsers.SelectedItem
			$AD_textboxUser.Text = $userName
			$AD_listboxDomains.SelectedItem=$currentDomainByLoggedInUser
			AD_buttonSearchStatus_Click
			$tabcontrolMain.SelectedIndex = 1
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show("Please select a username before trying to show status for it")
		}
	})

	#------------------------------------------------------------
	# comboboxFunction
	$IT_comboboxFunction.DropDownStyle = 'DropDownList'
	$IT_comboboxFunction.FormattingEnabled = $True
	$IT_comboboxFunction.Location = '426, 29'
	$IT_comboboxFunction.Name = 'comboboxFunction'
	$IT_comboboxFunction.Size = '251, 21'
	$IT_comboboxFunction.TabIndex = 5

	#------------------------------------------------------------
	# checkedlistboxDomains
	$IT_listboxUsers.FormattingEnabled = $True
	$IT_listboxUsers.Location = '6, 55'
	$IT_listboxUsers.Name = 'IT_listboxUsers'
	$IT_listboxUsers.Size = '294, 95'
	$IT_listboxUsers.Sorted = $True
	$IT_listboxUsers.TabIndex = 5

	#------------------------------------------------------------
	# richtextboxOutput
	$IT_richtextboxOutput.Location = '6, 154'
	$IT_richtextboxOutput.Name = 'IT_richtextboxOutput'
	$IT_richtextboxOutput.ReadOnly = $True
	$IT_richtextboxOutput.Size = '752, 282'
	$IT_richtextboxOutput.TabIndex = 4
	$IT_richtextboxOutput.Text = ''
	$IT_richtextboxOutput.add_TextChanged({IT_richtextboxOutput_TextChanged})
	$IT_richtextboxOutput.add_LinkClicked({IT_richtextboxOutput_LinkClicked})
	
	#------------------------------------------------------------
	# flowlayoutpanel1
	$IT_flowlayoutpanelQuickButtons.FlowDirection = 'TopDown'
	$IT_flowlayoutpanelQuickButtons.Location = '422, 52'
	$IT_flowlayoutpanelQuickButtons.Name = 'IT_flowlayoutpanelQuickButtons'
	$IT_flowlayoutpanelQuickButtons.Size = '350, 93'
	$IT_flowlayoutpanelQuickButtons.TabIndex = 3

	#------------------------------------------------------------
	# labelPleaseEnterUsernameO
	$IT_labelPleaseEnterComputername.AutoSize = $True
	$IT_labelPleaseEnterComputername.Location = '6, 10'
	$IT_labelPleaseEnterComputername.Name = 'IT_labelPleaseEnterComputername'
	$IT_labelPleaseEnterComputername.Size = '586, 13'
	$IT_labelPleaseEnterComputername.TabIndex = 1
	$IT_labelPleaseEnterComputername.Text = 'Please enter computername or searchphrase. (* is wildcard charater, Search in Name,DisplayName,DNSHostName)'
	
	#------------------------------------------------------------
	# textbox1
	$IT_textboxComputer.Location = '6, 29'
	$IT_textboxComputer.Name = 'IT_textboxComputer'
	$IT_textboxComputer.Size = '294, 20'
	$IT_textboxComputer.TabIndex = 0
	$IT_textboxComputer.add_TextChanged({IT_textboxComputer_TextChanged})
	$IT_textboxComputer.Add_KeyDown({
		if ($_.KeyCode -eq "Enter") 
		{
			IT_buttonSearchStatus_Click
		}
	}) 
	
	#############################################################
	# AD Helper object setup
	$tabpageADHelper.Controls.Add($AD_buttonRun)
	$tabpageADHelper.Controls.Add($AD_comboboxFunction)
	$tabpageADHelper.Controls.Add($AD_listboxDomains)
	$tabpageADHelper.Controls.Add($AD_buttonSearchStatus)
	$tabpageADHelper.Controls.Add($AD_richtextboxOutput)
	$tabpageADHelper.Controls.Add($AD_flowlayoutpanelQuickButtons)
	$tabpageADHelper.Controls.Add($AD_labelPleaseEnterUsernameO)
	$tabpageADHelper.Controls.Add($AD_textboxUser)
	$tabpageADHelper.Controls.Add($AD_textboxDomain)
	$tabpageADHelper.BackColor = 'Control'
	$tabpageADHelper.Location = '4, 22'
	$tabpageADHelper.Name = 'tabpageADHelper'
	$tabpageADHelper.Padding = '3, 3, 3, 3'
	$tabpageADHelper.Size = '764, 442'
	$tabpageADHelper.TabIndex = 1
	$tabpageADHelper.Text = 'ADHelper'
	
	#------------------------------------------------------------
	# buttonRun
	$AD_buttonRun.Location = '683, 29'
	$AD_buttonRun.Name = 'AD_buttonRun'
	$AD_buttonRun.Size = '75, 20'
	$AD_buttonRun.TabIndex = 6
	$AD_buttonRun.Text = 'Run'
	$AD_buttonRun.UseVisualStyleBackColor = $True
	$AD_buttonRun.add_Click({AD_buttonRun_Click})

	#------------------------------------------------------------
	# comboboxFunction
	$AD_comboboxFunction.DropDownStyle = 'DropDownList'
	$AD_comboboxFunction.FormattingEnabled = $True
	$AD_comboboxFunction.Location = '426, 29'
	$AD_comboboxFunction.Name = 'AD_comboboxFunction'
	$AD_comboboxFunction.Size = '251, 21'
	$AD_comboboxFunction.TabIndex = 5

	#------------------------------------------------------------
	# AD_listboxDomains
	$AD_listboxDomains.FormattingEnabled = $True
	$AD_listboxDomains.Location = '6, 55'
	$AD_listboxDomains.Name = 'AD_listboxDomains'
	$AD_listboxDomains.Size = '294, 95'
	$AD_listboxDomains.TabIndex = 5

	#------------------------------------------------------------
	# buttonSearchStatus
	$AD_buttonSearchStatus.Location = '306, 55'
	$AD_buttonSearchStatus.Name = 'AD_buttonSearchStatus'
	$AD_buttonSearchStatus.Size = '99, 20'
	$AD_buttonSearchStatus.TabIndex = 2
	$AD_buttonSearchStatus.Text = 'Search/Status'
	$AD_buttonSearchStatus.UseVisualStyleBackColor = $True
	$AD_buttonSearchStatus.add_Click({AD_buttonSearchStatus_Click})
	
	#------------------------------------------------------------
	# richtextboxOutput
	$AD_richtextboxOutput.Location = '6, 154'
	$AD_richtextboxOutput.Name = 'AD_richtextboxOutput'
	$AD_richtextboxOutput.ReadOnly = $True
	$AD_richtextboxOutput.Size = '752, 282'
	$AD_richtextboxOutput.TabIndex = 4
	$AD_richtextboxOutput.Text = ''
	$AD_richtextboxOutput.add_TextChanged({AD_richtextboxOutput_TextChanged})
	$AD_richtextboxOutput.add_LinkClicked({AD_richtextboxOutput_LinkClicked})
	
	#------------------------------------------------------------
	# flowlayoutpanel1
	$AD_flowlayoutpanelQuickButtons.FlowDirection = 'TopDown'
	$AD_flowlayoutpanelQuickButtons.Location = '422, 52'
	$AD_flowlayoutpanelQuickButtons.Name = 'AD_flowlayoutpanelQuickButtons'
	$AD_flowlayoutpanelQuickButtons.Size = '350, 93'
	$AD_flowlayoutpanelQuickButtons.TabIndex = 3

	#------------------------------------------------------------
	# labelPleaseEnterUsernameO
	$AD_labelPleaseEnterUsernameO.AutoSize = $True
	$AD_labelPleaseEnterUsernameO.Location = '6, 10'
	$AD_labelPleaseEnterUsernameO.Name = 'AD_labelPleaseEnterUsernameO'
	$AD_labelPleaseEnterUsernameO.Size = '586, 13'
	$AD_labelPleaseEnterUsernameO.TabIndex = 1
	$AD_labelPleaseEnterUsernameO.Text = 'Please enter username or searchphrase. (* is wildcard charater, Search in displayName/SAMName/Office/Email and Description)'
	
	#------------------------------------------------------------
	# textbox1
	$AD_textboxUser.Location = '6, 29'
	$AD_textboxUser.Name = 'AD_textboxUser'
	$AD_textboxUser.Size = '245, 20'
	$AD_textboxUser.TabIndex = 0
	$AD_textboxUser.add_TextChanged({AD_textboxUser_TextChanged})
	$AD_textboxUser.Add_KeyDown({
		if ($_.KeyCode -eq "Enter") 
		{
			AD_buttonSearchStatus_Click
		}
	}) 
	
	#------------------------------------------------------------
	# textbox1
	$AD_textboxDomain.Location = '255, 29'
	$AD_textboxDomain.Name = 'AD_textboxDomain'
	$AD_textboxDomain.Size = '150, 20'
	$AD_textboxDomain.Text = ""
	$AD_textboxDomain.ReadOnly = $True
	$AD_textboxDomain.TabStop = $False

	#############################################################
	#
	$tabpageADHelper.ResumeLayout()
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
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\ADPlugins")
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
# Unload all modules from itplugin folder
foreach ($moduleToUnLoad in Get-ChildItem -Path "$PSScriptRoot\ITPlugins")
{
	if ($moduleToUnLoad.Attributes -ne "Directory")
	{
		if ($moduleToUnLoad -like "*.psm1")
		{
			Remove-Module -EA SilentlyContinue $moduleToUnLoad.BaseName | out-null
		}
	}
}