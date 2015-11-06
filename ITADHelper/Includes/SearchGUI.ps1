# Script manifest
#----------------------------------------------------------------------------
# Version number of Script
# ScriptVersion = '0.5'

# Author of this module
# Author = 'Anderscode (git@c-solutions.se)'

# Description of the functionality provided by this module
# Description = 'Search result GUIs'
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

function searchForADComputerAndReturn($computers)
{
	#----------------------------------------------
	#region Form Objects
	#----------------------------------------------
	$label = New-Object 'System.Windows.Forms.Label'
	$listboxComputers = New-Object 'System.Windows.Forms.ListBox'
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
    $label.Text = "Select user:"

    # listboxComputers
	$listboxComputers.Location = '10, 30'
	$listboxComputers.Size = '345, 280'
	$listboxComputers.Name = "listboxComputers"
	$listboxComputers.FormattingEnabled = $True
	
    # ButtonOk
    $buttonOk.Location = '70,310'
    $buttonOk.Size = '75,25'
    $buttonOk.Text = "OK"
    $buttonOk.Add_Click({ 
		if ($listboxComputers.SelectedIndex -ne -1)
		{
			if ($computers -is [system.array])
			{
				$form.Tag = ($computers[$listboxComputers.SelectedIndex]).Name
			}
			else
			{
				$form.Tag = ($computers).Name
			}
		}
		else
		{
			$form.Tag = $null
		}
		$form.Close() 
	})
     
    # ButtonCancel
    $buttonCancel.Location = '220,310'
    $buttonCancel.Size = '75,25'
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Add_Click({
		$form.Tag = $null
		$form.Close()
	})
     
    # form
    $form.Text = "Search for AD computer"
    $form.Size = '370,370'
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $True
    $form.AcceptButton = $buttonOk
    $form.CancelButton = $buttonCancel
    $form.ShowInTaskbar = $true
     
    # Add controls to form.
    $form.Controls.Add($label)
    $form.Controls.Add($listboxComputers)
    $form.Controls.Add($buttonOk)
    $form.Controls.Add($buttonCancel)

	$form.Add_Shown({
		if ($computers -is [system.array])
		{
			Foreach ($computers in $computers)
			{
				$showName = ($computers).Name + " - " + ($computers).DNSHostName
				$listboxComputers.Items.Add($showName)
			}
		}
		else
		{
			$showName = ($computers).Name + " - " + ($Users).DNSHostName
			$listboxComputers.Items.Add($showName)
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

function searchForADUserAndReturn($Users)
{

	#----------------------------------------------
	#region Form Objects
	#----------------------------------------------
	$label = New-Object 'System.Windows.Forms.Label'
	$listBoxUsers = New-Object 'System.Windows.Forms.ListBox'
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
    $label.Text = "Select user:"

    # listBoxUsers
	$listBoxUsers.Location = '10, 30'
	$listBoxUsers.Size = '345, 280'
	$listBoxUsers.Name = "listBoxUsers"
	$listBoxUsers.FormattingEnabled = $True
	
    # buttonOk
    $buttonOk.Location = '70,310'
    $buttonOk.Size = '75,25'
    $buttonOk.Text = "OK"
    $buttonOk.Add_Click({ 
		if ($listBoxUsers.SelectedIndex -ne -1)
		{
			if ($users -is [system.array])
			{
				$form.Tag = ($users[$listBoxUsers.SelectedIndex]).SamAccountName
				$form.Text = ($users[$listBoxUsers.SelectedIndex]).Domain
			}
			else
			{
				$form.Tag = ($users).SamAccountName
			}
		}
		else
		{
			$form.Tag = $null
		}
		$form.Close() 
	})
     
    # ButtonCancel
    $ButtonCancel.Location = '220,310'
    $ButtonCancel.Size = '75,25'
    $ButtonCancel.Text = "Cancel"
    $ButtonCancel.Add_Click({
		$form.Tag = $null
		$form.Close()
	})
     
    # form
    $form.Text = "Search for AD user"
    $form.Size = '370,370'
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $True
    $form.AcceptButton = $buttonOk
    $form.CancelButton = $ButtonCancel
    $form.ShowInTaskbar = $true
     
    # Add controls to form.
    $form.Controls.Add($label)
    $form.Controls.Add($listBoxUsers)
    $form.Controls.Add($buttonOk)
    $form.Controls.Add($ButtonCancel)

	$form.Add_Shown({
		if ($users -is [system.array])
		{
			Foreach ($user in $users)
			{
				$showName = ($user).SamAccountName + " - " + ($user).DisplayName + " - " + ($user).Description
				$listBoxUsers.Items.Add($showName)
			}
		}
		else
		{
			$showName = ($users).SamAccountName + " - " + ($users).DisplayName + " - " + ($users).Description
			$listBoxUsers.Items.Add($showName)
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
