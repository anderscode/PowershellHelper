Here you can add any extra IT plugins containing functions that you want to run through this tool.
The files need to be named *.psm1 to be loaded. 

Important
* Any function inluded in these plugins need to end with "ForUser", "ForComputer" or "CreateButton" if you want a quickbutton if you want them usable in GUI
* Dont use write-host for gui text output only do "" and it ends up in the gui
* Some functions will only return errors in the powershell console handle accordingly


Examples
#-----------------------------------------------------------------------------------------------

# If you need access to computerName and userName variable in your function end the function name with ForUser
function RandomFuncForUser($computerName, $userName, $currentDomain)
{
	"$computerName"
	"$userName"
	"$currentDomain"
	# This will return an array of text that is listed in GUI
}
#-----------------------------------------------------------------------------------------------

# If you only need access to computerName variable in your function end the function name with ForComputer
function RandomFuncForComputer($computerName, $currentDomain)
{
	"$currentDomain : $computerName"
	# This will return text that is listed in GUI
}
#-----------------------------------------------------------------------------------------------

# If you want to add a quickbutton end the function name with CreateButton
# This ex will add a quickbutton with the text "RandomFunc" with calls RandomFuncForComputer
function RandomFuncCreateButton($computerName, $currentDomain)
{
	RandomFuncForComputer $computerName $currentDomain
}
#-----------------------------------------------------------------------------------------------


# Limit export of functions, they need to be the last lines of the file.
Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *ForComputer
Export-ModuleMember -Function *CreateButton