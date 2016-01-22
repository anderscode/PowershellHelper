Here you can add any extra AD plugins containing functions that you want to run through this tool.
The files need to be named *.psm1 to be loaded. 

Important
* Any function inluded in these plugins need to end with "ForUser" or "CreateButton" if you want a quickbutton if you want them usable in GUI
* Dont use write-host for gui text output only do "" and it ends up in the gui
* Some functions will only return errors in the powershell console handle accordingly


Examples
#-----------------------------------------------------------------------------------------------

# AD function plugins have userName and server variables to use and need to end with ForUser
function RandomFuncForUser($userName, $server)
{
	"$computerName"
	"$userName"
	"$currentDomain"
	# This will return an array of text that is listed in GUI
}
#-----------------------------------------------------------------------------------------------

# If you want to add a quickbutton end the function name with CreateButton
# This ex will add a quickbutton with the text "RandomFunc" with calls RandomFuncForUser
function RandomFuncCreateButton($computerName, $currentDomain)
{
	RandomFuncForUser $userName $server
}
#-----------------------------------------------------------------------------------------------


# Limit export of functions, they need to be the last lines of the file.
Export-ModuleMember -Function *ForUser
Export-ModuleMember -Function *CreateButton