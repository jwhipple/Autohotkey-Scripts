#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Warn UseUnsetLocal, Off


#Include Libraries\JRWTools.ahk


OS_Version := GetWindowsVersion()
OFFICEVER := GetOutlookVersion()
_SettingsINI = MainLogon.ini

;SetupOutlook() has three parameters:
	; Usage:  SetupOutlook(<ini settings file>, <ignore existing profile? 0|1>, <Section in ini file to look for outlook settings>)
	; Requires: GetOutlookVersion() OutlookProfileExist() RunBackgroundOutlook() OutlookProfileCount() ParseSignatureFiles() SetOutlookSignatureNames()

If OFFICEVER > 0  ;They have Office installed.
{
	
	if (UserInOU("OU=Westerville")) ; Westerville OU gets a different Outlook profile setup (local storage of PST's on their server)
	{
		; Westerville has a local store for their PSTs.
		SetupOutlook(_SettingsINI,0,"WESTERVILLE Outlook")
	} else {
		SetupOutlook(_SettingsINI)
	}

	If IsUserInADGroup("G CU.INT ServiceFolderAccess")
	{	
		; The people in the above group get an additional outlook profile named Service Folder.
		SetupOutlook(_SettingsINI,0,"SERVICE Outlook")
	} else {
		RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Service Folder
	}
}


