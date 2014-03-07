; MainLogon.AHK
; MainLogon.AHK
; Created by Joe Whipple for KEMBA Financial Credit Union on FEB 2009
; Licensed under the Creative Commons License
; Free to use/modify as long as this notice remains intact.
;
;
;
#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn UseUnsetLocal, Off


;Change these below to where you put the INI files on the network at.
_UsersINI = \\network\path\to\logonscripts\UsersINI\%A_UserName%.ini  ; The name and location of individual users INI files. This keeps track of things like signature versions and last computers logged into.
_SettingsINI = \\network\path\to\logonscripts\MainLogon.ini ; The main settings file location.
IfNotExist,%_SettingsINI% ; If we dont have settings, quit.
{
	ExitApp
}

;Display progress bar or not
IniRead, ProgressBarDisplay, %_SettingsINI%, LogonScript, ProgressBarDisplay









;########################################################################################
;##                                                                                    ##
;## PLEASE LOOK AT JRWTOOLS.AHK, THIS IS WHERE MOST OF THE WORK IS DONE IN THIS SCRIPT ##
;##                                                                                    ##
;########################################################################################
; You probably should include the full path for the two files below.
#Include JRWTools.ahk
#Include Crypt.ahk

RanManually =
; Test to see if they are on a computer that does not use logonscripts, if they are listed, exit.
IniRead, NoLogonScripts, %_SettingsINI%, LogonScript, NoLogonScripts
If ( searchRegexString(A_ComputerName,NoLogonScripts) )
{
	ExitApp
}

;################################
;# Logon Script Global Settings #
;# Configure variables below    #
;################################
ProgressMeter(1,"Loading settings...")

IniRead, DomainName, %_SettingsINI%, Network, DomainName
IniRead, NameServers, %_SettingsINI%, Network, NameServers
IniRead, DNSSearchOrder, %_SettingsINI%, Network, DNSSearchOrder
IniRead, TimeServer, %_SettingsINI%, Network, TimeServer

IniRead, AdminUser, %_SettingsINI%, Admin, AdminUser
IniRead, AdminPW, %_SettingsINI%, Admin, AdminPW
IniRead, AdminDomain, %_SettingsINI%, Admin, AdminDomain

; Because the hash key allows whoever has it to unencrypt your domain admin acct pw, do not leave this source code accessible to users.
; When you compile this to exe, it will be unreadable to the users.
AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"MyHAsh",5,1)
AdminPW := Crypt.Encrypt.StrDecrypt(AdminPW,"MyHAsh",5,1)

OS_Version := GetWindowsVersion()
OFFICEVER := GetOutlookVersion()





;#################################
;# Logonscript Functions below.  #
;#################################

;This is the meat and potatoes part. I commented out some functions we use in house that you may or may not want (look for a ;* to see what I commented out. Look at them before you enable them.
;I am only really going to show Outlook being set up.

ProgressMeter(2, "Getting version numbers.")
; Read the version that should have been passed to the script. If no version, use 0.
if ( %0% > 1 )
{
	LogonVersion = %1%
} else {
	LogonVersion = 0
}

; Main login script
;Events for every server/workstation:


; ###################### RAN EVERY TIME ###############################
;*ProgressMeter(5,"Updating login to computer.")
;*LoginTrack() ; Track whether they logged in or not.


;Events only for Workstations:
; This function looks for Windows XP and 7
If HasWorkstation()
{
; ###################### RAN EVERY TIME ON WORKSTATIONS ###############################
	;ALL OS Versions
	
	If ( OS_Version == "Windows 7" )
	{
		;ProgressMeter(35,"Setting up registry settings.")
		SetTimeWin7() ;Sync time on computer to adc01
	}
	
	If OFFICEVER != 0
	{
		ProgressMeter(40,"Configuring Outlook.")
		;This function starts the setup for Outlook based on the ini file it is passed.
		SetupOutlook(_SettingsINI)
		
		;Here is an example of checking for a group membership, and then setting up an additional Outlook profile based on it.
		;This does a recursive search for the user to see if they are in a group.
		If IsUserInADGroup("ServiceFolder Access AD Group Name Here")
		{	
			ProgressMeter(60,"Setting up Service email account access for you.")
			SetupOutlook(_SettingsINI,0,"SERVICE Outlook")
		} else {
			RemoveOfficeProfile("Service Folder")
			;RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Service Folder
		}
		
		
		; If they have more than one profile, prompt for which one.
		If OutlookProfileCount() > 1
		{
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 1
		} else {
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0
		}
	}

}



RegRead, UserLogonVersion, HKEY_CURRENT_USER, SOFTWARE\KEMBA, LogonVersion
; The following runs if the users logon version does not match. I use the version number which I store in a registry location to see if I need to do things that only require running once in a logon script.

If (LogonVersion != UserLogonVersion) or (LogonVersion = 0)
{


; ###################### RAN ON NEW LOGONSCRIPT VERSION ###############################
	;ProgressMeter(70,"Setting up DNS Servers")
	;SetDNSServers(NameServers, DomainName, DNSSearchOrder)  We are all DHCP now.

	If HasWorkstation()
	{
		; RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Applets\Tour\RunCount, 0 ;This was for Windows XP.

	
		;*ProgressMeter(90,"Removing DOC printers.")
		;*RemoveDocPrinters() ; Remove Office installed Doc printers
		
		;*ProgressMeter(90,"Disabling USB sleep.")
		;*USBNoSleep() ; Dont powersave USB
		
		;*ProgressMeter(90,"Disable WTime.")
		;*DisableWtime()
		
	}
	;Update registry logon version.
	RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\KEMBA, LogonVersion, %LogonVersion% 
}

; Remove blank lines from UserINI file.
IniFileCleanup()

ExitApp






; Functions Below
; #################################################################################################################################################################

;I used this function to move peoples favorites in IE to a network location when I applied a Group Policy policy that moved them.
CheckForFavorites()
{

	RegRead, FavFolder, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders, Favorites
	StringReplace, FavFolder, FavFolder,`%USERNAME`%,%A_UserName%, All
	StringLeft, FavDrive, FavFolder, 2
	
	;Only those with redirected favorites matter.
	If FavDrive != \\
	{

		return
	}
	
	IfNotExist,%FavFolder%
	{

		FileCreateDir,%FavFolder%
	}
	
	_homedir := GetUserHomeDir()
		
	IfExist, %_homedir%\Favorites
	{

		IfExist, %_homedir%\Favorites\FilesMoved.txt
		{

			FileRemoveDir, %_homedir%\Favorites, 1
		}
		else
		{

			FileMoveDir,%_homedir%\Favorites,%FavFolder%,1
			FileAppend, `n, %_homedir%\Favorites\FilesMoved.txt
		}
	}
	return
}



;Easy way to track the last logins a user made to a pc. It is stored in an ini file.
;An example of its output:
;
;	[LastLogin]
;	Date=02-24-2014
;	Time=13:16
;	Computer=HIL106
;	IPAddress=10.4.5.104
;	[LoginHistory]
;	PreviousComputer=LEND131
;	NextPrevious=LEND123
;
;
LoginTrack()
{
	global _UsersINI
	; Edits/creates the users INI file if not present and updates the login info.
	IniRead, PreviousLogin, %_UsersINI%, LastLogin, Computer
	IniRead, PreviousLogin2, %_UsersINI%, LoginHistory, PreviousComputer
	IniDelete, %_UsersINI%, LastLogin
	IniDelete, %_UsersINI%, LoginHistory
	
	Addresses := 
	
	Loop, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces, 1, 1
	{
		IP_Address := "0.0.0.0"
		StaleAdapter := A_LoopRegSubKey
		IsStale := 1
		
		StringReplace, StaleAdapter, StaleAdapter, Interfaces, DNSRegisteredAdapters, All
		
		if a_LoopRegName = DefaultGateway
		{
			RegRead, IPDomain, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, DhcpDomain
			; Change the domain below to the domain of your workstations.
			If ( IPDomain == "contoso.com" )
			{
				RegRead, IsStale, HKEY_LOCAL_MACHINE, %StaleAdapter%, StaleAdapter
				if ( IsStale = 0 )
				{
					RegRead, IP_Address, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, IPAddress
				}
			}
		}
		
		if a_LoopRegName = DhcpDefaultGateway
		{

			RegRead, IPDomain, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, DhcpDomain
			; Change the domain below to the domain of your workstations.
			If ( IPDomain = "contoso.com" )
			{
				RegRead, IsStale, HKEY_LOCAL_MACHINE, %StaleAdapter%, StaleAdapter
				if ( IsStale = 0 )
				{
					RegRead, IP_Address, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, DhcpIPAddress
				}
			}
		}
		
		If ( IP_Address != "0.0.0.0" ) and ( RegExMatch( IP_Address, "\b(25[0-5]|2[0-4][0-9]|[01]?[1-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b") )
		{
			If ( Addresses )
			{
				Addresses := Addresses . ", " . IP_Address 
			} else {
				Addresses := IP_Address 
			}
		}

	}
	; A simple way to track who logged into what computer when.
	IniWrite, %A_MM%-%A_DD%-%A_YYYY%, %_UsersINI%, LastLogin, Date
	IniWrite, %A_Hour%:%A_Min%, %_UsersINI%, LastLogin, Time
	IniWrite, %A_ComputerName%, %_UsersINI%, LastLogin, Computer
	IniWrite, %Addresses%, %_UsersINI%, LastLogin, IPAddress

	IniWrite, %PreviousLogin%, %_UsersINI%, LoginHistory, PreviousComputer
	IniWrite, %PreviousLogin2%, %_UsersINI%, LoginHistory, NextPrevious
}

;Make some desktop shortcuts for a user.
MakeShortcuts()
{
	global AdminUser, AdminPW, AdminDomain, NoDesktopIcons

	IfNotInString, NoDesktopIcons, %A_UserName%
	{
		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		Run, %A_WinDir%\system32\xcopy.exe /Y "\\network\location\Shortcuts\*.url" "c:\Documents and Settings\All Users\desktop\",,Hide
		RunAs,
	}
}

; Disable the windows time ntp client
DisableWtime()
{
	RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\W32Time, Start, 4
}

;Set the clock to the time of the time server
SetTimeWin7()
{
	global AdminUser, AdminPW, AdminDomain, TimeServer
	
	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	RunWait, net.exe time \\%TimeServer% /SET /Y ,,Hide
	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
}

;General ini file maint.
IniFileCleanup()
{
	global _UsersINI
	; Removes leading spaces generated sometimes.
	FileRead, OutputVar, %_UsersINI%
	OutputVar := RegExReplace(OutputVar, "m)(?:(\r?\n|$))+", "$1")
	FileDelete, %_UsersINI%
	FileAppend, %OutputVar%, %_UsersINI%
}
