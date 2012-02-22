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

_SettingsINI = \\CU.INT\logonscripts\MainLogon.ini
IfNotExist,%_SettingsINI%
{
	ExitApp
}
;Display progress bar or not
IniRead, ProgressBarDisplay, %_SettingsINI%, LogonScript, ProgressBarDisplay

#Include \\cu.int\logonscripts\Sources\JRWTools.ahk
#Include \\cu.int\logonscripts\Sources\Crypt.ahk

RanManually =



;################################
;# Logon Script Global Settings #
;# Configure variables below    #
;################################
ProgressMeter(1,"Loading settings...")

IniRead, DomainName, %_SettingsINI%, Network, DomainName
IniRead, NameServers, %_SettingsINI%, Network, NameServers
IniRead, DNSSearchOrder, %_SettingsINI%, Network, DNSSearchOrder
IniRead, TimeServer, %_SettingsINI%, Network, TimeServer

IniRead, SymitarVersion, %_SettingsINI%, Symitar, SymitarVersion
IniRead, SymitarInstaller, %_SettingsINI%, Symitar, SymitarInstaller
Symitar := "Episys Windows Interface " . SymitarVersion

IniRead, NoDesktopIcons, %_SettingsINI%, Desktop, DontCopyIcons

IniRead, AdminUser, %_SettingsINI%, Admin, AdminUser
IniRead, AdminPW, %_SettingsINI%, Admin, AdminPW
IniRead, AdminDomain, %_SettingsINI%, Admin, AdminDomain
AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"BooYah",5,1)
AdminPW := Crypt.Encrypt.StrDecrypt(AdminPW,"BooYah",5,1)

OS_Version := GetWindowsVersion()
OFFICEVER := GetOutlookVersion()





;#################################
;# Do not modify below this line #
;#################################







ProgressMeter(2, "Getting version numbers.")
; Read the version that should have been passed to the script. If no version, use 0.
if ( %0% > 1 )
{
	LogonVersion = %1%
} else {
	LogonVersion = 0
	;LogonVersion = 6.2 ; ***CHANGEME***
}

; Main Logonscript

;Events for every server/workstation:

; ###################### RAN EVERY TIME ###############################
ProgressMeter(5, "Setting your computers time.")
SetTime(TimeServer) ;Set system time.

ProgressMeter(10,"Tracking login to computer.")
LoginTrack() ; Track wether they logged in or not.

;Events only for Workstations:
If HasWorkstation()
{
; ###################### RAN EVERY TIME ON WORKSTATIONS ###############################
	;ALL OS Versions
	
	ProgressMeter(10,"Checking for Antivirus.")
	;CheckAV() ;Check and install antivirus.
	
	ProgressMeter(20,"Mapping network drives.")
	MapDrives() ;Map standard drives.
	
	ProgressMeter(25,"Initializing Symitar.")
	Clear_Symitar_Dropdowns()
	SymitarRegistry()
	
	If ( OS_Version == "Windows XP" )
	{
		ProgressMeter(30,"Setting up registry settings.")
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Applets\Tour\RunCount, 0
		MapPrinters() ;Register printers
		;DefragmentTask()
    
	}
	else If ( OS_Version == "Windows 7" )
	{
		ProgressMeter(35,"Setting up registry settings.")
	}
	
	If OFFICEVER != 0
	{
		; Remove old Scalix profiles as this will trigger the multi-profile selection box when starting Outlook.
		RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Scalix
		Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 2, 0
		{
			FoundPos := RegExMatch(A_LoopRegName, "Default(.*)")
			If FoundPos
			{
				RegDelete
			}
		}
	
		if (UserInOU("OU=Westerville"))
		{
			; Westerville has a local store for their PSTs.
			ProgressMeter(40,"Setting Outlook for WES user.")

			SetupOutlook(_SettingsINI,0,"WESTERVILLE Outlook")
		} else {
			ProgressMeter(40,"Configuring Outlook.")
			SetupOutlook(_SettingsINI)
		}
		
		If OFFICEVER < 12
		{
			ProgressMeter(50,"Installing Office 2007 converters.")
			O2k7_File_Format_Converters()
		}
		
		If IsUserInADGroup("G CU.INT ServiceFolderAccess")
		{	
			ProgressMeter(60,"Setting up Service email account access for you.")
			SetupOutlook(_SettingsINI,0,"SERVICE Outlook")
		} else {
			RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Service Folder
			If OutlookProfileCount() > 1
			{
				RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0
			}
		}
	}


;	If IsUserInADGroup("G CU.INT CSI Screen Recording")
	If A_ComputerName contains callcenter,collection,risk,msr,lending
	{
		ProgressMeter(70,"Checking for CSI software.")
		CSI_ScreenRecordingInstall()
	}
	If A_ComputerName contains jwhipple,gocconf200,gocconf201,gocconf204,GOCHRConf,gocboardroom
	{
		ProgressMeter(75,"Setting up power save settings.")
		ConfPowerSave()
	}
}



RegRead, UserLogonVersion, HKEY_CURRENT_USER, SOFTWARE\KEMBA, LogonVersion
; The following runs if the users logon version does not match.

If (LogonVersion != UserLogonVersion) or (LogonVersion = 0)
{


; ###################### RAN ON NEW LOGONSCRIPT VERSION ###############################
	
	SetDNSServers(NameServers, DomainName, DNSSearchOrder) 
	DisableNetBios()
	
	;Disabled because it doesnt like Outlook being set up.
	;ProgressMeter(75,"Inventorying PC...")
	;InventoryPC() ; Inventory PC
	
	If HasWorkstation()
	{

		If ( OS_Version == "Windows XP" )
		{
			RegOLE() ; Fix OLE32 Registration
			FixTIFFAssoc()
		}
		
		
		ProgressMeter(80,"Setting up COWWW viewer.")
		Alternatiff() ;Register and setup Alternatiff
	
		ProgressMeter(85,"Setting up Explorer Favorites.")
		CheckForFavorites()

		ProgressMeter(90,"Configuring Symitar.")
		SymitarInstall()		
	
		ProgressMeter(90,"Removing DOC printers.")
		RemoveDocPrinters() ; Remove Office installed Doc printers
		
		ProgressMeter(90,"Disabling USB sleep.")
		USBNoSleep() ; Dont powersave USB
		
		;PowerSave() ; Setup powersave settings
		
		ProgressMeter(90,"Disable WTime.")
		DisableWtime()
		
		;MakeShortcuts() ;Make user shortcuts
		
		If UsersDepartment("Member Services,Accounting Department,Contact Center,Branch Operations,Operations,Operations Administration,Lending,Branch Administration,Tellers")
		{
			ProgressMeter(99,"Installing TrueChecks.")
			TrueChecksInstall() ; Install check verification software.
		}
	}
}

ExitApp









; Functions Below
; #################################################################################################################################################################




CSI_ScreenRecordingInstall()
{
	global AdminUser, AdminPW, AdminDomain


	If Not IsInstalled("Virtual Observer Agent Client")
	{
		MsgBox, 292,CSI Virtual Observer Client, Do you wish to install the call recorder `nsoftware on this PC? `n(Select NO if this is not your normal PC),10
		IfMsgBox Yes
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, msiexec /i \\cu.int\logonscripts\Installers\CSI\Agent_Client.msi /qn,,
			RunAs,
		}
		else IfMsgBox Timeout
		{	

			return
		}
		else
		{

			return
		}
	}
	

	If Not IsInstalled("Virtual Observer E-Learning Client")
	{

		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		RunWait, msiexec /i \\cu.int\logonscripts\Installers\CSI\Elearning_Client.msi /qn,,
		RunAs,
	}
}

RegOLE()
{
	global AdminUser, AdminPW, AdminDomain

	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	RunWait, %comspec% /c Regsvr32.exe /s %A_WinDir%\System32\Ole32.dll,,
	RunAs,			
} 

TrueChecksInstall()
{
	global AdminUser, AdminPW, AdminDomain


		
	If Not IsInstalled("TrueChecks Notifier")
	{
		;RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		;Run, %comspec% /c DEL /Y `"%A_DesktopCommon%\TrueChecks*.lnk`",,
		;RunAs,

		;RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		RunWait, msiexec /passive /uninstall \\cu.int\logonscripts\Installers\TrueChecks\TrueChecksNotifierSetup.msi,,
		RunWait, msiexec /passive /i \\cu.int\logonscripts\Installers\TrueChecks\TrueChecksNotifierSetup.msi,,
		RunWait, \\cu.int\logonscripts\Installers\TrueChecks\TrueChecksSetup.exe -s -o,,
		;RunAs,


		;Move the links to the desktop.
		; AdminDesktop = %A_DesktopCommon%
		; StringReplace, AdminDesktop, AdminDesktop, All Users, %AdminUser%, ALL
		; RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		; Run, %comspec% /c MOVE /Y `"%AdminDesktop%\TrueChecks*.lnk`" `"%A_DesktopCommon%`",,
		; RunAs, 
	}

}


ETokenCheck()
{
	global AdminUser, AdminPW, AdminDomain

	If Not IsInstalled("eToken PKI Client 5.0 SP1")
	{

		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		RunWait, msiexec /qb /promptrestart /lvx* %A_Temp%/pki-install.log /i \\cu.int\logonscripts\Installers\CorpOneToken\Windows_x32_v5.00.msi PROP_REG_FILE=\\cu.int\logonscripts\Installers\CorpOneToken\settings.reg,,
		RunAs,			
	} else {

	}
}

Clear_Symitar_Dropdowns()
{

	RegDelete, HKEY_CURRENT_USER, SOFTWARE\Symitar\SFW\2.0\Account Manager\Recent RepGens
	RegDelete, HKEY_CURRENT_USER, SOFTWARE\Symitar\SFW\2.0\\Application Processing\RecentRepGens
	RegDelete, HKEY_CURRENT_USER, SOFTWARE\Symitar\SFW\2.0\Teller Transactions\RecentRepGens
}

FixTIFFAssoc()
{

	if A_OSVersion in WIN_XP
	{	

		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\,,TIFImage.Document
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\,Content Type,image/tif
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\,PerceivedType,image
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\OpenWithList\mspview.exe\,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\OpenWithProgids\TIFImage.Document,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tif\PersistentHandler\,,{58F2E3BB-72BD-46DF-B134-1B50628668FB}
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithProgids\,TIFImage.Document,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\,Progid,MSPaper.Document
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithList\,a,MSPVIEW.EXE
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithProgids\,TIFImage.Document,0
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithProgids\,MSPaper.Document,0
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithProgids\,Imaging.Document,0
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\UserChoice\,Progid,MSPaper.Document

		;Registration for .TIFF

		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\,,TIFImage.Document
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\,Content Type,image/tif
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\,PerceivedType,image
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\OpenWithList\mspview.exe\,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\OpenWithProgids\TIFImage.Document,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.tiff\PersistentHandler\,,{58F2E3BB-72BD-46DF-B134-1B50628668FB}
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithProgids\,TIFImage.Document,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\,Progid,MSPaper.Document
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithList\,a,MSPVIEW.EXE
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithProgids\,TIFImage.Document,0
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithProgids\,MSPaper.Document,0
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithProgids\,Imaging.Document,0
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\UserChoice\,Progid,MSPaper.Document

		;Registration for .MDI

		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.mdi\,,mdi_auto_file
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.mdi\,Content Type,image/vnd.ms-modi
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.mdi\,MSPaper.Document,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.mdi\MSPaper.Document\,ShellNew,
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Classes\.mdi\PersistentHandler\,,{58F2E3BB-72BD-46DF-B134-1B50628668FB}
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mdi\,Progid,MSPaper.Document
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mdi\OpenWithList\,a,MSPVIEW.EXE
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mdi\OpenWithProgids\mdi_auto_file,0
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mdi\UserChoice\,Progid,MSPaper.Document
	}
}

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

CheckAV()
{
	global AdminUser, AdminPW, AdminDomain

	; Install ESET NOD32 if not found.
	If Not IsInstalled("ESET NOD32 Antivirus")
	{

		SplashImage, \\cu.int\logonscripts\Pictures\logo.jpg, CWFFFFFF  h400 w500 b1 fs18, `n`nNow installing your anti-virus.`nAfter the install your PC will reboot automatically.
		ExitApp
		
		If OSBitVersion() = "x86"
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, \\cu.int\logonscripts\Installers\ESET\NOD32Settings.exe,,
			RunWait, msiexec.exe /passive /forcerestart /i \\cu.int\logonscripts\Installers\ESET\eavbe_nt32_enu.msi,,
			RunAs,
		}
		else
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, \\cu.int\logonscripts\Installers\ESET\NOD32Settings.exe,,
			RunWait, msiexec.exe /passive /forcerestart /i \\cu.int\logonscripts\Installers\ESET\eavbe_nt64_enu.msi,,
			RunAs,
		}
		SplashImage, Off
		Sleep, 10000
		Shutdown, 6

	} else {

	}

}

O2k7_File_Format_Converters()
{
	global AdminUser, AdminPW, AdminDomain

	If (HasOffice("11.0"))
	{
		If (Not IsInstalled("Compatibility Pack for the 2007 Office system"))
		{
			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, \\cu.int\logonscripts\Installers\OfficePatch\FileFormatConverters.exe /q,,
			RunWait, \\cu.int\logonscripts\Installers\OfficePatch\o2k7sp1pak.exe /q,,
			RunAs,
		}
	}
}



SymitarInstall()
{
	global OS_Version, Symitar, SymitarInstaller, AdminUser, AdminPW, AdminDomain

	
	If Not IsInstalled("SymForm API")
	{

		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		RunWait, %comspec% /c start msiexec /passive /i \\cu.int\GPOFiles$\Installs\symformapi\SymFormAPI.msi,,
		RunAs,
	}
	
	If Not IsInstalled(Symitar)
	{


		UninstallSoftware("Episys Windows Interface`%")
		FileDelete, C:\Documents and Settings\%A_UserName%\Application Data\Microsoft\Internet Explorer\Quick Launch\Symitar.lnk

		If ( OS_Version == "Windows XP" )
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, TASKKILL /F /IM RemoteAdminServer.exe,,Hide
			RunWait, %comspec% /c DEL "C:\Documents and Settings\All Users\Desktop\Symitar.lnk",,
			RunWait, %comspec% /c DEL "C:\Documents and Settings\%A_Username%\Application Data\Microsoft\Internet Explorer\Quick Launch\Episys*.lnk",,
			RunWait, %comspec% /c DEL "C:\Documents and Settings\%A_Username%\Application Data\Microsoft\Internet Explorer\Quick Launch\EWI*.lnk",,
			RunWait, %comspec% /c RMDIR /S /Q "C:\Documents and Settings\All Users\Start Menu\Programs\Episys",,
			RunAs,
		}
		else If ( OS_Version == "Windows 7" )
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, TASKKILL /F /IM RemoteAdminServer.exe,,Hide
			RunWait, %comspec% /c DEL "C:\Users\Public\Desktop\Symitar.lnk",,
			RunWait, %comspec% /c RMDIR /S /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Episys",,
			RunAs,
		}

		InstallSoftware(SymitarInstaller . " /quiet")
		RunWait, %comspec% /c copy "C:\Documents and Settings\All Users\Desktop\Epi*.lnk" "C:\Documents and Settings\%A_Username%\Application Data\Microsoft\Internet Explorer\Quick Launch\"

	} else {
		; Symitar is installed.
		Clear_Symitar_Dropdowns()
		SymitarRegistry()
	}
}


SymitarRegistry()
{

	;Adjusts margin settings for printing in Explorer.
	;### Changed to be group policy.
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,footer,
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,header,
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,margin_bottom, 0.50000
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,margin_left, 0.50000
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,margin_right, 0.50000
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Internet Explorer\PageSetup,margin_top, 0.50000

	;Start Page settings
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Browser Options,Startup Page,\\cu.int\Public\Symitar\html\startup.htm
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Browser Options,Information Page,http://inhouse.kemba.org/

	;Printer fix
	;RegDelete, HKEY_LOCAL_MACHINE, SOFTWARE\Symitar\SymForm, Printers
	
	; Symitar and Symform Fixes
	RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Account Manager, BlockEMailFunction, 0

	IfExist, \\cu.int\logonscripts\Registry\SymitarReg.exe
	{
		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		Run, \\cu.int\logonscripts\Registry\SymitarReg.exe
		RunAs,
	}
	
}



PowerSave()
{
	global AdminUser, AdminPW, AdminDomain

	if A_OSVersion in WIN_XP
	{
		RunWait, \\cu.int\logonscripts\Installers\Registry\powersave.bat,,HIDE
		; Run privileged
		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		Run, %A_WinDir%\system32\powercfg.exe /SETACTIVE "KEMBA",,HIDE
		RunAs,
	} else {

		return
	}
}

ConfPowerSave()
{

	if A_OSVersion in WIN_XP
	{
		Run, \\cu.int\logonscripts\Installers\Registry\confpowersave.bat,,HIDE
		; Run privileged
		;RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		;RunWait, %A_WinDir%\system32\powercfg.exe /SETACTIVE Presentation,,HIDE
		;RunAs,
	} else {

		return
	}
}

RegistrySetup()
{
	global AdminUser, AdminPW, AdminDomain

	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	RunWait, \\cu.int\logonscripts\Registry\RegRunAs.exe,,Hide
	RunAs,
}

MapPrinters()
{
	if A_OSVersion in WIN_XP
	{

		Run, \\CU.INT\LogonScripts\Printers\pushprinterconnections.exe,,Hide
	}
}

MapDrives()
{

		IfNotExist, I:\
			Run, %A_WinDir%\system32\net.exe use i: \\cu.int\KembaApps,,Hide
			;RunWait, %A_WinDir%\system32\net.exe use i: /delete /Y,,Hide
			
		IfNotExist, V:\
			Run, %A_WinDir%\system32\net.exe use v: \\cu.int\department,,Hide
			;RunWait, %A_WinDir%\system32\net.exe use v: /delete /Y,,Hide
}


Alternatiff()
{
	global AdminUser, AdminPW, AdminDomain

	IfNotExist, %A_WinDir%\system\alttiff.ocx
	{

		if not A_IsAdmin
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, %A_WinDir%\system32\xcopy.exe \\cu.int\logonscripts\tqcrunas\alttiff.ocx %A_WinDir%\system\ /y /c,,Hide
			Run, %A_WinDir%\system32\cmd.exe /C start "Please do not close this window" /belownormal /min regsvr32 /s %A_WinDir%\system\alttiff.ocx,,Hide
			RunAs,
		} else {

			RunWait, %A_WinDir%\system32\xcopy.exe \\cu.int\logonscripts\tqcrunas\alttiff.ocx %A_WinDir%\system\ /y /c,,Hide
			Run, %A_WinDir%\system32\cmd.exe /C start "Please do not close this window" /belownormal /min regsvr32 /s %A_WinDir%\system\alttiff.ocx,,Hide
		}
	}
	else
	{

		if not A_IsAdmin
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			Run, %A_WinDir%\system32\cmd.exe /C start "Please do not close this window" /belownormal /min regsvr32 /s %A_WinDir%\system\alttiff.ocx,,Hide
			RunAs,
		} else {
			Run, %A_WinDir%\system32\cmd.exe /C start "Please do not close this window" /belownormal /min regsvr32 /s %A_WinDir%\system\alttiff.ocx,,Hide
		}
	}
}

InventoryPC()
{
	; This uses the Asset Tracker for Networks package availible at http://www.alchemy-lab.com/products/atn/ for client inventory.
	IfNotExist, \\cu.int\kembaapps\Inventory\Data\%A_ComputerName%.XML
	{
		FileDelete, \\cu.int\kembaapps\Inventory\Data\%A_ComputerName%.XML
		Run, \\cu.int\kembaapps\Inventory\clientcon.exe echo-
	}
}

LoginTrack()
{
	; A simple way to track who logged into what computer when.
	FileAppend,
	(
		%A_Now% - %A_ComputerName%
	), \\cu.int\kembaapps\Inventory\Logins\%A_UserName%.txt
}


MakeShortcuts()
{
	global AdminUser, AdminPW, AdminDomain, NoDesktopIcons

	IfNotInString, NoDesktopIcons, %A_UserName%
	{

		RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
		Run, %A_WinDir%\system32\xcopy.exe /Y "\\cu.int\logonscripts\Installers\Shortcuts\*.url" "c:\Documents and Settings\All Users\desktop\",,Hide
		RunAs,
	}
}

DefragmentTask()
{
	global AdminUser, AdminPW, AdminDomain
	if A_OSVersion in WIN_XP
	{
		IfNotExist, %A_WinDir%\Installers\Tasks\DefragDriveC.job
		{
			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			Run, %A_WinDir%\system32\xcopy.exe /Y "\\cu.int\logonscripts\Installers\tasks\*.job" "%A_WinDir%\Tasks\",,Hide
			RunAs,
		}
		IfNotExist, %A_WinDir%\System32\jt.exe
		{
			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			Run, %A_WinDir%\system32\xcopy.exe /Y "\\cu.int\logonscripts\Installers\tasks\jt.exe" "%A_WinDir%\System32\",,Hide
			RunAs,
			Sleep, 1000
		}
		IfExist, %A_WinDir%\System32\jt.exe
			Run, %A_WinDir%\system32\cmd.exe /C "%A_WinDir%\System32\jt.exe /LJ %A_WinDir%\Tasks\DefragDriveC.job /SC %AdminUser% %AdminPW%",,Hide
	}
}

DisableWtime()
{

	RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\W32Time, Start, 4
}
	