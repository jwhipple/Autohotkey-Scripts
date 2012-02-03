; MainLogon.AHK
; Created by Joe Whipple for KEMBA Financial Credit Union on FEB 2009
; Licensed under the Creative Commons License
; Free to use/modify as long as this notice remains intact.
;
;
; THIS SHOULD ONLY BE USED AS AN EXAMPLE!
; THIS SCRIPT IS WHAT WE USE, IT WONT WORK FOR YOU UNLESS HEAVILY MODIFIED.
;
;
;

#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn UseUnsetLocal, Off

#Include JRWTools.ahk
#Include Crypt.ahk

RanManually =



;################################
;# Logon Script Global Settings #
;# Configure variables below    #
;################################

_SettingsINI = \\Network\path\to\MainLogon.ini

 


IfNotExist,%_SettingsINI%
{

	ExitApp
}

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

;NOTE!
;If you want to encrypt the username/password in the ini files, uncomment the next two lines... You have to put the encrypted strings in the INI file.
;To make an encrypted string, use the following:
;         MsgBox % Crypt.Encrypt.StrEncrypt(AdminUser,"ENTER A CRYPT HASH HERE",5,1)
;
;AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"ENTER A CRYPT HASH HERE",5,1)
;AdminPW := Crypt.Encrypt.StrDecrypt(AdminPW,"ENTER A CRYPT HASH HERE",5,1)

OS_Version := GetWindowsVersion()
OFFICEVER := GetOutlookVersion()





;#################################
;# Do not modify below this line #
;#################################


;Since this exact script should only be a reference, we kill it before it does anything.
ExitApp




SetTime(TimeServer) ;Set system time.

If HasWorkstation()
{
; ###################### RAN EVERY TIME ON WORKSTATIONS ###############################
	;ALL OS Versions
	
	MapDrives() ;Map standard drives.
	Clear_Symitar_Dropdowns()
	SymitarRegistry()
	
	If ( OS_Version == "Windows XP" )
	{
		RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Applets\Tour\RunCount, 0
		MapPrinters() ;Register printers
		;DefragmentTask()
    
	}
	else If ( OS_Version == "Windows 7" )
	{
		; Windows 7 specific tasks here
	}
	
	If OFFICEVER > 0
	{
		
		if (UserInOU("OU=Westerville"))
		{
			; Westerville has a local store for their PSTs.
			SetupOutlook(_SettingsINI,0,"WESTERVILLE Outlook")
		} else {
			SetupOutlook(_SettingsINI)
		}
		
		If OFFICEVER < 12
		{
			O2k7_File_Format_Converters()
		}
		
		If IsUserInADGroup("Group That Gets ServiceFolderAccess")
		{	
			; The people in the above group get an additional outlook profile named Service Folder.
			SetupOutlook(_SettingsINI,0,"SERVICE Outlook")
		} else {
			RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Service Folder
		}
	}
}

	
	SetDNSServers(NameServers, DomainName, DNSSearchOrder) 
	DisableNetBios()
		
	
	If HasWorkstation()
	{

		If ( OS_Version == "Windows XP" )
		{
			RegOLE() ; Fix OLE32 Registration
			FixTIFFAssoc()
		}
		
		
		Alternatiff() ;Register and setup Alternatiff
	
		CheckForFavorites()

		SymitarInstall()		
	
		RemoveDocPrinters() ; Remove Office installed Doc printers
		USBNoSleep() ; Dont powersave USB
		;PowerSave() ; Setup powersave settings
		DisableWtime()
		;MakeShortcuts() ;Make user shortcuts
		If UsersDepartment("Member Services,Accounting Department,Contact Center,Branch Operations,Operations,Operations Administration,Lending,Branch Administration,Tellers")
		{

			TrueChecksInstall() ; Install check verification software.
		}
	}
}

ExitApp









; Functions Below
; #################################################################################################################################################################

; Most of these functions below are KEMBA specific... however I left them here so you can see how some things are done (like installing an MSI file etc...)


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

		SplashImage, \\Network\Path\To\Pictures\logo.jpg, CWFFFFFF  h400 w500 b1 fs18, `n`nNow installing your anti-virus.`nAfter the install your PC will reboot automatically.
		ExitApp
		
		If OSBitVersion() = "x86"
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, \\Network\Path\To\Installers\ESET\NOD32Settings.exe,,
			RunWait, msiexec.exe /passive /forcerestart /i \\Network\Path\To\Installers\ESET\eavbe_nt32_enu.msi,,
			RunAs,
		}
		else
		{

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, \\cu.int\logonscripts\Installers\ESET\NOD32Settings.exe,,
			RunWait, msiexec.exe /passive /forcerestart /i \\Network\Path\To\Installers\ESET\eavbe_nt64_enu.msi,,
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
			RunWait, \\Network\Path\To\Installers\OfficePatch\FileFormatConverters.exe /q,,
			RunWait, \\Network\Path\To\Installers\OfficePatch\o2k7sp1pak.exe /q,,
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
		RunWait, %comspec% /c start msiexec /passive /i \\Network\Path\To\symformapi\SymFormAPI.msi,,
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
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Browser Options,Startup Page,\\Network\Path\To\Public\Symitar\html\startup.htm
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Browser Options,Information Page,http://inhouse.kemba.org/

	;Printer fix
	;RegDelete, HKEY_LOCAL_MACHINE, SOFTWARE\Symitar\SymForm, Printers
	
	; Symitar and Symform Fixes
	RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Symitar\SFW\2.0\Account Manager, BlockEMailFunction, 0
}


MapPrinters()
{
	if A_OSVersion in WIN_XP
	{
		Run, \\Network\Path\To\pushprinterconnections.exe,,Hide
	}
}




DisableWtime()
{

	RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\W32Time, Start, 4
}
	