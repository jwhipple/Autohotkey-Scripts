; JRWTools.ahk
; Created by Joe Whipple for KEMBA Financial Credit Union on FEB 2009
; Licensed under the Creative Commons License
; Free to use/modify as long as this notice remains intact.
; Parts of this script shamelessly stolen from others on the web and modified for my own purposes.
;
; Variables:
;
; %OVer1% = Outlook Version Number (eg: 11.0)

;#####################################
;## Global Variables                ##
;#####################################
_CurUsersGroups =  ;Contains the groups a user belongs to.
OSBitNumber =   ;Contains x86 for 32bit systems, AMD64 for 64 bit.
_distinguishedName = ;Contains the distinguished name of current user.
_userDepartment = ;Contains the department a user is in.
InstalledApps = ;Contains a list of all installed applications on the computer.



UserObjectADQuery(_object)
{
	global _distinguishedName, _userDepartment

	if (_object = "distinguishedName") && _distinguishedName
		return %_distinguishedName%

	if (_object = "department") && _userDepartment
		return %_userDepartment%
	
	objRootDSE := ComObjGet("LDAP://rootDSE")
	strDomain := objRootDSE.Get("defaultNamingContext")
	strADPath := "LDAP://" . strDomain
	objDomain := ComObjGet(strADPath)
	objConnection := ComObjCreate("ADODB.Connection")
	objConnection.Open("Provider=ADsDSOObject")
	objCommand := ComObjCreate("ADODB.Command")
	objCommand.ActiveConnection := objConnection
	;objFileSystem := ComObjCreate("Scripting.FileSystemObject")
	objCommand.CommandText := "<" . strADPath . ">" . ";(&(objectCategory=person)(objectClass=user)(sAMAccountName=" . A_UserName . "))" . ";" . _object . ";subtree"
	try
	{
		objRecordSet := objCommand.Execute
		objRecordCount := objRecordSet.RecordCount
	}
	catch
	{
		objRelease(objRootDSE)
		objRelease(objDomain)
		objRelease(objConnection)
		objRelease(objCommand)
		return
	}

	objOutputVar :=
	strObjectDN :=
	While !objRecordSet.EOF
	{
		strObjectDN := objRecordSet.Fields.Item(_object).value
		if %strObjectDN%
		{
			objRelease(objRootDSE)
			objRelease(objDomain)
			objRelease(objConnection)
			objRelease(objCommand)
			
			if (_object == "distinguishedName")
				_distinguishedName = %strObjectDN%

			if (_object == "department")
				_userDepartment = %strObjectDN%
				
				
			return %strObjectDN%
		}
		objRecordSet.MoveNext
	}
	objRelease(objRootDSE)
	objRelease(objDomain)
	objRelease(objConnection)
	objRelease(objCommand)
	return
}


InstallSoftware(_Software)
{
	global AdminUser, AdminPW, AdminDomain

	MyCommand := "msiexec /i " . _Software . " /quiet"
	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	RunWait, %comspec% /c start msiexec /i %_Software%,,
	RunAs,

}

UninstallSoftware(_Software)
{
	global AdminUser, AdminPW, AdminDomain
	

	objWMIService := ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\" . A_ComputerName . "\root\cimv2")
	WQLQuery = Select IdentifyingNumber, Name  FROM Win32_Product WHERE Name LIKE "%_Software%"
	colFolder := objWMIService.ExecQuery(WQLQuery)._NewEnum    	
	While colFolder[objFolder] {
		Text := objFolder.Name
			MyCommand := "msiexec /passive /uninstall " . objFolder.IdentifyingNumber . " /quiet"

			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RunWait, %MyCommand%
			RunAs,

	}
}


Check_Network_Alive(_PingHost)
{
	RunWait, %comspec% /c ping -n 1 -w 1000 %_PingHost% | find "Received = 1",,hide
	if ErrorLevel
		return false
	else
		return true
}

GetWindowsVersion()
{

	RegRead, ProductName, HKEY_LOCAL_MACHINE, Software\Microsoft\Windows NT\CurrentVersion, ProductName
	IfInString, ProductName, Vista
	{

		return "Windows Vista"
	}
	IfInString, ProductName, Windows 7
	{

		return "Windows 7"
	}
	IfInString, ProductName, XP
	{

		return "Windows XP"
	}
	IfInString, ProductName, 2003
	{

		return "Windows 2003"
	}
	IfInString, ProductName, 2000
	{

		return "Windows 2000"
	}
	return %ProductName%
}

GetUserHomeDir()
{

	IfExist, C:\Users\%A_UserName%
	{
		X = C:\Users\%A_UserName%

		return %X%
	}
	IfExist, C:\Documents and Settings\%A_UserName%
	{
		X = C:\Documents and Settings\%A_UserName%

		return %X%
	}
	return
}


TurnOnNumlock()
{

	SetNumLockState, On
}




HasWorkstation()
{

	if A_OSVersion in WIN_95,WIN_98,WIN_ME,WIN_XP,WIN_VISTA,WIN_7  ; Note: No spaces around commas.
	{

		return true
	}

	return false
}

IsLocalDrive(_Path)
{
	If RegExMatch(_Path, "^\\")
		return 0

	StringLeft, _Path, _Path, 1
	DriveGet, NetworkDrives, List, NETWORK
	
	IfInString, NetworkDrives, %_Path%
		return 0
	else
		return 1
}

Hex(Inp,UC = 0)
{
   ; UC = 0 standard ascii, UC = 1 unicode, UC = 2 unicode with appended nulls for outlook registry settings.
   Result =
   OldFmt = %A_FormatInteger%
   SetFormat, Integer, hex

   Loop, Parse, Inp
   {
      TransForm, Asc, Asc, %A_LoopField%
      Asc += 0
      StringTrimLeft, Hex, Asc, 2
      IfEqual, UC, 0
         Result = %Result%%Hex%
      Else
         Result = %Result%%Hex%00
   }
   IfEqual, UC, 2
     Result = %Result%0000
	 
   SetFormat, Integer, %OldFmt%
   StringUpper, Result, Result

   Return Result
}

Asc(Inp,UC = 0)
{
   Result =
   StringLen, Len, Inp
   Len /= 2

   OldFmt = %A_FormatInteger%
   SetFormat, Integer, D

   Loop, %Len%
   {
      StringLeft, Hex, Inp, 2
      IfEqual, UC, 0
         StringTrimLeft, Inp, Inp, 2
      Else
         StringTrimLeft, Inp, Inp, 4
      
      Hex = 0x%Hex%
      Hex += 0
      
      TransForm, Chr, Chr, %Hex%
      Result = %Result%%Chr%
   }
   SetFormat, Integer, %OldFmt%

   Return Result
}

SetDNSServers(_NameServers, _DomainName, _DNSSearchOrder)
{

	
	;Set the search order for DNS suffixes.
	RegWrite, REG_SZ, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\Tcpip\Parameters, SearchList, %_DNSSearchOrder% 
	;This will change the DNS servers for the TCP interface with a gateway present.
	Loop, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces, 1, 1
	{
		if a_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		if a_LoopRegName = DefaultGateway
		{
			If (value <> "")
			{
				RegWrite, REG_SZ, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, NameServer, %_NameServers% ;Setting the nameservers.
				RegWrite, REG_SZ, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, Domain, %_DomainName% ;Setting the default domain.
			}	
		}
	}
}

DisableNetBios()
{	
	; This disables NetBIOS on all network interfaces.
	; KEMBA uses DNS only to resolve computer names... This is good practice on large networks.
	global AdminUser, AdminPW, AdminDomain

	;Disables NetBios on ALL network adapters.
	Loop, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces, 1, 1
	{
		if a_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		if a_LoopRegName = NetbiosOptions
		{
			RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
			RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, %A_LoopRegSubKey%, NetBiosOptions, 2
			RunAs,
		}
	}
}

OSBitVersion()
{
	; Returns if a system is 32bit "x86" or 64bit "AMD64"
	global OSBitNumber
	If not OSBitNumber
	{
		RegRead, OSBitNumber, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Session Manager\Environment, Processor_Architecture
	}
	return OSBitNumber
}

UsersDepartment(_Dept)
{
	strObjectDN := UserObjectADQuery("department")
	Loop, parse, _Dept, `,
	{
		If A_LoopField = %strObjectDN%
		{
			return true
		}
	}
	return false
}


UsersExt()
{
	; We use the ipphone value in Active Directory as the extension.
	return UserObjectADQuery("ipphone")
}

UsersHomeLocation()
{
	; Returns the physical office name set in Active Directory.
	return UserObjectADQuery("physicalDeliveryOfficeName")
}

UserIsMemberOf(_User)
{
	;Returns a carriage return seperated list of the group(s) the user belongs to. 

	_UsersGroups =
	
	StringLeft, UserNameStart, _User, 3
	StringUpper, UserNameStart, UserNameStart
	If UserNameStart != "CN=" ; We were given a simple name for the group so we find the distinguished name.
	{
		UserName := FindDistinguishedName(_User)

	} else {
		UserName := %_User%
	}
	
	objRootDSE := ComObjGet("LDAP://rootDSE")
	strDomain := objRootDSE.Get("defaultNamingContext")
	strADPath := "LDAP://" . strDomain
	objDomain := ComObjGet(strADPath)
	objConnection := ComObjCreate("ADODB.Connection")
	objConnection.Open("Provider=ADsDSOObject")
	objCommand := ComObjCreate("ADODB.Command")
	objCommand.ActiveConnection := objConnection
	objCommand.CommandText := "<" . strADPath . ">" . ";(&(&(&(objectCategory=group)(member=" . UserName . "))));Name;subtree"
	objRecordSet := objCommand.Execute
	objRecordCount := objRecordSet.RecordCount
	objOutputVar :=
	While !objRecordSet.EOF
	{
		strObjectDN := objRecordSet.Fields.Item("Name").value
		_UsersGroups = %_UsersGroups%`n%strObjectDN%
		objRecordSet.MoveNext
	}
	objRelease(objRootDSE)
	objRelease(objDomain)
	objRelease(objConnection)
	objRelease(objCommand)
	if _UsersGroups
	{
		_UsersGroups = %_UsersGroups%`nDomain Users
	}
	return _UsersGroups 
}


IsUserInADGroup(_GroupName)
{
	;Determines if a user is in _GroupName
	global _CurUsersGroups

	
	If not _CurUsersGroups
	{
		_CurUsersGroups := UserIsMemberOf(A_Username)
	}
	
	Loop, parse, _CurUsersGroups, `n
	{
		IfEqual,_GroupName,%A_LoopField%
		{
			return true
		}
	}

	return false
}



UserInOU(_OUName)
{
	; Determines if a user is in an active directory OU.
	;Ext := FindDistinguishedName(A_UserName)
	Ext := UserObjectADQuery("distinguishedName")
	IfInString, Ext, %_OUName%
	{
		return true
	}
	return false
}


; Map network drive based on AD group membership
MapDrivesByGroup(_GroupName,_DriveLetter,_DrivePath)
{
	;If user is in _GroupName, map the drive letter _DriveLetter to _DrivePath

		If IsUserInADGroup(_GroupName)
		{

			RegExMatch(_DriveLetter, "[e-zE-Z]",_DriveVar)
			IfNotExist, %_DriveVar%:\
			{

				RunWait, %A_WinDir%\system32\net.exe use %_DriveVar%: %_DrivePath%,,Hide
			}
		}
}

; Remove mapped drive as needed by drive letter
RemoveMappedDrive(_RDriveLetter)
{


	RegExMatch(_RDriveLetter, "[e-zE-Z]",_RDriveVar)
		IfExist, %_RDriveVar%:\
		{

			RunWait, %A_WinDir%\system32\net.exe use /DELETE %_RDriveVar%: ,,Hide
		}
}


RemoveDocPrinters()
{
	; Removes those annoying Microsoft Office "fake" printers.
	global AdminUser, AdminPW, AdminDomain

	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	Run, %A_WinDir%\system32\cscript.exe //T:10 c:\windows\system32\prnmngr.vbs -d -p "Microsoft Office Live Meeting Document Writer",,Hide
	Run, %A_WinDir%\system32\cscript.exe //T:10 c:\windows\system32\prnmngr.vbs -d -p "Microsoft Office Document Image Writer",,Hide
	Run, %A_WinDir%\system32\cscript.exe //T:10 c:\windows\system32\prnmngr.vbs -d -p "Microsoft XPS Document Writer",,Hide
	RunAs,
}

USBNoSleep()
{
	global AdminUser, AdminPW, AdminDomain

	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\services\USB, DisableSelectiveSuspend, 1
	RunAs,
}

SetTime(_NetworkTimeServer)
{
	global AdminUser, AdminPW, AdminDomain

	RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
	Run, %A_WinDir%\system32\net.exe time \\%_NetworkTimeServer% /SET /Y,,Hide
	RunAs,
}

StdoutToVar_CreateProcess(sCmd, bStream = "", sDir = "", sInput = "")
{
	DllCall("CreatePipe", "UintP", hStdInRd , "UintP", hStdInWr , "Uint", 0, "Uint", 0)
	DllCall("CreatePipe", "UintP", hStdOutRd, "UintP", hStdOutWr, "Uint", 0, "Uint", 0)
	DllCall("SetHandleInformation", "Uint", hStdInRd , "Uint", 1, "Uint", 1)
	DllCall("SetHandleInformation", "Uint", hStdOutWr, "Uint", 1, "Uint", 1)
	VarSetCapacity(pi, 16, 0)
	NumPut(VarSetCapacity(si, 68, 0), si)	; size of si
	NumPut(0x100	, si, 44)		; STARTF_USESTDHANDLES
	NumPut(hStdInRd	, si, 56)		; hStdInput
	NumPut(hStdOutWr, si, 60)		; hStdOutput
	NumPut(hStdOutWr, si, 64)		; hStdError
	If Not	DllCall("CreateProcess", "Uint", 0, "Uint", &sCmd, "Uint", 0, "Uint", 0, "int", True, "Uint", 0x08000000, "Uint", 0, "Uint", sDir ? &sDir : 0, "Uint", &si, "Uint", &pi)	; bInheritHandles and CREATE_NO_WINDOW
		ExitApp
	DllCall("CloseHandle", "Uint", NumGet(pi,0))
	DllCall("CloseHandle", "Uint", NumGet(pi,4))
	DllCall("CloseHandle", "Uint", hStdOutWr)
	DllCall("CloseHandle", "Uint", hStdInRd)
	If	sInput <>
	DllCall("WriteFile", "Uint", hStdInWr, "Uint", &sInput, "Uint", StrLen(sInput), "UintP", nSize, "Uint", 0)
	DllCall("CloseHandle", "Uint", hStdInWr)
	bStream+0 ? (bAlloc:=DllCall("AllocConsole"),hCon:=DllCall("CreateFile","str","CON","Uint",0x40000000,"Uint",bAlloc ? 0 : 3,"Uint",0,"Uint",3,"Uint",0,"Uint",0)) : ""
	VarSetCapacity(sTemp, nTemp:=bStream ? 64-nTrim:=1 : 4095)
	Loop
		If	DllCall("ReadFile", "Uint", hStdOutRd, "Uint", &sTemp, "Uint", nTemp, "UintP", nSize:=0, "Uint", 0)&&nSize
		{
			NumPut(0,sTemp,nSize,"Uchar"), VarSetCapacity(sTemp,-1), sOutput.=sTemp
			If	bStream
				Loop
					If	RegExMatch(sOutput, "[^\n]*\n", sTrim, nTrim)
						bStream+0 ? DllCall("WriteFile", "Uint", hCon, "Uint", &sTrim, "Uint", StrLen(sTrim), "UintP", 0, "Uint", 0) : %bStream%(sTrim), nTrim+=StrLen(sTrim)
					Else	Break
		}
		Else	Break
	DllCall("CloseHandle", "Uint", hStdOutRd)
	bStream+0 ? (DllCall("Sleep","Uint",1000),hCon+1 ? DllCall("CloseHandle","Uint",hCon) : "",bAlloc ? DllCall("FreeConsole") : "") : ""
	Return	sOutput
}

CreateFilePath(_FilePath)
{

	IfNotExist, %_FilePath%
	{

		FileCreateDir, %_FilePath%
	}
}

FindDistinguishedName(_Item)
{
	;This finds a full DN name from a short name or a samaccount name.
	If _Item = %A_UserName%
	{
		;Pass to a more efficient function.
		_Item := UserObjectADQuery("distinguishedName")
	} else {
		;MsgBox, FindDistinguishedName(_Item)
		MembersOfGroup := Object()
		objRootDSE := ComObjGet("LDAP://rootDSE")
		strDomain := objRootDSE.Get("defaultNamingContext")
		strADPath := "LDAP://" . strDomain
		objDomain := ComObjGet(strADPath)
		objConnection := ComObjCreate("ADODB.Connection")
		objConnection.Open("Provider=ADsDSOObject")
		objCommand := ComObjCreate("ADODB.Command")
		objCommand.ActiveConnection := objConnection

		objCommand.CommandText := "<" . strADPath . ">;(|(name=" . _Item . ")(sAMAccountName=" . _Item . "));distinguishedName;subtree"
		objRecordSet := objCommand.Execute
		objRecordCount := objRecordSet.RecordCount
		objOutputVar :=
		While !objRecordSet.EOF
		{
		  _Item := objRecordSet.Fields.Item("distinguishedName").value
		  objRecordSet.MoveNext
		}
		objRelease(objRootDSE)
		objRelease(objDomain)
		objRelease(objConnection)
		objRelease(objCommand)
	}
	return _Item
}


IsInstalled(_Software)
{
	;Check registry for installed program.
	;Must match the name set in DisplayName in the registry path HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
	global InstalledApps, _SettingsINI

	StringLower, _Software, _Software
	IniRead, InvHelperApp, %_SettingsINI%, LogonScript, InvHelperApp
	
	IfExist, %InvHelperApp%
	{
		_Use64BitHelper = 1
	}
	
	If not InstalledApps
	{
		;Since this can be a time consuming task, we gather all the results into the InstalledApps variable so it is quicker next time its ran.
		
		If (OSBitVersion() = "x86" or _Use64BitHelper = 0)
		{
			REGPATH =SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
			Loop, HKEY_LOCAL_MACHINE, %REGPATH%, 1, 1
			{
				If A_LoopRegName = DisplayName
				{
					RegRead, value
					StringLower, value, value
					InstalledApps = %InstalledApps%%value%`n
				}
			}
			
		} else {
		
			DetectHiddenWindows, on
			DetectHiddenText, on
			Try
			{
				Run %InvHelperApp%, , , MyPID
				WinWait, Search64BitReg.exe,,5
				If Not ErrorLevel
				{
					ControlGet, List, List,, ListBox1, Search64BitReg.exe
					Loop, Parse, List, `n
						InstalledApps = %InstalledApps%%A_LoopField%`n
				}
				WinClose, Search64BitReg.exe
				;MsgBox, %InstalledApps%
			}
			catch
			{
				;MsgBox, Could not run helper.
			}
		}
	}
	
	Loop, Parse, InstalledApps, `n
	{
		If _Software = %A_LoopField%
		{

			return true
		}
	}
	
	

	return false
}



ProgressMeter(_Percent=0, _Message="Message Not Set", _2Message="We recommend waiting until this is complete before using your PC.", _TitleBar="KEMBA Login Script")
{
	global ProgressBarDisplay
	If ProgressBarDisplay
		Progress, %_Percent%, %_Message%, %_2Message%, %_TitleBar%
	return
}


; ####################### MS OFFICE FUNCTIONS ####################### 
; ####################### MS OFFICE FUNCTIONS ####################### 
; ####################### MS OFFICE FUNCTIONS ####################### 
; ####################### MS OFFICE FUNCTIONS ####################### 
; ####################### MS OFFICE FUNCTIONS ####################### 
; ####################### MS OFFICE FUNCTIONS ####################### 


SetOfficeTemplatesLocation(_OfficeTemplates)
{
	global OFFICEVER
	;This sets a network location for Word so everyone can have access to things like Fax templates, letterhead, etc...
	OFFICEVER := GetOutlookVersion()
	RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Office\%OFFICEVER%\Common\General, SharedTemplates, %_OfficeTemplates%
}

HasOffice(_OfficeVersion=-1)
{
	; We create a Word object, then check for the version... Thanks to the site: http://blogs.technet.com/b/heyscriptingguy/archive/2005/01/10/how-can-i-determine-which-version-of-word-is-installed-on-a-computer.aspx

	try
	{
		objCommand := ComObjCreate("Word.Application")
		_OVersion := objCommand.Version
		objRelease(objCommand)
	} catch e {
		_OVersion = 0
	}
	If _OfficeVersion == -1 and _OVersion > 0
		return true
	if _OVersion = %_OfficeVersion%
		return true

	return false
}



; ####################### OUTLOOK FUNCTIONS ####################### 
; ####################### OUTLOOK FUNCTIONS ####################### 
; ####################### OUTLOOK FUNCTIONS ####################### 
; ####################### OUTLOOK FUNCTIONS ####################### 
; ####################### OUTLOOK FUNCTIONS ####################### 
; ####################### OUTLOOK FUNCTIONS ####################### 


CleanUpServiceFolderAccess()
{
	;Old Installs Cleanup.
	Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 1, 0
	{
		If A_LoopRegType = KEY
		{
			IfEqual, A_LoopRegName, Service Folder
			{
				;Keep this folder
			} else IfEqual, A_LoopRegName, Exchange2007
			{
				;Also keep this folder
			} else {
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\%A_LoopRegName%
			}
		}
	}
}



; ###### Completed Functions Below ######

CleanOutlookTemp()
{	

	Loop, HKEY_CURRENT_USER, Software\Microsoft\Office, 1, 1
	{
		if a_LoopRegName = OutlookSecureTempFolder
		{
			if a_LoopRegType = key
				value =
			else
			{
				RegRead, value
				if ErrorLevel
					value = *error*
			}
			FileSetAttrib, -R, %value%*.*
			FileDelete, %value%*.*
		}
	}
}

GetOutlookVersion()
{

	If A_UserName = administrator
	{
		; We generally do not want to set up outlook for administrator users. Comment out the next two lines if you do.

		return 0
	}
	
	try
	{
		RegRead,OutlookPath,HKEY_LOCAL_MACHINE, SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE, Path
		SplitPath, OutlookPath, name, dir, ext, name_no_ext, drive
		StringUpper, dir, dir
		FoundPos := RegExMatch(dir, "OFFICE([0-9]+)", SubPat)
		FoundPos := RegExMatch(SubPat, "([0-9]+)", SubPat)

		If Not SubPat
		{
			return 0
		}
		else
		{
			SubPat = %SubPat%.0
			return SubPat
		}

	} catch e {
		SubPat = 0
	}
	return 0
}


CloseOutlook()
{

	OulookPID = 0
	Process, Exist, OUTLOOK.EXE
	if ErrorLevel != 0
	{
		OutlookPID = %ErrorLevel%

	}

	if (OutlookPID > 0)
	{

		WinClose, ahk_pid %OutlookPID%,,10
		Process, Exist, %OutlookPID%
		if ErrorLevel != 0
		{

			WinKill, ahk_pid %OutlookPID%,,10
		} else {

		}

		
		Process, WaitClose, %OutlookPID%, 20
		if ErrorLevel 
		{

			Run, taskkill /IM OUTLOOK.EXE
			loop
			{
				Process, Exist, %OutlookPID%
				If ( ErrorLevel == 0 )
				{
					break

				}

				sleep, 1000
			}
		}
		
	} else {

		SetTitleMatchMode, 2
		WinClose, Outlook

		WinWaitClose, Outlook,, 20
		If ErrorLevel
		{

			Run, taskkill /IM OUTLOOK.EXE
			loop
			{
				Process, Exist, OUTLOOK.EXE
				If (ErrorLevel == 0)
				{
					break

				}

				sleep, 1000
			}
		}
	}
}

SetupOutlook(_SettingsINI, wipe=0, _MailSectionName="Outlook")
{
	; Usage:  SetupOutlook(<ini settings file>, <ignore existing profile? 0|1>, <Section in ini file to look for outlook settings>)
	; Requires: GetOutlookVersion() OutlookProfileExist() RunBackgroundOutlook() OutlookProfileCount() ParseSignatureFiles() SetOutlookSignatureNames()
	global OFFICEVER
	
	; Check to make sure settings file exists.
	IfNotExist,%_SettingsINI%
	{
		MsgBox,Can't Find %_SettingsINI%!`nPlease let MIS know.`n
		return
	}

	IniRead, MailServer, %_SettingsINI%, %_MailSectionName%, MailServer
	IniRead, DoAutoArchive, %_SettingsINI%, %_MailSectionName%, DoAutoArchive
	IniRead, AutoArchiveFile, %_SettingsINI%, %_MailSectionName%, AutoArchiveFile
	IniRead, prfTemplate, %_SettingsINI%, %_MailSectionName%, prfTemplate
	IniRead, UsersPRF, %_SettingsINI%, %_MailSectionName%, UsersPRFPath
	IniRead, MailProfile, %_SettingsINI%, %_MailSectionName%, ProfileName 
	IniRead, DefaultProfile, %_SettingsINI%, %_MailSectionName%, DefaultProfile
	IniRead, OverwriteProfile, %_SettingsINI%, %_MailSectionName%, OverwriteProfile
	IniRead, MailboxName, %_SettingsINI%, %_MailSectionName%, MailboxName
	SplitPath, UsersPRF,, UserPRFdir
	StringReplace, UserPRFdir, UserPRFdir, `%USERNAME`%, %A_UserName%, ALL
	
	IfNotExist, %UserPRFdir%
	{
		FileCreateDir, %UserPRFdir%
	}
	IfNotExist, %UserPRFdir%
	{
		progress, Off
		SplashImage, Off
		MsgBox, Cant create PRF directory!
		return
	}
	
	UsersPRF := UsersPRF . MailProfile . ".prf"
		
	StringReplace, AutoArchiveFile, AutoArchiveFile, `%USERNAME`%, %A_UserName%, ALL
	StringReplace, prfTemplate, prfTemplate, `%USERNAME`%, %A_UserName%, ALL
	StringReplace, UsersPRF, UsersPRF, `%USERNAME`%, %A_UserName%, ALL
	StringReplace, MailProfile, MailProfile, `%USERNAME`%, %A_UserName%, ALL 
	StringReplace, MailboxName, MailboxName, `%USERNAME`%, %A_UserName%, ALL

	If wipe
	{
		Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 2, 0
		{
			RegDelete
		}
		Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Deleted Profiles, 2, 0
		{
			RegDelete
		}
	}
	
	If OutlookProfileExist(MailProfile)
	{
		IfExist, %UsersPRF%
		{
			; We have Outlook configured and also a PRF file... life is good.
			
		} else {
		
			; NO PRF, but we got an Outlook configured.
			; Create a PRF file for user.

			FileRead,CustomPRFFile,%prfTemplate%

			StringReplace, CustomPRFFile, CustomPRFFile, _PROFILENAME_, %MailProfile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _AUTOARCHIVEFILE_, %AutoArchiveFile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _MAILBOXNAME_, %MailboxName%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _MAILSERVERNAME_, %MailServer%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _DOAUTOARCHIVE_, %DoAutoArchive%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _DEFAULTPROFILE_, %DefaultProfile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _OVERWRITEPROFILE_, %OverwriteProfile%, ALL
			
			StringSplit, PRFFileArray, CustomPRFFile,~

			

			PSTCount := 4
			Loop, HKEY_CURRENT_USER, Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\%MailProfile%, 1, 1
			{
				if a_LoopRegName = 001f6700
				{
					RegRead, value
					jResult := Asc(value)
					FoundPos := RegExMatch(jResult, "^[a-eA-E]:\\")
					; PST was a network accessible one.
					IfExist, %jResult%  ;If the PST actually exists on the drive, we add it to the Outlook config.
					{

						StringUpper, jResult, jResult
						PRFFileArray1 := PRFFileArray1 . "`nService" . PSTCount . "=Personal Folders"
						PRFFileArray2 := PRFFileArray2 . "`n[Service" . PSTCount . "]"
						PRFFileArray2 := PRFFileArray2 . "`nUniqueService=No"
						PRFFileArray2 := PRFFileArray2 . "`nPathToPersonalFolders=" . jResult . "`n"
						PSTCount++
					}
					else
					{

						PSTHasBeenMoved = 1 ; The PST that is listed does not exist, so Outlook config is invalid.
					}
				}	
			}

			PRFFileArray1 := PRFFileArray1 . "`n`n"
			CustomPRFFile  := PRFFileArray1 . PRFFileArray2 . PRFFileArray3
			StringReplace, CustomPRFFile, CustomPRFFile, ~,, All 
			

			
			file := FileOpen(UsersPRF, "w")
			file.write(CustomPRFFile)
			file.close()				
		}
		
		
	} else {
	
		;No Outlook profile configured, check for PRF.
		
		; Check for Outlook.
		OFFICEVER := GetOutlookVersion()
		If Not OFFICEVER
		{
			return
		}
	
		Run, taskkill /T /F /IM OUTLOOK.EXE,,Hide
	
		RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0
		
		IfExist, %UsersPRF%
		{
			; We have a PRF file but no Outlook Profile. Tell Outlook to use the PRF file.
			loop, 4
			{
				;RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,First-Run
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,FirstRun
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,ImportPRF
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,CreateWelcome
				RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0
				RegRead,CurrentDefaultProfile,HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile 

				If CurrentDefaultProfile
				{
					;MsgBox, CurrentDefaultProfile
				} else {
					RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile,%MailProfile%
				}

				RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,ImportPRF,%UsersPRF%
				RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\Common\,QMEnable,0
				RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,CreateWelcome,4294967295
				
				Sleep, 2000 
				If RunBackgroundOutlook()
				{
					break
				}
			}			
			

			
			
		} else {
		
			; No Outlook Profile and no PRF file, use a generic one.
			
			FileRead,CustomPRFFile,%prfTemplate%
			StringReplace, CustomPRFFile, CustomPRFFile, _PROFILENAME_, %MailProfile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _AUTOARCHIVEFILE_, %AutoArchiveFile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _MAILBOXNAME_, %MailboxName%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _MAILSERVERNAME_, %MailServer%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _DOAUTOARCHIVE_, %DoAutoArchive%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _DEFAULTPROFILE_, %DefaultProfile%, ALL
			StringReplace, CustomPRFFile, CustomPRFFile, _OVERWRITEPROFILE_, %OverwriteProfile%, ALL

			file := FileOpen(UsersPRF, "w")
			file.write(CustomPRFFile)
			file.close()			
			
			
			loop, 4
			{

		

				;RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,First-Run
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,FirstRun
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,ImportPRF
				RegDelete, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,CreateWelcome
				
				RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0

				RegRead,CurrentDefaultProfile,HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile 
				If CurrentDefaultProfile
				{

				} else {

					RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, DefaultProfile,%MailProfile%
				}

				RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,ImportPRF,%UsersPRF%
				RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\Common\,QMEnable,0
				RegWrite, REG_DWORD, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Outlook\Setup\,CreateWelcome,4294967295
				
				Sleep, 2000 
				If RunBackgroundOutlook()
				{
					break
				}
			}
			
		}

		If OutlookProfileCount() > 1
		{
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 1
		}
		
		; Set up name and initials so it wont prompt.
		Company = UserObjectADQuery("company")
		FullName := UserObjectADQuery("name")
		Initials =
		StringSplit, Names, FullName, %A_Space%
		Loop, %Names0%
		{
			this_color := Names%a_index%
			Initial := SubStr(this_color, 1, 1)
			Initials = %Initials%%Initial%
		}
		
		If OFFICEVER >= 14.0
		{
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\Common\UserInfo, Company, %Company%
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\Common\UserInfo, UserInitials, %Initials%
			RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\Common\UserInfo, UserName, %FullName%		
		} else {
			Company  := Hex(Company,2)
			Initials := Hex(Initials,2)
			FullName := Hex(FullName,2)
			RegWrite, REG_BINARY, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Common\UserInfo, Company, %Company%
			RegWrite, REG_BINARY, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Common\UserInfo, UserInitials, %Initials%
			RegWrite, REG_BINARY, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%OFFICEVER%\Common\UserInfo, UserName, %FullName%	
		}
	}
	
	;Setup Signatures Here
	IniRead, signatures, %_SettingsINI%, %_MailSectionName%, signatures
	IniRead, NoSignatures, %_SettingsINI%, %_MailSectionName%, NoSignatures
	IniRead, fullsignaturetemplate, %_SettingsINI%, %_MailSectionName%, fullsignaturetemplate
	IniRead, replysignaturetemplate, %_SettingsINI%, %_MailSectionName%, replysignaturetemplate
	
	If signatures
	{
		IfNotInString, NoSignatures, %A_UserName%
		{

			;Set up signatures if the flag is set in the INI file and they are not listed as no signatures.
			ParseSignatureFiles(fullsignaturetemplate)
			ParseSignatureFiles(replysignaturetemplate)
			SplitPath, fullsignaturetemplate, fullname
			SplitPath, replysignaturetemplate, replyname
			SetOutlookSignatureNames(MailProfile, fullname, replyname)
		}
	}

	CleanOutlookTemp() ;If this temp file fills up, then they wont be able to open attachments in Outlook.
	
	return
}

CreateOutlookPRF()
{


	; Create a PRF file for user.


	FileRead,CustomPRFFile,%prfTemplate%

	StringReplace, CustomPRFFile, CustomPRFFile, _PROFILENAME_, %MailProfile%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _AUTOARCHIVEFILE_, %AutoArchiveFile%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _MAILBOXNAME_, %MailboxName%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _MAILSERVERNAME_, %MailServer%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _DOAUTOARCHIVE_, %DoAutoArchive%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _DEFAULTPROFILE_, %DefaultProfile%, ALL
	StringReplace, CustomPRFFile, CustomPRFFile, _OVERWRITEPROFILE_, %OverwriteProfile%, ALL
	
	StringSplit, PRFFileArray, CustomPRFFile,~

	

	PSTCount := 4
	Loop, HKEY_CURRENT_USER, Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\%MailProfile%, 1, 1
	{
		if a_LoopRegName = 001f6700
		{
			RegRead, value
			jResult := Asc(value)
			FoundPos := RegExMatch(jResult, "^[a-eA-E]:\\")
			; PST was a network accessible one.
			IfExist, %jResult%  ;If the PST actually exists on the drive, we add it to the Outlook config.
			{

				StringUpper, jResult, jResult
				PRFFileArray1 := PRFFileArray1 . "`nService" . PSTCount . "=Personal Folders"
				PRFFileArray2 := PRFFileArray2 . "`n[Service" . PSTCount . "]"
				PRFFileArray2 := PRFFileArray2 . "`nUniqueService=No"
				PRFFileArray2 := PRFFileArray2 . "`nPathToPersonalFolders=" . jResult . "`n"
				PSTCount++
			}
			else
			{

				PSTHasBeenMoved = 1 ; The PST that is listed does not exist, so Outlook config is invalid.
			}
		}
		
	}
	

	PRFFileArray1 := PRFFileArray1 . "`n`n"
	CustomPRFFile  := PRFFileArray1 . PRFFileArray2 . PRFFileArray3
	StringReplace, CustomPRFFile, CustomPRFFile, ~,, All 
	

	
	file := FileOpen(UsersPRF, "w")
	file.write(CustomPRFFile)
	file.close()


}
OutlookProfileCount()
{
	;If we have more than one profile, set the flag that allows us to choose which profile we want.
	X = 0
	Loop, HKEY_CURRENT_USER, Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 2, 0
	{
		X++
	}
	return X
}

RunBackgroundOutlook()
{

	try
	{
		;Tell Outlook to configure itself using COM objects.
		;SplashImage, \\cu.int\logonscripts\Pictures\logo.jpg, CWFFFFFF  h600 w800 b1 fs18,`n`n`nNow configuring your e-mail client.`n`nPLEASE WAIT`n`nDo not use your keyboard or mouse yet.

		;The FlipOutlookLoginBit function is for people that set the "Always prompt for login credentials" setting. It bypasses this so we can run Outlook in the background without prompts.
		
		objCommand := ComObjCreate("Outlook.Application")
		objNameSpace := objCommand.GetNameSpace("MAPI")
		objFolder := objNameSpace.GetDefaultFolder(6)
		objFolder.unreaditemcount
		objRelease(objCommand)
		return 1
	}
	catch e
	{

		sleep 1000
		return 0
	}

}

OutlookProfileExist(_ProfileName) ; Check to see if the user has the specified profile in Outlook.
{

	; I changed the following because contact center could have multiple profiles.
	Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 1, 0
	{
		If A_LoopRegType = KEY
		{
			If A_LoopRegName = %_ProfileName%
			{

				return true
			}
		}
	}

	return false
}

RestoreOutlookLoginBit(_RegistryKeys)
{
	return
	Loop, parse, _RegistryKeys,`n
	{
		RegRead, value, HKEY_CURRENT_USER, %A_LoopField%, 00036601		
		If value
		{
			StringMid, x, value, 2, 1
			StringLeft, y, value, 1
			StringRight, z, value, 6
			
			if x = 4
			{
				x = %y%C%z%
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopField%, 00036601, %x%
			}
			if x = 0 
			{
				x = %y%8%z%	
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopField%, 00036601, %x%
			}
		}
	}

}

FlipOutlookLoginBit()
{
	;The FlipOutlookLoginBit function is for people that set the "Always prompt for login credentials" setting. It bypasses this so we can run Outlook in the background without prompts.
	regvals =
	Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 1, 1
	{
		IfEqual, A_LoopRegName,00036601
		{
			RegRead, value
			StringMid, x, value, 2, 1
			StringLeft, y, value, 1
			StringRight, z, value, 6
			if x = C
			{
				x = 4
				x = %y%%x%%z%
				regvals = %regvals%%A_LoopRegSubKey%`n
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopRegSubKey%, 00036601, %x%
				sleep, 1000
			}
			if x = 8 
			{
				x = 0
				x = %y%%x%%z%
				regvals = %regvals%%A_LoopRegSubKey%`n
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopRegSubKey%, 00036601, %x%
				sleep, 1000
			}
		}
	}
	return %regvals%
}

; ####################### OUTLOOK SIGNATURE FUNCTIONS ####################### 
; ####################### OUTLOOK SIGNATURE FUNCTIONS ####################### 
; ####################### OUTLOOK SIGNATURE FUNCTIONS ####################### 
; ####################### OUTLOOK SIGNATURE FUNCTIONS ####################### 
; ####################### OUTLOOK SIGNATURE FUNCTIONS ####################### 


ParseSignatureFiles(_FileTemplateBaseName)
{
	;Where Outlook keeps users local copy of signatures.
	SignaturePath = %A_AppData%\Microsoft\Signatures\
	
	IfNotExist, %SignaturePath%
	{
		;If the signature directory is not there, create it.
		FileCreateDir, %SignaturePath%
	}
	;Base filename of the file.
	SplitPath, _FileTemplateBaseName, fullname

	;These are the file extensions supported by Outlook.

		FileRead,filecontents1,%_FileTemplateBaseName%.txt
		FileRead,filecontents2,%_FileTemplateBaseName%.htm
		FileRead,filecontents3,%_FileTemplateBaseName%.rtf

		
		;Flag keeps track to see if we have hit a variable that needs replaced.
		flag = 0
		a =
		word =
		
		Loop, parse, filecontents1
		{
			x = %A_LoopField%
			if x=_ 
			{
				if flag
				{
					flag = 0
					a := UserObjectADQuery(word)
					StringReplace, filecontents1, filecontents1, _%word%_, %a%, ALL
					StringReplace, filecontents2, filecontents2, _%word%_, %a%, ALL
					StringReplace, filecontents3, filecontents3, _%word%_, %a%, ALL
					word =
					a =
				} else 
				{
					flag = 1
				}
				continue
			}
			if flag
			{
				word = %word%%x%
			}
		}
		;Write the new signature file out using the users path.
		file := FileOpen(SignaturePath . fullname . ".txt", "w")
		file.write(filecontents1)
		file.close()
		
		file := FileOpen(SignaturePath . fullname . ".htm", "w")
		file.write(filecontents2)
		file.close()
		
		file := FileOpen(SignaturePath . fullname . ".rtf", "w")
		file.write(filecontents3)
		file.close()
		
	return
}

SetOutlookSignatureNames(_Profile, _Fullsig, _Replysig)
{
	; The values for the signature files names are currently hard set for full and reply. Figure out the vaules for fullsig and replysig if you need to change them.


	Loop, HKEY_CURRENT_USER, Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\%_Profile%\9375CFF0413111d3B88A00104B2A6676, 1, 1
	{
		RegRead, value
		If Asc(value) = "MicrosoftExchange" || Asc(value) = "MSEMS"
		{
			_full := Hex(_Fullsig,2)
			_reply := Hex(_Replysig,2)

			RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopRegSubKey% , New Signature, %_full%
			RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopRegSubKey% , Reply-Forward Signature, %_reply%
		}
	}
}
