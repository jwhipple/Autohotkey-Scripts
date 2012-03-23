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

#Include \\cu.int\logonscripts\Sources\Crypt.ahk

_SettingsINI = \\CU.INT\logonscripts\MainLogon.ini
IfNotExist,%_SettingsINI%
{
	ExitApp
}

IniRead, AdminUser, %_SettingsINI%, Admin, AdminUser
IniRead, AdminPW, %_SettingsINI%, Admin, AdminPW
IniRead, AdminDomain, %_SettingsINI%, Admin, AdminDomain
AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"BooYah",5,1)
AdminPW := Crypt.Encrypt.StrDecrypt(AdminPW,"BooYah",5,1)

RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
RunWait, \\cu.int\logonscripts\Installers\DelProf\delprof.exe /q /i /d:30,,
RunAs,
		