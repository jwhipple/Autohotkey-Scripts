#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn UseUnsetLocal, Off ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include \\cu.int\logonscripts\Sources\Crypt.ahk

_SettingsINI = \\CU.INT\logonscripts\MainLogon.ini
IniRead, AdminUser, %_SettingsINI%, Admin, AdminUser
IniRead, AdminPW, %_SettingsINI%, Admin, AdminPW
IniRead, AdminDomain, %_SettingsINI%, Admin, AdminDomain
AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"BooYah",5,1)
AdminPW := Crypt.Encrypt.StrDecrypt(AdminPW,"BooYah",5,1)

; Install spark.
RunAs, %AdminUser%, %AdminPW%, %AdminDomain%
RunWait,msiexec /i \\cu.int\logonscripts\Installers\Spark\Spark.msi,,
RunAs,		
	

; Copy default config over.

FileCopy, \\cu.int\logonscripts\Installers\Spark\config\*.*, %A_AppData%\Spark\, 1
