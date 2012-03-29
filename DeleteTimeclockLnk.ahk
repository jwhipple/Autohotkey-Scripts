#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance
;#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

IfExist, C:\Users\Public\Desktop
{
	; Delete the All Users icon
	FileDelete, C:\Users\Public\Desktop\TimeClock.url
	FileDelete, C:\Users\Public\Desktop\Time Clock.url
	FileDelete, C:\Users\Public\Desktop\TimeClock.lnk
	FileDelete, C:\Users\Public\Desktop\Time Clock.url
}

IfExist, C:\Documents and Settings\All Users
{
	; Delete the All Users icon
		FileDelete, C:\Documents and Settings\All Users\Desktop\TimeClock.url
		FileDelete, C:\Documents and Settings\All Users\Desktop\Time Clock.url
		FileDelete, C:\Documents and Settings\All Users\Desktop\TimeClock.lnk
		FileDelete, C:\Documents and Settings\All Users\Desktop\Time Clock.url
}
