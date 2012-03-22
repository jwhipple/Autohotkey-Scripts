#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance
;#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

IfExist, C:\Users\Public
{
	; Delete the All Users icon
	Loop, C:\Users\TimeClock.url,,1
	{
		FileDelete, %A_LoopFileFullPath%
	}
	
}

IfExist, C:\Documents and Settings
{
	; Delete the All Users icon
	Loop, C:\Documents and Settings\TimeClock.url,,1
	{
		FileDelete, %A_LoopFileFullPath%
	}
	

}
