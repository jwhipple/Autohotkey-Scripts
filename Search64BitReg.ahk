#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance


; ######################################################################################
; THIS SCRIPT MUST BE COMPILED BY THE 64 BIT VERSION OF AUTOHOTKEY TO WORK AS INTENDED!
; ######################################################################################


FindInstalled()
{

	REGPATH =SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
	Loop, HKEY_LOCAL_MACHINE, %REGPATH%, 1, 1
	{
		If A_LoopRegName = DisplayName
		{
			RegRead, value
			StringLower, value, value
			GuiControl,, MyListBox, %value%
		}
	}

	REGPATH =SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
	Loop, HKEY_LOCAL_MACHINE, %REGPATH%, 1, 1
	{
		If A_LoopRegName = DisplayName
		{
			RegRead, value
			StringLower, value, value
			GuiControl,, MyListBox, %value%
		}
	}
}


Gui, Add, ListBox, vMyListBox gMyListBox w640 r10
Gui, Add, Button, Default, OK
FindInstalled()

Gui, Hide
SetTimer, GuiEscape, 8000

return

MyListBox:
if A_GuiEvent <> DoubleClick
    return
; Otherwise, the user double-clicked a list item, so treat that the same as pressing OK.
; So fall through to the next label.
ButtonOK:
GuiClose:
GuiEscape:
ExitApp

