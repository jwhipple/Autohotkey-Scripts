#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance
;#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include \\CU.INT\logonscripts\Sources\JRWTools.ahk


SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
RanManually =
;clipboard =

;################################
;# Logon Script Global Settings #
;# Configure variables below    #
;################################
_SettingsINI = \\CU.INT\logonscripts\MainLogon.ini

A := GetAllGroups()

MsgBox, %A%