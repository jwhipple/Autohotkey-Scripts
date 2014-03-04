#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn UseUnsetLocal, Off

#Include JRWTools.ahk
#Include Crypt.ahk

; This is how to create a password, and decrypt a password. You probably want to change MyHAsh to something else, think of it as the password to decrypt.

AdminUser = TestThisPasswordNow

AdminUser := Crypt.Encrypt.StrEncrypt(AdminUser,"MyHAsh",5,1)

MsgBox, %AdminUser%

AdminUser := Crypt.Encrypt.StrDecrypt(AdminUser,"MyHAsh",5,1)

MsgBox, %AdminUser%
