REM @ECHO OFF
REM ****
REM ****
REM ****
REM **** NOTICE: YOU MUST INCREMENT THE NUMBER BELOW 
REM **** BY .1 TO RUN ANY CHANGES TO THIS LOGON SCRIPT!
REM ****
REM ****
REM ****
REM ****

SET VERSION="6.2"
REM SET OUTLOOKVER="1.2"

REM ****
REM ****
REM ****
REM ****

REM The following must be run regardless of version


"C:\Program Files\AutoHotkey\AutoHotkey.exe" \\CU.INT\LogonScripts\Sources\MainLogon.ahk %VERSION% 
REM  %OUTLOOKVER%



REM ******** VERSION DEPENDANT BELOW **************
REM Check version, skip if version matches.
reg query "hkcu\Software\KEMBA" /v LogonVersion | find %VERSION% > nul
if %ERRORLEVEL% == 0 GOTO FINISH
ECHO NoChecks.vbs
%WinDir%\system32\cscript.exe //T:10 //B \\CU.INT\LogonScripts\Installers\Registry\NoChecks.vbs

REM ECHO NIC Power Settings
REM \\CU.INT\LogonScripts\tqcrunas\tqcrunascmd.exe -f \\cu.int\logonscripts\tqcrunas\DisableNICPower.tqc

REM ECHO Password Safe Install
REM if not exist "C:\Documents and Settings\All Users\Start Menu\Programs\Password Safe\Password Safe.lnk" \\CU.INT\LogonScripts\tqcrunas\tqcrunascmd.exe -f \\cu.int\logonscripts\tqcrunas\PasswordSafe.tqc

ECHO Finished
:FINISH

REM Add version key to registry so shorter logon next time.
Reg add "hkcu\Software\KEMBA" /v LogonVersion /t REG_SZ /d %VERSION% /f

:EXIT

exit


