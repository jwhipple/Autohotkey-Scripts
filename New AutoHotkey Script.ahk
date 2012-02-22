#SingleInstance force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance
;#Warn  ; Recommended for catching common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include \\CU.INT\logonscripts\Sources\JRWTools.ahk
;#Include \\cu.int\logonscripts\Sources\Crypt.ahk


SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
RanManually =
;clipboard =

;################################
;# Logon Script Global Settings #
;# Configure variables below    #
;################################
_SettingsINI = \\CU.INT\logonscripts\MainLogon.ini

;RegWrite, REG_BINARY, HKEY_CURRENT_USER,SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Exchange2007\4dcab55718a0c141b87e48e84340e30d, 00036601, 0C100000
;sleep, 1000
;x := FlipOutlookLoginBit()
;clipboard = %x%
;RestoreOutlookLoginBit(x)


;OFFICEVER := GetOutlookVersion()
_MailSectionName = Outlook


Loop, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles, 2, 0
{
	FoundPos := RegExMatch(A_LoopRegName, "Default(.*)")
	If FoundPos
	{
		RegDelete
	}
}

ExitApp

IGetOutlookVersion()
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
		FoundPos := RegExMatch(dir, "Office([0-9]+)", SubPat)
		FoundPos := RegExMatch(SubPat, "([0-9]+)", SubPat)
		clipboard = %FoundPos%
		
		If Not SubPat
		{
			return 0
		}
		else
		{
			SubPat = %SubPat%.0
			return SubPat
		}

		;Removed below because users who never logged into a PC can't use this.
		
		;RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Exchange\Client\Options, PickLogonProfile, 0
		; objCommand := ComObjCreate("Word.Application")
		; _OVersion := objCommand.Version
		; objRelease(objCommand)
		; StringSplit, MyArray, _OVersion, .
		; _OVersion = %MyArray1%.%MyArray2%
		; return _OVersion
	} catch e {
		_OVersion = 0
	}
	return 0
}


IRestoreOutlookLoginBit(_RegistryKeys)
{
	Loop, parse, _RegistryKeys,`n
	{
		;SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Exchange2007\4dcab55718a0c141b87e48e84340e30d

		RegRead, value, HKEY_CURRENT_USER, %A_LoopField%, 00036601		
		If value
		{
			StringMid, x, value, 2, 1
			StringLeft, y, value, 1
			StringRight, z, value, 6
			
			if x = 4
			{
				x = %y%C%z%
				;RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopField%, 00036601, %x%
			}
			if x = 0 
			{
				x = %y%8%z%	
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopField%, 00036601, %x%
			}
		}
	}

}

IFlipOutlookLoginBit()
{
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
			}
			if x = 8 
			{
				x = 0
				x = %y%%x%%z%
				regvals = %regvals%%A_LoopRegSubKey%`n
				RegWrite, REG_BINARY, HKEY_CURRENT_USER, %A_LoopRegSubKey%, 00036601, %x%
			}
		}
	}
	return %regvals%
}


toInt(b, s = 0, c = 0) {
   Loop, % l := StrLen(b) - c
      i += SubStr(b, ++c, 1) * 1 << l - c
   Return, i - s * (1 << l)
}

toBin(i, s = 0, c = 0) {
   l := StrLen(i := Abs(i + u := i < 0))
   Loop, % Abs(s) + !s * l << 2
      b := u ^ 1 & i // (1 << c++) . b
   Return, b
}


pad(x,len) { ; pad with 0's from left to len chars
   IfLess x,0, Return "-" pad(SubStr(x,2),len-1)
   VarSetCapacity(p,len,Asc("0"))
   Return SubStr(p x,1-len)
}

IHex(Inp,UC = 0)
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

IAsc(Inp,UC = 0)
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




; AutoHotkey v2 alpha

; format(format_string, ...)
;   Equivalent to format_v(format_string, Array(...))
format(f, v*) {
    return format_v(f, v)
}

; format_v(format_string, values)
;   - format_string:
;       String of literal text and placeholders, as described below.
;   - values:
;       Array or map of values to insert into the format string.
;
; Placeholder format: "{" id ":" format "}"
;   - id:
;       Numeric or string literal identifying the value in the parameter
;       list.  For example, {1} is values[1] and {foo} is values["foo"].
;   - format:
;       A format specifier as accepted by printf but excluding the
;       leading "%".  See "Format Specification Fields" at MSDN:
;           http://msdn.microsoft.com/en-us/library/56e442dc.aspx
;       The "*" width specifier is not supported, and there may be other
;       limitations.
;
; Examples:
;   MsgBox % format("0x{1:X}", 4919)
;   MsgBox % format("Computation took {2:.9f} {1}", "seconds", 3.2001e-5)
;   MsgBox % format_v("chmod {mode:o} {file}", {mode: 511, file: "myfile"})
;
format_v(f, v)
{
    local out, arg, i, j, s, m, key, buf, c, type, p
    out := "" ; To make #Warn happy.
    VarSetCapacity(arg, 8), j := 1, VarSetCapacity(s, StrLen(f)*2.4)  ; Arbitrary estimate (120% * size of Unicode char).
    O_ := A_AhkVersion >= "2" ? "" : "O)"  ; Seems useful enough to support v1.
    while i := RegExMatch(f, O_ "\{((\w+)(?::([^*`%{}]*([scCdiouxXeEfgGaAp])))?|[{}])\}", m, j)  ; For each {placeholder}.
    {
        out .= SubStr(f, j, i-j)  ; Append the delimiting literal text.
        j := i + m.Len[0]  ; Calculate next search pos.
        if (m.1 = "{" || m.1 = "}") {  ; {{} or {}}.
            out .= m.2
            continue
        }
        key := m.2+0="" ? m.2 : m.2+0  ; +0 to convert to pure number.
        if !v.HasKey(key) {
            out .= m.0  ; Append original {} string to show the error.
            continue
        }
        if m.3 = "" {
            out .= v[key]  ; No format specifier, so just output the value.
            continue
        }
        if (type := m.4) = "s"
            NumPut((p := v.GetAddress(key)) ? p : &(s := v[key] ""), arg)
        else if InStr("cdioux", type)  ; Integer types.
            NumPut(v[key], arg, "int64") ; 64-bit in case of something like {1:I64i}.
        else if InStr("efga", type)  ; Floating-point types.
            NumPut(v[key], arg, "double")
        else if (type = "p")  ; Pointer type.
            NumPut(v[key], arg)
        else {  ; Note that this doesn't catch errors like "{1:si}".
            out .= m.0  ; Output m unaltered to show the error.
            continue
        }
        ; MsgBox % "key=" key ",fmt=" m.3 ",typ=" m.4 . (m.4="s" ? ",str=" NumGet(arg) ";" (&s) : "")
        if (c := DllCall("msvcrt\_vscwprintf", "wstr", "`%" m.3, "ptr", &arg)) >= 0  ; Determine required buffer size.
          && DllCall("msvcrt\_vsnwprintf", "wstr", buf, "ptr", VarSetCapacity(buf, ++c*2)//2, "wstr", "`%" m.3, "ptr", &arg) >= 0 {  ; Format string into buf.
            out .= buf  ; Append formatted string.
            continue
        }
    }
    out .= SubStr(f, j)  ; Append remainder of format string.
    return out
}