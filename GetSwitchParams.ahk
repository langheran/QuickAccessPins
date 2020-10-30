/*
NbParams := GetParams(ParamArrayName[, MaxParams = ""])

	Retrieves the non-switch-parameters (i.e the parameters not starting with a '-' or a '/') into an array :

	ParamArrayName : string or variable containing a string that will be used as a prefix for an array containing the parameters : %ParamArrayName%1 contains 1st param, %ParamArrayName%2 contains 2nd, etc. %ParamArrayName%0 will contain the number of items found.
	MaxParams : the maximum number of parameters that should be saved into the array.

	returns : the number of parameters found (same as %ParamArrayName%0)



NbSubparams := GetSwitch(Switch[, CaseSens = 0, SubparamArrayName = "", MaxSubparams = "", FistParseChar = ":", ParseChar = ""])

	Detects a switch and optionally retrieve its subparameters (e.g. /sw1:SubParam1:SubParam2) into an array :

	Switch : name of the switch (without - or / prefix), can be a Perl-compatible regular expression (PCRE) WITHOUT pattern options.
	CaseSens : Set this to 1 if you want the switch detection to be case sensitive
	SubparamArrayName : string or variable containing a string that will be used as a prefix for an array containing the subparameters : %SubparamArrayName%1 contains 1st param, %SubparamArrayName%2 contains 2nd, etc. %SubparamArrayName%0 will contain the number of items found.
	MaxSubparams : the maximum number of sub parameters that should be saved into the array. If set to -1, the whole string of sub parameters will be retrieved.
	FirstParseChar : The character separating the switch from its sub params. Default is ":". It can be empty.
	ParseChar : The character between two sub params. By default, it is set to empty and internally resolves to ":" and will retrieve correctly the paths (e.g "c:\program files" won't be split into "c" and "\program file"). If explicitly set to ":" or any other char, the paths will be split as well.

	returns : 	0 if the switch was not found ;
				-1 if the switch was found but no parameters ;
				the number of subparameters found (same as %SubparamArrayName%0)





The parameters passed to the script must match the following rules :

Non-switch-parameters, e.g. file paths, must be enclosed in quotes if they contain spaces. They will be saved in the order in which they appear.
Switches can be placed before, inbetween or after the non-switch-parameters, their order doesn't matter.
Switches can be indicated either with '-' or '/'. their subparameters must be separated by ':'
Example :
/i /time:5:10 -date:11:08:2009
File paths will be parsed correctly even if they contain DRIVELETTER:\ pattern
Example :
/WorkingDir:M:\Documents\
If a parameter contains spaces, the spaces must be enclosed in quotes
Example :
/AHKDir:"C:\Program Files\AutoHotkey\" OR /AHKDir:C:\Program" "Files\AutoHotkey\
*/



GetParams(ParamArrayName, MaxParams = "")
{
	Local Params
	Params = 1
	Loop, %0%
	{
		If RegExMatch(%A_Index%, "S)(?:^[^-/].*)", %ParamArrayName%%Params%)
			Params++
		If ((MaxParams != "") && (Params > MaxParams))
			Break
	}
	Params--
	%ParamArrayName%0 := Params
	Return %Params%
}

GetSwitchParams(SwitchName, CaseSens = 0, ParamArrayName = "", MaxParams = "", FirstParseChar = ":", ParseChar = "")
{
	Local Params, Params1, ParseMode
	Static ParamList
	If !ParamList ;Initializing a variable containing all the parameters passed to the script
	{
		Loop, %0%
			ParamList .= "`n" %A_Index%
		ParamList .= "`n"
	}
	If FirstParseChar
		FirstParseChar := ( InStr("\.*?+[{|()^$", SubStr(FirstParseChar, 1, 1)) ? "\" SubStr(FirstParseChar, 1, 1) : SubStr(FirstParseChar, 1, 1) )
	If (ParamArrayName = "") ;if the switch does not require parameters
	{
		If RegExMatch(ParamList, ( CaseSens ? "\n[-/]" SwitchName "(?:\n|" FirstParseChar "\n)" : "i)\n[-/]" SwitchName "(?:\n|" FirstParseChar "\n)"))
			Return -1
		Else
			Return 0
	}
	If (MaxParams = -1) ;the whole content is extracted
	{
		If RegExMatch(ParamList, ( CaseSens ? "\n[-/]" SwitchName "(?:\n|" FirstParseChar "([^\n]*))" : "i)\n[-/]" SwitchName "(?:\n|" FirstParseChar "([^\n]*))" ), %ParamArrayName%)
		{
			%ParamArrayName%0 = 1
			Return 1
		}
		Else
		{
			%ParamArrayName%0 = 0
			Return 0
		}
	}
	If (ParseChar = "")
	{
		ParseChar := ":"
		ParseMode := -2 ;test for files
	}
	Else
	{
		ParseChar := ( InStr("\.*?+[{|()^$", SubStr(ParseChar, 1, 1)) ? "\" SubStr(ParseChar, 1, 1) : SubStr(ParseChar, 1, 1) )
		ParseMode := -3 ;don't test for files
	}
	If RegExMatch(ParamList, ( CaseSens ? "\n[-/]" SwitchName "(?:\n|" FirstParseChar "([^" ParseChar "\n]*(" ParseChar "[^" ParseChar "\n]*)*))" : "i)\n[-/]" SwitchName "(?:\n|" FirstParseChar "([^" ParseChar "\n]*(" ParseChar "[^" ParseChar "\n]*)*))" ) , Params )
	{
		If !(Params1)
		{
			%ParamArrayName%0 = 0
			Return -1 ;switch found but no subparams
		}
		Params = 0
		Loop, Parse, Params1, % ( SubStr(ParseChar, 1, 1) = "\" ?  SubStr(ParseChar, 2) : ParseChar )
		{
			Params++
			If ((ParseMode = -2) && RegExMatch(A_LoopField, "^\\")) ;Managing paths containing LETTER:\..., as they will otherwise be parsed at ":"
			{
				Params--
				If RegExMatch(%ParamArrayName%%Params%, "^[A-Za-z]$")
					%ParamArrayName%%Params% .= ":" A_LoopField
				Else If ((MaxParams != "") && (Params >= MaxParams))
					Break
				Else
				{
					Params++
					%ParamArrayName%%Params% := A_LoopField
				}
			}
			Else If ((MaxParams != "") && (Params > MaxParams))
			{
				Params--
				Break
			}
			Else
				%ParamArrayName%%Params% := A_LoopField
		}
		%ParamArrayName%0 := Params
		Return %Params%
	}
	Else
		Return 0
}

GetCommandLineArgs(){
	Local Args := []
	Loop, %0%
		Args.Push(%A_Index%)
	return Args
}

GetCommandLineArgsString(){
	Local Args := ""
	Loop, %0%
		Args.=" " %A_Index%
	return Args
}

;
;
;Example : try command line : Autohotkey.exe scriptname.ahk file1 /sw1 /sw2:c:\documents:param2 file2
;
;

; ;retrieve the non-switch parameters
; GetParams("File")

; ;Help dialogbox : in this case there must be at least one non-switch parameter
; If (!file1 || GetSwitchParams("(h(elp|)|\?)"))
; {
; 	MsgBox, , PROGRAMNAME [%A_ScriptName%], PROGRAMNAME help :`ncommand : %A_ScriptName% [param1 [param2|/sw1] [/sw2]] [/h]`nParam1:`n%A_Tab%Relative or absolute path and name of file to proceed.`n%A_Tab%Should be enclosed in quotes if containing spaces.`nSwitches :`n%A_Tab%/h :%A_Tab%Displays this help.`n%A_Tab%/sw1 :%A_Tab%Proceeds script in a way.`n%A_Tab%/sw2:option1:option2:... :`n%A_Tab%%A_Tab%Proceeds script in another way.
; ;	ExitApp
; }

; ;Simple switch
; If GetSwitchParams("sw1")
; 	MsgBox sw1 was used.

; ;Switch with subparameters
; If GetSwitchParams("sw2", 0, "sw2_")
; {
; 	If %sw2_0%
; 	{
; 		msg = sw2 was used, with parameters :
; 		Loop, %sw2_0%
; 			msg .= "`n'" sw2_%A_Index% "'"
; 	}
; 	Else
; 		msg = sw2 was used with no parameters.
; 	msgbox %msg%
; }

; msg = Number of files : %File0%`n
; Loop, %File0%
; 	msg .= "file" A_Index " : '" File%A_Index% "'`n"
; msgbox %msg%