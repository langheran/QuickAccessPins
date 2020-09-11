Loop, %0%
{
    param := %A_Index%
	num = %A_Index%
	args := args . param
}

IniRead, skip, %A_ScriptDir%\QuickAccessPins.ini, Settings,skip, 2
IniWrite, %skip%, %A_ScriptDir%\QuickAccessPins.ini, Settings,skip

if(GetKeyState("Control", "P") || GetKeyState("Shift", "P"))
{
	if(FileExist(args))
	{
		count:=0
		Loop, read, %args%
		{
			if(FileExist(A_LoopReadLine))
			{
				count:=count+1
				if(count>skip)
					Run, %A_LoopReadLine%
			}
		}
	}
	ExitApp
}

paths:=""
shell := ComObjCreate("Shell.Application")
for e in shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items()
{
	if(InStr(FileExist(e.Path), "D"))
	{
		paths.=e.Path "`n"
	}
}

FormatTime, CurrTime,,dMMM
FileSelectFile, CurrentPinsPath, S16, %A_Desktop%\%CurrTime%.pins, Save Pins as..., Quick Access Pins (*.pins)
if(!ErrorLevel)
{
	FileAppend, %paths%, %CurrentPinsPath%
	if(FileExist(args))
	{
		count:=0
		for e in shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items()
		{
			if(InStr(FileExist(e.Path), "D")){
				count:=count+1
				if(count>skip)
					e.InvokeVerb("unpinfromhome")
			}
		}
		Loop, read, %args%
		{
			if(InStr(FileExist(A_LoopReadLine), "D")){
				shell.Namespace(A_LoopReadLine).Self.InvokeVerb("pintohome")
			}
		}
	}
}