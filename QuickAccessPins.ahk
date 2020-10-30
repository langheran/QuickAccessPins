#Include GetSwitchParams.ahk

global openfilepath := 0
If GetSwitchParams("openfile", 0, "openfile_",,"!","!")
{
	If %openfile_0%
	{
		openfilepath :=openfile_1
	}
}

global savefilepath := 0
If GetSwitchParams("savefile", 0, "savefile_",,"!","!")
{
	If %savefile_0%
	{
		savefilepath :=savefile_1
	}
}

shiftIsPressed:=GetKeyState("Shift", "P")
controlIsPressed:=GetKeyState("Control", "P")

Loop, %0%
{
    param := %A_Index%
	num = %A_Index%
	args := args . param
}

if(InStr(FileExist(args), "D"))
	openfilepath:=args

global skip:=2
IniRead, skip, %A_ScriptDir%\QuickAccessPins.ini, Settings,skip, %skip%
IniWrite, %skip%, %A_ScriptDir%\QuickAccessPins.ini, Settings,skip

if(controlIsPressed)
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

global shell := ComObjCreate("Shell.Application")
if(shiftIsPressed && !FileExist(args))
{
	EmptyPins()
	ExitApp
}

global paths:=GetCurrentPaths()

if(openfilepath || savefilepath)
{
	if(savefilepath)
	{
		SplitPath, savefilepath,,,OutExtension
		if(FileExist(savefilepath) && OutExtension=="pins")
		{
			CurrentPinsPath:=savefilepath
			if(FileExist(CurrentPinsPath))
				FileDelete, %CurrentPinsPath%
			FileAppend, %paths%, %CurrentPinsPath%
		}
	}

	if(openfilepath)
	{
		SplitPath, openfilepath,,,OutExtension
		if(FileExist(openfilepath))
		{
			if(OutExtension=="pins")
			{
				OpenPins(openfilepath)
			}
			else
			{
				if(InStr(FileExist(openfilepath), "D"))
				{
					SplitPath, openfilepath, OutFileName, OutDir
					MsgBox, 4,Create QuickAccessPins?,Create QuickAccessPins in "%OutFileName%"? (Si o No)
					IfMsgBox Yes
					{
						EmptyPins()
						shell.Namespace(openfilepath).Self.InvokeVerb("pintohome")
						paths:=GetCurrentPaths()
						CurrentPinsPath:= openfilepath . "\" . OutFileName . ".pins"
						if(!FileExist(CurrentPinsPath))
							FileAppend, %paths%, %CurrentPinsPath%
					}
				}
			}
		}
	}
}
else
{
	if(!shiftIsPressed)
	{
		if(CurrentPinsPath:=GetSaveFileName())
		{
			if(FileExist(CurrentPinsPath))
				FileDelete, %CurrentPinsPath%
			FileAppend, %paths%, %CurrentPinsPath%
		}
	}

	if(FileExist(args) && (CurrentPinsPath || shiftIsPressed))
	{
		OpenPins(args)
	}
}

return

OpenPins(filepath)
{
	EmptyPins()
	Loop, read, %filepath%
	{
		if(InStr(FileExist(A_LoopReadLine), "D")){
			shell.Namespace(A_LoopReadLine).Self.InvokeVerb("pintohome")
		}
	}
}

GetSaveFileName()
{
	FormatTime, CurrTime,,dMMM
	FileSelectFile, CurrentPinsPath, S16, %A_Desktop%\%CurrTime%.pins, Save Pins as..., Quick Access Pins (*.pins)
	if(!ErrorLevel)
	{
		return CurrentPinsPath
	}
	return 0
}

EmptyPins()
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
}

GetCurrentPaths()
{
	for e in shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items()
	{
		if(InStr(FileExist(e.Path), "D"))
		{
			_paths.=e.Path "`n"
		}
	}
	return _paths
}