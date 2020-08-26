Dim Program, x, ObjShell

REM Program = "steam://rungameid/730" (Experiment)
Program = "WinWord.exe"

set ObjShell = CreateObject("WScript.Shell")
ObjShell.RUN Program

WScript.Sleep 3000
ObjShell.SendKeys("{Enter}")
ObjShell.SendKeys("Blablabla")
ObjShell.SendKeys("%{F4}") 'Alt + F4

WScript.Sleep 3000
ObjShell.SendKeys("{Enter}") 'Save
WScript.Sleep 3000
ObjShell.SendKeys("{Enter}") 'Save

