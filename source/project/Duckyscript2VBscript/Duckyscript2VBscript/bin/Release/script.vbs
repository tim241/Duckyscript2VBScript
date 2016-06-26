set wshshell = wscript.createobject("wscript.shell")
set objshell = createobject("shell.application")
WScript.Sleep 6000
For i = 1 To 6
	WshShell.SendKeys("hi")
Next
