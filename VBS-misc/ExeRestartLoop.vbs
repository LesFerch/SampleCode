Set oShell = CreateObject("WScript.Shell")
Do While True
	oShell.Run Chr(34) & "C:\Programs\Borderless Gaming\BorderlessGaming.exe" & Chr(34), 2, 0
	WScript.Sleep 1800000
	oShell.Run "TaskKill /im BorderlessGaming.exe /f", 0, 1
Loop
