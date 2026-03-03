Set WshShell = CreateObject("WScript.Shell")
WshShell.Run Chr(34) & WshShell.CurrentDirectory & "\launch.bat" & Chr(34), 0, False
