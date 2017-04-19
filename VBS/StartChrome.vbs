Set oShell = CreateObject("WScript.Shell")
Command = "cmd /K start chrome google.com & exit"
oShell.Run Command,1,True