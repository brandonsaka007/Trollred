WSCript.Sleep 30000
set WshShell = CreateObject("WScript.Shell")
WshShell.Run "%SystemRoot%\System32\SndVol.exe"
WScript.Sleep 400
WshShell.AppActivate "Volume Mixer"
WshShell.SendKeys "{PGUP} {PGUP} {PGUP} {PGUP} {PGUP} {PGUP}"
WshShell.SendKeys "%{F4}"  ' Alt+F4
WshShell.SendKeys "%{F4}"
WSCript.Sleep 1000
WshShell.AppActivate "Google Chrome"
Dim iURL 
Dim objShell

iURL = "https://www.youtube.com/watch?v=g-sgw9bPV4A"

set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 2000
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "Google Chrome"
WshShell.SendKeys "f"
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 70000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 80000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1
WSCript.Sleep 50000
set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "chrome.exe", iURL, "", "", 1