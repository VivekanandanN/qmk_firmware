Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad.exe",9
WshShell.AppActivate "Notepad"
WScript.Sleep 2000
WshShell.Sendkeys "Hi"
WshShell.Sendkeys "{ENTER}"
WshShell.Sendkeys "{F5}"

for i = 1 to 800

 Dim dteWait
 dteWait = DateAdd("s", 30, Now())
 Do Until (Now() > dteWait)
 Loop
 WshShell.Sendkeys "Hi"
 WshShell.Sendkeys "{ENTER}"
 WshShell.Sendkeys "{F5}"
next
	