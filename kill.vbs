Set s = Wscript.CreateObject("WScript.Shell")
s.Run "taskkill.exe /F /IM Wscript.exe /T", , True
