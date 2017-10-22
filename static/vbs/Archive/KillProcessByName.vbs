Set shell = CreateObject("WScript.Shell")
shell.run "cmd /k C:\KillProcessByName.bat", 1, True
set shell=nothing
Msgbox "Remote Machines are Ready for execution."