@echo off
start cmd /C
set /p input=<KillTaskOnSpecificMachine_MachineName.txt
TaskKill /F /S "%input%" /U Global\automationqateam /P Welcome2 /IM uft.exe /FI "USERNAME eq automationqateam"

TaskKill /F /S "%input%" /U Global\automationqateam /P Welcome2 /IM chrome.exe /FI "USERNAME eq automationqateam"




