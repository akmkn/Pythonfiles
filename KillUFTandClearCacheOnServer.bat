TaskKill /F /IM uft.exe 

Set Temp = C:\Users\%USERNAME%\AppData\Local\Temp\*.*
set ChromeCache=C:\Users\%USERNAME%\AppData\Local\Google\Chrome\User Data\Default\Cache\*.*

RD /s /q "%Temp%"
del /q /s /f "%ChromeCache%"