@echo off

Set Temp = C:\Users\automationqateam\AppData\Local\Temp\*.*
set ChormeHistory=C:\Users\automationqateam\AppData\Local\Google\Chrome\User Data\Default\History.*
set ChromeCookies=C:\Users\automationqateam\AppData\Local\Google\Chrome\User Data\Default\Cookies.*
set ChromeCache=C:\Users\automationqateam\AppData\Local\Google\Chrome\User Data\Default\Cache\*.*

RD /s /q "%Temp%"
del /q /s /f "%ChormeHistory%"
del /q /s /f "%ChromeCookies%"
del /q /s /f "%ChromeCache%"