@Echo Off
CD /D %~dp0
MD AppData 2>Nul
Move Test.txt AppData\Test.ini 2>Nul
PowerShell -ExecutionPolicy Bypass .\Test.ps1 AppData\Test.ini