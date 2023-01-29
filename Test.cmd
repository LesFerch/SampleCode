@Echo Off
Md AppData >Nul 2>Nul
Move Test.txt AppData\Test.ini >Nul 2>Nul
PowerShell -ExecutionPolicy Bypass .\Test.ps1 AppData\Test.ini