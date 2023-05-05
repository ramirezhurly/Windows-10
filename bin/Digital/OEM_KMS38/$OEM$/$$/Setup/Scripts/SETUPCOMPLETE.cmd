@echo off
:checkPrivileges 
NET FILE 1>NUL 2>&1
if '%errorlevel%' == '0' ( goto gotPrivileges
) else ( start /min powershell "saps -filepath '%0' -verb runas" >nul 2>&1)
exit /b
:gotPrivileges
cd %~dp0
cd /d "%~dp0"
cd C:\Windows\%windir%\Setup\Scripts\
start /min KMS38.cmd -kms38
TIMEOUT /T 6
rmdir /s /q %windir%\Setup >nul 2>&1