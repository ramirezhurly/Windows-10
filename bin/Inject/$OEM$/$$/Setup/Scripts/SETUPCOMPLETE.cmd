@echo off
:checkPrivileges 
NET FILE 1>NUL 2>&1
if '%errorlevel%' == '0' ( goto gotPrivileges
) else ( start /min powershell "saps -filepath '%0' -verb runas" >nul 2>&1)
exit /b
:gotPrivileges
cd %~dp0
cd /d "%~dp0"
IF /I "%PROCESSOR_ARCHITECTURE%" EQU "AMD64" (set xOS=x64) else (set xOS=x86)
if not exist "%windir%\KMS\KMSInject.cmd" (
xcopy /cryi bin\* %windir%\KMS\bin >nul 2>&1
copy /y run.cmd %windir%\KMS\KMSInject.cmd >nul 2>&1
schtasks /create /tn "KMS_Aktivasyon" /xml "%~dp0bin\KMS.xml" /f >nul 2>&1
)
del /f /q %windir%\KMS\bin\*.xml >nul 2>&1
IF /I "%PROCESSOR_ARCHITECTURE%" EQU "AMD64" (del /f /q C:\Windows\KMS\bin\x86.dll) >nul 2>&1 else (del /f /q C:\Windows\KMS\bin\x64.dll) >nul 2>&1
start /min %windir%\KMS\KMSInject.cmd -act && exit