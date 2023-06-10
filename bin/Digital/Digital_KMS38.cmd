@echo off
cd /d "%~dp0"


if "[%1]" == "[49127c4b-02dc-482e-ac4f-ec4d659b7547]" goto :MAINMENU
REG QUERY HKU\S-1-5-19\Environment >NUL 2>&1 && goto :MAINMENU

set command="""%~f0""" 49127c4b-02dc-482e-ac4f-ec4d659b7547
SETLOCAL ENABLEDELAYEDEXPANSION
set "command=!command:'=''!"

IF "%~1"=="-suite" GOTO :MAINMENU

powershell -NoProfile Start-Process -FilePath '%COMSPEC%' ^
-ArgumentList '/c """!command!"""' -Verb RunAs 2>NUL

IF %ERRORLEVEL% GTR 0 (
    echo.
    echo =============================================================
    echo This script needs to be run as administrator.
    echo =============================================================
    echo.
pause
)

SETLOCAL DISABLEDELAYEDEXPANSION
goto :EOF
::===============================================================================================================
:MAINMENU
title Digital KMS Licenses Office and Windows 10 Activation v8.8 - Hurly
set ver=v8.8
mode con cols=70 lines=1
for /f "tokens=2 delims==" %%a in ('wmic path Win32_OperatingSystem get BuildNumber /value') do (
  set /a WinBuild=%%a
)
del %temp%\msg.vbs /f /q >nul 2>&1
echo Set WshShell = CreateObject("WScript.Shell"^) >> %temp%\msg.vbs
echo x = WshShell.Popup ("Not detected Windows 10. Digital License/KMS38 Activation is Not Supported. The process will be terminated in 5 seconds.",5, "WARNING") >> %temp%\msg.vbs
if %winbuild% LSS 10240 (
call %temp%\msg.vbs
del %temp%\msg.vbs /f /q >nul 2>&1
del /f /q "bin\editions" >nul 2>&1 
RD %TEMP%\. /S /Q >nul
RD %windir%\TEMP\. /S /Q >nul
echo.
exit
)
set Auto=0
IF "%~1"=="-digi" set Auto=1
IF "%~1"=="-kms38" set Auto=1
IF "%~1"=="-digi" GOTO :Digital
IF "%~1"=="-kms38" GOTO :KMS38
::===============================================================================================================
setlocal enabledelayedexpansion
setlocal EnableExtensions
pushd "%~dp0"
mode con cols=84 lines=41
cd /d "%~dp0"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get Caption /format:LIST"')do (set NameOS=%%a) >nul 2>&1
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get CSDVersion /format:LIST"')do (set SP=%%a) >nul 2>&1
for /f "tokens=2 delims==" %%a IN ('"wmic Path Win32_OperatingSystem Get Version /format:LIST"')do (set Version=%%a) >nul 2>&1
for /f "tokens=2* delims= " %%a in ('reg query "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" /v "PROCESSOR_ARCHITECTURE"') do if "%%b"=="AMD64" (set vera=x64) else (set vera=x86)
color f2
echo ====================================================================================
set yy=%date:~-4%
set mm=%date:~-7,2%
set dd=%date:~-10,2%
set MYDATE=%yy%%mm%%dd%
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a:%%b)
echo                                                            %dd%.%mm%.%yy% ^- %mytime%
echo.
echo   Digital KMS Licenses Office and Windows 10 Activation v8.8 - Hurly
echo.
echo   SUPPORT MICROSOFT PRUDUCTS:
echo   Windows 10 (all versions)
echo.
echo          Operating System : %NameOS% %SP% %xOS%
reg.exe query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v DigitalProductId >nul 2>&1
echo          Version : %Version%
echo          Architecture : %PROCESSOR_ARCHITECTURE%
echo          Identifier : %computername%
echo ====================================================================================
echo.
Echo.     [1] DIGITAL ACTIVATION START FOR WINDOWS 10 
Echo.
Echo.     [2] KMS38 ACTIVATION START FOR WINDOWS 10 
Echo.
Echo.     [3] DIGITAL or KMS38 $OEM$ ACTIVATION FOLDER EXTRACT TO %systemdrive%\ DIRECTORY
Echo.
Echo.     [4] WINDOWS ^& OFFICE ACTIVATION STATUS CHECK
Echo.
Echo.     [5] DIGITAL ACTIVATION VISIT WEBSITE (TNCTR)
Echo.
Echo.     [6] EXIT 
Echo.
Echo.     [7] RETURN KMS SUITE MENU
echo.
echo ====================================================================================
choice /C:1234567 /N /M "YOUR CHOICE : "
if errorlevel 7 goto :KMSSuite
if errorlevel 6 goto :Exit 
if errorlevel 5 goto :TNCTR
if errorlevel 4 goto :Check
if errorlevel 3 goto :OEM
if errorlevel 2 goto :KMS38
if errorlevel 1 goto :Digital
::===============================================================================================================
:Digital
set "DIGI=1"
if exist "bin\editions" del /f /q "bin\editions" >nul 2>&1 
pushd "%~dp0bin\"
if exist "%~dp0bin\gatherosstatemodified.exe" del /f /q "%~dp0bin\gatherosstatemodified.exe" >nul 2>&1
cd..
set _1=ClipSVC
set _2=wlidsvc
set _3=sppsvc
set _4=wuauserv
echo.
echo Checking Windows Services, please wait...
echo.
for %%# in (%_1% %_2% %_3% %_4%) do call :ServiceCheck %%#

set "CLecho=Checking %_1% [Service Status -%Cl_state%] [Startup type -%Cl_start_type%]"
set "wlecho=Checking %_2% [Service Status -%wl_state%] [Startup type -%wl_start_type%]"
set "specho=Checking %_3% [Service Status -%sp_state%] [Startup type -%sp_start_type%]"
set "wuecho=Checking %_4% [Service Status -%wu_state%] [Startup type -%wu_start_type%]"

if not "%Cl_start_type%"=="Demand"      ("%CLecho%" &echo: & set Clst_e=1) else (echo %CLecho%)
if not "%wl_start_type%"=="Demand"      ("%wlecho%" &echo: & set wlst_e=1) else (echo %wlecho%)
if not "%sp_start_type%"=="Delayed-Auto" ("%specho%" &echo: & set spst_e=1) else (echo %specho%)

if "%wu_start_type%"=="Disabled" (set _C=4F) else (set _C=8F) >nul 2>&1
if not "%wu_start_type%"=="Auto" (%_C% "%wuecho%" &echo: & set wust_e=1) else (echo %wuecho%) 

if defined Clst_e (sc config %_1% start= Demand %nul%       && set Clst_s=%_1%-Demand || set Clst_u=%_1%-Demand )
if defined wlst_e (sc config %_2% start= Demand %nul%       && set wlst_s=%_2%-Demand || set wlst_u=%_2%-Demand )
if defined spst_e (sc config %_3% start= Delayed-Auto %nul% && set spst_s=%_3%-Delayed-Auto || set spst_u=%_3%-Delayed-Auto )
if defined wust_e (sc config %_4% start= Auto %nul%         && set wust_s=%_4%-Auto || set wust_u=%_4%-Auto )

for %%# in (Clst_s,wlst_s,spst_s,wust_s) do if defined %%# set st_s=1
if defined st_s (
echo Changing the service start type: [%Clst_s%%wlst_s%%spst_s%%wust_s%] [SUCCESSFUL]
)

for %%# in (Clst_u,wlst_u,spst_u,wust_u) do if defined %%# set st_u=1
if defined st_u (
echo Unable to change service start type: [%Clst_u%%wlst_u%%spst_u%%wust_u%]
)

if not "%Cl_state%"=="Running" (Powershell -NoProfile start-service %_1% %nul% && set Cl_s=%_1% || set Cl_u=%_1% )
if not "%wl_state%"=="Running" (Powershell -NoProfile start-service %_2% %nul% && set wl_s=%_2% || set wl_u=%_2% )
if not "%sp_state%"=="Running" (Powershell -NoProfile start-service %_3% %nul% && set sp_s=%_3% || set sp_u=%_3% )
if not "%wu_state%"=="Running" (Powershell -NoProfile start-service %_4% %nul% && set wu_s=%_4% || set wu_u=%_4% )

for %%# in (Cl_s,wl_s,sp_s,wu_s) do if defined %%# set s_s=1
if defined s_s (
echo Starting service: [%Cl_s%%wl_s%%sp_s%%wu_s%] [SUCCESSFUL]
)

for %%# in (Cl_u,wl_u,sp_u,wu_u) do if defined %%# set s_u=1
if defined s_u (
echo Service failed to start: [%Cl_u%%wl_u%%sp_u%%wu_u%]
)

if defined wust_u (
echo A Windows Update blocking program has safely disabled wuauserv.
)
goto HWIDActivate
:KMS38
set "KMS38=1"
if exist "bin\editions" del /f /q "bin\editions" >nul 2>&1 
pushd "%~dp0bin\"
if exist "%~dp0bin\gatherosstatemodified.exe" del /f /q "%~dp0bin\gatherosstatemodified.exe" >nul 2>&1
cd..
echo.
echo Checking Windows Services, please wait...
echo.
set _1=ClipSVC
set _3=sppsvc

for %%# in (%_1% %_3%) do call :Servicecheck %%#

set "CLecho=Checking %_1% [Service Status -%Cl_state%] [Startup type -%Cl_start_type%]"
set "specho=Checking %_3% [Service Status -%sp_state%] [Startup type -%sp_start_type%]"

if not "%Cl_start_type%"=="Demand"       ("%CLecho%" &echo: & set Clst_e=1) else (echo %CLecho%) 
if not "%sp_start_type%"=="Delayed-Auto" ("%specho%" &echo: & set spst_e=1) else (echo %specho%) 

echo.
if defined Clst_e (sc config %_1% start= Demand %nul%       && set Clst_s=%_1%-Demand || set Clst_u=%_1%-Demand ) 
if defined spst_e (sc config %_3% start= Delayed-Auto %nul% && set spst_s=%_3%-Delayed-Auto || set spst_u=%_3%-Delayed-Auto ) 

for %%# in (Clst_s,spst_s) do if defined %%# set st_s=1
if defined st_s (
echo Changing the service start type: [ %Clst_s%%spst_s%] [BASARILI]
)

for %%# in (Clst_u,spst_u) do if defined %%# set st_u=1
if defined st_u (
echo Unable to change service start type: [ %Clst_u%%spst_u%]
)

if not "%Cl_state%"=="Running" (Powershell -NoProfile start-service %_1% %nul% && set Cl_s=%_1% || set Cl_u=%_1% ) 
if not "%sp_state%"=="Running" (Powershell -NoProfile start-service %_3% %nul% && set sp_s=%_3% || set sp_u=%_3% )

for %%# in (Cl_s,sp_s) do if defined %%# set s_s=1 >nul
if defined s_s (
echo Starting service: [ %Cl_s%%sp_s%] [BASARILI]
)

for %%# in (Cl_u,sp_u) do if defined %%# set s_u=1
if defined s_u (
echo Service failed to start: [ %Cl_u%%sp_u%]
)
goto HWIDActivate
:HWIDActivate
mode con cols=83 lines=33
CALL :DetectEdition

If defined KMS38 (
set "A2=KMS38"
set "A3=GVLK"
set "A4=Volume:GVLK"
set "A5=gatherosstate.exe"
set "A6= >nul 2>&1"
)
If defined DIGI (
set "B2=Digital License"
set "B3=Retail-OEM_Key"
set "B4=Retail"
set "B5=gatherosstate.exe"
)
::===========================================================
call:%A3%%B3%

for /f "tokens=1-4 usebackq" %%a in ("bin\editions") do (if ^[%%a^]==^[%osedition%^] (
    set edition=%%a
    set sku=%%b 
    set Key=%%c
    goto:parseAndPatch))
echo %osedition% %vera% %A2%%B2% Activation is Not Supported.
del /f /q "bin\editions" >nul 2>&1 
echo.
echo Devam etmek icin bir tusa basin...
pause >nul
CLS
GOTO MAINMENU
::===========================================================
:parseAndPatch
echo.
echo ===================================================================================
echo               Windows 10 %osedition% %vera% %A2%%B2% Activation
echo ===================================================================================
echo.
echo Cleaning ClipsSVC tokens...
sc stop clipsvc >nul 2>&1
del /f /s /q "%allusersprofile%\Microsoft\Windows\ClipSVC\tokens.dat" >nul 2>&1
If defined KMS38 (
cscript /nologo %windir%\system32\slmgr.vbs -ckms >nul 2>&1
echo.
echo Rearming Windows Application ID...
cscript /nologo %windir%\system32\slmgr.vbs -rearm-app 55c92734-d682-4d71-983e-d6ec3f16059f >nul 2>&1
echo.
for /f "tokens=2 delims==" %%G in ('"wmic path SoftwareLicensingProduct where (ProductKeyID like '%%-%%' AND Description like '%%Windows%%') get ID /value"') do (set winapp=%%G) >nul 2>&1
echo Rearming Windows SKU ID...
cscript //nologo "%systemroot%\System32\slmgr.vbs" /rearm-sku %winapp% >nul 2>&1
)
echo.
echo Installing key %key%
cscript /nologo %windir%\system32\slmgr.vbs -ipk %key% >nul 2>&1
echo.

echo Create GenuineTicket.XML file for Windows 10 %edition% %vera%
pushd "%~dp0bin\"
rundll32 "%~dp0bin\slc.dll",PatchGatherosstate >nul 2>&1
start /wait gatherosstatemodified.exe 
timeout /t 3 >nul 2>&1
cd..

echo.
echo GenuineTicket.XML file is installing for Windows 10 %edition% %vera%
clipup -v -o -altto bin\
echo.

del /f /q "bin\editions" >nul 2>&1
pushd "%~dp0bin\"
if exist "%~dp0bin\gatherosstatemodified.exe" del /f /q "%~dp0bin\gatherosstatemodified.exe" >nul 2>&1
cd..
echo Activating...
echo.
cscript /nologo %windir%\system32\slmgr.vbs -ato >nul 2>&1
cscript /nologo %windir%\system32\slmgr.vbs -xpr
if "%Auto%"=="1" (
     GOTO :Out 
) ELSE (
     GOTO :Done 
)

for /f "tokens=2 delims==" %%A in ('"wmic path SoftwareLicensingProduct where (Description like '%%KMSCLIENT%%' and Name like 'Windows%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" ') do set "gpr=%%A"
if %gpr% GTR 259200 echo Windows 10 %edition% %vera% activated with KMS38.

if "%Auto%"=="1" (
     GOTO :Out 
) ELSE (
     GOTO :Done 
)
echo.
if %gpr% LEQ 259200 Goto:Rearm
:Rearm
echo Windows 10 %edition% %vera% KMS38 is not activated.
echo.
echo Applying slmgr /rearm to fix activation...
cscript /nologo %windir%\system32\slmgr.vbs -rearm >nul 2>&1
echo.
echo Restarting the system...
shutdown.exe /r /soft
echo.
::===============================================================================================================
:Done
echo.
echo Press any key to continue...
pause >nul
CLS
GOTO MAINMENU
::===============================================================================================================
:Out
TIMEOUT /T 2 
exit
::===============================================================================================================
:DetectEdition
FOR /F "TOKENS=2 DELIMS==" %%A IN ('"WMIC PATH SoftwareLicensingProduct WHERE (Name LIKE 'Windows%%' AND PartialProductKey is not NULL) GET LicenseFamily /VALUE"') DO IF NOT ERRORLEVEL 1 SET "osedition=%%A"
if not defined osedition (FOR /F "TOKENS=3 DELIMS=: " %%A IN ('DISM /English /Online /Get-CurrentEdition 2^>nul ^| FIND /I "Current Edition :"') DO SET "osedition=%%A")

if %winbuild% EQU 10240 (
if "%osedition%"=="EnterpriseS" set "osedition=EnterpriseS2015"
if "%osedition%"=="EnterpriseSN" set "osedition=EnterpriseSN2015"
)
if %winbuild% EQU 14393 (
if "%osedition%"=="EnterpriseS" set "osedition=EnterpriseS2016"
if "%osedition%"=="EnterpriseSN" set "osedition=EnterpriseSN2016"
)
if %winbuild% GEQ 17763 (
if "%osedition%"=="EnterpriseS" set "osedition=EnterpriseS2019"
if "%osedition%"=="EnterpriseSN" set "osedition=EnterpriseSN2019"
)
exit /b
::===============================================================================================================
:Retail-OEM_Key
rem              Edition          SKU            Retail/OEM_Key         
(                                                                       
echo Core                         101      YTMG3-N6DKC-DKB77-7M9GH-8HVX7
echo CoreN                         98      4CPRK-NM3K3-X6XXQ-RXX86-WXCHW
echo CoreCountrySpecific           99      N2434-X9D7W-8PF6X-8DV9T-8TYMD
echo CoreSingleLanguage           100      BT79Q-G7N6G-PGBYW-4YWX6-6F4BT
echo Education                    121      YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY
echo EducationN                   122      84NGF-MHBT6-FXBX8-QWJK7-DRR8H
echo Enterprise                     4      XGVPP-NMH47-7TTHJ-W3FW7-8HV2C
echo EnterpriseN                   27      3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT
echo EnterpriseS2015              125      FWN7H-PF93Q-4GGP8-M8RF3-MDWWW
echo EnterpriseSN2015             126      8V8WN-3GXBH-2TCMG-XHRX3-9766K
echo EnterpriseS2016              125      NK96Y-D9CD8-W44CQ-R8YTK-DYJWX
echo EnterpriseSN2016             126      2DBW3-N2PJG-MVHW3-G7TDK-9HKR4
echo Professional                  48      VK7JG-NPHTM-C97JM-9MPGT-3V66T
echo ProfessionalN                 49      2B87N-8KFHP-DKV6R-Y2C8J-PKCKT
echo ProfessionalEducation        164      8PTT6-RNW4C-6V7J2-C2D3X-MHBPB
echo ProfessionalEducationN       165      GJTYN-HDMQY-FRR76-HVGC7-QPF8P
echo ProfessionalWorkstation      161      DXG7C-N36C4-C4HTG-X4T3X-2YV77
echo ProfessionalWorkstationN     162      WYPNQ-8C467-V2W6J-TX4WX-WT2RQ
echo ServerRdsh                   175      NJCF7-PW8QT-3324D-688JX-2YV66
echo IoTEnterprise                188      XQQYW-NFFMW-XJPBH-K8732-CKFFD
                                                                        
) > "bin\editions" &exit /b                                                                                 
::===============================================================================================================
:GVLK
rem              Edition          SKU                  GVLK             
(                                                                       
echo Core                         101      TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
echo CoreN                         98      3KHY7-WNT83-DGQKR-F7HPR-844BM
echo CoreCountrySpecific           99      PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
echo CoreSingleLanguage           100      7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
echo Education                    121      NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
echo EducationN                   122      2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
echo Enterprise                     4      NPPR9-FWDCX-D2C8J-H872K-2YT43
echo EnterpriseN                   27      DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
echo EnterpriseS2016              125      DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
echo EnterpriseSN2016             126      QFFDN-GRT3P-VKWWX-X7T3R-8B639
echo EnterpriseS2019              125      M7XTQ-FN8P6-TTKYV-9D4CC-J462D
echo EnterpriseSN2019             126      92NFX-8DJQP-P6BBQ-THF9C-7CG2H
echo Professional                  48      W269N-WFGWX-YVC9B-4J6C9-T83GX
echo ProfessionalN                 49      MH37W-N47XK-V7XM9-C7227-GCQG9
echo ProfessionalEducation        164      6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
echo ProfessionalEducationN       165      YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
echo ProfessionalWorkstation      161      NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
echo ProfessionalWorkstationN     162      9FNHH-K3HBT-3W4TD-6383H-6XYWF
echo ServerStandard                 7      N69G4-B89J2-4G8F4-WWYCC-J464C
echo ServerStandardCore            13      N69G4-B89J2-4G8F4-WWYCC-J464C
echo ServerDatacenter               8      WMDGN-G9PQG-XVVXX-R3X43-63DFG
echo ServerDatacenterCore          12      WMDGN-G9PQG-XVVXX-R3X43-63DFG
echo ServerSolution                52      WVDHN-86M7X-466P6-VHXV7-YY726
echo ServerSolutionCore            53      WVDHN-86M7X-466P6-VHXV7-YY726
echo ServerRdsh                   175      7NBT4-WGBQX-MP4H7-QXFF8-YP3KX

) > "bin\editions" &exit /b   
::===============================================================================================================
:OEM
Echo.
Echo.   [1] DIGITAL $OEM$       [2] KMS 2038 $OEM$
Echo.
choice /C:12 /N /M "YOUR CHOICE : "
if errorlevel 2 goto :KMS38OEM 
if errorlevel 1 goto :DIGIOEM

:DIGIOEM
IF EXIST %systemdrive%\$OEM$ (
echo.
echo =============================================
echo $OEM$ folder already exists on %systemdrive%\ directory
echo =============================================
echo. 
echo Press any key to continue...
pause >nul
CLS
GOTO MAINMENU
) ELSE (
md %systemdrive%\$OEM$
cd /d "%~dp0"
md %systemdrive%\$OEM$\$$\Setup\Scripts\bin\
xcopy OEM_Digital\* %systemdrive%\ /s /i /y >nul 2>&1
xcopy /cryi bin\* %systemdrive%\$OEM$\$$\Setup\Scripts\bin\) >nul 2>&1 
echo MSGBOX "DIGITAL ACTIVATION $OEM$ FOLDER EXTRACT TO %systemdrive%\ DIRECTORY" > %temp%\TEMPmessage.vbS
call %temp%\TEMPmessage.vbs
del %temp%\TEMPmessage.vbs /f /q
CLS
GOTO MAINMENU
::===============================================================================================================
:KMS38OEM
IF EXIST %systemdrive%\$OEM$ (
echo.
echo =============================================
echo $OEM$ folder already exists on %systemdrive%\ directory
echo =============================================
echo. 
echo Press any key to continue...
pause >nul
CLS
GOTO MAINMENU
) ELSE (
md %systemdrive%\$OEM$
cd /d "%~dp0"
md %systemdrive%\$OEM$\$$\Setup\Scripts\bin\
xcopy OEM_KMS38\* %systemdrive%\ /s /i /y >nul 2>&1
xcopy /cryi bin\* %systemdrive%\$OEM$\$$\Setup\Scripts\bin\) >nul 2>&1 
echo MSGBOX "KMS 2038 ACTIVATION $OEM$ FOLDER EXTRACT TO %systemdrive%\ DIRECTORY" > %temp%\TEMPmessage.vbS
call %temp%\TEMPmessage.vbs
del %temp%\TEMPmessage.vbs /f /q
CLS
GOTO MAINMENU
::===============================================================================================================
:ServiceCheck
for /f "tokens=1,3 delims=: " %%a in ('sc query %1') do (if /i %%a==state set "state=%%b")
for /f "tokens=1-4 delims=: " %%a in ('sc qc %1') do (if /i %%a==start_type set "start_type=%%c %%d")

if /i "%state%"=="STOPPED" set state=Stopped
if /i "%state%"=="RUNNING" set state=Running

if /i "%start_type%"=="auto_start (delayed)" set start_type=Delayed-Auto
if /i "%start_type%"=="auto_start "          set start_type=Auto
if /i "%start_type%"=="demand_start "        set start_type=Demand
if /i "%start_type%"=="disabled "            set start_type=Disabled

for %%i in (%*) do (
if /i "%%i"=="%_4%" set "wu_start_type=%start_type%" & set "wu_state=%state%"
if /i "%%i"=="%_3%" set "sp_start_type=%start_type%" & set "sp_state=%state%"
if /i "%%i"=="%_1%" set "Cl_start_type=%start_type%" & set "Cl_state=%state%"
if /i "%%i"=="%_2%" set "wl_start_type=%start_type%" & set "wl_state=%state%"
)
exit /b
::===============================================================================================================
:Check
del /f /q %temp%\check.txt >nul 2>&1
del /f /q %temp%\check.vbs >nul 2>&1
echo textFilePath = "check.txt" >> %temp%\check.vbs
echo set objFSO = createobject("Scripting.FileSystemObject") >> %temp%\check.vbs
echo set objTextFile = objFSO.opentextfile(textFilePath) >> %temp%\check.vbs
echo wscript.echo objTextFile.ReadAll >> %temp%\check.vbs
echo objTextFile.Close >> %temp%\check.vbs

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set osps=OfficeSoftwareProtectionService
set winApp=55c92734-d682-4d71-983e-d6ec3f16059f
set o14App=59a52881-a989-479d-af46-f275c6370663
set o15App=0ff1ce15-a989-479d-af46-f275c6370663
for %%# in (spp_get,ospp_get,W1nd0ws,sppw,0ff1ce15,sppo,osppsvc,ospp14,ospp15) do set "%%#="
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "spp_get=Description, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, EvaluationEndDate, GracePeriodRemaining, ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceProductKeyID, LicenseStatus, LicenseStatusReason, Name, PartialProductKey, ProductKeyID, VLActivationInterval, VLRenewalInterval"
set "ospp_get=%spp_get%"
if %winbuild% geq 9200 set "spp_get=%spp_get%, DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled"

set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

call :PKey %spp% %winApp% W1nd0ws sppw
if %winbuild% geq 9200 call :PKey %spp% %o15App% 0ff1ce15 sppo
wmic path %osps% get Version 1>nul 2>nul && (
call :PKey %ospp% %o14App% osppsvc ospp14
if %winbuild% lss 9200 call :PKey %ospp% %o15App% osppsvc ospp15
)

:SPP
echo ..:: WINDOWS ACTIVATION STATUS ::.. >> %temp%\check.txt
if not defined W1nd0ws (
echo. >> %temp%\check.txt
echo Error: product key not found.>> %temp%\check.txt
echo. >> %temp%\check.txt
goto :SPPo
)
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (ApplicationID='%winApp%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%spp%" "%sps%" "%spp_get%"
  call :Output
)

:SPPo
set verbose=1
if not defined 0ff1ce15 (
if defined osppsvc goto :OSPP
goto :End
)
echo. >> %temp%\check.txt
echo ..:: OFFICE ACTIVATION STATUS ::.. >> %temp%\check.txt
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%spp%" "%sps%" "%spp_get%"
  call :Output
)
set verbose=0
if defined osppsvc goto :OSPP
goto :End

:OSPP
echo. >> %temp%\check.txt
if %verbose%==1 (
echo. >> %temp%\check.txt
echo ..:: OFFICE ACTIVATION STATUS ::.. >> %temp%\check.txt
)
if defined ospp15 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%ospp%" "%osps%" "%ospp_get%"
  call :Output
)
if defined ospp14 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o14App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :Property "%ospp%" "%osps%" "%ospp_get%"
  call :Output
  echo.
)
goto :End

:PKey
wmic path %1 where (ApplicationID='%2' and PartialProductKey is not null) get ID /value 2>nul | findstr /i ID 1>nul && (set %3=1&set %4=1)
exit /b

:Property
for %%# in (%~3) do set "%%#="
if %~1 equ %ospp% for %%# in (DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled) do set "%%#="
set "KmsClient="
for /f "tokens=* delims=" %%# in ('"wmic path %~1 where (ID='%chkID%') get %~3 /value" ^| findstr ^=') do set "%%#"

set /a gprDays=%GracePeriodRemaining%/1440
echo %Description%| findstr /i VOLUME_KMSCLIENT 1>nul && (set KmsClient=1)
call cmd /c exit /b %LicenseStatusReason%
set "LicenseReason=%=ExitCode%"

if %LicenseStatus%==0 (
set "License=Unlicensed"
set "LicenseMsg="
)
if %LicenseStatus%==1 (
set "License=Licensed"
set "LicenseMsg="
if not %GracePeriodRemaining%==0 set "LicenseMsg=Volume activation expiration: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==2 (
set "License=Initial grace period"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==3 (
set "License=Additional grace period (KMS license expired or hardware out of tolerance)"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==4 (
set "License=Non-genuine grace period."
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==6 (
set "License=Extended grace period"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus%==5 (
set "License=Notification"
  if "%LicenseReason%"=="C004F200" (set "LicenseMsg=Notification Reason: 0xC004F200 (non-genuine)."
  ) else if "%LicenseReason%"=="C004F009" (set "LicenseMsg=Notification Reason: 0xC004F009 (grace time expired)."
  ) else (set "LicenseMsg=Notification Reason: 0x%LicenseReason%"
  )
)
if %LicenseStatus% gtr 6 (
set "License=Unknown"
set "LicenseMsg="
)
if not defined KmsClient exit /b

if %KeyManagementServicePort%==0 set KeyManagementServicePort=1688
set "KmsReg=Registered KMS machine name: %KeyManagementServiceMachine%:%KeyManagementServicePort%"
if "%KeyManagementServiceMachine%"=="" set "KmsReg=Registered KMS machine name: KMS name not available"

if %DiscoveredKeyManagementServiceMachinePort%==0 set DiscoveredKeyManagementServiceMachinePort=1688
set "KmsDns=KMS machine name from DNS: %DiscoveredKeyManagementServiceMachineName%:%DiscoveredKeyManagementServiceMachinePort%"
if "%DiscoveredKeyManagementServiceMachineName%"=="" set "KmsDns=DNS auto-discovery: KMS name not available" >nul 2>&1

for /f "tokens=* delims=" %%# in ('"wmic path %~2 get ClientMachineID, KeyManagementServiceHostCaching /value" ^| findstr ^=') do set "%%#"
if /i %KeyManagementServiceHostCaching%==True (set KeyManagementServiceHostCaching=Enabled) else (set KeyManagementServiceHostCaching=Disabled)

if %winbuild% lss 9200 exit /b
if %~1 equ %ospp% exit /b

if "%DiscoveredKeyManagementServiceMachineIpAddress%"=="" set "DiscoveredKeyManagementServiceMachineIpAddress=not available"

if "%KeyManagementServiceLookupDomain%"=="" set "KeyManagementServiceLookupDomain="

if %VLActivationTypeEnabled%==3 (
set VLActivationType=Token
) else if %VLActivationTypeEnabled%==2 (
set VLActivationType=KMS
) else if %VLActivationTypeEnabled%==1 (
set VLActivationType=AD
) else (
set VLActivationType=All
)
exit /b

:Output
echo. >> %temp%\check.txt
echo Name: %Name% >> %temp%\check.txt
echo Activation ID: %ID% >> %temp%\check.txt
echo Extended PID: %ProductKeyID% >> %temp%\check.txt
if defined ProductKeyChannel echo Product Key Channel: %ProductKeyChannel% >> %temp%\check.txt
echo Partial Product Key: %PartialProductKey% >> %temp%\check.txt
echo License Status: %License% >> %temp%\check.txt
if defined LicenseMsg echo %LicenseMsg% >> %temp%\check.txt
if not %LicenseStatus%==0 if not %EvaluationEndDate:~0,8%==16010101 echo Evaluation End Date: %EvaluationEndDate:~0,4%-%EvaluationEndDate:~4,2%-%EvaluationEndDate:~6,2% %EvaluationEndDate:~8,2%:%EvaluationEndDate:~10,2% UTC >> %temp%\check.txt
if not defined KmsClient exit /b >> %temp%\check.txt
if defined DiscoveredKeyManagementServiceMachineIpAddress echo KMS Server IP address : %DiscoveredKeyManagementServiceMachineIpAddress% >> %temp%\check.txt
echo. >> %temp%\check.txt
if not %LicenseStatus%==1 (
echo Please activate the product in order to update KMS client information values.
exit /b 
) >> %temp%\check.txt
exit /b

:End
pushd "%temp%\"
check.vbs
del /f /q %temp%\check.txt >nul 2>&1
del /f /q %temp%\check.vbs >nul 2>&1
CLS
GOTO MAINMENU
::===============================================================================================================
:TNCTR
echo.
start https://www.tnctr.com/topic/450916-kms-dijital-online-aktivasyon-suite-v52/
CLS
GOTO MAINMENU
::===============================================================================================================
:Exit
cls
echo MSGBOX "SPECIAL THANKS : TNCTR Family - Hurly, CODYQX4, abbodi1406, qewlpal, s1ave77, cynecx, qad, Mouri_Naruto (MDL), WindowsAddict, mspaintmsi", vbInformation,"..:: mephistooo2 | TNCTR ::.."  > %temp%\TEMPmessage.vbs
call %temp%\TEMPmessage.vbs
rd /s /q %windir%\temp & md %windir%\temp
rd /s /q %TEMP% & md %TEMP%
exit
::===============================================================================================================
:KMSSuite
cd..
cd..
call KMS_Suite.cmd -suite
::===============================================================================================================
