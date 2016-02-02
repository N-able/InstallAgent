@echo off
rem Agent install script launcher
rem by Tim Wiser, GCI Managed IT (March 2015)

rem This file should be called by a GPO with three parameters:

rem   SITE_CODE
rem   DOMAIN_NAME or AUTO
rem   STARTUP or SHUTDOWN
rem eg. 100 ORCHID STARTUP



set SITE_CODE=%1%
set DOMAIN=%2%
set MODE=%3%


rem Change the name of the source if required
set EVENT_SOURCE=N-central agent installer

rem If you've changed the strAgentFolder variable in the VBS file, change this line as well
set AGENT_FOLDER=Agent

rem If the domain name is set as AUTO, we can now automatically detect it
if "%DOMAIN%"=="AUTO" ( wmic computersystem get domain | findstr /i . > %TEMP%\IA_dom.txt
			for /f "tokens=*" %%d in (%TEMP%\IA_dom.txt) do SET DOMAIN=%%d )



color 1f && cls
title :: %EVENT_SOURCE% ::


if NOT EXIST \\%DOMAIN%\netlogon\%AGENT_FOLDER%\installagent.vbs ( eventcreate /L APPLICATION /T ERROR /ID 998 /SO "%EVENT_SOURCE%" /D "The InstallAgent.vbs could not be reached on the network." 1>nul 2>nul
							) else ( %windir%\system32\cscript.exe \\%DOMAIN%\netlogon\%AGENT_FOLDER%\InstallAgent.vbs //nologo /site:%SITE_CODE% /mode:%MODE% )

set IAEXITCODE=%errorlevel%
if %IAEXITCODE% NEQ 10 eventcreate /L APPLICATION /SO "%EVENT_SOURCE%" /ID 998 /D "The %EVENT_SOURCE% script exited unexpectedly on this device.  This may be causing the agent to not install or upgrade properly.  The exit code was %IAEXITCODE%." /T ERROR 1>nul 2>nul
reg ADD "HKLM\Software\N-able Technologies\InstallAgent" /v "LastOperation" /d %IAEXITCODE% /t REG_DWORD /f 1>nul 2>nul
