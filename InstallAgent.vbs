' INSTALLAGENT.VBS

' Installation and maintenance script for the N-able agent software
' by Tim Wiser, GCI Managed IT


' 1.00 28/09/12 	- 	First release
' 1.10 02/10/12 	- 	Allows the repair process to run for seven mins
'				Recognises agent installs onto non C drives and amends verify stage accordingly
' 2.00 22/10/12 	-	Amended verify fail time to be 180 secs instead of 30 secs
'				Clarified that SXS path needs to be a mapped drive, not a UNC path
'				Added support for proxy.cfg file which contains proxy configuration to be used when installing the agent
'				Fixed null string bug in proxy code (thanks, Jon Czerwinski)
'				Added visible msgbox when ping test fails in interactive mode
' 3.00 16/11/12 	-  	Slight rephrasing of some event log messages
'				Agent is no longer verified immediately after installation
'				Now installs .NET 4 instead of .NET 2, due to 9.1.0.105 (9.1 beta) agent requiring this version
'				Removed support for installing .NET 3.5 on Windows 8 and Server 2012, .NET 4 is built into these platforms
'				Server Core now recognised and handled during .NET installer. Cannot yet deploy .NET to this OS as AgentCleanup doesn't work on this platform
'				Now forces elevation when double clicked
' 3.01 10/12/12 	-	Fixed a bug with OperatingSystemSKU not being recognised by 2003/XP devices
' 3.02 23/01/13 	-	Updated for agent 9.1.0.345
' 3.03 18/02/13 	-	Fixed issue where the script thought the agent had installed, yet hadn't
'				Updated to recognise 9.1.0.458 (9.1 GA) agent
' 3.04 21/02/13 	- 	Updated to resolve issue with .NET causing immediate reboot
'				Now properly detects the Full package of .NET 4 instead of seeing the Client Profile as adequate for the agent installer
' 3.05 02/04/13		-	Now displays the path of the agent if it's found to be installed when running in interactive mode
'				Updated to work with 9.2.0.142 agent installer
'				Now uses File Version of the installer EXE instead of checking the file size
' 3.06			- 	Added /nonetwork parameter which skips the test for Internet access on networks that block ping
'				No longer attempts to install on Windows 2000 unless bolNoWin2K is set to True
'				Now warns if Windows XP without SP3 is detected
' 3.07			-	Powershell awareness.  Writes a warning event if Powershell is not installed or is out of date
' 3.08			-	Double checks agent services when they're found and warns if registration isn't correct
' 3.09 01/08/13		-	Now correctly repairs agents with corrupted service registration
'				No longer alerts that .NET 4 is "INSTALLED" when running in interactive mode.  Only alerts if it's "NOT_INSTALLED"
'				Alternative source now works for all files, not just the agent source
'				Now copies AgentCleanup4.exe if the file versions differ between the local and network versions (ie: a new version is available)
'				Automatically installs Windows Imaging Components if not present
'				No longer tries to install .NET 4 on a Windows XP device if it doesn't have SP3 installed
' 3.10 24/09/13 	-	Small update to improve performance of WMI queries
' 3.11 23/06/14		-	Can no longer potentially downgrade the agent, causing a zombie device
'				Fixed error when agent file version is incorrect
'				Now tries to start up agent and agentmaint services if they're found to be stopped
'				Tidier messages by using spinner instead of dotted lines
'				Now logs if a proxy.cfg file is NOT found, which can be useful when trounleshooting in a proxy environment
'				Fixed problem with /source parameter not working properly
' 3.12 04/09/2014	-	Added in bolWarningsOn value which can prevent alerting when agent fails to install
' 3.13 07/01/2015	-	Added CleanQuit function which returns exit code 10 when the script exits cleanly, so any other code can be picked up by the batch file
' 3.14 10/02/2015	-	Changed code to allow broken service code to work on Windows XP
'				Added DirtyQuit fuction which returns passable exit code when the script exits prematurely. All Reg values are now written within this script
' 3.15 05/03/2015	-	Categorised exit codes (see documentation)
'			-	Fixed bug which prevented exit code from being written to Registry properly
' 3.16 02/07/2015	-	Exit code is now accompanied with a comment on what the code means
'				Appliance ID is read and written to event log for informational purposes
'				Fixed bug that caused agent version comparison check to fail on N-central 10 agents
'				Added ability to adjust the tolerance of the Connectivity (ping) test.  Change the CONST values if you need allow for dropped packets
' 4.00 02/11/2015	-	Formatting changes and more friendly startup message
'				Dirty exit now shows error message and contact information on console
'				Added 'Checking files' bit to remove confusing delay at that stage. No spinner though, unfortunately
'				This is the final release by Tim :o(
'				First version committed to git - Jon Czerwinski
' 4.01 20151109		-	Corrected agent version zombie check
'


Option Explicit

' Declare and define our variables and objects
Dim objFSO, output, objReg, objArguments, objArgs, objNetwork, objWMI, objShell, objEnv, objAppsList, objDomainInfo
Dim objCleanup, objServices, objInstallerProcess, objAgentInstaller, strAgentInstallerShortName, objMarkerFile
Dim objQuery, objDotNetInstaller, objFile, objExecObject, objWIC, colService, objCmd
Dim strComputer, strAgentPath, strAgentBin, strVersion, strCheckForFilme, strMSIflags, strNETversion, strInstallFlags
Dim strSiteID, strMode, strDomain, strInstallSource, strMessage, strWindowsFolder, strInstallCommand, strLine
Dim strApplianceConfig, arrApplianceConfig
Dim strKey, strType, strValue, strExitComment
Dim strOperatingSystem, strReqDotNetKey, strAltInstallSource, strTemp, strProxyString, strProxyConfigFile
Dim strBaselineListOfPIDs, strInstallationListOfPIDs, strInstallProxyString, strArchitecture, strDotNetInstallerFile
Dim strOperatingSystemSKU, strNoNetwork, strPSValue, strWICinstaller, strString, strDecimal, strSpin
Dim bolInteractiveMode, bolAgentServiceFound, bolAgentMaintServiceFound
Dim intServicePackLevel, intPSValue, intValue
Dim item, response, count, service
Dim intPingFailureCount



strComputer = "."
strAgentPath = ""
strDomain = "UNKNOWN.TIM"		' This is a dummy domain name, do NOT change
strSiteID = ""
strMode = ""
strSpin = "|"
strVersion = "4.01"				' The current version of this script
bolInteractiveMode = False
intPingFailureCount = 0
CONST HKEY_LOCAL_MACHINE = &H80000002


' ***********************************************************************************************************************************************************
' Define some constants for your environment
CONST strServerAddress = "YOUR.FQDN.HERE"			' The FQDN or IP address of your N-central server
CONST strBranding = "N-able"					' Branding label, eg. 'ACME MSP monitoring'.  Used in popup messages as the window title
CONST strAgentFolder = "Agent"					' This is the name of the folder within NETLOGON where you have placed the agent installer

' Update these when a new agent is released
CONST strRequiredAgent = "10.0.1.274"				' This should be the N-central version of the agent that has been deployed out to customers
CONST strAgentFileVersion = "10.0.10274.0"			' This is the 'File Version' of the agent installer file itself, obtained from the Details tab in Explorer
CONST bolNoWin2K = True						' Set this to False if you want to attempt installs on Windows 2000, which is only supported on pre-9.2 agents
CONST bolWarningsOn = False					' If the agent fails, the user gets a popup message.  Change this to False if you don't want them to be informed

CONST strPingAddress = "8.8.8.8"				' Address to ping to check connectivity.  Usually this will be your N-central server's FQDN
CONST strSOAgentEXE = "SOAgentSetup.exe"			' The name of the SO-level agent installer file that is stored within NETLOGON\strAgentFolder
CONST strContactAdmin = "Please contact your MSP for assistance."		' Your contact info

CONST intPingTolerance = 20							' % value.  If you have networks that drop packets under normal conditions, raise this value
CONST intPingTestCount = 10							' Increase this to make the script perform more pings during the Connectivity test
' ***********************************************************************************************************************************************************



' Create objects that we're going to be using
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("PROCESS")
Set objNetwork = CreateObject("WScript.Network")
Set output = WScript.Stdout
Set objArguments = WScript.Arguments
Set objArgs = objArguments.Named

' Get the mandatory parameters
strSiteID = objArgs("site")
strMode = objArgs("mode")
strNoNetwork = objArgs("nonetwork")
strAltInstallSource	= objArgs("source")
strMode = UCase(strMode)
strNoNetwork = UCase(strNoNetwork)

' Make sure that the Registry key that we write to actually exists
objReg.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent"

' Work out which scripting engine we're running under as we behave differently according to which one we're using
If Instr(1, WScript.FullName, "CScript", vbTextCompare) = 0 Then
	bolInteractiveMode = True
	objShell.LogEvent 0, "The agent installer script is running in interactive mode."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastMode", "Interactive"
	' Ensure that the script runs in elevated mode
	If WScript.Arguments.Named.Exists("elevated") = False Then
		objShell.LogEvent 2, "The agent installer script is elevating its execution"
		CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
		WScript.Quit
	End If
Else
	bolInteractiveMode = False
	objShell.LogEvent 0, "The agent installer script is running in non-interactive (scripted) mode."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastMode", "Unattended"
	' Support the /? parameter on the command line to get a little bit of help
	If objArguments(0) = "/?" Then
		WScript.Stdout.Write "InstallAgent.vbs version " & strVersion & vbCRLF & "by Tim Wiser, GCI Managed IT (tim.wiser@gcicom.net)" & vbCRLF & vbCRLF
		WScript.Stdout.Write "Syntax: InstallAgent.vbs /site:[ID] /mode:[STARTUP|SHUTDOWN] (/source:[ALTERNATIVE SOURCE PATH]) (/nonetwork:yes)" & vbCRLF & vbCRLF
		CleanQuit
	End IF
End If

' Say we've started execution
objShell.LogEvent 0, "The agent installer script being executed is version v" & strVersion
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastRun", "" & Now()
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "Version", strVersion
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastExitComment", "Script started successfully"
WRITETOCONSOLE("Please wait whilst the " & strBranding & " agent is checked." & vbCRLF & vbCRLF)




' Get the domain that this device is a member of
Set objDomainInfo = objWMI.ExecQuery("SELECT Domain FROM Win32_ComputerSystem")
For Each item In objDomainInfo
	strDomain = item.Domain
Next
			
If strDomain = "UNKNOWN.TIM" Then
	objShell.LogEvent 2, "The agent installer script was unable to determine what the local domain is for this device.  If the script was run on a home or workgroup device, please run the standard agent installer executable instead."
	DIRTYQUIT 0
Else
	strInstallSource = "\\" & strDomain & "\Netlogon\" & strAgentFolder & "\"
End If



' Perform a connectivity test
If strNoNetwork <> "YES" Then
	For count = 1 to intPingTestCount
		Set objQuery = objWMI.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address='" & strPingAddress & "'")
		WRITETOCONSOLE("Checking connectivity ............ " & SPIN & vbCR)
		For Each item in objQuery
			If item.StatusCode = 0 Then
				' successful ping, great stuff, happy days!
			Else
				' ping dropped, bad times
				intPingFailureCount = intPingFailureCount + 1
			End If
		Next
	Next
	
	' Terminate if the ping test failed the tolerance test
	If ((intPingFailureCount / intPingTestCount) * 100) > intPingTolerance Then
		WRITETOCONSOLE("Checking connectivity ............ failed!" & vbCRLF)
		objShell.LogEvent 1, "The agent installer script has detected that this device does not have access to the central server.  The ping check to " & strPingAddress & " failed with a fault rate of " & (intPingFailureCount / intPingTestCount) * 100 & "%.  Please check connectivity and try again.  If network conditions are generally poor you can adjust the tolerance.  Refer to the script documentation for further assistance.  The script will now terminate."
		If bolInteractiveMode = True Then
			MsgBox "This device cannot ping the external test address of " & strPingAddress & ".  Please check connectivity and try again.", vbOKOnly + vbCritical, strBranding
		End If
		DIRTYQUIT 2
	End If
	
	' Warn if packets were dropped	
	If ((intPingFailureCount / intPingTestCount) * 100) > 0 Then
		WRITETOCONSOLE("Checking connectivity ............ done!" & vbCRLF)
		objShell.LogEvent 2, "The agent installer script has detected that this device has connectivity to " & strPingAddress & " but is dropping packets.  The ping test had a fault rate of " & (intPingFailureCount / intPingTestCount) * 100 & "%.  If the fault rate passes " & intPingTolerance & "% the script will not run."
	Else	
		' All was well, we got 100% ping connectivity
		WRITETOCONSOLE("Checking connectivity ............ done!" & vbCRLF)
		objShell.LogEvent 0, "The agent installer script has detected that this device has reliable connectivity to " & strPingAddress & "."
	End If
End If



' Lets you override the automatic assumption that you are installing from NETLOGON by specifying an alternative path in /source
If strAltInstallSource <> "" Then
	strInstallSource = strAltInstallSource
	If Right(strAltInstallSource, 1) <> "\" Then strAltInstallSource = strAltInstallSource & "\" End If
	
	If objFSO.FileExists(strAltInstallSource & strSOAgentEXE) = False Then
		objShell.LogEvent 1, "The agent installer script could not validate the alternative path to the agent installer that was specified, which was " & strInstallSource
	Else
		objShell.LogEvent 0, "The agent installer script will use the agent installer file " & strInstallSource & strSOAgentEXE
	End If
End If


' Ensure that the install source path has a slash at the end of it, otherwise it won't form a valid path
If Right(strInstallSource, 1) <> "\" Then strInstallSource = strInstallSource & "\" End If

' Write the strInstallSource path into the Registry for the custom service to read
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "Path", strInstallSource

' Validate that the agent installer matches the size specified in the CONST variables at the top of this script
If (objFSO.FileExists(strInstallSource & strSOAgentEXE)) = False Then	
	strMessage = "The agent cannot be installed on this computer.  The install source " & strInstallSource & strSOAgentEXE & " is missing or invalid.  Additional information can be found in the Application event log." & vbCRLF & vbCRLF & strContactAdmin
	objShell.LogEvent 1, strMessage
	If (bolWarningsOn = True or bolInteractiveMode = True) Then
		Msgbox strMessage, vbOKOnly + vbCritical, strBranding
	End If
	DIRTYQUIT 3
End If



' Check the File Version of the available installer file and check that it's the right version
WRITETOCONSOLE("Checking files ................... ")
If objFSO.GetFileVersion(strInstallSource & strSOAgentEXE) <> strAgentFileVersion Then
	strMessage = "The agent installer script has found that the agent installer at '" & strInstallSource & strSOAgentEXE & "' is the wrong version.  The expected version is " & strAgentFileVersion & ".  The detected version is " & objFSO.GetFileVersion(strInstallSource & strSOAgentEXE) & ".  The installation will not continue."
	WRITETOCONSOLE("failed!" & vbCRLF)
	objShell.LogEvent 1, strMessage
	If bolInteractiveMode = True Then
			MsgBox strMessage & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
	End If
	DIRTYQUIT 3
Else
	WRITETOCONSOLE("done!" & vbCRLF)
	objShell.LogEvent 0, "The agent installer script has validated that the available installer file (version " & objFSO.GetFileVersion(strInstallSource & strSOAgentEXE) & ") is the required version, " & strAgentFileVersion
End If



' See if we're wanting to use a proxy server when installing the agent.  If so, read the proxy configuration file
strProxyConfigFile = "\\" & strDomain & "\Netlogon\" & strAgentFolder & "\proxy.cfg"
If objFSO.FileExists(strProxyConfigFile)= True Then
	objShell.LogEvent 0, "The agent installer script has found a proxy.cfg file on the network"
	WRITETOCONSOLE("Configuring proxy ................ ")
	Set objFile = objFSO.OpenTextFile(strProxyConfigFile)
		strProxyString = objFile.ReadLine
	objFile.Close
	objShell.LogEvent 0, "The agent installer script will install the agent using: " & strProxyString
	WRITETOCONSOLE("done!" & vbCRLF)
Else
	' No proxy.cfg found, so log for troubleshooting purposes
	objShell.LogEvent 0, "The agent installer script did not find a proxy configuration file to use"
End If
	


' Copy agentcleanup4.exe locally so we have it available later on.  NOTE! This version will only run if .NET 4 is installed
strWindowsFolder = objShell.ExpandEnvironmentStrings("%WINDIR%")
If objFSO.FileExists(strWindowsFolder & "\AgentCleanup4.exe") = False Then
	COPYAGENTCLEANUP
Else
	' Check that the version of the cleanup tool is the same as the one on the network share.  If it's different, copy it locally
	If objFSO.GetFileVersion(strWindowsFolder & "\AgentCleanup4.exe") <> objFSO.GetFileVersion(strInstallSource & "\AgentCleanup4.exe") Then
		COPYAGENTCLEANUP
	End If
End If


' Pop up a message asking if the device should be excluded from agent deployment in the future if running interactively
If bolInteractiveMode = True Then
	' Check to see if the nable_disable.mrk file is present and, if it is, ask if it should be deleted
	If objFSO.FileExists(strWindowsFolder & "\nable_disable.mrk") = True Then
		' The agent is currently prevented from deploying on this device, so ask if this should remain the case
		response = MsgBox("The agent is currently prevented from deploying onto this device.  Do you want to enable deployment in the future?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Agent istaller")
		Select Case response
		Case vbYes		:		objFSO.DeleteFile(strWindowsFolder & "\nable_disable.mrk")
		Case vbCancel	:		MsgBox "Agent installation has been cancelled.", 48, strBranding
								CleanQuit
		End Select
	Else
		response = MsgBox("Do you want to PREVENT the agent from deploying onto this device in the future?", vbYesNo + vbQuestion + vbDefaultButton2, strBranding)
		Select Case response
		Case vbYes		:		' Create the nable_disable.mrk file which prevents the agent from installing in the future
								Set objMarkerFile = objFSO.CreateTextFile(strWindowsFolder & "\nable_disable.mrk")
								objMarkerFile.Write "The N-able agent will not deploy onto this device."
								objMarkerFile.Close
								MsgBox "The agent will not be installed on this device in the future via this script.  It can still be installed manually." & vbCRLF & vbCRLF & "Exiting now.", vbExclamation + vbOKOnly, strBranding
								CleanQuit
		End Select
	End If
End If


' Is the N-able marker file present?  If so, stop immediately
If objFSO.FileExists(strWindowsFolder & "\nable_disable.mrk") = True Then
	objShell.LogEvent 2, "nable_disable.mrk is present inside the " & strWindowsFolder & " folder, so this device should not have the agent installed.  Exiting now."
	CleanQuit
End If	

strAgentBin = GETAGENTPATH										' Contains the full path of agent.exe including quotes
If Len(strAgentBin) > 1 Then
	strAgentPath = Left(strAgentBin, LEN(strAgentBin)-15)				
	strAgentPath = Right(strAgentPath, Len(strAgentPath)-1)		' Contains the full path of the Window Agent folder, excluding quotes
Else
	strAgentPath = ""
End If


' Detect the OS and type
Set objQuery = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem WHERE Caption LIKE '%Windows%'")
For Each item In objQuery
	strOperatingSystem = item.Caption
	intServicePackLevel = item.ServicePackMajorVersion
Next


' Detect and warn if PowerShell 2 is not installed
objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\", "PowerShellVersion", strPSValue
If IsNull(strPSValue) Then
	' Warn that no version of Powershell at all is installed on the device
	objShell.LogEvent 1, "The agent installer script cannot find Microsoft PowerShell installed on this device.  For use of advanced automation within N-central a device must have PowerShell v2 or greater installed."
Else
	intPSValue = CInt(Left(strPSValue,1))
	If intPSValue < 2 Then
		' PSH 1 isn't good enough for N-able Automation Manager
		objShell.LogEvent 1, "The agent installer script found an outdated version of PowerShell installed on this device. For use of advanced automation within N-central a device must have PowerShell v2 or greater installed."
	Else
		' PSH 2 onwards is suitable for N-able Automation Manager
		objShell.LogEvent 0, "The agent installer script found PowerShell v" & intPSValue & " installed on this device.  This version is suitable for use with N-able Automation Manager policies (.AMP files)"
	End If
End If



' Complain if Windows 2000 is detected and bolNoWin2K is set to True (default)
If bolNoWin2K = True Then
	If Instr(strOperatingSystem, "2000")>0 Then
		strMessage = "The agent cannot be installed on Windows 2000."
		objShell.LogEvent 1, "The agent installer script has detected that this device is running Windows 2000.  The agent cannot be installed onto this operating system due to the pre-requisites of Microsoft .NET Framework 4.  Exiting now."
		WRITETOCONSOLE(strMessage & vbCRLF)
		If bolInteractiveMode = True Then
			MsgBox strMessage, vbOKOnly + vbCritical, strBranding
		End If
		DIRTYQUIT 1
	End If
Else
	WRITETOCONSOLE(strOperatingSystem & " detected" & vbCRLF)
End If

' Warn if the device is running Windows XP and isn't patched to SP3 level
If Instr(strOperatingSystem, "XP")>0 Then
	If intServicePackLevel < 3 Then
		objShell.LogEvent 1, "The agent installer script has detected that this device is running Windows XP but is not patched with Service Pack 3.  This will prevent .NET 4 from installing which will in turn prevent the agent being able to be deployed or maintained.  Please install Service Pack 3 on this device.  The script will now terminate."
		WRITETOCONSOLE("Checking Windows XP SP3 .......... failed!" & vbCRLF & vbCRLF & "Please install Service Pack 3 on this computer." & vbCRLF)
		DIRTYQUIT 3
	End If
End If

' If the script is running in interactive mode, check to see if the agent is installed and bug out if it is
If (bolInteractiveMode = True And strAgentBin <> "") Then
	objShell.LogEvent 0, "The agent installer found that the agent is already installed on this device at " & strAgentBin
	response = MsgBox("The agent is installed on this device at " & strAgentBin & vbCRLF & vbCRLF & "Do you want to uninstall it?  This can take up to five minutes to complete.", vbYesNo+vbQuestion, strBranding)
	If response = vbNo Then
		CleanQuit
	Else
		If DOTNETPRESENT("v4\Full")="NOT_INSTALLED" Then
			' The AgentCleanup4 utility requires .NET 4 so if it's not installed, let's just try a clean removal using Windows Installer
			objShell.LogEvent 2, "The agent installer script initiated a clean removal of the existing agent using MSIEXEC on user request."
			Call objShell.Run("msiexec /X {07BA9781-16A5-4066-A0DF-5DBA3484FDB2} /passive /norestart",,True)
		Else
			' Do an uninstall using AgentCleanup4.exe as we have .NET 4 installed on the device
			objShell.LogEvent 2, "The agent installer script initiated a cleanup of the existing agent on user request."
			Call objShell.Run("cmd /c agentcleanup4.exe writetoeventlog",,True)
		End If
	End If
	CleanQuit
End If


' Validate the mandatory parameters if we're running in non-interactive mode
strMessage = "The agent installer script was passed an insufficient number of parameters.  Both /site:[ID] and /mode:[STARTUP|SHUTDOWN] are required for this script to function. The agent installation cannot continue on this device.  Please check the GPO configuration."
If bolInteractiveMode = False Then	
	If (strSiteID = "" Or strMode = "") Then
		If bolInteractiveMode = True Then
			MsgBox strMessage, 36, strBranding
		End If
		objShell.LogEvent 1, strMessage
		DIRTYQUIT 0
	End If
End If

	


' Decide whether we need to install a new agent or verify an existing agent
If strAgentBin = "" Then
	' AGENT IS NOT INSTALLED, SO INSTALL IT
	objShell.LogEvent 2, "The agent installer script did not find the agent on this device so will attempt to install it."

	' Pop up a prompt for interactive users to enter the site ID code manually
	If bolInteractiveMode = True Then
		strSiteID = InputBox("Welcome to the agent installer script. © Tim Wiser, GCI Managed IT" & vbCRLF & vbCRLF & "Please enter the numerical site code which is found by navigating to the SO level and selecting Administration -> Customers/Sites." & vbCRLF & vbCRLF & "NOTE: If you do not understand this message, please click the Cancel button and contact your IT administrator.", strBranding)
		If strSiteID = "" Then
			objShell.LogEvent 1, "The agent installer script aborted by user request."
			MsgBox "Agent installation has been aborted.", 48, strBranding
			CleanQuit
		End If
		strMode = "STARTUP"
	End If
	

	' We can't install the agent on shutdown due to Windows killing off the script before the installer has completed, so we create
	' a registry key which will perform the installation the next time it starts up, whether it's on the domain or not.
	If strMode = "SHUTDOWN" Then
		WRITETOCONSOLE("Shutting down")
		objShell.LogEvent 2, "The agent installer script is not able to deploy the agent whilst Windows is shutting down."
		CleanQuit
	End If

	
	' Check for .NET Framework v4
	If DOTNETPRESENT("v4\Full") = "NOT_INSTALLED" Then
	
		' Decide which installer to use (deprecated Nov 2012, we always install the same one now)
		strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
		strArchitecture = objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
		strDotNetInstallerFile = ""
		Select Case strArchitecture
			Case "AMD64"	: strDotNetInstallerFile = "dotNetFx40_Full_x86_x64.exe"
			Case "x86"		: strDotNetInstallerFile = "dotNetFx40_Full_x86_x64.exe"		' Originally dotNetFx40_Full_x86.exe, but the 64 bit includes 32 bit support as well
		End Select
		
		' Firstly, check to see if we're running on a Core installation of Windows.  We cannot install the .NET Framework onto Server Core automatically 
		If (Instr(strOperatingSystem, "2008")>0 Or Instr(strOperatingSystem, "2012")>0) Then
			Set objQuery = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem WHERE Caption LIKE '%Windows%'")
			For Each item In objQuery
				strOperatingSystemSKU = item.OperatingSystemSKU
			Next
			Select Case strOperatingSystemSKU
				Case 12, 39, 14, 41, 13, 40, 29	:	WRITETOCONSOLE(vbCRLF & "Server Core detected!" & vbCRLF & "Please refer to the Event Log for futher details" & vbCRLF)
													objShell.LogEvent 1, "The agent installer script detected that this device is running Windows Server Core and that Microsoft .NET Framework 4 is not currently installed.  This needs to be installed before the agent can be installed but the script is unable to install it automatically." & vbCRLF & vbCRLF & "Please install the .NET Framework 4 for Server Core from http://www.microsoft.com/en-us/download/details.aspx?id=22833 or from within N-central and run the script again."
													strDotNetInstallerFile = "dotNetFx40_Full_x86_x64_SC.exe"
													DIRTYQUIT 1
			End Select
		End If
		
		' Check to see if the WindowsCodecs.dll file is present, which is an indication of the Windows Imaging Components feature which is a pre-req for .NET 4
		strWindowsFolder = objShell.ExpandEnvironmentStrings("%WINDIR%")
		If (objFSO.FileExists(strWindowsFolder & "\System32\WindowsCodecs.dll")=False) Then
			' WIC is not installed, so let's see if we can install it ourselves from strInstallSource
			Select Case strArchitecture
				Case "AMD64"	: strWICinstaller = "wic_x64_enu.exe"
				Case "x86"		: strWICinstaller = "wic_x86_enu.exe"
			End Select
			
			If objFSO.FileExists(strInstallSource & strWICinstaller) = True Then
				' Install it
				objShell.LogEvent 1, "The agent installer script has detected that Windows Imaging Components is not installed on this device and will attempt to install it from " & strInstallSource & strWICinstaller
				Set objWIC = objShell.Exec("cmd /c " & strInstallSource & strWICinstaller & " /quiet")
				Do While objWIC.Status = 0
					WRITETOCONSOLE("Installing Windows Imaging Components " & strArchitecture & " ... " & SPIN & vbCR)
					WScript.Sleep 100
				Loop
				WRITETOCONSOLE("Installing Windows Imaging Components " & strArchitecture & " ... done!" & vbCR)
			Else
				' We can't install WIC as the required file isn't present in strInstallSource
				objShell.LogEvent 2, "The agent installer script has detected that Windows Imaging Components may not be installed on this device.  For this reason, Microsoft .NET Framework 4 may fail to install.  You can download the 32-bit WIC installer from http://www.microsoft.com/en-us/download/details.aspx?id=32.  The script is capable of installing WIC automatically as long as the WIC_x86_enu.exe and WIC_x64_enu.exe files are present inside " & strInstallSource
			End If
		End If
		
		
		objShell.LogEvent 2, "The agent installer script detected that Microsoft .NET Framework 4 Full Package is not currently installed.  This needs to be installed before the agent can be installed.  The script will now attempt to install Microsoft .NET Framework 4 Full Package on this device."
		
		objEnv("SEE_MASK_NOZONECHECKS") = 1
		
		WRITETOCONSOLE("Checking for .NET v4 installer ... ")
		If (objFSO.FileExists(strInstallSource & strDotNetInstallerFile) = False) Then
			' The .NET installer file that we need to run doesn't exist, so error out
			strMessage = "The agent installer could not install Microsoft .NET Framework 4 as the installer file, " & strDotNetInstallerFile & ", does not exist at " & strInstallSource & strDotNetInstallerFile & ".  Please download the installer from N-central and try again, or install Microsoft .NET Framework 4 manually on this device.  The script will now terminate."
			If strDotNetInstallerFile = "dotNetFx40_Full_x86_x64_SC.exe" Then strMessage = strMessage & vbCRLF & vbCRLF & "NOTE:  As this server is running a Core edition of Windows you will need to download the '.NET Framework 4 Server Core - x64' installer from within N-central and store it in the " & strInstallSource & " folder.  Once this is done, .NET should install automatically the next time this script runs."
			If bolInteractiveMode = True Then
				MsgBox strMessage, vbCritical + vbOKOnly, strBranding
			End If
			objShell.LogEvent 1, strMessage
			WRITETOCONSOLE("failed!" & vbCRLF)
			DIRTYQUIT 1
		Else
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
			
		' Copy the .NET installer file locally before installation
		WRITETOCONSOLE("Copying .NET v4 installer ........ ")
		objFSO.CopyFile strInstallSource & strDotNetInstallerFile, strTemp & "\dotnetfx.exe"
		If (objFSO.FileExists(strTemp & "\dotnetfx.exe"))=True Then 
			objShell.LogEvent 0, "The agent installer script successfully copied the Microsoft .NET 4 Full Package installer for " & strArchitecture & ", " & strDotNetInstallerFile & " into " & strTemp & " as dotnetfx.exe"
			WRITETOCONSOLE("done!" & vbCRLF)
		Else
			objShell.LogEvent 1, "The agent installer script could not copy " & strDotNetInstallerFile & " into " & strTemp & " as dotnetfx.exe so cannot continue with the installation.  Check file permissions and that the " & strDotNetInstallerFile & " file exists within \\" & strDomain & "\Netlogon\" & strAgentFolder
			WRITETOCONSOLE("failed!" & vbCRLF)
			DIRTYQUIT 1
		End If
		
		' Now start the installer up from the local copy
		objShell.LogEvent 0, "The agent installer script started the installation of Microsoft .NET Framework 4 Full Package on this device."
		Set objDotNetInstaller = objShell.Exec(strTemp & "\dotnetfx.exe /passive /norestart /l c:\dotnetsetup.htm" & Chr(34))
		Do Until objDotNetInstaller.Status <> 0
			WRITETOCONSOLE("Installing .NET v4 Full Package .. " & SPIN & vbCR)
			WScript.Sleep 100
		Loop
		
		objEnv.Remove("SEE_MASK_NOZONECHECKS")
		' We usually terminate here, as .NET seems to kill off any script that calls it for some reason.  The code below is included in case we're running
		' in interactive mode or to support rare cases where the .NET installer allows the script to continue running.
		If DOTNETPRESENT("v4\Full") = "NOT_INSTALLED" Then
			' .NET failed to install, so pop a message up and log an event, then exit
			WRITETOCONSOLE("Installing .NET v4 Full Package ... failed!" & vbCRLF)
			strMessage = "The agent installer script failed to install Microsoft .NET Framework 4 Full Package on this device." 
			objShell.LogEvent 1, strMessage & vbCRLF & vbCRLF & "If this computer is running Windows XP then check that it is running SP3 and has Windows Installer 3.1 installed.  Please check other supported operating systems and pre-requisities with Microsoft at http://www.microsoft.com/en-gb/download/details.aspx?id=17718 and try again."
			Msgbox strMessage & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
			DIRTYQUIT 4
		Else
			' .NET 4 Full Package installed successfully
			WRITETOCONSOLE("Installing .NET v4 Full Package .. done!" & vbCRLF)
			objShell.LogEvent 0, "The agent installer script successfully installed Microsoft .NET Framework 4 Full Package on this device."
		End If
	Else
		objShell.LogEvent 0, "The agent installer script found that the Microsoft .NET Framework Full Package is installed on this device."
	End If

		
	' Branch off for an install.  We no longer do a verify of the agent afterwards -  we assume that the agent is installed and do checks during the Function
	INSTALLAGENT strSiteID, strDomain, strMode
	
Else
	' AGENT IS INSTALLED, SO DO A VERIFY
	strMessage = "The agent installer script has determined that the agent is already installed on this device.  The binary was found at " & strAgentBin
	If bolInteractiveMode = True Then
		MsgBox strMessage, 64, strBranding
	End If
	objShell.LogEvent 0, strMessage
	' Branch off for a verify
	VERIFYAGENT
	
	' Check to see if the agent services exist.  If not, the agent verify stage has caused a removal of the (old or corrupted) agent, so let's reinstall it now
	Set objServices = objWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
	If objServices.Count = 0 Then
		' The agent services are no longer present
		INSTALLAGENT strSiteID, strDomain, strMode
	End If
		
	
	
End IF


CleanQuit
' THIS IS THE END OF THE MAIN SCRIPT BODY


' ***************************************************************
' Function - InstallAgent - Performs an installation of the agent
' ***************************************************************
Function INSTALLAGENT(strSiteID, strDomain, strMode)
	strMessage = "Install agent for N-central site ID " & strSiteID & " on this device which is in domain " & strDomain & " from " & strInstallSource & " in " & strMode & " mode"
	' If we're in interactive mode, pop up a question to confirm that we want to start the installation
	If bolInteractiveMode = True Then
		response = MsgBox(strMessage & "?", 36, strBranding)
		If response = vbNo Then
			MsgBox "Agent installation aborted.", 48, strBranding
			CleanQuit
		End If
	End If
			
	
	' Run a cleanup in preparation for the agent to be installed and wait for it to finish
	VERIFYAGENT
	
	' If we're running in interactive mode then make the agent installation less hands-off
	strMSIflags = " /S /v" & Chr(34) & " /qn "
	If bolInteractiveMode = True Then
		strMSIflags = " /v" & Chr(34) & " /qb "
	End If
	

	' Build the install command string
	strInstallProxyString = " "
	If strProxyString <> "" Then
		strInstallProxyString = " AGENTPROXY=" & strProxyString & " "
	End If
	
	strInstallFlags = strMSIflags & "CUSTOMERID=" & strSiteID & " CUSTOMERSPECIFIC=1" & strInstallProxyString & "SERVERPROTOCOL=HTTPS SERVERADDRESS=" & strServerAddress & " SERVERPORT=443" & Chr(34)
	strInstallCommand = strInstallSource & strSOAgentEXE & strInstallFlags
	
	' Launch the agent installer and wait for it to finish
	objShell.LogEvent 0, "The agent installer script has started an installation of the agent using the following command:  " & strInstallCommand
	
	objEnv("SEE_MASK_NOZONECHECKS") = 1				' Turn off the Open File confirmation window
	
	WRITETOCONSOLE(vbCRLF & "Starting installer ............... ")
	Call objShell.Exec(strInstallCommand)
	WRITETOCONSOLE("done!" & vbCRLF)
	
	' Wait for the Windows Agent Maintenance Service to become present
	Do While count < 3600	' allow 6 mins for installation to complete (based on 100 millisecond interval)
		Set objQuery = objWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Caption='Windows Agent Maintenance Service'")
		If objQuery.Count = 1 Then
			WRITETOCONSOLE("Waiting for services ............. done!" & vbCRLF)
			Exit Do
		End If
		WRITETOCONSOLE("Waiting for services ............. " & SPIN & vbCR)
		count = count + 1
	Loop
	
	objEnv.Remove("SEE_MASK_NOZONECHECKS")			' Turn on the Open File confirmation window
	
	Set objQuery = objWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Caption='Windows Agent Maintenance Service'")
	If objQuery.Count <> 1 Then	
		WRITETOCONSOLE("Waiting for services ............. failed!")
		objShell.LogEvent 1, "The agent installer script failed to install the agent on this computer within the permitted timeframe."
		If (bolWarningsOn = True Or bolInteractiveMode = True) Then
			MsgBox "The agent could not be installed on this computer." & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
		End If
		DIRTYQUIT 5
	Else
		objShell.LogEvent 0, "The agent installer detected that the Windows Agent Maintenance service has registered on this computer.  The agent is now installed and will register in N-central shortly."
	End If
	
End Function


' ***********************************************************************************************
' Function - DotNetPresent - Checks to see if .NET Framework of a particular version is installed
' ***********************************************************************************************
Function DOTNETPRESENT(strReqDotNetKey)

	objReg.GetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\NET Framework Setup\NDP\" & strReqDotNetKey, "Install", strNETversion
	If strNETversion <> "" Then
		strNETversion = "INSTALLED"
	Else
		strNETversion = "NOT_INSTALLED"
		
		' Tell the interactive user that it's not installed
		If bolInteractiveMode = True Then
			msgbox ".NET Framework " & strReqDotNetKey & " is not installed on this device.", vbOKOnly + vbInformation, strBranding
		End If
	End If
	
	DOTNETPRESENT = strNETversion
End Function



' ***********************************************************************
' Function - GetAgentPath - Returns the true path to the agent.exe binary
' ***********************************************************************
Function GETAGENTPATH
	Set objServices = objWMI.ExecQuery("SELECT Pathname FROM Win32_Service WHERE Name = 'Windows Agent Service'")
	strAgentBin = ""
	If objServices.Count <> 0 Then
		' At least one service (hopefully only one!) was returned by the query, so let's find the path to that service
		For Each item In objServices
			strAgentBin = item.Pathname
			objShell.LogEvent 0, "The agent installer script found the Windows Agent Service's path to be:  " & strAgentBin
		Next
	Else
		' The agent isn't installed - no service was returned by the WQL query.  An empty string here indicates no agent installed on the device to other sections of this script
		strAgentBin = ""
	End If
	
	
	' Check to see if this script could potentially downgrade the agent, which N-central doesn't like AT ALL :-(
	WRITETOCONSOLE("Checking downgrade ... ")
	If strAgentBin <> "" Then
		If IsDowngrade(objFSO.GetFileVersion(Mid(strAgentBin,2, Len(strAgentBin)-2)), strRequiredAgent) Then
			objShell.LogEvent 1, "The agent installer script found that this device already has agent version " & objFSO.GetFileVersion(Mid(strAgentBin,2, Len(strAgentBin)-2)) & " installed which is newer than the version available for installing, which is " & strRequiredAgent & ".  If maintenance was carried out on this device it could potentially effectively downgrade the agent which would result in a zombie device.  Therefore, this script will now terminate."
			WRITETOCONSOLE("failed!" & vbCRLF)
			DIRTYQUIT 6
		Else
			objShell.LogEvent 0, "The agent installer found that the installed agent is suitable for maintenance."
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
	End If

	' The agent isn't installed, there's no risk of a downgrade
	If strAgentBin = "" Then
		objShell.LogEvent 0, "The agent installer found that there is no risk of a downgrade occurring."
	End If
	
	WRITETOCONSOLE("done!" & vbCRLF)
	
	GetAgentPath = strAgentBin
End Function


' ***********************************************************************
' Function - StripAgent - Uses AgentCleanup4 to entirely remove the agent
' ***********************************************************************
Function STRIPAGENT
	If DOTNETPRESENT("v4\Full")="NOT_INSTALLED" Then
		WRITETOCONSOLE(".NET 4 is not installed - cannot run cleanup" & vbCRLF)
		objShell.LogEvent 1, "The agent installer script cannot perform a cleanup of the agent as Microsoft .NET Framework 4 is not installed on this device.  The script will now terminate."
		DIRTYQUIT 1
	Else		
		' Now we proceed and do a cleanup of the agent
		strString = ""
		objShell.LogEvent 1, "The agent installer script is running the cleanup process.  This happens when either the agent is not installed (in which case the cleanup is being done in preparation for a fresh installation of the agent) or if not all the services are present (in which case the agent is going to be reinstalled)."
		Set objCleanup = objShell.Exec("cmd /c AgentCleanup4.exe")
		
		Do While objCleanup.Status = 0
			'WRITETOCONSOLE(".")
			WRITETOCONSOLE("Preparing agent .................. " & SPIN & vbCR)
			WScript.Sleep 100	' interval delay in milliseconds
		Loop
		WRITETOCONSOLE("Preparing agent .................. done!" & vbCRLF)
	End If
End Function


' ***********************************************
' Sub function - SPIN - Provides a spinning wheel
' ***********************************************
Function SPIN
	Select Case strSpin
		Case "|"	: strSpin = "/"
		Case "/"	: strSpin = "-"
		Case "-"	: strSpin = "\"
		Case "\"	: strSpin = "||"
		Case "||"	: strSpin = "/"
		Case "/"	: strSpin = "--"
		Case "--"	: strSpin = "\\"
		Case "\\"	: strSpin = "|"
	End Select
	SPIN = Left(strSpin, 1)
End Function


' ********************************************************************************************
' Sub function - VerifyAgent - Runs the agentcleanup command to verify the health of the agent
' ********************************************************************************************
Sub VERIFYAGENT

	' Firstly, check to see that we've got two services present
	Set objServices = objWMI.ExecQuery("SELECT State,Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
	If objServices.Count = 2 Then
		' OK, so the two services are installed.
		objShell.LogEvent 0, "The agent installer found both agent services installed on this device."

		For Each item In objServices
			If item.State <> "Running" Then
			
				' Report that one of the two services isn't running, and try to start it
				objShell.LogEvent 2, "The agent installer script found that the '" & item.Name & "' service is present but is in state '" & item.State & "' at the moment.  This is not necessarily a problem."
				objShell.LogEvent 0, "The agent installer script is trying to start the " & item.Name & " service."
				item.StartService()
				
				' Wait a bit of time for the service to start up
				count = 0
				Do While count < 100
					WRITETOCONSOLE("Starting " & item.Name & " ... " & SPIN & vbCR)
					WScript.Sleep 100
					count = count + 1
				Loop
								
				' Check that the service held in item.Name is running
				Set colService = objWMI.ExecQuery("SELECT Name,State FROM Win32_Service WHERE Name = '" & item.Name & "'")
				For Each service In colService
					'output.writeline service.name & service.State
					If service.state <> "Running" Then
						WRITETOCONSOLE("Starting " & item.Name & " ... failed!" & vbCRLF)
						objShell.LogEvent 2, "The '" & item.Name & "' service could not be started."
					Else
						WRITETOCONSOLE("Starting " & item.Name & " ... done!" & vbCRLF)
						objShell.LogEvent 0, "The '" & item.Name & "' service was started successfully."
					End If
				Next
				
				
			Else
				' The service is running
				objShell.LogEvent 0, "The '" & item.Name & "' service is present and started."
			End If
		Next
			
		' See if both agent services are installed and actually registered with SC (developed in conjunction with Willem Zeeman @ Mostware.nl)
		bolAgentServiceFound = False
		bolAgentMaintServiceFound = False

		WRITETOCONSOLE("Checking services ................ ")
		Set objExecObject = objShell.Exec("cmd /c sc query | findstr /i Agent")

		Do While Not objExecObject.StdOut.AtEndOfStream
			strLine = objExecObject.StdOut.ReadLine()
			If Instr(strLine, "Windows Agent Service")>0 Then bolAgentServiceFound = True End If
			If Instr(strLine, "Windows Agent Maintenance Service")>0 Then bolAgentMaintServiceFound = True End If	
		Loop
		
		If (bolAgentServiceFound = True And bolAgentMaintServiceFound = True) Then
			WRITETOCONSOLE("done!" & vbCRLF)
			objShell.LogEvent 0, "The agent installer script found that both agent services are properly registered on this device."
		Else
			' The services are not registered properly so we need to remove the agent
			WRITETOCONSOLE("failed!" & vbCRLF)
			objShell.LogEvent 1, "The agent installer script found a problem with the registration of the agent services on this device.  This agent is now considered to be corrupt and will be uninstalled."
			If bolInteractiveMode = True Then
					msgbox "The agent services appear to be corrupted on this device.  Therefore the agent will be uninstalled.", vbOKOnly + vbCritical, strBranding
			End If
			
			' Remove the agent
			STRIPAGENT
		End If
	
	Else
	
		' Only one service is listed in WMI (services.msc) so the agent is totally broken, so let's remove it
		STRIPAGENT
	End If


	
	' Check to see that we have .NET 4 installed, as we can only run a verify if it is
	If DOTNETPRESENT("v4\Full")="NOT_INSTALLED" Then
		objShell.LogEvent 1, "The agent installer script cannot perform a verify of the agent as Microsoft .NET Framework 4 is not installed on this device.  It has probably been uninstalled.  The script will now terminate."
		DIRTYQUIT 1
	End If
	

	' If the agent appears to be installed, let's check the appliance ID and see if it's a valid one
	'output.writeline "Agent path is " & strAgentPath
	If objFSO.FileExists(strAgentPath & "\config\ApplianceConfig.xml") Then
		' Check the appliance ID of the agent and write it to the event log
		WRITETOCONSOLE("Checking appliance ............... ")
		Set objFile = objFSO.OpenTextFile(strAgentPath & "\config\ApplianceConfig.xml")
		strApplianceConfig = ""
		Do Until objFile.AtEndOfStream
			strLine = objFile.Readline
			If Instr(strLine, "ApplianceID")>0 Then
				strApplianceConfig = strLine
			End If
		Loop
		objFile.Close
	
		' Strip the XML headers out of the line so we're left with just the number
		arrApplianceConfig = Split(strApplianceConfig, "<")
		strApplianceConfig = Right(arrApplianceConfig(1), Len(arrApplianceConfig(1))-12)
		If CLng(strApplianceConfig) < 1000 Then
			' definately a bad number
			objShell.LogEvent 2, "The agent installer script found that the installed agent has an invalid ApplianceID, " & strApplianceConfig & ".  This agent may not be able to check into " & strServerAddress & " correctly."
			WRITETOCONSOLE("failed!" & vbCRLF)
			'STRIPAGENT			' We don't want to strip the agent out as -1 is not ALWAYS a bad agent
		Else
			' Just write the appliance ID to the event log
			objShell.LogEvent 0, "The agent installer script found that the installed agent has an ApplianceID of " & strApplianceConfig & "."
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
	End If
		
	

	' Check to see if the agent is installed onto a non-C drive.  If it is, don't proceed as AgentCleanup4 doesn't support this and will
	' cause repeated uninstalls and installs of the agent
	If (Mid(strAgentBin, 2,1) <> "C" And strAgentBin <> "")Then
		WRITETOCONSOLE("Checking health .................. skipped" & vbCRLF)
		objShell.LogEvent 2, "The agent installer cannot perform a verify of the agent on this device as the agent is not installed onto the C drive. The script will now terminate."
		DIRTYQUIT 0
	End If
	
	' The agent is installed onto the C drive so we can do a proper verify
	objShell.LogEvent 0, "The agent installer script started a verify of the agent."
	' Launch the cleanup utility
	
	Set objCleanup = objShell.Exec("cmd /c AgentCleanup4.exe " & strRequiredAgent & " writetoeventlog")
		
	' Wait for the cleanup to finish
	count = 0
	Do While objCleanup.Status = 0
		count = count +1
	
		If count < 900 Then
			WRITETOCONSOLE("Checking health .................. " & SPIN & vbCR)
		End If
		
		If count > 900 Then
			WRITETOCONSOLE("Repairing ........................ " & SPIN & "         " & vbCR)
		End If
		
		' If the repair takes over 90 seconds then a fullscale remove is probably being done
		If count = 900 Then
			WRITETOCONSOLE("Checking health .................. failed!" & vbCRLF)
			' agentcleanup4 is doing a repair
			objShell.LogEvent 2, "The agent installer script has detected that the verify stage is probably performing a full cleanup of the agent. The agent is probably not communicating back to N-central or is outdated and is therefore being removed in preparation for a reinstallation."
		End If
	
		' The repair is taking too long, so bug out of this loop
		If count > 4200 Then
			objShell.LogEvent 1, "The agent installer script tried to repair the agent but the process exceed the permitted timeframe of six minutes."
			Exit Do
		End If
	
		WScript.Sleep 100
	
	Loop
	
	' Report on the various states that we could be in now the loop is exited
	If count < 90 Then
		WRITETOCONSOLE("Checking health .................. done!         " & vbCRLF)
	End If
	
	If count > 90 and count < 420 Then
		WRITETOCONSOLE("Repairing ........................ done!               " & vbCRLF)
	End If
	
	If count > 420 Then
		WRITETOCONSOLE("Repairing ........................ failed!             " & vbCRLF)
	End If
		
		
	' Did anything happen during the cleanup?
	If objCleanup.ExitCode <> 0 Then
		objShell.LogEvent 2, "The agent installer verify process found a problem with the agent and may have removed it."
	Else
		objShell.LogEvent 0, "The agent installer script has finished a cleanup/verify"
	End If
	
End Sub


' *************************************************************************************
' Sub Function - CopyAgentCleanup - Copies the AgentCleanup4.exe from network to device
' *************************************************************************************
Sub COPYAGENTCLEANUP

	WRITETOCONSOLE("Copying cleanup utility .......... ")
	
	' First we make sure that it's actually available to copy
	If objFSO.FileExists(strInstallSource & "AgentCleanup4.exe")=False Then
		WRITETOCONSOLE("failed!" & vbCRLF)
		objShell.LogEvent 1, "The agent installer script could not find the AgentCleanup4.exe utility at " & strInstallSource & "AgentCleanup4.exe and cannot proceed."
		DIRTYQUIT 3
	End If	
	
	' Delete local version if already present (ie: wrong version)
	'If objFSO.FileExists(strWindowsFolder & "\AgentCleanup4.exe")=True Then
	'	objFSO.DeleteFile(strWindowsFolder & "\AgentCleanup4.exe")
	'End If
	
	' Now copy it from the network to the local device and check that it succeeded
	objFSO.CopyFile strInstallSource & "Agentcleanup4.exe", strWindowsFolder & "\AgentCleanup4.exe"
	
	If objFSO.FileExists(strWindowsFolder & "\agentcleanup4.exe") = False Then
		WRITETOCONSOLE(" failed!" & vbCRLF)
		strMessage = "The agent installer script could not not copy AgentCleanup4.exe from " & strInstallSource & " into " & strWindowsFolder
		If bolInteractiveMode = True Then
			msgbox strMessage & vbCRLF & vbCRLF & strContactAdmin, 16, strBranding
		End If
		objShell.LogEvent 1, strMessage
		DIRTYQUIT 3
	End If
	WRITETOCONSOLE(" done!" & vbCRLF)
End Sub



' *****************************************************************************************
' Sub Function - WriteToConsole - Writes a string to the console if interactive mode is off
' *****************************************************************************************
Sub WRITETOCONSOLE(strMessage)
	If bolInteractiveMode = False Then
		WScript.StdOut.Write strMessage
	End If
End Sub


' ***********************************************************************
' Function - IsDowngrade- Compares current and proposed agent versions
'	and returns whether the proposed version would downgrade the current
'	version installed
' ***********************************************************************
Function IsDowngrade(strCurrent, strProposed)
	Dim bValid
	Dim currMajor, currMinor, currSP, currBuild
	Dim propMajor, propMinor, propSP, propBuild
	Dim strSplit
	
	strSplit = Split(strCurrent, ".")
	currMajor = CInt(strSplit(0))
	currMinor = CInt(strSplit(1))
	currSP = CInt(strSplit(2))
	currBuild = CInt(strSplit(3))
	
	strSplit = Split(strProposed, ".")
	propMajor = CInt(strSplit(0))
	propMinor = CInt(strSplit(1))
	propSP = CInt(strSplit(2))
	propBuild = CInt(strSplit(3))
	
	bValid = True

	If propMajor < currMajor Then
		bValid = False
	Else
		If propMajor = currMajor Then
			If propMinor < currMinor Then
				bValid = False
			Else
				If propMinor = currMinor Then
					If propSP < currSP Then
						bValid = False
					Else
						If propSP = currSP Then
							If propBuild < currBuild Then
								bValid = False
							End If
						End If
					End If
				End If
			End If
		End If
	End If
	
	IsDowngrade = Not bValid
End Function


' *****************************************************************************
' Function - WriteRegValue - Writes a value into the Registry with minimal fuss
' *****************************************************************************
Function WRITEREGVALUE(strKey, strType, strValue)
	WRITETOCONSOLE("Writing " & strKey & " " & strType & " " & strValue)
	If Instr(strValue, " ")>0 Then strValue = Chr(34) & strValue & Chr(34) End If
	Set objCmd = objShell.Exec("cmd /c reg add " & Chr(34) & "HKLM\Software\Tim Wiser\InstallAgent" & Chr(34) & " /v " & strKey & " /t " & strType & " /d " & strValue & " /f")
End Function


' *********************************************************************************************************************
' Function - CleanQuit - Exits the script cleanly with error code 10, which can be picked up by the launcher batch file
' *********************************************************************************************************************
Function CLEANQUIT
	objShell.LogEvent 0, "The agent installer script v" & strVersion & " has finished running."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastSuccessfulRun", "" & Now()
	objReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastOperation", 10
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastExitComment", "Successful execution"
	WScript.Quit 10
End Function

' ************************************************************************
' Function - DirtyQuit - Exits the script with an error code and Reg value
' ************************************************************************
Function DIRTYQUIT(intValue)
	objReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastOperation", intValue
	Select Case intValue
		Case 0	: strExitComment = "Internal error"
		Case 1	: strExitComment = ".NET 4 ot installed / Win 2K detected / Win XP not at SP3"
		Case 2	: strExitComment = "No external network access (cannot ping " & strPingAddress & ")"
		Case 3	: strExitComment = "Install source is missing or wrong version"
		Case 4	: strExitComment = "Failed to install .NET 4"
		Case 5	: strExitComment = "Failed to install agent"
		Case 6	: strExitComment = "Zombie agent detected!"
		Case 10	: strExitComment = "Successful completion"
		Case Else	: strExitComment = "Undefined exit code"
	End Select
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Tim Wiser\InstallAgent\", "LastExitComment", strExitComment
	WRITETOCONSOLE(vbCRLF & "An error occurred. " & strExitComment & vbCRLF & strContactAdmin & vbCRLF)
	objShell.LogEvent 1, "The agent installer script experienced a problem and exited prematurely.  The exit code was " & intValue & " (" & strExitComment & ")"
	WScript.Quit intValue
End Function


' ############## Additional padding to make file size different
