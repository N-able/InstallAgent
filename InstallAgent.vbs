' INSTALLAGENT.VBS

' Installation and maintenance script for the N-able agent software
' Created by Tim Wiser, GCI Managed IT
' Maintained by N-Able partners


' 4.00 02/11/2015
'				Formatting changes and more friendly startup message
'				Dirty exit now shows error message and contact information on console
'				Added 'Checking files' bit to remove confusing delay at that stage. No spinner though, unfortunately
'				This is the final release by Tim :o(
'				First version committed to git - Jon Czerwinski
'
' 4.01 20151109
'				Corrected agent version zombie check
'
' 4.10 20151115
'				Refactored code - moved mainline code to subroutines, replaced literals with CONSTs
'				Aligned XP < SP3 exit code with documentation (was 3, should be 1)
'				Added localhost zombie checking
'				Changed registry location to HKLM:Software\N-Central
'
'				NOTE ON REFACTORING - Jon Czerwinski
' 				The intent of the refactoring is:
'
' 				1. Shorten and simplify the mainline of code by moving larger sections of mainline code to subroutines
' 				2. Replace areas where the code quit from subroutines and functions with updates to runState variable
'    			and flow control in the mainline.  The script will quit the mainline with its final runState.
' 				3. Remove the duplication of code
' 				4. Remove inaccessible code
'
' 				This code relies heavily on side-effects.  These have been documented at the top of each
' 				function or subroutine.
'
' 4.20 20170119
'				Moved partner-configured parameters out to AgentInstall.ini
'				Removed Windows 2000 checks
'				Cleaned up agent checks to eliminate redundant calls to StripAgent
'				Remove STARTUP|SHUTDOWN mode
'
' 4.21 20170126
'				Error checking for missing or empty configuration file.
'
' 4.22 20170621
'				Close case where service is registered but executable is missing.
'

Option Explicit

' Script Version
CONST strVersion = "4.22"		' The current version of this script

' Declare and define our variables and objects
Dim objFSO, output, objReg, objArguments, objArgs, objNetwork, objWMI, objShell, objEnv, objAppsList
Dim objCleanup, objServices, objInstallerProcess, objAgentInstaller, strAgentInstallerShortName, objMarkerFile
Dim objQuery, objDotNetInstaller, objFile, objExecObject, objWIC, colService, objCmd
Dim strComputer, strAgentPath, strAgentBin, strCheckForFilme, strMSIflags, strNETversion, strInstallFlags
Dim strSiteID, strDomain, strInstallSource, strMessage, strWindowsFolder, strInstallCommand, strLine
Dim strAgentConfig, strApplianceConfig, arrApplianceConfig, strServerConfig
Dim strKey, strType, strValue, strExitComment
Dim strOperatingSystem, strAltInstallSource, strTemp, strProxyString
Dim strBaselineListOfPIDs, strInstallationListOfPIDs, strInstallProxyString, strArchitecture, strDotNetInstallerFile
Dim strOperatingSystemSKU, strNoNetwork, strPSValue, strWICinstaller, strString, strDecimal, strSpin
Dim bolInteractiveMode, bolAgentServiceFound, bolAgentMaintServiceFound
Dim intServicePackLevel, intPSValue, intValue
Dim item, response, count, service
Dim	runState

CONST HKEY_LOCAL_MACHINE = &H80000002
CONST evtSuccess = 0
CONST evtError = 1 
CONST evtWarning = 2 
CONST evtInfo = 4 
CONST strDotNetRegKey = "V4\Full"
CONST strDummyDom = "UNKNOWN.TIM"
CONST strRegBase = "N-Central"	' Change this to "Tim Wiser" to use the traditional registry location for monitoring
								' Replace "InstallAgentStatus.amp" with "InstallAgentStatus - Tim Wiser.amp"

' Exit code definitions
CONST errInternal = 0			' Internal error
CONST errPrereq = 1 			' .NET4 not installed / Windows 2000 detected / Windows XP not at SP3
CONST errNetwork = 2 			' No external network access detected
CONST errSource = 3 			' Install source is missing / wrong version
CONST errNET4 = 4 				' Failure to install .NET4
CONST errAgent = 5 				' Failure to install agent
CONST errZombie = 6 			' Installed agent higher version than source
CONST errNormal = 10 			' Normal execution

strComputer = "."
strAgentPath = ""
strDomain = strDummyDom			' This is a dummy domain name, do NOT change
strSiteID = ""
strSpin = "|"
bolInteractiveMode = False
runState = errNormal			' Tracks the error condition in the script and determines if we should continue

' ***********************************************************************************************************************************************************
' Define some constants for your environment
CONST bolWarningsOn = False					' If the agent fails, the user gets a popup message.  Change this to False if you don't want them to be informed

CONST strPingAddress = "8.8.8.8"				' Address to ping to check connectivity.  Usually this will be your N-central server's FQDN
CONST strSOAgentEXE = "SOAgentSetup.exe"			' The name of the SO-level agent installer file that is stored within NETLOGON\strAgentFolder

CONST intPingTolerance = 20							' % value.  If you have networks that drop packets under normal conditions, raise this value
CONST intPingTestCount = 10							' Increase this to make the script perform more pings during the Connectivity test
' ***********************************************************************************************************************************************************


' ***********************************************************************************************************************************************************
' Define INI File Keys
CONST strINIFile = "InstallAgent.ini"
CONST strSection = "General"
CONST strDefaultValue = "-"
CONST strZeroBytes = "0"
CONST strNull = ""
CONST strKeyContactAdmin = "ContactAdminMsg"
CONST strKeyServerAddress = "ServerAddress"
CONST strKeyBranding = "Branding"
CONST strKeyAgentVersion = "SOAgentVersion"
CONST strKeyAgentFileVersion = "SOAgentFileVersion"
CONST strKeyAgentFolder = "AgentFolder"
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

strWindowsFolder = objShell.ExpandEnvironmentStrings("%WINDIR%")


' Get contents of the INI file As a string
DIM strScriptPath : strScriptPath = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
DIM INIContents : INIContents = GetFile(strScriptPath & strINIFile)

' If the configuration file is empty or missing, exit with an error
If Len(INIContents) = 0 Then
	objShell.LogEvent evtError, "The agent installer script could not load the configuration file " & strScriptPath & strINIFile & "."	
	DIRTYQUIT errInternal
End If

' ***********************************************************************************************************************************************************
' Load program values from INI file
DIM strServerAddress: strServerAddress = GetINIString(INIContents, strSection, strKeyServerAddress, strDefaultValue)
DIM strBranding: strBranding = GetINIString(INIContents, strSection, strKeyBranding, strDefaultValue)
DIM strContactAdmin: strContactAdmin = GetINIString(INIContents, strSection, strKeyContactAdmin, strDefaultValue)
DIM strAgentFolder : strAgentFolder = GetINIString(INIContents, strSection, strKeyAgentFolder, strDefaultValue)
DIM strRequiredAgent: strRequiredAgent = GetINIString(INIContents, strSection, strKeyAgentVersion, strDefaultValue)
DIM strAgentFileVersion: strAgentFileVersion = GetINIString(INIContents, strSection, strKeyAgentFileVersion, strDefaultValue)
' ***********************************************************************************************************************************************************

' Get the mandatory parameters
strSiteID = objArgs("site")
strNoNetwork = UCase(objArgs("nonetwork"))
strAltInstallSource	= objArgs("source")

' WScript.StdOut.Write "Server Address: " & strServerAddress & vbCRLF & _
'	"Branding: " & strBranding & vbCRLF & _
'	"Contact Admin: " & strContactAdmin & vbCRLF & _
'	"Agent Folder: " & strAgentFolder & vbCRLF & _
'	"Required Version: " & strRequiredAgent & vbCRLF & _
'	"File Version: " & strAgentFileVersion & vbCrlf

' Make sure that the Registry key that we write to actually exists
objReg.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent"

' Work out which scripting engine we're running under as we behave differently according to which one we're using
If Instr(1, WScript.FullName, "CScript", vbTextCompare) = 0 Then
	bolInteractiveMode = True
	objShell.LogEvent evtSuccess, "The agent installer script is running in interactive mode."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastMode", "Interactive"
	' Ensure that the script runs in elevated mode
	If WScript.Arguments.Named.Exists("elevated") = False Then
		objShell.LogEvent evtWarning, "The agent installer script is elevating its execution"
		CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
		WScript.Quit
	End If
Else
	bolInteractiveMode = False
	objShell.LogEvent evtSuccess, "The agent installer script is running in non-interactive (scripted) mode."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastMode", "Unattended"
	' Support the /? parameter on the command line to get a little bit of help
	If objArguments.Count > 0 Then
		If objArguments(0) = "/?" Then
			WScript.Stdout.Write "InstallAgent.vbs version " & strVersion & vbCRLF & "by Tim Wiser, GCI Managed IT (tim.wiser@gcicom.net)" & vbCRLF & vbCRLF
			WScript.Stdout.Write "Syntax: InstallAgent.vbs /site:[ID] /mode:[STARTUP|SHUTDOWN] (/source:[ALTERNATIVE SOURCE PATH]) (/nonetwork:yes)" & vbCRLF & vbCRLF
			CleanQuit
		End If
	End IF
End If

' Say we've started execution
objShell.LogEvent evtSuccess, "The agent installer script being executed is version v" & strVersion
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastRun", "" & Now()
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "Version", strVersion
objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastExitComment", "Script started successfully"
WRITETOCONSOLE("Please wait whilst the " & strBranding & " agent is checked." & vbCRLF & vbCRLF)

' Validate the mandatory parameter if we're running in non-interactive mode
If bolInteractiveMode = False Then	
	If strSiteID = "" Then
		strMessage = "The agent installer script was passed an insufficient number of parameters.  /site:[ID] is required for this script to function. The agent installation cannot continue on this device.  Please check the GPO configuration."
		objShell.LogEvent evtError, strMessage
		runState = errInternal 
	End If
End If
If runState <> errNormal Then DIRTYQUIT runState End If	' Placeholder until chunks are moved to subroutines

' Get Operating System information
GetOSInfo

' Verify Operation System prerequisites.
' Cannot install on Windows 2000 due to .NET requirements
' Windows XP must be at SP3
VerifyOSPrereq
If runState <> errNormal Then DIRTYQUIT runState End If	' Placeholder until chunks are moved to subroutines

' Configure Install Source
ConfigureSource
If runState <> errNormal Then DIRTYQUIT runState End If	' Placeholder until chunks are moved to subroutines

' Perform a connectivity test
TestNetwork
If runState <> errNormal Then DIRTYQUIT runState End If	' Placeholder until chunks are moved to subroutines

' Set strProxyString according to presence of proxy.cfg file
ConfigureProxy

' Copy agentcleanup4.exe locally so we have it available later on.  NOTE! This version will only run if .NET 4 is installed
CopyAgentCleanup

' Pop up a message asking if the device should be excluded from agent deployment in the future if running interactively
If bolInteractiveMode = True Then
	' Check to see if the nable_disable.mrk file is present and, if it is, ask if it should be deleted
	If objFSO.FileExists(strWindowsFolder & "\nable_disable.mrk") = True Then
		' The agent is currently prevented from deploying on this device, so ask if this should remain the case
		response = MsgBox("The agent is currently prevented from deploying onto this device.  Do you want to enable deployment in the future?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Agent istaller")
		Select Case response
		Case vbYes		:		objFSO.DeleteFile(strWindowsFolder & "\nable_disable.mrk")
		Case vbCancel	:		MsgBox "Agent installation has been cancelled.", vbExclamation, strBranding
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
	objShell.LogEvent evtWarning, "nable_disable.mrk is present inside the " & strWindowsFolder & " folder, so this device should not have the agent installed.  Exiting now."
	CleanQuit
End If	

' Get the path to the agent service executable.  strAgentBin will have this path or "" if not found.
'	strAgentPath will be strAgentBin with quotes stripped off the path or "" if agent is not found.
GetAgentPath

' Detect and warn if PowerShell 2 is not installed
CheckPowerShell

' If the script is running in interactive mode, check to see if the agent is installed and bug out if it is
If (bolInteractiveMode = True And strAgentBin <> "") Then
	objShell.LogEvent evtSuccess, "The agent installer found that the agent is already installed on this device at " & strAgentBin
	response = MsgBox("The agent is installed on this device at " & strAgentBin & vbCRLF & vbCRLF & "Do you want to uninstall it?  This can take up to five minutes to complete.", vbYesNo+vbQuestion, strBranding)
	If response = vbNo Then
		CleanQuit
	Else
		If DOTNETPRESENT(strDotNetRegKey)="NOT_INSTALLED" Then
			' The AgentCleanup4 utility requires .NET 4 so if it's not installed, let's just try a clean removal using Windows Installer
			objShell.LogEvent evtWarning, "The agent installer script initiated a clean removal of the existing agent using MSIEXEC on user request."
			Call objShell.Run("msiexec /X {07BA9781-16A5-4066-A0DF-5DBA3484FDB2} /passive /norestart",,True)
		Else
			' Do an uninstall using AgentCleanup4.exe as we have .NET 4 installed on the device
			objShell.LogEvent evtWarning, "The agent installer script initiated a cleanup of the existing agent on user request."
			Call objShell.Run("cmd /c agentcleanup4.exe writetoeventlog",,True)
		End If
	End If
	CleanQuit
End If

' Process agent
ProcessAgent

CleanQuit
' *****************************************************************************
' THIS IS THE END OF THE MAIN SCRIPT BODY
' *****************************************************************************


' *****************************************************************************
' Sub - ConfigureSource - Sets install source and tests for presence and correct
'		version of agent install file.
'
'
' Side Effects - updates
'	runState
'	strDomain
'	strInstallSource
' 
' *****************************************************************************
Sub ConfigureSource
	Dim objDomainInfo, strInstallAgentVersion

	' Get the domain that this device is a member of
	Set objDomainInfo = objWMI.ExecQuery("SELECT Domain FROM Win32_ComputerSystem")
	For Each item In objDomainInfo
		strDomain = item.Domain
	Next
				
	If strDomain = strDummyDom Then
		objShell.LogEvent evtWarning, "The agent installer script was unable to determine what the local domain is for this device.  If the script was run on a home or workgroup device, please run the standard agent installer executable instead."
		runState = errInternal
	Else
		strInstallSource = "\\" & strDomain & "\Netlogon\" & strAgentFolder & "\"

		' Lets you override the automatic assumption that you are installing from NETLOGON by specifying an alternative path in /source
		If strAltInstallSource <> "" Then
			' Ensure that the install source path has a slash at the end of it, otherwise it won't form a valid path
			If Right(strAltInstallSource, 1) <> "\" Then
				strAltInstallSource = strAltInstallSource & "\"
			End If
			strInstallSource = strAltInstallSource
		End If

		' Write the strInstallSource path into the Registry for the custom service to read
		objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "Path", strInstallSource

		WRITETOCONSOLE("Checking files ................... ")
		' Validate that the agent installer exists
		If (objFSO.FileExists(strInstallSource & strSOAgentEXE)) = True Then	
			' Get the file version of the available installer file and check that it's the right version
			strInstallAgentVersion = objFSO.GetFileVersion(strInstallSource & strSOAgentEXE)
			If strInstallAgentVersion <> strAgentFileVersion Then
				strMessage = "The agent installer script has found that the agent installer at '" & strInstallSource & strSOAgentEXE & "' is the wrong version (version " & strInstallAgentVersion & ").  The expected version is " & strAgentFileVersion & ".  The installation will not continue."
				WRITETOCONSOLE("failed!" & vbCRLF)
				objShell.LogEvent evtError, strMessage
				If bolInteractiveMode = True Then
						MsgBox strMessage & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
				End If
				runState = errSource
			Else
				WRITETOCONSOLE("done!" & vbCRLF)
				objShell.LogEvent evtSuccess, "The agent installer script has validated that the available installer file (version " & strInstallAgentVersion & ") is the required version, " & strAgentFileVersion
			End If
		Else
			strMessage = "The agent cannot be installed on this computer.  The install source " & strInstallSource & strSOAgentEXE & " is missing or invalid.  Additional information can be found in the Application event log." & vbCRLF & vbCRLF & strContactAdmin
			objShell.LogEvent evtError, strMessage
			If (bolWarningsOn = True or bolInteractiveMode = True) Then
				Msgbox strMessage, vbOKOnly + vbCritical, strBranding
			End If
			runState = errSource		
		End If
	End If
End Sub ' ConfigureSource


' *****************************************************************************
' Sub - TestNetwork - Ping test to internet
'
' Side Effects - updates
'	runState
' 
' *****************************************************************************
Sub TestNetwork
	Dim count, item, intPingFailurePct
	Dim intPingFailureCount : intPingFailureCount = 0
	
	If strNoNetwork <> "YES" Then
		For count = 1 to intPingTestCount
			Set objQuery = objWMI.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address='" & strPingAddress & "'")
			WRITETOCONSOLE("Checking connectivity ............ " & SPIN & vbCR)
			For Each item in objQuery
				If item.StatusCode <> 0 Then
					' ping dropped, accumulate failure
					intPingFailureCount = intPingFailureCount + 1
				End If
			Next
		Next
		
		intPingFailurePct = (intPingFailureCount / intPingTestCount) * 100
		
		If intPingFailurePct > intPingTolerance Then
			' Terminate if the ping test failed the tolerance test
			WRITETOCONSOLE("Checking connectivity ............ failed!" & vbCRLF)
			objShell.LogEvent evtError, "The agent installer script has detected that this device does not have access to the central server.  The ping check to " & strPingAddress & " failed with a fault rate of " & (intPingFailureCount / intPingTestCount) * 100 & "%.  Please check connectivity and try again.  If network conditions are generally poor you can adjust the tolerance.  Refer to the script documentation for further assistance.  The script will now terminate."
			If bolInteractiveMode = True Then
				MsgBox "This device cannot ping the external test address of " & strPingAddress & ".  Please check connectivity and try again.", vbOKOnly + vbCritical, strBranding
			End If
			runState = errNetwork
		ElseIf intPingFailurePct > 0 Then
			' Warn if packets were dropped	
			WRITETOCONSOLE("Checking connectivity ............ done!" & vbCRLF)
			objShell.LogEvent evtWarning, "The agent installer script has detected that this device has connectivity to " & strPingAddress & " but is dropping packets.  The ping test had a fault rate of " & (intPingFailureCount / intPingTestCount) * 100 & "%.  If the fault rate passes " & intPingTolerance & "% the script will not run."
		Else	
			' All was well, we got 100% ping connectivity
			WRITETOCONSOLE("Checking connectivity ............ done!" & vbCRLF)
			objShell.LogEvent evtSuccess, "The agent installer script has detected that this device has reliable connectivity to " & strPingAddress & "."
		End If
	End If
End Sub ' TestNetwork


' *****************************************************************************
' Sub - ConfigureProxy - Sets proxy string if proxy.cfg file exists
'
' Side Effects - updates
'	strProxyString
' 
' *****************************************************************************
Sub ConfigureProxy
	Dim strProxyConfigFile, objFile
	
	' See if we're wanting to use a proxy server when installing the agent.  If so, read the proxy configuration file
	strProxyConfigFile = "\\" & strDomain & "\Netlogon\" & strAgentFolder & "\proxy.cfg"
	
	If objFSO.FileExists(strProxyConfigFile)= True Then
		objShell.LogEvent evtSuccess, "The agent installer script has found a proxy.cfg file on the network"
		WRITETOCONSOLE("Configuring proxy ................ ")
		Set objFile = objFSO.OpenTextFile(strProxyConfigFile)
			strProxyString = objFile.ReadLine
		objFile.Close
		objShell.LogEvent evtSuccess, "The agent installer script will install the agent using: " & strProxyString
		WRITETOCONSOLE("done!" & vbCRLF)
	Else
		' No proxy.cfg found, so log for troubleshooting purposes
		strProxyString = ""
		objShell.LogEvent evtInfo, "The agent installer script did not find a proxy configuration file to use"
	End If
End Sub ' ConfigureProxy


' *****************************************************************************
' Sub - CheckPowerShell - Checks the installed version of Powershell, and warns
'		if missing or below Version 2
'
' Side Effects - none
' 
' *****************************************************************************
Sub CheckPowerShell
	Dim strPSValue, intPSValue

	objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\", "PowerShellVersion", strPSValue
	If IsNull(strPSValue) Then
		' Warn that no version of Powershell at all is installed on the device
		objShell.LogEvent evtError, "The agent installer script cannot find Microsoft PowerShell installed on this device.  For use of advanced automation within N-central a device must have PowerShell v2 or greater installed."
	Else
		intPSValue = CInt(Left(strPSValue,1))
		If intPSValue < 2 Then
			' PSH 1 isn't good enough for N-able Automation Manager
			objShell.LogEvent evtError, "The agent installer script found an outdated version of PowerShell installed on this device. For use of advanced automation within N-central a device must have PowerShell v2 or greater installed."
		Else
			' PSH 2 onwards is suitable for N-able Automation Manager
			objShell.LogEvent evtSuccess, "The agent installer script found PowerShell v" & intPSValue & " installed on this device.  This version is suitable for use with N-able Automation Manager policies (.AMP files)"
		End If
	End If
End Sub ' CheckPowerShell


' *****************************************************************************
' Sub - GetOSInfo - Get the Operating System, Service Pack Level, and the
'		SKU.  For Server 2008 or 2012, this returns the OperatingSystemSKU
'		value - used to check for Server Core.  For anything else, it returns 0
'
' Side Effects - updates
'	intServicePackLevel
'	strOperatingSystem
'	strOperatingSystemSKU
' 
' *****************************************************************************
Sub GetOSInfo
	' Detect the OS and type
	Set objQuery = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem WHERE Caption LIKE '%Windows%'")
	For Each item In objQuery
		strOperatingSystem = item.Caption
		intServicePackLevel = item.ServicePackMajorVersion
		If (Instr(strOperatingSystem, "2008")>0 Or Instr(strOperatingSystem, "2012")>0) Then
			strOperatingSystemSKU = item.OperatingSystemSKU
		Else
			strOperatingSystemSKU = 0
		End If
	Next
End Sub 'GetOSInfo


' *****************************************************************************
' Sub - VerifyOSPrereq - 
'
' Side Effects - updates
'	runState
' 
' *****************************************************************************
Sub VerifyOSPrereq
	WRITETOCONSOLE(strOperatingSystem & " detected" & vbCRLF)
		
	' Warn if the device is running Windows XP and isn't patched to SP3 level
	If Instr(strOperatingSystem, "XP")>0 Then
		If intServicePackLevel < 3 Then
			objShell.LogEvent evtError, "The agent installer script has detected that this device is running Windows XP but is not patched with Service Pack 3.  This will prevent .NET 4 from installing which will in turn prevent the agent being able to be deployed or maintained.  Please install Service Pack 3 on this device.  The script will now terminate."
			WRITETOCONSOLE("Checking Windows XP SP3 .......... failed!" & vbCRLF & vbCRLF & "Please install Service Pack 3 on this computer." & vbCRLF)
			runState = errPrereq
		End If
	End If
End Sub ' VerifyOSPrereq


' *****************************************************************************
' Sub - ProcessAgent - Handles verifying an installed agent or installing the
'	agent if it's not installed
'
' Side Effects - updates
'	runState
' 
' *****************************************************************************
Sub ProcessAgent
	Dim strInstalledVersion

	' Decide whether we need to install a new agent or verify an existing agent
	If strAgentBin = "" Then	' Agent is not installed, so install it
		objShell.LogEvent evtInfo, "The agent installer script did not find the agent on this device so will attempt to install it."

		' Pop up a prompt for interactive users to enter the site ID code manually
		If bolInteractiveMode = True Then
			strSiteID = InputBox("Welcome to the agent installer script. © Tim Wiser, GCI Managed IT" & vbCRLF & vbCRLF & "Please enter the numerical site code which is found by navigating to the SO level and selecting Administration -> Customers/Sites." & vbCRLF & vbCRLF & "NOTE: If you do not understand this message, please click the Cancel button and contact your IT administrator.", strBranding)
			If strSiteID = "" Then
				objShell.LogEvent evtError, "The agent installer script aborted by user request."
				MsgBox "Agent installation has been aborted.", vbExclamation, strBranding
				CleanQuit
			End If
		End If

		' Check for .NET Framework v4
		If DOTNETPRESENT(strDotNetRegKey) = "NOT_INSTALLED" Then
		
			' Decide which installer to use (deprecated Nov 2012, we always install the same one now)
			strDotNetInstallerFile = ""
			strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
			strArchitecture = objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
			Select Case strArchitecture
				Case "AMD64"	: strDotNetInstallerFile = "dotNetFx40_Full_x86_x64.exe"
				Case "x86"		: strDotNetInstallerFile = "dotNetFx40_Full_x86_x64.exe"		' Originally dotNetFx40_Full_x86.exe, but the 64 bit includes 32 bit support as well
			End Select
			
			' Firstly, check to see if we're running on a Core installation of Windows.  We cannot install the .NET Framework onto Server Core automatically 
			Select Case strOperatingSystemSKU
				Case 12, 39, 14, 41, 13, 40, 29	:	WRITETOCONSOLE(vbCRLF & "Server Core detected!" & vbCRLF & "Please refer to the Event Log for futher details" & vbCRLF)
													objShell.LogEvent evtError, "The agent installer script detected that this device is running Windows Server Core and that Microsoft .NET Framework 4 is not currently installed.  This needs to be installed before the agent can be installed but the script is unable to install it automatically." & vbCRLF & vbCRLF & "Please install the .NET Framework 4 for Server Core from http://www.microsoft.com/en-us/download/details.aspx?id=22833 or from within N-central and run the script again."
													strDotNetInstallerFile = "dotNetFx40_Full_x86_x64_SC.exe"
													DIRTYQUIT errPrereq
			End Select
			
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
					objShell.LogEvent evtError, "The agent installer script has detected that Windows Imaging Components is not installed on this device and will attempt to install it from " & strInstallSource & strWICinstaller
					Set objWIC = objShell.Exec("cmd /c " & strInstallSource & strWICinstaller & " /quiet")
					Do While objWIC.Status = 0
						WRITETOCONSOLE("Installing Windows Imaging Components " & strArchitecture & " ... " & SPIN & vbCR)
						WScript.Sleep 100
					Loop
					WRITETOCONSOLE("Installing Windows Imaging Components " & strArchitecture & " ... done!" & vbCR)
				Else
					' We can't install WIC as the required file isn't present in strInstallSource
					objShell.LogEvent evtWarning, "The agent installer script has detected that Windows Imaging Components may not be installed on this device.  For this reason, Microsoft .NET Framework 4 may fail to install.  You can download the 32-bit WIC installer from http://www.microsoft.com/en-us/download/details.aspx?id=32.  The script is capable of installing WIC automatically as long as the WIC_x86_enu.exe and WIC_x64_enu.exe files are present inside " & strInstallSource
				End If
			End If
			
			
			objShell.LogEvent evtWarning, "The agent installer script detected that Microsoft .NET Framework 4 Full Package is not currently installed.  This needs to be installed before the agent can be installed.  The script will now attempt to install Microsoft .NET Framework 4 Full Package on this device."
			
			objEnv("SEE_MASK_NOZONECHECKS") = 1
			
			WRITETOCONSOLE("Checking for .NET v4 installer ... ")
			If (objFSO.FileExists(strInstallSource & strDotNetInstallerFile) = False) Then
				' The .NET installer file that we need to run doesn't exist, so error out
				strMessage = "The agent installer could not install Microsoft .NET Framework 4 as the installer file, " & strDotNetInstallerFile & ", does not exist at " & strInstallSource & strDotNetInstallerFile & ".  Please download the installer from N-central and try again, or install Microsoft .NET Framework 4 manually on this device.  The script will now terminate."
				If strDotNetInstallerFile = "dotNetFx40_Full_x86_x64_SC.exe" Then strMessage = strMessage & vbCRLF & vbCRLF & "NOTE:  As this server is running a Core edition of Windows you will need to download the '.NET Framework 4 Server Core - x64' installer from within N-central and store it in the " & strInstallSource & " folder.  Once this is done, .NET should install automatically the next time this script runs."
				If bolInteractiveMode = True Then
					MsgBox strMessage, vbCritical + vbOKOnly, strBranding
				End If
				objShell.LogEvent evtError, strMessage
				WRITETOCONSOLE("failed!" & vbCRLF)
				DIRTYQUIT errPrereq
			Else
				WRITETOCONSOLE("done!" & vbCRLF)
			End If
				
			' Copy the .NET installer file locally before installation
			WRITETOCONSOLE("Copying .NET v4 installer ........ ")
			objFSO.CopyFile strInstallSource & strDotNetInstallerFile, strTemp & "\dotnetfx.exe"
			If (objFSO.FileExists(strTemp & "\dotnetfx.exe"))=True Then 
				objShell.LogEvent evtSuccess, "The agent installer script successfully copied the Microsoft .NET 4 Full Package installer for " & strArchitecture & ", " & strDotNetInstallerFile & " into " & strTemp & " as dotnetfx.exe"
				WRITETOCONSOLE("done!" & vbCRLF)
			Else
				objShell.LogEvent evtError, "The agent installer script could not copy " & strDotNetInstallerFile & " into " & strTemp & " as dotnetfx.exe so cannot continue with the installation.  Check file permissions and that the " & strDotNetInstallerFile & " file exists within " & strInstallSource
				WRITETOCONSOLE("failed!" & vbCRLF)
				DIRTYQUIT errPrereq
			End If
			
			' Now start the installer up from the local copy
			objShell.LogEvent evtSuccess, "The agent installer script started the installation of Microsoft .NET Framework 4 Full Package on this device."
			Set objDotNetInstaller = objShell.Exec(strTemp & "\dotnetfx.exe /passive /norestart /l c:\dotnetsetup.htm" & Chr(34))
			Do Until objDotNetInstaller.Status <> 0
				WRITETOCONSOLE("Installing .NET v4 Full Package .. " & SPIN & vbCR)
				WScript.Sleep 100
			Loop
			
			objEnv.Remove("SEE_MASK_NOZONECHECKS")
			' We usually terminate here, as .NET seems to kill off any script that calls it for some reason.  The code below is included in case we're running
			' in interactive mode or to support rare cases where the .NET installer allows the script to continue running.
			If DOTNETPRESENT(strDotNetRegKey) = "NOT_INSTALLED" Then
				' .NET failed to install, so pop a message up and log an event, then exit
				WRITETOCONSOLE("Installing .NET v4 Full Package ... failed!" & vbCRLF)
				strMessage = "The agent installer script failed to install Microsoft .NET Framework 4 Full Package on this device." 
				objShell.LogEvent evtError, strMessage & vbCRLF & vbCRLF & "If this computer is running Windows XP then check that it is running SP3 and has Windows Installer 3.1 installed.  Please check other supported operating systems and pre-requisities with Microsoft at http://www.microsoft.com/en-gb/download/details.aspx?id=17718 and try again."
				Msgbox strMessage & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
				DIRTYQUIT errNET4
			Else
				' .NET 4 Full Package installed successfully
				WRITETOCONSOLE("Installing .NET v4 Full Package .. done!" & vbCRLF)
				objShell.LogEvent evtSuccess, "The agent installer script successfully installed Microsoft .NET Framework 4 Full Package on this device."
			End If
		Else
			objShell.LogEvent evtSuccess, "The agent installer script found that the Microsoft .NET Framework Full Package is installed on this device."
		End If

			
		' Branch off for an install.  We no longer do a verify of the agent afterwards -  we assume that the agent is installed and do checks during the Function
		INSTALLAGENT strSiteID, strDomain
		
	Else	' Agent is installed, do a verify
		strMessage = "The agent installer script has determined that the agent is already installed on this device.  The binary was found at " & strAgentBin
		If bolInteractiveMode = True Then
			MsgBox strMessage, vbInformation, strBranding
		End If
		objShell.LogEvent evtSuccess, strMessage
		
		' Check to see if this script could potentially downgrade the agent, which N-central doesn't like AT ALL :-(
		WRITETOCONSOLE("Checking downgrade ............... ")
		strInstalledVersion = objFSO.GetFileVersion(strAgentPath)
		If IsDowngrade(strInstalledVersion, strRequiredAgent) Then
			objShell.LogEvent evtError, "The agent installer script found that this device already has agent version " & strInstalledVersion & " installed which is newer than the version available for installing, which is " & strRequiredAgent & ".  If maintenance was carried out on this device it could potentially effectively downgrade the agent which would result in a zombie device.  Therefore, this script will now terminate."
			WRITETOCONSOLE("failed!" & vbCRLF)
			DIRTYQUIT errZombie
		Else
			objShell.LogEvent evtSuccess, "The agent installer found that the installed agent is suitable for maintenance."
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
		
		' Branch off for a verify
		VERIFYAGENT
		
		' Check to see if the agent services exist.  If not, the agent verify stage has caused a removal of the (old or corrupted) agent, so let's reinstall it now
		Set objServices = objWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
		If objServices.Count = 0 Then
			' The agent services are no longer present
			INSTALLAGENT strSiteID, strDomain
		End If
			

	End IF
End Sub ' ProcessAgent


' ***************************************************************
' Function - InstallAgent - Performs an installation of the agent
' ***************************************************************
Function INSTALLAGENT(strSiteID, strDomain)
	strMessage = "Install agent for N-central site ID " & strSiteID & " on this device which is in domain " & strDomain & " from " & strInstallSource
	' If we're in interactive mode, pop up a question to confirm that we want to start the installation
	If bolInteractiveMode = True Then
		response = MsgBox(strMessage & "?", vbQuestion + vbYesNo, strBranding)
		If response = vbNo Then
			MsgBox "Agent installation aborted.", vbExclamation, strBranding
			CleanQuit
		End If
	End If

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
	objShell.LogEvent evtSuccess, "The agent installer script has started an installation of the agent using the following command:  " & strInstallCommand
	
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
		objShell.LogEvent evtError, "The agent installer script failed to install the agent on this computer within the permitted timeframe."
		If (bolWarningsOn = True Or bolInteractiveMode = True) Then
			MsgBox "The agent could not be installed on this computer." & vbCRLF & vbCRLF & strContactAdmin, vbOKOnly + vbCritical, strBranding
		End If
		DIRTYQUIT errAgent
	Else
		objShell.LogEvent evtSuccess, "The agent installer detected that the Windows Agent Maintenance service has registered on this computer.  The agent is now installed and will register in N-central shortly."
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


' *****************************************************************************
' Sub - GetAgentPath - Returns the true path to the agent.exe binary
'
' Side Effects - updates
'	strAgentBin
'	strAgentPath
'	strAgentConfig		Path to agent config folder, with trailing '\'
' 
' *****************************************************************************
Sub GetAgentPath
	Set objServices = objWMI.ExecQuery("SELECT Pathname FROM Win32_Service WHERE Name = 'Windows Agent Service'")
	
	If objServices.Count <> 0 Then
		' At least one service (hopefully only one!) was returned by the query, so let's find the path to that service
		For Each item In objServices
			strAgentBin = item.Pathname
			strAgentPath = Mid(strAgentBin,2, Len(strAgentBin)-2)	' Strip quotes
			strAgentConfig = Left(strAgentPath, instr(1, strAgentPath, "\bin", vbTextCompare)) + "config\"
			objShell.LogEvent evtInfo, "The agent installer script found the Windows Agent Service's path to be:  " & strAgentBin
		Next
		
		'Check for service exists, but EXE is not present (some incomplete removal has occurred and agent requires reinstall)
		If Not objFSO.FileExists(strAgentBin) Then
			'Agent EXE not found, assume its not installed.
			WRITETOCONSOLE("Agent service found, but EXE not present. Deleting service and reset agent status to not installed." & vbCrLf)
			strAgentBin = ""
			strAgentPath = ""
			objShell.LogEvent evtInfo, "The agent installer script found the Windows Agent service, but not the EXE."
		End If
	Else
		' The agent isn't installed - no service was returned by the WQL query.  An empty string here indicates no agent installed on the device to other sections of this script
		strAgentBin = ""
		strAgentPath = ""
		objShell.LogEvent evtInfo, "The agent installer script found the Windows Agent is not installed."
	End If
End Sub ' GetAgentPath


' ***********************************************************************
' Function - StripAgent - Uses AgentCleanup4 to entirely remove the agent
' ***********************************************************************
Function STRIPAGENT
	If DOTNETPRESENT(strDotNetRegKey)="NOT_INSTALLED" Then
		WRITETOCONSOLE(".NET 4 is not installed - cannot run cleanup" & vbCRLF)
		objShell.LogEvent evtError, "The agent installer script cannot perform a cleanup of the agent as Microsoft .NET Framework 4 is not installed on this device.  The script will now terminate."
		DIRTYQUIT errPrereq
	Else		
		' Now we proceed and do a cleanup of the agent
		strString = ""
		objShell.LogEvent evtWarning, "The agent installer script is running the cleanup process.  This happens when either the agent is not installed (in which case the cleanup is being done in preparation for a fresh installation of the agent) or if not all the services are present (in which case the agent is going to be reinstalled)."
		Set objCleanup = objShell.Exec("cmd /c AgentCleanup4.exe")
		
		Do While objCleanup.Status = 0
			'WRITETOCONSOLE(".")
			WRITETOCONSOLE("Removing agent ................... " & SPIN & vbCR)
			WScript.Sleep 100	' interval delay in milliseconds
		Loop
		WRITETOCONSOLE("Removing agent ................... done!" & vbCRLF)
	End If
End Function


' ********************************************************************************************
' Sub function - VerifyAgent - Runs the agentcleanup command to verify the health of the agent
' ********************************************************************************************
Sub VERIFYAGENT

	' Check to see that we have .NET 4 installed, as we can only run a verify if it is
	If DOTNETPRESENT(strDotNetRegKey)="NOT_INSTALLED" Then
		objShell.LogEvent evtError, "The agent installer script cannot perform a verify of the agent as Microsoft .NET Framework 4 is not installed on this device.  It has probably been uninstalled.  The script will now terminate."
		DIRTYQUIT errPrereq
	End If

	' Verify both Agent and Agent Maintenance services are present and running
	If VerifyServices Then
	
		' Check the value of appliance id in the config\ApplianceConfig.xml file
		If VerifyApplianceID Then

			' Check the value of the ServerIP field in config\ServerConfig.xml
			If VerifyServerConfig Then
			
				' Check to see if the agent is installed onto a non-C drive.  If it is, don't proceed as AgentCleanup4 doesn't support this and will
				' cause repeated uninstalls and installs of the agent
				If strAgentBin <> "" And Mid(strAgentBin, 2,1) <> "C" Then
					WRITETOCONSOLE("Checking health .................. skipped" & vbCRLF)
					objShell.LogEvent evtWarning, "The agent installer cannot perform a verify of the agent on this device as the agent is not installed onto the C drive. The script will now terminate."
					DIRTYQUIT errInternal
				End If
				
				' The agent is installed onto the C drive so we can do a proper verify
				objShell.LogEvent evtSuccess, "The agent installer script started a verify of the agent."
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
						objShell.LogEvent evtWarning, "The agent installer script has detected that the verify stage is probably performing a full cleanup of the agent. The agent is probably not communicating back to N-central or is outdated and is therefore being removed in preparation for a reinstallation."
					End If
				
					' The repair is taking too long, so bug out of this loop
					If count > 4200 Then
						objShell.LogEvent evtError, "The agent installer script tried to repair the agent but the process exceed the permitted timeframe of six minutes."
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
					objShell.LogEvent evtWarning, "The agent installer verify process found a problem with the agent and may have removed it."
				Else
					objShell.LogEvent evtSuccess, "The agent installer script has finished a cleanup/verify"
				End If
			End If ' VerifyServerConfig
		End If ' VerifyApplianceID
	End If ' VerifyServices
End Sub


' *************************************************************************************
' Sub Function - CopyAgentCleanup - Copies the AgentCleanup4.exe from network to device
' *************************************************************************************
Sub CopyAgentCleanup
	Dim bolCopy

	WRITETOCONSOLE("Copying cleanup utility .......... ")
	
	' First we make sure that it's actually available to copy
	If objFSO.FileExists(strInstallSource & "AgentCleanup4.exe")=False Then
		WRITETOCONSOLE("failed!" & vbCRLF)
		objShell.LogEvent evtError, "The agent installer script could not find the AgentCleanup4.exe utility at " & strInstallSource & "AgentCleanup4.exe and cannot proceed."
		runState = errSource
	Else
		' If the local copy does not exist or is not the same version as on the network, copy the network locally
		If objFSO.FileExists(strWindowsFolder & "\AgentCleanup4.exe") = False Then
			' No local copy exists.  Flag to copy
			bolCopy = True
		ElseIf objFSO.GetFileVersion(strWindowsFolder & "\AgentCleanup4.exe") <> objFSO.GetFileVersion(strInstallSource & "\AgentCleanup4.exe") Then
			' Local copy is the wrong version.  Delete and flag to copy
			objFSO.DeleteFile(strWindowsFolder & "\AgentCleanup4.exe")
			bolCopy = True
		Else
			' Local copy exists and matches network version.  No need to copy
			bolCopy = False
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
		
		If bolCopy = True Then
			objFSO.CopyFile strInstallSource & "Agentcleanup4.exe", strWindowsFolder & "\AgentCleanup4.exe"
			
			If objFSO.FileExists(strWindowsFolder & "\agentcleanup4.exe") = False Then
				WRITETOCONSOLE(" failed!" & vbCRLF)
				strMessage = "The agent installer script could not not copy AgentCleanup4.exe from " & strInstallSource & " into " & strWindowsFolder
				If bolInteractiveMode = True Then
					msgbox strMessage & vbCRLF & vbCRLF & strContactAdmin, 16, strBranding
				End If
				objShell.LogEvent evtError, strMessage
				runState = errSource
			Else
				WRITETOCONSOLE("done!" & vbCRLF)
			End If
		End If

	End If
End Sub ' CopyAgentCleanup


' *****************************************************************************
' Function - VerifyServices - Checks the state of Agent and Agent Maintenance
'	services.
'
' Side Effects - none
' 
' *****************************************************************************
Function VerifyServices
	Dim bValid : bValid = True

	' Firstly, check to see that we've got two services present
	Set objServices = objWMI.ExecQuery("SELECT State,Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
	If objServices.Count = 2 Then
		' OK, so the two services are installed.
		objShell.LogEvent evtSuccess, "The agent installer found both agent services installed on this device."

		For Each item In objServices
			If item.State <> "Running" Then
			
				' Report that one of the two services isn't running, and try to start it
				objShell.LogEvent evtWarning, "The agent installer script found that the '" & item.Name & "' service is present but is in state '" & item.State & "' at the moment.  This is not necessarily a problem."
				objShell.LogEvent evtSuccess, "The agent installer script is trying to start the " & item.Name & " service."
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
						objShell.LogEvent evtWarning, "The '" & item.Name & "' service could not be started."
					Else
						WRITETOCONSOLE("Starting " & item.Name & " ... done!" & vbCRLF)
						objShell.LogEvent evtSuccess, "The '" & item.Name & "' service was started successfully."
					End If
				Next
				
				
			Else
				' The service is running
				objShell.LogEvent evtSuccess, "The '" & item.Name & "' service is present and started."
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
			objShell.LogEvent evtSuccess, "The agent installer script found that both agent services are properly registered on this device."
		Else
			' The services are not registered properly so we need to remove the agent
			WRITETOCONSOLE("failed!" & vbCRLF)
			objShell.LogEvent evtError, "The agent installer script found a problem with the registration of the agent services on this device.  This agent is now considered to be corrupt and will be uninstalled."
			If bolInteractiveMode = True Then
					msgbox "The agent services appear to be corrupted on this device.  Therefore the agent will be uninstalled.", vbOKOnly + vbCritical, strBranding
			End If
			
			' Remove the agent
			bValid = False
			STRIPAGENT
		End If
	Else
		' Only one service is listed in WMI (services.msc) so the agent is totally broken, so let's remove it
		bValid = False
		STRIPAGENT
	End If	
	
	VerifyServices = bValid
End Function ' VerifyServices	
	

' *****************************************************************************
' Function - VerifyApplianceID - Checks the appliance ID in ApplianceConfig.XML
'
' Side Effects - none
' 
' *****************************************************************************
Function VerifyApplianceID
	Dim bValid : bValid = True
	Dim strApplianceConfig, strApplianceID
	Dim xmlDoc, objNode
	
	strApplianceConfig = strAgentConfig & "ApplianceConfig.xml"

	WRITETOCONSOLE("Checking appliance ............... ")

	' If the agent appears to be installed, let's check the appliance ID and see if it's a valid one
	'output.writeline "Agent path is " & strAgentPath
	If objFSO.FileExists(strApplianceConfig) Then
		' Check the appliance ID of the agent and write it to the event log
		
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.load(strApplianceConfig)

		Set objNode = xmlDoc.selectSingleNode("//ApplianceConfig/ApplianceID")

		If Not objNode Is Nothing Then
			strApplianceID = CLng(objNode.Text)
		Else
			strApplianceID = 0 
		End If

		If strApplianceID < 1000 Then
			' Definitely a bad number
			objShell.LogEvent evtWarning, "The agent installer script found that the installed agent has an invalid ApplianceID, " & strApplianceID & ".  This agent may not be able to check into " & strServerAddress & " correctly."
			WRITETOCONSOLE("failed!" & vbCRLF)
			'STRIPAGENT			' We don't want to strip the agent out as -1 is not ALWAYS a bad agent
		Else
			' Just write the appliance ID to the event log
			objShell.LogEvent evtSuccess, "The agent installer script found that the installed agent has an ApplianceID of " & strApplianceID & "."
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
	Else ' ApplianceConfig.XML does not exist
		objShell.LogEvent evtError, "The agent installer script could not find the appliance config XML file at " & strApplianceConfig & ".  This indicates a likely corrupt installation."
		WRITETOCONSOLE("failed!" & vbCRLF)
		bValid = False
		STRIPAGENT		' If there's not a ApplianceConfig.xml file, then strip the agent so it can be reinstalled
	End If

	VerifyApplianceID = bValid
End Function ' VerifyApplianceID


' *****************************************************************************
' Function - VerifyServerConfig - Checks the server in ServerConfig.xml
'	This handles the 'localhost' zombie
'
'	Normally, the agent service self-heals the ServerConfig.xml file.  On
'	start, if the agent can connect to then N-Central server at the host
'	defined in ServerIP, its value is saved into BackupServerIP.  If the agent
'	cannot connect to the N-Central server, then BackupServerIP is copied
'	into ServerIP.
'
'	Zombies are created when, for whatever reason, localhost is assigned to
'	both ServerIP and BackupServerIP.
'
'	This procedure checks ServerIP for localhost, and if it is found, uses
'	strServerAddress in this script to rewrite the ServerConfig.xml file
'	and restart the agent.  We assume BackupServerIP reads localhost as well
'	or the agent would have self-healed.
'
' Side Effects - none
' 
' *****************************************************************************
Function VerifyServerConfig
	Dim bValid : bValid = True
	CONST strLH = "localhost"

	Dim strServerConfig, strConfigAddress
	Dim xmlDoc, objNode
	
	strServerConfig = strAgentConfig & "ServerConfig.xml"

	WRITETOCONSOLE("Checking server config ........... ")

	' If the agent appears to be installed, let's check the N-Central server
	' and correct it if it's "localhost"
	If objFSO.FileExists(strServerConfig) Then
		' Check the N-Central server of the agent and write it to the event log
		
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.load(strServerConfig)

		Set objNode = xmlDoc.selectSingleNode("//ServerConfig/ServerIP")

		If Not objNode Is Nothing Then
			strConfigAddress = objNode.Text
		Else
			strConfigAddress = "" 
		End If
		
		If LCase(strConfigAddress) = LCase(strLH) Then
			' Read "localhost" from ServerConfig.XML  Log the issue and remediate
			objShell.LogEvent evtWarning, "The agent installer script found that the installed agent has localhost in the ServerIP field of ServerConfig.xml.  The script will attempt to remediate this."
			WRITETOCONSOLE("Remediation needed" & vbCRLF)
			'
			' We need to stop the agent services here and write the value of strServerAddress into objNode, save the xml file back, and restart the services.

			' Stop both services
			objShell.LogEvent evtInfo, "Stopping both agent services so the ServerConfig.xml file can be updated."
			Set objServices = objWMI.ExecQuery("SELECT State,Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
			For Each item In objServices
				item.StopService()
					
				' Wait a bit of time for the service to stop
				count = 0
				Do While count < 100
					WRITETOCONSOLE("Stopping " & item.Name & " ... " & SPIN & vbCR)
					WScript.Sleep 100
					count = count + 1
				Loop
									
				' Check that the service held in item.Name is running
				Set colService = objWMI.ExecQuery("SELECT Name,State FROM Win32_Service WHERE Name = '" & item.Name & "'")
				For Each service In colService
					'output.writeline service.name & service.State
					If service.state = "Running" Then
						WRITETOCONSOLE("Stopping " & item.Name & " ... failed!" & vbCRLF)
						objShell.LogEvent evtWarning, "The '" & item.Name & "' service could not be stopped."
					Else
						WRITETOCONSOLE("Stopping " & item.Name & " ... done!" & vbCRLF)
						objShell.LogEvent evtSuccess, "The '" & item.Name & "' service was stopped successfully."
					End If
				Next
			Next
			
			' Wait up to 60 seconds for the file handles to be released
			count = 600
			Do While count > 0 and Not IsWriteAccessible(strServerConfig)
				If count Mod 10 = 0 Then
					WRITETOCONSOLE("Waiting for file handles to be released ... " & count/10 & " " & vbCR)
				End If
				WScript.Sleep 100
				count = count - 1
			Loop
			
			If IsWriteAccessible(strServerConfig) Then
				WRITETOCONSOLE("Waiting for file handles to be released ... done!" & vbCRLF)
				
				' Modify and save ServerConfig.XML
				WRITETOCONSOLE("Updating ServerConfig.xml ... ")
				objNode.Text = strServerAddress
				xmldoc.save(strServerConfig)
				objShell.LogEvent evtInfo, "The ServerIP field in " & strServerConfig & " was updated to " & strServerAddress
				WRITETOCONSOLE("done!" & vbCRLF)
				
				' Start the services
				Set objServices = objWMI.ExecQuery("SELECT State,Name FROM Win32_Service WHERE Name LIKE 'Windows Agent%'")
				If objServices.Count = 2 Then
					' OK, so the two services are installed.
					objShell.LogEvent evtSuccess, "The agent installer found both agent services installed on this device."

					For Each item In objServices
						objShell.LogEvent evtSuccess, "The agent installer script is trying to start the " & item.Name & " service."
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
								objShell.LogEvent evtWarning, "The '" & item.Name & "' service could not be started."
							Else
								WRITETOCONSOLE("Starting " & item.Name & " ... done!" & vbCRLF)
								objShell.LogEvent evtSuccess, "The '" & item.Name & "' service was started successfully."
							End If
						Next
					Next			
				End If
				
				WRITETOCONSOLE("Checking server config ........... done!" & vbCRLF)				
			Else
				WRITETOCONSOLE("Waiting for file handles to be released ... failed!" & vbCRLF)
			End If
		ElseIf strConfigAddress = "" Then
			objShell.LogEvent evtError, "The agent installer script could not find the ServerIP field in " & strServerConfig & ".  This indicates a likely corrupt installation."
			WRITETOCONSOLE("failed!" & vbCRLF)
			bValid = False
			STRIPAGENT		' If there's not a ServerConfig.xml file, then strip the agent so it can be reinstalled	
		Else
			objShell.LogEvent evtSuccess, "The agent installer script found the ServerIP field in " & strServerConfig & " is " & strConfigAddress
			WRITETOCONSOLE("done!" & vbCRLF)
		End If
	Else ' ServerConfig.XML does not exist
		objShell.LogEvent evtError, "The agent installer script could not find the server config XML file at " & strServerConfig & ".  This indicates a likely corrupt installation."
		WRITETOCONSOLE("failed!" & vbCRLF)
		bValid = False
		STRIPAGENT		' If there's not a ServerConfig.xml file, then strip the agent so it can be reinstalled
	End If
	
	VerifyServerConfig = bValid
End Function ' VerifyServerConfig


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
End Function ' IsDowngrade


' *****************************************************************************
' Function - IsWriteAccessible - Returns if the file has no locks for other
'	processes
' 
' *****************************************************************************
Function IsWriteAccessible(sFilePath)
    ' Strategy: Attempt to open the specified file in 'append' mode.
    ' Does not appear to change the 'modified' date on the file.
    ' Works with binary files as well as text files.

	On Error Resume Next
	
    Const ForAppending = 8

    IsWriteAccessible = False

    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
	
    Dim oFile : Set oFile = oFso.OpenTextFile(sFilePath, ForAppending)
    If Err.Number = 0 Then
        oFile.Close
        If Not Err Then
            IsWriteAccessible = True
        End if
    End If
	
	On Error Goto 0
End Function ' IsWriteAccessible


' *****************************************************************************************
' Sub Function - WriteToConsole - Writes a string to the console if interactive mode is off
' *****************************************************************************************
Sub WRITETOCONSOLE(strMessage)
	If bolInteractiveMode = False Then
		WScript.StdOut.Write strMessage
	End If
End Sub


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


' *****************************************************************************
' Function - WriteRegValue - Writes a value into the Registry with minimal fuss
' *****************************************************************************
Function WRITEREGVALUE(strKey, strType, strValue)
	WRITETOCONSOLE("Writing " & strKey & " " & strType & " " & strValue)
	If Instr(strValue, " ")>0 Then strValue = Chr(34) & strValue & Chr(34) End If
	Set objCmd = objShell.Exec("cmd /c reg add " & Chr(34) & "HKLM\Software\" & strRegBase & "\InstallAgent" & Chr(34) & " /v " & strKey & " /t " & strType & " /d " & strValue & " /f")
End Function


' *********************************************************************************************************************
' Function - CleanQuit - Exits the script cleanly with error code 10, which can be picked up by the launcher batch file
' *********************************************************************************************************************
Function CLEANQUIT
	objShell.LogEvent evtSuccess, "The agent installer script v" & strVersion & " has finished running."
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastSuccessfulRun", "" & Now()
	objReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastOperation", 10
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastExitComment", "Successful execution"
	WScript.Quit 10
End Function


' ************************************************************************
' Function - DirtyQuit - Exits the script with an error code and Reg value
' ************************************************************************
Function DIRTYQUIT(intValue)
	objReg.SetDWORDValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastOperation", intValue
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
	objReg.SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & strRegBase & "\InstallAgent\", "LastExitComment", strExitComment
	WRITETOCONSOLE(vbCRLF & "An error occurred. " & strExitComment & vbCRLF & strContactAdmin & vbCRLF)
	objShell.LogEvent evtError, "The agent installer script experienced a problem and exited prematurely.  The exit code was " & intValue & " (" & strExitComment & ")"
	WScript.Quit intValue
End Function


' *****************************************************************************************
' Work with INI files In VBS (ASP/WSH)
' v1.00
' 2003 Antonin Foller, PSTRUH Software, http://www.motobit.com
' Modified by Jon Czerwinski, Cohn Consulting, 20130324
' *****************************************************************************************
' **************************************************************************
' Sub - WriteINIString - Writes an INI value to file, creating section
'       and value as necessary
' **************************************************************************
Sub WriteINIString(FileName, ByRef INIContents, Section, KeyName, Value)
  Dim PosSection, PosEndSection
  
  ' Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    ' Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    ' Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    ' Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    ' Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    ' Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Line = KeyName & "=" & Value
        Found = True
      End If
      NewsContents = NewsContents & Line & vbCrLf
    Next

    If isempty(Found) Then
      ' key Not found - add it at the end of section
      NewsContents = NewsContents & KeyName & "=" & Value
    Else
      ' remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    ' Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    ' Section Not found. Add section data at the end of file contents.
    If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
      INIContents = INIContents & vbCrLf 
    End If
    INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
      KeyName & "=" & Value
  end if'if PosSection>0 Then
  PutFile FileName, INIContents
End Sub


' **************************************************************************
' Function - GetINIString - Retrieves value from INI file
' **************************************************************************
Function GetINIString(IniContents, Section, KeyName, Default)
  Dim PosSection, PosEndSection, sContents, Value, Found
  
  ' Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    ' Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    ' Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    ' Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      ' Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    End If
  End If
  If isempty(Found) Then Value = Default
  GetINIString = Value
End Function


' **************************************************************************
' Function - SeparateField - Extracts field from sFrom, between sStart
'       and sEnd
' **************************************************************************
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function


' **************************************************************************
' Function - GetFile - Loads INI file
' **************************************************************************
Function GetFile(ByVal FileName)
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")

  On Error Resume Next

  GetFile = FS.OpenTextFile(FileName).ReadAll
End Function


' **************************************************************************
' Function - PutFile - Saves INI file
' **************************************************************************
Function PutFile(ByVal FileName, ByVal Contents)
  
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")

  Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 2, True)
  OutStream.Write Contents
End Function

