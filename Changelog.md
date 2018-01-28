# Changelog

| Version | Date | Notes |
|---------|------|-------|
| 1.00 | 20120928 | First release |
| 1.10 | 20121002 |	Allows the repair process to run for seven mins <br/>Recognises agent installs onto non C drives and amends verify stage accordingly |
| 2.00 | 20121022| Amended verify fail time to be 180 secs instead of 30 secs<br/>Clarified that SXS path needs to be a mapped drive, not a UNC path<br/>Added support for proxy.cfg file which contains proxy configuration to be used when installing the agent<br/>Fixed null string bug in proxy code (thanks, Jon Czerwinski)<br/>Added visible msgbox when ping test fails in interactive mode
| 3.00 | 20121116 |	Slight rephrasing of some event log messages<br/>Agent is no longer verified immediately after installation<br/>Now installs .NET 4 instead of .NET 2, due to 9.1.0.105 (9.1 beta) agent requiring this version<br/>Removed support for installing .NET 3.5 on Windows 8 and Server 2012, .NET 4 is built into these platforms<br/>Server Core now recognised and handled during .NET installer. Cannot yet deploy .NET to this OS as AgentCleanup doesn't work on this platform<br/>Now forces elevation when double clicked
| 3.01 | 20121210 | Fixed a bug with OperatingSystemSKU not being recognised by 2003/XP devices
| 3.02 | 20130123 | Updated for agent 9.1.0.345
| 3.03 | 20130218 |	Fixed issue where the script thought the agent had installed, yet hadn't<br/>Updated to recognise 9.1.0.458 (9.1 GA) agent
| 3.04 | 20130221 | Updated to resolve issue with .NET causing immediate reboot<br/>Now properly detects the Full package of .NET 4 instead of seeing the Client Profile as adequate for the agent installer
| 3.05 | 20130402 |	Now displays the path of the agent if it's found to be installed when running in interactive mode<br/>Updated to work with 9.2.0.142 agent installer<br/>Now uses File Version of the installer EXE instead of checking the file size
| 3.06 | |Added /nonetwork parameter which skips the test for Internet access on networks that block ping<br/>No longer attempts to install on Windows 2000 unless bolNoWin2K is set to True<br/>Now warns if Windows XP without SP3 is detected
| 3.07 | |Powershell awareness.  Writes a warning event if Powershell is not installed or is out of date
| 3.08 | |Double checks agent services when they're found and warns if registration isn't correct
| 3.09 | 20130801 | Now correctly repairs agents with corrupted service registration<br/>No longer alerts that .NET 4 is "INSTALLED" when running in interactive mode.  Only alerts if it's "NOT_INSTALLED"<br/>Alternative source now works for all files, not just the agent source<br/>Now copies AgentCleanup4.exe if the file versions differ between the local and network versions (ie: a new version is available)<br/>Automatically installs Windows Imaging Components if not present<br/>No longer tries to install .NET 4 on a Windows XP device if it doesn't have SP3 installed
| 3.10 | 20130924 | Small update to improve performance of WMI queries
| 3.11 | 20140623 |	Can no longer potentially downgrade the agent, causing a zombie device<br/>Fixed error when agent file version is incorrect<br/>Now tries to start up agent and agentmaint services if they're found to be stopped<br/>Tidier messages by using spinner instead of dotted lines<br/>Now logs if a proxy.cfg file is NOT found, which can be useful when trounleshooting in a proxy environment<br/>Fixed problem with /source parameter not working properly
| 3.12 | 20140904 |Added in bolWarningsOn value which can prevent alerting when agent fails to install
| 3.13 | 20150107 | Added CleanQuit function which returns exit code 10 when the script exits cleanly, so any other code can be picked up by the batch file
| 3.14 | 20150210 | Changed code to allow broken service code to work on Windows XP<br/>Added DirtyQuit fuction which returns passable exit code when the script exits prematurely. All Reg values are now written within this script
| 3.15 | 20150305 | Categorised exit codes (see documentation)<br/>Fixed bug which prevented exit code from being written to Registry properly
| 3.16 | 20150702	| Exit code is now accompanied with a comment on what the code means<br/>Appliance ID is read and written to event log for informational purposes<br/>Fixed bug that caused agent version comparison check to fail on N-central 10 agents<br/>Added ability to adjust the tolerance of the Connectivity (ping) test.  Change the CONST values if you need allow for dropped packets
| 4.00 | 02/11/2015 |Formatting changes and more friendly startup message<br />Dirty exit now shows error message and contact information on console<br />Added 'Checking files' bit to remove confusing delay at that stage. No spinner though, unfortunately<br />This is the final release by Tim :o(<br />First version committed to git - Jon Czerwinski
| 4.01 | 20151109 | Corrected agent version zombie check
| 4.10 | 20151115 | Refactored code - moved mainline code to subroutines, replaced literals with CONSTs<br />Aligned XP < SP3 exit code with documentation (was 3, should be 1)<br />Added localhost zombie checking<br />Changed registry location to HKLM:Software\N-Central<br />NOTE ON REFACTORING - Jon Czerwinski<br/><br/>The intent of the refactoring is:<br />1. Shorten and simplify the mainline of code by moving larger sections of mainline code to subroutines<br/>2. Replace areas where the code quit from subroutines and functions with updates to runState variable and flow control in the mainline.  The script will quit the mainline with its final runState.<br/>3. Remove the duplication of code<br/>4. Remove inaccessible code<br/><br/>This code relies heavily on side-effects.  These have been documented at the top of each function or subroutine.
| 4.20 | 20170119 | Moved partner-configured parameters out to AgentInstall.ini<br />Removed Windows 2000 checks<br />Cleaned up agent checks to eliminate redundant calls to StripAgent<br />Remove STARTUP / SHUTDOWN mode
| 4.21 | 20170126 |	Error checking for missing or empty configuration file. |
| 4.22 | 20170621 |	Close case where service is registered but executable is missing. |
| 4.23 | 20171002 | Bug fix on checking executable path - thanks Rod Clark |
| 4.24 | 20171016 | Rebased on .NET 4.5.2; reorganized prerequisite checks
| 4.25 | 20180128 | Detect whether .ini file is saved with ASCII encoding.  Log error and exit if not.
