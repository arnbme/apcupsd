'///Windows testing: open a command window: winkey + R, enter Cmd, then execute: cscript C:"\apcupsd\etc\apcupsd\onbattery.vbs ///
'///Comment out the next line when testing, then change back when done. ///
dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.Exec("cscript C:\apcupsd\etc\apcupsd\smartUPS.vbs doshutdown")
set WshShell = nothing
WScript.Sleep(30000)	'///wait for hub to send messages and shutdown
'/// Exit with a 0 error level to ensure the apccontrol.bat continues ///
Wscript.Quit 0