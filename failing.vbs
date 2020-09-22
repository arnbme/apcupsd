'///Windows testing: open a command window: winkey + R, enter Cmd, then execute: cscript C:"\apcupsd\etc\apcupsd\failing.vbs ///
'///Comment out the next line when testing, then change back when done. ///
dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.Exec("cscript C:\apcupsd\etc\apcupsd\smartUPS.vbs failing")
set WshShell = nothing
'/// Exit with a 0 error level to ensure the apccontrol.bat continues ///
Wscript.Quit 0