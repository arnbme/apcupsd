'///Windows testing: open a command window: winkey + R, enter Cmd, then execute: cscript C:"\apcupsd\etc\apcupsd\onbattery.vbs ///
'///Comment out the next line when testing, then change back when done. ///
On Error Resume Next
Dim objXmlHttpMain , hubitatHubIp, strJSONToSend
hubitatHubIp = "http://192.168.0.106:39501/notify"				'///adjust IP address for your system
strJSONToSend = "{""data"":{""event"":""onbattery""}}"
Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP") 
objXmlHttpMain.open "POST",hubitatHubIp, False 
objXmlHttpMain.setRequestHeader "Content-Type","text/html"
objXmlHttpMain.setRequestHeader "Content-Length","Len(strJSONToSend)"
objXmlHttpMain.setRequestHeader "Accept", "*/*"
objXmlHttpMain.setRequestHeader "Referer","apcupsd"
objXmlHttpMain.setRequestHeader "VBReferer","apcupsd"	'///VB refuses to send Referer header, needs the modified version of SmartUPS.groovy
objXmlHttpMain.setRequestHeader "Connection", "Close"		'///Stop KeepAlive
objXmlHttpMain.send strJSONToSend
objXmlHttpMain = nothing 
set hubitatHubIp = nothing
set strJSONToSend = nothing
'/// Exit with a 0 error level to ensure the apccontrol.bat continues ///
Wscript.Quit 0