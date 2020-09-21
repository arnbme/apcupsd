' This is pure alpha code
' You MUST change IP address to your hub IP address at send to Hubitat below	
'///Windows testing: open a command window: winkey + R, enter Cmd, then execute: cscript C:"\apcupsd\etc\apcupsd\smartUPS.vbs ///
'///Test apcaccess: C:\apcupsd\bin\apcaccess.exe
'///Comment out the next line when testing, then change back when done. ///
'On Error Resume Next		'//not needed for stand alone script
'//	Get information from apsupsd
Const WshFinished = 1
Const WshFailed = 2
strCommand = "C:\apcupsd\bin\apcaccess.exe"
Set WshShell = CreateObject("WScript.Shell")
Set WshShellExec = WshShell.Exec(strCommand)		'// cant make WshShell.Run (strCommand,0,1) function, so use exec with a small wait
looper = 0
While WshShellExec.Status = 0
	WScript.Sleep(200)
	looper = looper + 1
	If looper > 15	Then			'// 3 seconds max allowed
		Wscript.Quit 99
	End if	
Wend
'MsgBox WshShellExec.Status		
'MsgBox looper
Select Case WshShellExec.Status
   Case WshFinished
       strOutput = WshShellExec.StdOut.ReadAll
       
   Case WshFailed
       strOutput = WshShellExec.StdErr.ReadAll
	   MsgBox strOutput
	   Wscript.Quit 4095
End Select
'/// format into JSON data no VBS methods AFAIK 
strJSONToSend = "{""data"":{""device"":{"
a=Split(strOutput,vbCrLf)
For each x in a
	if (x>"    ") then
		mySplit = Split(x,":",2)
	    strJSONToSend = strJSONToSend + """"+LCase(trim(mySplit(0)))+""":"""+trim(mySplit(1))+""","
    End if
next

'//Remove trailing comma and close out string
strJSONToSend = Left(strJSONToSend, Len(strJSONToSend) - 1) +"}}}"
'MsgBox strJSONToSend                'write results in a message box

'// send to hubitat
hubitatHubIp = "http://192.168.0.106:39501/notify"				'///adjust IP address for your system
Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP") 
objXmlHttpMain.open "POST",hubitatHubIp, False 
objXmlHttpMain.setRequestHeader "Content-Type","text/html"
objXmlHttpMain.setRequestHeader "Content-Length","Len(strJSONToSend)"
objXmlHttpMain.setRequestHeader "Accept", "*/*"
objXmlHttpMain.setRequestHeader "Referer","apcupsd"
objXmlHttpMain.setRequestHeader "VBReferer","apcupsd"	'///VB refuses to send Referer header, needs the modified version of SmartUPS.groovy
objXmlHttpMain.setRequestHeader "Connection", "Close"		'///Stop KeepAlive
objXmlHttpMain.send strJSONToSend