<a name="top"></a>
# SmartUPS VBS Version 
## Table of Contents 
[&ensp;1. Purpose](#purpose)<br />
[&ensp;2. Requirements](#require)<br />
[&ensp;3. Features](#features)<br />
[&ensp;4. Donate](#support)<br />
[&ensp;5. Installation Overview](#installOver)<br />
[&ensp;6. SmartUPS driver and VBS modules](#modules)<br />
[&ensp;7. Create Virtual Device](#vdevice)<br />
[&ensp;8. Testing](#testing)<br />
[&ensp;9. Create a Windows Scheduler Task](#windowstask)<br />
[10. Create RM Power Control Rule(s)](#rules)<br />
[11. Restarting the Hub after a graceful shutdown](#restartHub)<br />
[12. Restarting the Windows system after a shutdown](#restartWin)<br />
[13. Uninstalling](#uninstall)<br />
[14. Get Help, report an issue, or contact information](#help)<br />
[15. Known Issues](#issues)

<a name="purpose"></a>
## 1. Purpose
**Perform a graceful Hub shutdown when power is lost.** 

Developed for, and tested on a Windows 10 system, but may work on any system supporting Visual Basic Script. 

This version, maintained by Arn Burkhoff, was derived from Steve Wright's APC UPS Monitor Driver release, but does not use or require a PHP server.

[:arrow_up_small: Back to top](#top)

<a name="require"></a>
## 2. Requirements
* Hubitat Hub plugged into an APC UPS battery backup plug
* The APC communication cable plugged into a Windows machine's USB port
* apcupsd.org's package installed on same Windows machine

Should your APC UPS support WiFi consider [LG Kahn's release](https://community.hubitat.com/t/apc-smartups-status-device/50456)<br />
For Non-windows systems consider [Steve Wright's APC UPS Monitor Driver](https://community.hubitat.com/t/release-apc-ups-monitor-driver/13092)

[:arrow_up_small: Back to top](#top)

<a name="features"></a>
## 3. Features<br />

* Reports UPS Device Events in  Hubitat virtual device attribute, lastEvent
  * onbattery - mains power restored
  * offbattery - mains power down
  * failing - UPS about to shutdown
  * powerout - ?????
* Sends UPS Device Statistics: every "user defined" minutes, using a repeating Windows Scheduled Task.<br />
* Support modules are Visual Basic Script, no Windows server required or used<br />
* Executes without being logged in to Windows

[:arrow_up_small: Back to top](#top)
<a name="support"></a>
## 4. Support this project
This app is free. However, if you like it, derived benefit from it, and want to express your support, donations are appreciated.
* Paypal: https://www.paypal.me/arnbme 

[:arrow_up_small: Back to top](#top)

<a name="installOver"></a>
## 5. Installation Overview
1. Uninstall APC PowerChute, if installed

2. Connect APC UPS supplied cable to a USB port
3. Connect Hub power connector to a UPS Battery Backup plug
   * Place a Wifi plug between the UPS and the Hub power connector, insuring a remote hub restart in some scenarios. I use a TP-Link Kasa plug.
3. [Install apcupds app](http://www.apcupsd.org), then setup apcupsd
4. [Install module SmartUPS.groovy](#modules) from Github repository into Hub's Drivers 
5. [Copy the five VBS modules](#modules) from Github repository to Windows directory C:/apcupsd/etc/apcupsd<br />
Edit your hub's IP address in module smartUPS.VBS
6. [Create a virtual device using SmartUps driver,](#vdevice) then set IP address to your Windows machine IP address. This IP address should be permanently reserved in your router. 
7. [Create a Windows Scheduled Task](#windowstask)
8. Reboot Windows system, then test

[:arrow_up_small: Back to top](#top)

<a name="modules"></a>
## 6. SmartUPS Driver and VBS modules

There are five VBS scripts and a one Groovy Device Handler (DH) associated with this app stored in a Github respitory. You may also install he Groovy module using the [Hubitat Package Manager](https://community.hubitat.com/t/beta-hubitat-package-manager/38016). The VBS scripts must be copied to C:/apcupsd/etc/apcupsd from Github
* After or prior to installing smartUPS.vbs
  * edit the module
  * change hubitatHubIp = "http://192.168.0.106:39501/notify" to the your hub's IP address
  * For example, your hub's IP is 192.168.1.3 the new statement is hubitatHubIp = "http://192.168.1.3:39501/notify" 
 <table style="width:100%">
  <tr>
    <th>Module Name</th>
    <th>Function</th>
    <th>Install Location</th>
  </tr>
  <tr>
    <td>SmartUPS.groovy</td>
    <td>UPS device handler. Controls communication from UPS to Hub</td>
    <td>Hubitat Drivers</td>
  </tr>
  <tr>
    <td>smartUPS.vbs</td>
    <td>Gets UPS devices statistics and sends information to the Hub, also sends all apcupsd event handler information to the Hub</td>
    <td>Windows C:apcupsd/etc/apsupsd</td>
  </tr>
  <tr>
    <td>onbattery.vbs</td>
    <td>apcupsd onbattery event handler</td>
    <td>Windows C:apcupsd/etc/apsupsd</td>
  </tr>  <tr>
    <td>offbattery.vbs</td>
    <td>apcupsd offbattery event handler</td>
    <td>Windows C:apcupsd/etc/apsupsd</td>
  </tr>
  <tr>
    <td>failing.vbs</td>
    <td>apcupsd failing event handler</td>
    <td>Windows C:apcupsd/etc/apsupsd</td>
  </tr>
  <tr>
    <td>powerout.vbs</td>
    <td>apcupsd powerout event hanler</td>
    <td>Windows C:apcupsd/etc/apsupsd</td>
  </tr>
</table> 

[:arrow_up_small: Back to top](#top)


<a name="vdevice"></a>
## 7. Create a Virtual Device


[:arrow_up_small: Back to top](#top)

<a name="testing"></a>
## 8. Testing

1. Open a command window
   * Right click on Windows "Start" icon, menu opens
   * Click on Run
   * Enter cmd, click OK, the command window opens
   
2. Test smartUPS.vbs as follows
   * in command window enter: cscript C:\apcupsd\etc\apcupsd\smartUPS.vbs then click Enter key 
   * verify hub's APC UPS device statistics were updated
3. Test the four event scripts
   * on command line enter: cscript C:\apcupsd\etc\apcupsd\onbattery.vbs then click Enter key<br />
     Hub device attribute results: PowerSouce: battery; lastEvent: onbattery
     
   * on command line enter: cscript C:\apcupsd\etc\apcupsd\offbattery.vbs then click Enter key<br />
     Hub device attribute results: PowerSouce: mains; lastEvent: offbattery
   * on command line enter: cscript C:\apcupsd\etc\apcupsd\powerout.vbs then click Enter key<br />
     Hub device attribute results: PowerSouce: battery; lastEvent: onbattery   
   * on command line enter: cscript C:\apcupsd\etc\apcupsd\failing.vbs then click Enter key<br />
     Hub device attribute results: PowerSouce: battery; lastEvent: failing<br />
     When the RM example rule is installed the hub gracefully shuts down. Restart the Hub with a power cycle: power off, power on.   

[:arrow_up_small: Back to top](#top)

<a name="windowstask"></a>
## 9. Create a Windows Scheduled Task.

Open the Windows Task Scheduler.
1. on right side of screenc Click on Create Task
show image
2. Set name SmartUps, select "run weater user is logged in or not"
3. Click Triggers, then click New
show image
4. Set Begin Task at Startup, set repeat every 10 minutes, Indefinitly, Set enabled, Click OK
5. On top of screen click Actions, then click New
Show image
6. set Program/script to: cscript  
7. set Arguments to C:\apcupsd\etc\apcupsd\smartUPS.VBS click OK
show image
8. Uncheck Start the task only if computer is on AC power
9. Click OK on bottom of window, enter your windows password (not the pin)
10. The task is created, test it by clicking Run, then reboot system to active

* Note: graceful Hub shutdown works without this task  
  
[:arrow_up_small: Back to top](#top)
<a name="rules"></a>
## 10. Prepare RM Power Control rule(s)
Additional notification rules for events onbattery and offbattery are strongly suggested. 

![image RM Power](https://github.com/arnbme/apcupsd/blob/master/images/RMPower.png)


[:arrow_up_small: Back to top](#top)
<a name="keypadDH"></a>
## 11. Restarting the Hub after a graceful shutdown

A Wifi plug between the UPS plug and the HE Hub power connector allows for a remote hub restart in case the Hub must be power cycled to restart after a graceful shutdown. This occurs when the Hub is gracefully shutdown, but never loses power. The hub must be power cycled to restart when power is restored.

When the Hub loses power, it will automatically restart when power is restored.

<a name="restartWin"></a>
## 12. Restarting the Windows system after a shutdown

[:arrow_up_small: Back to top](#top)

<a name="uninstall"></a>
## 13. Uninstalling
1. Delete scheduled task<br />
2. Uninstall apcupsd<br />
3. Remove SmartUPS virtual device<br />
4. Remove SmartUPS from Devices code

[:arrow_up_small: Back to top](#top)
<a name="help"></a>
## 14. Get Help, report an issue, and contact information
* [Use the HE Community's Nyckelharpa forum](https://community.hubitat.com/t/release-nyckelharpa/15062) to request assistance, or to report an issue. Direct private messages to user @arnb

[:arrow_up_small: Back to top](#top)

<a name="issues"></a>
## 15. Known Issues
* The
SmartUPS device Refresh command does nothing because no server is available for communications

[:arrow_up_small: Back to top](#top)
