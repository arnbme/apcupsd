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
[10. Adjust Windows Power Settings](#sleep)<br />
[11. Create RM Power Control Rule(s)](#rules)<br />
[12. Restarting the Hub after a graceful shutdown](#restartHub)<br />
[13. Restarting the Windows system after a shutdown](#restartWin)<br />
[14. Should I plug my computer into the UPS' battery backup?](#plugin)<br />
[15. Uninstalling](#uninstall)<br />
[16. Get Help, report an issue, or contact information](#help)<br />
[17. Known Issues](#issues)

<a name="purpose"></a>
## 1. Purpose
**Perform a graceful Hub shutdown when power is lost.** 

Developed for, and tested on a Windows 10 system, but may work on any system supporting apcupsd and Visual Basic Script. 

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
  * onbattery - mains power failed
  * offbattery - mains power restored
  * failing - UPS about to shutdown
  * powerout - UPS switched to batteries (for any reason)
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
7. [Test the VBS scripts and Hubitat SmartUPS device](#testing)
7. [Create a Windows Scheduled Task](#windowstask)
8. [Adjust Windows Power setings](#sleep)
8. Reboot Windows system, then verify smartUPS.vbs is running on your selected schedule.
9. [Review: Should I plug my computer into the UPS' battery backup][#plugin]

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
## 7. Create and setup a Virtual Device
   * On Hubitat main menu, click Devices
   * Click: Add Virtual Device
   * Set Device Name and Device Label to: APC UPS (or whatever you want)
   * Type: Select User driver - SmartUPS
   * Click button: Save Device
   ----------------------------
   * Set IP Address of system running apcupsd
   * Click button: Save Preferences
   ----------------
   * Set Event History to 5 (optional)
   * Click button: Save Device   


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
## 9. Create a Windows Scheduled Task
1. Open the Windows Task Scheduler
   * Click on Windows task bar "search" icon
   * Enter: Task Scheduler
   * Click on Run
   ------------
2. Task Scheduler opens   
   * On right side of screen, Click on Create Task
   * Set name: SmartUps
   * Select "run whether user is logged in or not"
   * On top of window click Triggers, then click New
   ------------
3.  Set Triggers 
    * Set Begin the task selector to: At Startup
    * Check Delay task for, key in 3 minutes   
    * Check Repeat task for, 10 minutes, for a duration of: Indefinitely 
    * Enabled should be set, Set enabled if not
    * Click OK
    * On top of screen click Actions, then click New
    ----------
4. Set Actions
   * Set Program/script to: cscript  
   * Set Arguments to C:\apcupsd\etc\apcupsd\smartUPS.VBS
   * click OK 
   * On top of screen click Conditions
   -----------
5. Set Conditions
   * Uncheck Start the task only if computer is on AC power
   ------------
6. Save then test the Task   
   * Click OK on bottom of window
   * Enter your windows password (not the pin)
   * The task is created
   * Test it by clicking Run on right side of window
   -------------
7. Activate task by Rebooting system

[:arrow_up_small: Back to top](#top)

<a name="sleep"></a>
## 10. Adjust Windows Power Settings
* On Power and Batteries change "Put Computer to Sleep" to Never
Unless you can figure out a way to wake the machine as needed

[:arrow_up_small: Back to top](#top)


<a name="rules"></a>
## 11. Prepare RM Power Control rule(s)
Additional notification rules for events onbattery and offbattery are strongly suggested. 

![image RM Power](https://github.com/arnbme/apcupsd/blob/master/images/RMPower.png)


[:arrow_up_small: Back to top](#top)
<a name="keypadDH"></a>
## 12. Restarting the Hub after a graceful shutdown

A Wifi plug between the UPS plug and the HE Hub power connector allows for a remote hub restart in case the Hub must be power cycled to restart after a graceful shutdown. This occurs when the Hub is gracefully shutdown, but never loses power. The hub must be power cycled to restart when power is restored.

When the Hub loses power, it will automatically restart when power is restored.

<a name="restartWin"></a>
## 13. Restarting the Windows system after a shutdown
Some links
   * https://www.technewsworld.com/story/78930.html
   * https://www.technewsworld.com/story/86034.html

[:arrow_up_small: Back to top](#top)

<a name="plugin"></a>
## 14. Should I plug my computer into the UPS' battery backup?

<table><tr><th>
<th>Computer has Batteries<th>Computer Plugged into UPS<th>Results <tr><td>1.<td>Yes<td>Yes<td>Ideal, computer shuts down at event 'failing" or when computer batteries reach low level after "failing" event<tr><td>2.<td>Yes<td>No<td>May work, may not. Computer's batteries must last longer than UPS power<tr><td>3.<td>No<td>No<td>It will never work, computer dies at power failure <tr><td>4.<td>No<td>Yes<td>Works, computer shuts down at event 'failing" </tr></table>
 
 [:arrow_up_small: Back to top](#top)

<a name="uninstall"></a>
## 15. Uninstalling
1. Delete scheduled task<br />
2. Uninstall apcupsd<br />
3. Remove SmartUPS virtual device<br />
4. Remove SmartUPS from Devices code

[:arrow_up_small: Back to top](#top)
<a name="help"></a>
## 16. Get Help, report an issue, and contact information
* [Use the HE Community's SmartUps VBS Version forum](https://community.hubitat.com/t/release-nyckelharpa/15062) to request assistance, or to report an issue. Direct private messages to user @arnb

[:arrow_up_small: Back to top](#top)

<a name="issues"></a>
## 17. Known Issues
* The SmartUPS device Refresh command does nothing because no server is available for communications
* The device's non-functional commands: Cancel, Pause, Set Time Remaining, Start, and Stop are inserted by the Hubitat system, and throw an error when clicked.   


[:arrow_up_small: Back to top](#top)
