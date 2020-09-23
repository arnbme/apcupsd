<a name="top"></a>
# SmartUPS V1 Sep 22, 2020 
# The Windows Centric Version (A work in progress)
## Table of Contents 
[&ensp;1. Purpose](#purpose)<br />
[&ensp;2. Features](#features)<br />
[&ensp;3. Donate](#support)<br />
[&ensp;4. Installation Overview](#installOver)<br />
[&ensp;5. SmartUPS driver and VBS modules](#modules)<br />
[&ensp;6. Create Virtual Device](#vdevice)<br />
[&ensp;7. Create a Windows Schedule Task](#windowstask)<br />
[&ensp;8. Testing](#testing)<br />
[&ensp;9. Create RM Power Control Rule(s)](#rules)<br />
[10. Restarting the Hub after a graceful shutdown](#restartHub)<br />
[11. Restarting the Windows system after a shutdown](#restartWin)<br />
[12. Uninstalling](#uninstall)<br />
[13. Get Help, report an issue, or contact information](#help)<br />
[14. Known Issues](#issues)

<a name="purpose"></a>
## 1. Purpose
SmartUPS, Windows Centric Version, allows a Hubitat Hub plugged into an APC UPS, and using apcupsd.org's package on a Windows machine, to gracefully shut down during a power outage. This version is maintained by Arn Burkhoff was derived from Steve Wright's SmartUPS release.

[:arrow_up_small: Back to top](#top)

<a name="features"></a>
## 2. Features<br />

* Reports UPS Device Events: onbattery - mains restored, offbattery - mains down , failing - UPS about to shutdown, and powerout - ?????<br />
* Sends UPS Device Statistics: every "user defined" minutes, using a repeating Windows Scheduled Task<br />
* Simplified installation and setup when compared to original version using Windows<br /> 
* Windows modules are VBS, a Windows PHP server is not used or required<br />
* Executes without being logged in to Windows

[:arrow_up_small: Back to top](#top)
<a name="support"></a>
## 3. Support this project
This app is free. However, if you like it, derived benefit from it, and want to express your support, donations are appreciated.
* Paypal: https://www.paypal.me/arnbme 

[:arrow_up_small: Back to top](#top)

<a name="installOver"></a>
## 4. Installation Overview
1. Uninstall APC PowerChute, if installed

2. Connect APC UPS supplied cable to a USB port when necesssary 
3. [Install apcupds app](http://www.apcupsd.org), then setup apcupsd
4. [Install module SmartUPS.groovy](#modules) from Github repository into Hub's Drivers 
5. [Copy the five VBS modules](#modules) from Github repository to Windows directory C:/apcupsd/etc/apcupsd
6. Create a virtual device using SmartUps driver, then set IP address to your Windows machine IP address. This should be a permanently reserved address in router. 
7. Create a Windows Scheduled Task
8. Reboot Windows system, then test

[:arrow_up_small: Back to top](#top)

<a name="modules"></a>
## 5. SmartUPS Driver and VBS modules

There are five VBS scripts and a one Groovy Device Handler (DH) associated with this app stored in a Github respitory. You may also install he Groovy module using the [Hubitat Package Manager](https://community.hubitat.com/t/beta-hubitat-package-manager/38016). The VBS scripts must be copied to C:/apcupsd/etc/apcupsd from Github  
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
## 6. Create a Virtual Device


[:arrow_up_small: Back to top](#top)

<a name="windowstask"></a>
## 7. Create a Windows Scheduled Task.

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
  
[:arrow_up_small: Back to top](#top)
<a name="testing"></a>
## 8. Testing

[:arrow_up_small: Back to top](#top)
<a name="rules"></a>
## 9. Prepare RM Power Control rule(s)

[:arrow_up_small: Back to top](#top)
<a name="keypadDH"></a>
## 10. Restarting the Hub after a graceful shutdown

<a name="restartWin"></a>
## 11. Restarting the Windows system after a shutdown

[:arrow_up_small: Back to top](#top)

<a name="uninstall"></a>
## 12. Uninstalling
1. Delete scheduled task<br />
2. Uninstall apcupsd<br />
3. Remove SmartUPS virtual device<br />
4. Remove SmartUPS from Devices code

[:arrow_up_small: Back to top](#top)
<a name="help"></a>
## 13. Get Help, report an issue, and contact information
* [Use the HE Community's Nyckelharpa forum](https://community.hubitat.com/t/release-nyckelharpa/15062) to request assistance, or to report an issue. Direct private messages to user @arnb

[:arrow_up_small: Back to top](#top)

<a name="issues"></a>
## 14. Known Issues
* The SmartUPS device Refresh command does nothing because no server is available for communications
[:arrow_up_small: Back to top](#top)
