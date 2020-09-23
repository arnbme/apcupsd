<a name="top"></a>
# SmartUPS V1 Sep 22, 2020 
# The Windows Centric Version 
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
[16. Uninstalling](#uninstall)<br />
[17. Get Help, report an issue, or contact information](#help)<br />
[18. Known Issues](#issues)

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

* https://docs.hubitat.com/index.php?title=How_to_Install_Custom_Apps <br />
* let's begin by installing the Nyckelharpa parent module into Hubitat (HE) from this Github repository.<br />
* OAuth is not required and should be skipped.<br /> 
* Save the the Nyckeharpa module
* Using Install's Import button: each module's Github raw address is availabe at the beginning of the module.<br />
* Then install Modefix, Talker, and User, ignore OAuth, and do not add these modules as User Apps.
* In Apps: click the "Add User App" button, select the Nycklharpa, click Done

* Should you be using the user Centralite Keypad driver follow these directions<br />
https://docs.hubitat.com/index.php?title=How_to_Install_Custom_Drivers

*  Next step: Quick Setup Guide

[:arrow_up_small: Back to top](#top)


<a name="vdevice"></a>
## 6. Create a Virtual Device

Global Settings is reached by: clicking Apps in the menu, then click the Nyckelharpa app, scroll down to Global Settings, then click  "click to show" 
1. Select all the keypads used for arming HSM
* When devices are selected, default options for valid and invalid pin message routing are shown

2. <b>Prepare for Forced Arming:</b> <i>For each armState</i> select real contact sensor devices that will allow HSM arming when the device is Open.
* _When Global Settings is saved, each selectd contact generates a child Virtual Contact Sensor named NCKL-contact-sensor-name that must be used to Adjust HSM Settings for Forced HSM Arming
* Specify optional destinations for "Arming canceled contact open" and "Arming forced messages: Push, SMS, Talk. Optional, but must be set to output these messages
3. Select any contact to be monitored for Open / Close Talker messages only, that are not used with Forced HSM Arming

4. Select any alarms and beeps as required
5. Set the Virtual Child Device prefix, Default NCKL. Once set, it displays but cannot be changed.

6. Set any Hubitat PhoneApp and Pushover messaging devices

7. *When finished, click Next, then click Done*

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

Modefix processes HSM armState changes, and optionally sets the Hubitat HSM mode. _It must be created even when the optional mode change data is empty._ 

(Optional) For each armState set:
* all valid modes for the armState
* the default mode for the armState

Caution: improper armState/mode choices, creates havoc with the system.

[:arrow_up_small: Back to top](#top)
<a name="rules"></a>
## 9. Prepare RM Power Control rule(s)

Table with Reason Issued and Message Issued. 
* Pin messages, arming canceled, and arming forced, do not allow for text adjustment. 
 <table style="width:100%">
  <tr>
    <th>Reason Issued</th>
    <th>Default Message</th>
   <th>Destinations</th> 
   <th>Issueing Module</th> 
  </tr>
  <tr>
    <td>Contact Sensor Opens, arm state disarmed</td>
    <td>%device is now open</td>
   <td>TTS, Speaker</td>
   <td>Nyckelharpa</td>
  </tr>
  <tr>
    <td>Contact Sensor Closes, arm state disarmed</td>
   <td>%device is now closed</td>
   <td>TTS, Speaker</td>
   <td>Nyckelharpa</td>
 </tr>
  <tr>
    <td>Exit Delay</td>
    <td>Alarm system is arming in %nn seconds. Please exit the facility</td>
   <td>TTS, Speaker</td>
   <td>Nyckelharpa Modefix</td>
  </tr>
  <tr>
    <td>Entry Delay</td>
    <td>Disarm system, or police will arrive shortly</td>
    <td>TTS, Speaker</td>
   <td>Nyckelharpa</td>
  </tr>
  <tr>
    <td>System Armed</td>
    <td>Alarm System is now armed in %hsmStatus Mode</td>
   <td>TTS, Speaker</td>
   <td>Nyckelharpa Modefix</td>
  </tr>
  <tr>
    <td>System Disarmed</td>
    <td>System Disarmed</td>
   <td>TTS, Speaker</td>
   <td>Nyckelharpa Modefix</td>
  </tr>
  <tr>
    <td>Valid Pin Entered<br>Centralitex driver only</td>
    <td>%keypad.displayname set HSM state to %armState with pin for %userName</td>
   <td>User Defined in global Settings</td>
   <td>Nyckelharpa</td>
  </tr> 
  <tr>
    <td>Bad Pin Entered<br>Centralitex driver only</td>
    <td>%keypad.displayname Invalid pin entered: %pinCode</td>
    <td>User Defined in global Settings</td>
    <td>Nyckelharpa</td>
  </tr>
   <tr>
    <td>Arming Canceled Open Contact (1)</td>
    <td>Arming Canceled %contact name(s) is open. Rearming within 15 seconds will force arming </td>
    <td>User Defined in global Settings</td>
    <td>Nyckelharpa</td>    
  </tr> 
  <tr>
    <td>Arming Forced Open Contact<br>Centralitex driver only</td>
    <td>Arming Forced %contact name(s) is open.</td>
    <td>User Defined in global Settings</td>
   </td><td>Nyckelharpa</td>
  </tr>
  <tr>
    <td>Intrusion Message</td>
   <td>Defined in HSM</td>
    <td>User Defined HSM</td>
   </td><td>HSM</td>
  </tr>
<tr>
    <td>Panic Message</td>
   <td>Defined in HSM Custom Rule</td>
    <td>User Defined Custom HSM Rule (see Section 12)</td>
   </td><td>HSM</td>
  </tr>
  </table>
  
1. In order to get the Nyckelharpa contacts open message and forced arming when using anything other the the Centralite Keypad driver: you must create an alert in HSM's Configure Arming/Disarming/Cancel --> Configure Alerts for Arming Failures. Please read Section 7, paragraph 2. 

[:arrow_up_small: Back to top](#top)
<a name="keypadDH"></a>
## 10. Restarting the Hub after a graceful shutdown

The Centalitex Keypad Device Handler was created by Mitch Pond on SmartThings, where it is still used by a few Smartapps including SHM Delay. With Mitch's assistance and Zigbee skills it was ported to HE, then I added the Alarm capability that sounds a fast high pitch tone until set off on the Iris V2, and beeps for 255 seconds on the Centralite, and the compatabilitty with the Hubitat keypad device drivers, HSM and Lock Code Manager. 

_This DH may be used with the Centralite/Xfinity 3400, Centralite 3400-G, Iris V2, Iris V3 and UEI keypads_

1. After installing the keypad DH, edit keypad devices changing Type to Centralitex Keypad, Save Device

2. Remove keypads using Centralitex driver from HSM. In HSM section Configure Arming/Disarming/Cancel Options --> Use keypad(s) to arm/disarm: optionally remove keypads using the Centralitex driver.

3. Add keypad to Nyckelharpa Global Settings

4. When using Nyckelharpa pins: Create User pin profiles. When using an Iris V3 User pin code 0000 is required and used for instant arming, but will not disarm. This keypad does not send a pin, even if entered, when arming.

5. When using Lock Code Manager Pins: in this device's setting set "Use Lock Code Manager Pins" on, save settings

5. Create HSM Custom Panic Rule

6. When using an Iris V2/V3 keypad set if Partial key creates Home (default) or Night arming mode

[:arrow_up_small: Back to top](#top)
<a name="restartWin"></a>
## 11. Restarting the Windows system after a shutdown

When using the app's keypad Device Handler and User Pin Module
* For each valid user pin, create a User pin profile

* Pin codes may be restricted by date/time, use count (burnable pins), and keypad device

* To use the Iris V2's instant arming, no pin required, create a User profile with pin code 0000. It is not accepted for OFF

* *The Iris V3 requires a User profile with pin code 0000, or it will not arm.* It is not accepted for OFF.

* You may define "Panic Pins" designed for use on keypads without a Panic key, but may be used on any keypad

[:arrow_up_small: Back to top](#top)

<a name="uninstall"></a>
## 16. Uninstalling
1. If using forced arming, change HSM settings NCKL-child devices to real devices<br />
2. If using Panic Key or Panic pins, remove custom Panic rule from HSM<br />
3. it is now safe to remove Nyckelharpa, child devices are deleted during removal process

[:arrow_up_small: Back to top](#top)
<a name="help"></a>
## 17. Get Help, report an issue, and contact information
* [Use the HE Community's Nyckelharpa forum](https://community.hubitat.com/t/release-nyckelharpa/15062) to request assistance, or to report an issue. Direct private messages to user @arnb

[:arrow_up_small: Back to top](#top)

<a name="issues"></a>
## 18. Known Issues
* Messages need individual custom destination settings

* SMS was disabled by Hubitat, but is still defined as a destination. Do not use SMS

* Iris V3 Keypad Issue: Lights remain on when device is sitting upright on a table or flat surface.<br /> 
Cause: Keypad's motion or proximity sensor is activated.<br /> 
Solution: Move keypad to edge of table, lay it flat on table or surface, or mount it on a wall. 

* Do not intermix Centralitex keypad drivers with system keypad drivers. Use one or the other

[:arrow_up_small: Back to top](#top)
