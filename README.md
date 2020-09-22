<a name="top"></a>
# SmartUPS V1 Sep 22, 2020 
# The Windows Centric Version 
## Table of Contents 
[&ensp;1. Purpose](#purpose)<br />
[&ensp;2. Features](#features)<br />
[&ensp;3. Donate](#support)<br />
[&ensp;4. Installation Overview](#installOver)<br />
[&ensp;5. SmartUPS driver and VBS modules](#modules)<br />
[&ensp;6. Global Settings](#globals)<br />
[&ensp;7. Forced Arming, Adjust HSM Settings](#adjustHSM)<br />
[&ensp;8. Modefix Setup and Usage](#modefix)<br />
[&ensp;9. Talker messages](#talker)<br />
[10. Centralitex Keypad Device Handler](#keypadDH)<br />
[11. User/Pin Profiles with Centralitex driver](#userpin)<br />
[12. Lock Code Manager Pins with Centralitex driver](#lcmpin)<br />
[13. Create Custom HSM Panic Rule](#panicrules)<br />
[14. Debugging](#testing)<br />
[15. Arming From Dashboard](#dboard)<br />
[16. Uninstalling](#uninstall)<br />
[17. Get Help, report an issue, or contact information](#help)<br />
[18. Known Issues](#issues)

<a name="purpose"></a>
## 1. Purpose
SmartUPS, Windows Centric Version, allows an Hubitat Hub plugged into an APC UPS, and using apcupsd.org's package on a Windows machine, to gracefully shut down during a power outage. This version is maintained by Arn Burkhoff was derived from Steve Wright's SmartUPS release.
[:arrow_up_small: Back to top](#top)

<a name="features"></a>
## 2. Features<br />

* Supported events: onbattery - mains restored, offbattery - mains down , failing - UPS about to shutdown, and powerout - ?????<br />
* Device Statistics: reported every "user defined" minutes, using a repeating Windows Scheduled Task<br />
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
3. Install [apcupds.com app](#modefix) then setup apcupsd
4. Install module SmartUPS.groovy into Hub Drivers
5. Create virtual device using SmartUps driver, then set IP address to your Windows machine. Should be a permanently reserved address in roouter 
6. Copy the five VBS modules to directory C:/apcupsd/etc/apcupsd
7. Create a Windows Scheduled Task
8. Reboot Windows system, then test

[:arrow_up_small: Back to top](#top)

<a name="modules"></a>
## 5. SmartUPS Driver and VBS modules

There are five VBS scripts and a one Groovy Device Handler (DH) associated with this app. You may also install he Groovy module using the [Hubitat Package Manager](https://community.hubitat.com/t/beta-hubitat-package-manager/38016). The VBS scripts must be copied to C:/apcupsd/etc/apcupsd from Github  
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


<a name="globals"></a>
## 6. Global Settings. Includes preparation information needed for Forced HSM Arming.

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

<a name="adjustHSM"></a>
## 7. Forced Arming, Adjust HSM Settings.

Forced Arming is a two step process: An standard initial HSM arming that fails normally, followed by a second arming within 15 seconds that forces HSM to arm. It works from any arming source, including: keypads, locks, dashboards, and the HSM app

1. Required Basic Setup: 
* Follow instructions in Section 6 above, generating the NCKL-child-contact-sensors 

2. <b>Important!</b> When arming from the Dashboard HSM Status, a keypad using the HE System Keypad drivers, or anything other than a keypad using the Centralite keypad driver: an alert must be created in HSM's Configure Arming/Disarming/Cancel --> Configure Alerts for Arming Failures (contacts open) section, or HSM immediately arms ignoring all open contacts. 
* Suggested solution:<br /> 
Create a dummy Virtual Switch<br />
Edit dummy Virtual device's settings, disable debug and text logging<br />
In HSM's Configure Arming/Disarming/Cancel --> Configure Alerts --> Light Alerts: turn on the "dummy Virtual Switch".

3. Setup HSM's devices for Forced Arming: 
* In Intrusion-Away, Intrusion-Home, and Intrusion-Night, "Contact Sensors": replace the real contact-sensor-name(s) with the virtual NCKL-contact-sensor-name(s)
* In "Configure/Arming/Disarming/Cancel Options", "Delay only for selected doors": replace the real contact-sensor-name(s) with the virtual NCKL-contact-sensor-name(s)

4. How to Force Arm, a two step process: Arming that fails normally, then Arming again within 15 seconds
* Arm system as you would normally. When there is an open contact sensor monitored by Nyckelharpa, the system will not arm as is normal for HSM
* At the initial arm fail:<br /> 
Talker issues an alert message including the open sensor(s) and the 15 second forced rearm time<br />
Keypads using Centalitex driver:<br />
Centralite/Xfinity-reject tone, delay, two beeps<br />
Iris V2-reject tone, delay, two beeps with old beep or three or four chirps new beep<br />
Iris V3-3 beeps old beep then the reject tone, unable to test new beep my device has old firmware
* Arming the system again, after a minimum of 3 seconds, to a maximum of 15 seconds from the initial arming failure, forces the HSM system to Arm. When using the Centralitex Keypad driver an "Arming Forced" message is issued.

5. Notice: Force Arming fails when attempting to arm using the Dashboard Mode. 
  
[:arrow_up_small: Back to top](#top)
<a name="modefix"></a>
## 8. Modefix setup and usage

Modefix processes HSM armState changes, and optionally sets the Hubitat HSM mode. _It must be created even when the optional mode change data is empty._ 

(Optional) For each armState set:
* all valid modes for the armState
* the default mode for the armState

Caution: improper armState/mode choices, creates havoc with the system.

[:arrow_up_small: Back to top](#top)
<a name="talker"></a>
## 9. Talker messages

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
## 10. Centralitex Keypad Device Handler

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
<a name="userpin"></a>
## 11. Nyckelharpa User pin profiles with Centralitex Keypad driver 

When using the app's keypad Device Handler and User Pin Module
* For each valid user pin, create a User pin profile

* Pin codes may be restricted by date/time, use count (burnable pins), and keypad device

* To use the Iris V2's instant arming, no pin required, create a User profile with pin code 0000. It is not accepted for OFF

* *The Iris V3 requires a User profile with pin code 0000, or it will not arm.* It is not accepted for OFF.

* You may define "Panic Pins" designed for use on keypads without a Panic key, but may be used on any keypad

[:arrow_up_small: Back to top](#top)
<a name="lcmpin"></a>
## 12. Lock Code Manager pins with Centralitex Keypad driver 

When using the app's keypad Device Handler
* In the keypad Preferences set "Use Lock Manager Pins" to On/True. If you want to be be able to see the pins for the device set "Enable Lock Manager Code encryption" to Off; tap/click button "Save Preferences"

* Pins may be maintained using the Lock Code Manager app, or device's pin manangement buttons

* A Panic pin is created by by placing case independent text "panic" as part or all of the pin's "Name", allowing a panic response when using a keypad without a Panic key.

[:arrow_up_small: Back to top](#top)

<a name="panicrules"></a>
## 13. Create Custom HSM Panic Rule

*A custom HSM Rule is required* to force an HSM response to a Panic key press, or Panic pin entry, enabling an instant Panic response even when the system is disarmed

1. Click on Apps-->then click Hubitat Safety Monitor 

2. Click on Custom

3. Click Create New Monitoring Rule --> Name this Custom Monitoring Rule-- enter Panic -->

4. Rule settings
What kind of device do you want to use: select Contact Sensor<br />
Select Contact Sensors: *check the Keypad Devices using Centralitex Keypad Driver, and %prfx%-Panic Id device when using HE drivers*, click Update<br />
What do you want to monitor?: Set Sensor Opens on/true<br />
For how long? (minutes): must be empty or 0<br />
Set Alerts for Text, Audio, Siren and Lights<br />
Click the "Arm This Rule" button<br />
Click Done

5. *Important: verify the Panic rule is "Armed" or it will not work*

6. Do a Panic test: Press the Iris keypad's Panic button, on Centralite 3400-G simultaneously press both "Police Icon" buttons,  or enter a Panic pin number on Centralite 3400 / UEI.

7. The Panic Alert may be stopped by entering a valid user pin on Centralite / UEI, or a valid pin and OFF on an Iris; or the "Cancel Alerts" button from HE App HSM options

[:arrow_up_small: Back to top](#top)
<a name="testing"></a>
## 14. Debugging
1. No entry delay tones on keypad<br />
Keypad may be selected as an Optional Alarm device. Remove it as an Alarm device

2. No exit delay tones<br />
Create and save a Modefix profile

3. Forced arming does not occur<br />
A user reported the Snapshot app somehow interfered with Nyclelharpa's forced arming, and removing or disabling Snapshot fixed the issue. This does not make sense to me, merely reporting what I was told by the user.

[:arrow_up_small: Back to top](#top)
<a name="dboard"></a>
## 15. Arming From Dashboard
* Always arm and disarm using HSM Status. Forced arming is supported and alert messages are created. 

* Mode will generally work, however when there is an alert, the mode remains in the entered mode, but the HSM Status does not change.

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
