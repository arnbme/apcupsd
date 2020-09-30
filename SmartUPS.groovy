/**
 *  Smart UPS apcupsd Monitor DeviceType
 *
 *  Copyright 2016 Steve White
 *
 *  Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 *  in compliance with the License. You may obtain a copy of the License at:
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software distributed under the License is distributed
 *  on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License
 *  for the specific language governing permissions and limitations under the License.
 *
 *	The following changes by Arn Burkhoff
 *  2020-09-29 V0.0.7 Enable Refresh command by using EventGhost on Windows to run smartUPS.vbs. Readme document has setup instructions
 *	                           add setting for the port number to use something other than 80 when another server runs on the windows machine
 *                            add settings EventGhost rerfresh timing in minutes, eliminating need for windows task scheduler								
 *  2020-09-26 V0.0.6 Update lastEvent on Device status update, add doshutdown event, default unknown event power to battery  
 *  2020-09-24 V0.0.5 Change namespace to arnbme  
 *  2020-09-23 V0.0.4 Adjust setings positions and text  
 *  2020-09-22 V0.0.3 Add support for UPS "failing" event, UPS is about to shut down, add attribute lastEvent and store last event status
 *  2020-09-21 V0.0.2 Add input setting for using VBS modules in windows 
 *                            Mostly informational for user, but cannot at this time send command to hub for info			
 *  2020-09-19 V0.0.1 Add support for direct send of VBS scripts from Windows apcupsd trigger scripts eliminating need for a Windows PHP server
 *                            add log.error when message is not accepted.
 *                            Note VBS wont send Referer header. Accept a VBReferer header
 */
import groovy.json.JsonSlurper
metadata
{
    definition (name: "SmartUPS", namespace: "arnbme", author: "Steve White")
    {
        capability "Battery"
		capability "Voltage Measurement"
		capability "Power Source"
		capability "Power Meter"
		capability "Timed Session"
		capability "Refresh"
		
		attribute "loadPercent", "string"
        attribute "model", "string"
		attribute "serial", "string"
		attribute "upsStatus", "string"
		attribute "batteryVoltage", "string"
		attribute "loadPercent", "string"
		attribute "timeOnBattery", "string"
		attribute "lifetimeOnBattery", "string"
		attribute "lowTransVolts", "string"
		attribute "highTransVolts", "string"
		attribute "nomPower", "string"
		attribute "sensitivity", "string"
		attribute "lastUpdate", "string"
		attribute "batteryDate", "string"
		attribute "lastPowerFail", "string"
		attribute "lastPowerRestore", "string"
		attribute "lastPowerFailReason", "string"
		attribute "batteryRuntime", "string"
		attribute "lastEvent", "string"
        
		command "refresh"
	}

	preferences
	{	
		section("Device")
		{
			input("ip", "string", title:"IP Address of apcupsd host system (Windows computer)", defaultValue: "192.168.nnn.nnn" ,required: true)		
			input("port", "number", title:"Port Number used by optional Windows EventGhost web server", defaultValue: 80, range: 1..65535, required: true)		
		}
		section
		{
			input "prefEventGhost", "bool", required: true, defaultValue: false,
				title: "ON: EventGhost is installed on Windows, enables Refresh command, and Hub controlled statistics updates<br />OFF (Default): Use Windows Task Scheduler for statistics, Refresh command disabled"
			if (prefEventGhost)
				input("prefRefreshMinutes", "number", title:"EventGhost update time in minutes, zero disables EventGhost updates. For example when using Windows Scheduler updates", defaultValue: 5, range: 0..20, required: true)		
		}
		section
		{
			input "enableDebug", "bool", title: "Enables debug logging for 30 minutes", defaultValue: false, required: false
		}
	}
}


def installed()
{
	log.info "Shackrats apcupsd UPS Monitor - UPS Installed"
	initialize()
}


def updated()
{
	log.info "Shackrats apcupsd UPS Monitor - UPS Updated"
	initialize()
	refresh()
}


def initialize()
{
    unschedule()
	log.info "Shackrats apcupsd UPS Monitor - UPS Initialized."

	if (enableDebug)
	{
		log.info "Verbose logging has been enabled for the next 30 minutes."
		runIn(1800, logsOff)
	}
}


def reset()
{
	initialize()
}


def refresh()
{
	if (enableDebug) log.debug "Entered Refresh"
	if (settings.prefEventGhost)
		{
		def egCommand = java.net.URLEncoder.encode("HE.smartUPS")
		def egHost = "${settings.ip}:${settings.port}" 
		if (enableDebug) log.debug "Sending Refresh to Windows EventGhost at $egHost"
		sendHubCommand(new hubitat.device.HubAction("""GET /?$egCommand HTTP/1.1\r\nHOST: $egHost\r\n\r\n""", hubitat.device.Protocol.LAN))
		if (prefRefreshMinutes && prefRefreshMinutes > 0)
			{
			if (enableDebug) log.debug "Refresh scheduled $prefRefreshMinutes ${prefRefreshMinutes * 60}"
			unschedule(refresh)
			runIn(prefRefreshMinutes*60, refresh) 

			}
		}
}


def parse(String description)
{
	def msg = parseLanMessage(description)
	if (enableDebug) log.debug "PARSED LAN EVENT Received: " + msg
	
	// Update DNI if changed
	if (msg?.mac && msg.mac != device.deviceNetworkId)
	{
		if (enableDebug) log.debug "Updating DNI to MAC ${msg.mac}..."
		device.deviceNetworkId = msg.mac
	}
	// Response to a push notification
	if ((msg?.headers?.Referer == "apcupsd" || msg?.headers?.VBReferer == "apcupsd") && msg?.body)
    {
    	if (enableDebug) log.trace "Processing LAN event notification..."
    	def body = msg?.body
		def sluper = new JsonSlurper();
		def json = sluper.parseText(body)

		// Response to an event notification
        if (json?.data?.event)
		{
			log.info "SmartUPS Push notification for UPS event [${json?.data?.event}] detected."
			
			// Update the child device if it's monitored
			updatePowerStatus(json.data.event)
		}
		// Response to a getUPSstatus call
		else if (json?.data?.device)
		{
			if (enableDebug) log.info "Device update received."
			updateDeviceStatus(json.data.device)
		}
		else log.error "SmartUPS ABORTING DUE TO UNKNOWN EVENT"
	}
    else
        if (enableDebug) log.error "SmartUps Unknown message received Referer:${msg?.headers?.Referer}  VBReferer:${msg?.headers?.VBReferer} body: ${msg?.body}" 
}


def updatePowerStatus(status)
{
	def powerSource =
		status == "mainsback" ? "mains" : 
			status == "powerout" ? "battery" : 
				status == "onbattery" ? "battery" : 
					status == "offbattery" ? "mains" : 
						status == "failing" ? "battery" : 
							status == "doshutdown" ? "battery" : 
						"battery"

	sendEvent(name: "lastEvent", value: status,displayed: this.currentLastEvent != lastEvent ? true : false)
	sendEvent(name: "powerSource", value: powerSource, displayed: this.currentPowerSource != powerSource ? true : false)
    
    if (powerSource == "mains") sendEvent(name: "sessionStatus", value: "stopped", displayed: this.currentSessionStatus != sessionStatus ? true : false)
    else if (powerSource == "battery") sendEvent(name: "sessionStatus", value: "running", displayed: this.currentSessionStatus != sessionStatus ? true : false)
}


def updateDeviceStatus(data)
{
	def power = 0
	def timeLeft = Math.round(Float.parseFloat(data.timeleft.replace(" Minutes", "")))
	def battery = Math.round(Float.parseFloat(data.bcharge.replace(" Percent", "")))
	def voltage = Float.parseFloat(data.linev.replace(" Volts", ""))
	def batteryVoltage = Float.parseFloat(data.battv.replace(" Volts", ""))
	def lowTransVolts = Float.parseFloat(data.lotrans.replace(" Volts", ""))
	def highTransVolts = Float.parseFloat(data.hitrans.replace(" Volts", ""))
	def loadPercent = Float.parseFloat(data.loadpct.replace(" Percent", ""))
    def nomPower = Float.parseFloat(data.nompower.replace(" Watts", ""))
	def powerSource =
    	data.status == "ONLINE" ? "mains" : 
        	data.status == "ONBATT" ? "battery" : 
            	"mains"

//	update lastEvent as necessary
	def lastEvent
	if (powerSource=='mains')
		lastEvent = 'offbattery'
	else
		lastEvent = 'onbattery'
	if (lastEvent != device.currentValue('lastEvent'))	
		sendEvent(name: "lastEvent", value: lastEvent)

	// Calculate wattage as a percentage of nominal load
    power = ((loadPercent / 100) * nomPower)
    
	sendEvent(name: "powerSource", value: powerSource, displayed: this.currentPowerSource != powerSource ? true : false)
	sendEvent(name: "timeRemaining", value: timeLeft, displayed: false)  
	sendEvent(name: "upsStatus", value: data.status.toLowerCase(), displayed: this.currentUpsStatus != data.status ? true : false)
	sendEvent(name: "model", value: data.model, displayed: this.currentModel != data.model ? true : false)
	sendEvent(name: "serial", value: data.serialno, displayed: this.currentSerial != data.serialno ? true : false)
	sendEvent(name: "battery", value: battery, displayed: this.currentBattery != battery ? true : false)
	sendEvent(name: "batteryRuntime", value: data.timeleft, displayed: this.currentRunTimeRemain != data.timeleft ? true : false)
	sendEvent(name: "batteryVoltage", value: batteryVoltage, displayed: this.currentBatteryVoltage != batteryVoltage ? true : false)
	sendEvent(name: "voltage", value: voltage, descriptionText: "Line voltage is ${voltage} volts.", displayed: this.currentVoltage != voltage ? true : false)
	sendEvent(name: "loadPercent", value: loadPercent, displayed: this.currentLoadPercent != loadPercent ? true : false)
	sendEvent(name: "timeOnBattery", value: data.tonbatt, displayed: this.currentTimeOnBattery != data.tonbatt ? true : false)
	sendEvent(name: "lifetimeOnBattery", value: data.cumonbatt, displayed: this.currentLifetimeOnBattery != data.cumonbatt ? true : false)
	sendEvent(name: "lowTransVolts", value: lowTransVolts, displayed: this.currentLowTransVolts != lowTransVolts ? true : false)
	sendEvent(name: "highTransVolts", value: highTransVolts, displayed: this.currentHighTransVolts != highTransVolts ? true : false)
	sendEvent(name: "sensitivity", value: data.sense, displayed: this.currentSensitivity != data.sense ? true : false)
	sendEvent(name: "batteryDate", value: data.battdate, displayed: this.currentBatteryDate != data.battdate ? true : false)
	sendEvent(name: "lastUpdate", value: data.date, displayed: false)
	sendEvent(name: "nomPower", value: nomPower, displayed: this.currentNomPower != nomPower ? true : false)
	sendEvent(name: "power", value: power, displayed: this.currentPower != power ? true : false)
	sendEvent(name: "lastPowerFail", value: data.xonbatt, displayed: this.currentNomPower != data.xonbatt ? true : false)
	sendEvent(name: "lastPowerRestore", value: data.xoffbatt, displayed: this.currentPower != data.xoffbatt ? true : false)
	sendEvent(name: "lastPowerFailReason", value: data.lastxfer, displayed: this.currentLastPowerFailReason != data.lastxfer ? true : false)
}


/*
	logsOff
    
	Disables debug logging.
*/
def logsOff()
{
    log.warn "debug logging disabled..."
	device.updateSetting("enableDebug", [value:"false",type:"bool"])
}