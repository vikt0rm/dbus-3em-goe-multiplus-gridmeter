#!/usr/bin/env python
 
# import normal packages
import platform 
import logging
import sys
import os
import sys
if sys.version_info.major == 2:
    import gobject
else:
    from gi.repository import GLib as gobject
import sys
import time
import requests # for http GET
import configparser # for config/ini file
 
# our own packages from victron
sys.path.insert(1, os.path.join(os.path.dirname(__file__), '/opt/victronenergy/dbus-systemcalc-py/ext/velib_python'))
from vedbus import VeDbusService


class DbusShelly3emService:
  def __init__(self, servicename, deviceinstance, paths, productname='Shelly 3EM', connection='Shelly 3EM HTTP JSON service'):
    self._dbusservice = VeDbusService("{}.http_{:02d}".format(servicename, deviceinstance))
    self._paths = paths
 
    logging.debug("%s /DeviceInstance = %d" % (servicename, deviceinstance))
 
    # Create the management objects, as specified in the ccgx dbus-api document
    self._dbusservice.add_path('/Mgmt/ProcessName', __file__)
    self._dbusservice.add_path('/Mgmt/ProcessVersion', 'Unkown version, and running on Python ' + platform.python_version())
    self._dbusservice.add_path('/Mgmt/Connection', connection)
 
    # Create the mandatory objects
    self._dbusservice.add_path('/DeviceInstance', deviceinstance)
    #self._dbusservice.add_path('/ProductId', 16) # value used in ac_sensor_bridge.cpp of dbus-cgwacs
    #self._dbusservice.add_path('/ProductId', 0xFFFF) # id assigned by Victron Support from SDM630v2.py
    self._dbusservice.add_path('/ProductId', 45069) # found on https://www.sascha-curth.de/projekte/005_Color_Control_GX.html#experiment - should be an ET340 Engerie Meter
    self._dbusservice.add_path('/DeviceType', 345) # found on https://www.sascha-curth.de/projekte/005_Color_Control_GX.html#experiment - should be an ET340 Engerie Meter
    self._dbusservice.add_path('/ProductName', productname)
    self._dbusservice.add_path('/CustomName', productname)    
    self._dbusservice.add_path('/Latency', None)    
    self._dbusservice.add_path('/FirmwareVersion', 0.1)
    self._dbusservice.add_path('/HardwareVersion', 0)
    self._dbusservice.add_path('/Connected', 1)
    self._dbusservice.add_path('/Role', 'grid')
    self._dbusservice.add_path('/Position', 0) # normaly only needed for pvinverter
    self._dbusservice.add_path('/Serial', self._getShellySerial())
    self._dbusservice.add_path('/UpdateIndex', 0)
    self._dbusservice.add_path('/Debug', "")
 
    # add path values to dbus
    for path, settings in self._paths.items():
      self._dbusservice.add_path(
        path, settings['initial'], gettextcallback=settings['textformat'], writeable=True, onchangecallback=self._handlechangedvalue)
 
    # last update
    self._lastUpdate = 0
    
    # correct metering
    self._lastMeterSumIn = 0;
    self._lastMeterSumOut = 0;
    self._lastResultIn = 0;
    self._lastResultOut = 0;
    self._lastMeterCounter = 0;
    self._lastMeterTime = time.time()
 
    # add _update function 'timer'
    gobject.timeout_add(1000, self._update) # pause 1000ms before the next request
    
    # add _signOfLife 'timer' to get feedback in log every 5minutes
    gobject.timeout_add(self._getSignOfLifeInterval()*60*1000, self._signOfLife)
 
  def _getShellySerial(self):
    meter_data = self._getShellyData()  
    
    if not meter_data['mac']:
        raise ValueError("Response does not contain 'mac' attribute")
    
    serial = meter_data['mac']
    return serial
 
 
  def _getConfig(self):
    config = configparser.ConfigParser()
    config.read("%s/config.ini" % (os.path.dirname(os.path.realpath(__file__))))
    return config;
 
 
  def _getSignOfLifeInterval(self):
    config = self._getConfig()
    value = config['DEFAULT']['SignOfLifeLog']
    
    if not value: 
        value = 0
    
    return int(value)
  
  
  def _getShellyStatusUrl(self):
    config = self._getConfig()
    accessType = config['DEFAULT']['AccessType']
    
    if accessType == 'OnPremise': 
        URL = "http://%s:%s@%s/status" % (config['ONPREMISE']['Username'], config['ONPREMISE']['Password'], config['ONPREMISE']['Host'])
        URL = URL.replace(":@", "")
    else:
        raise ValueError("AccessType %s is not supported" % (config['DEFAULT']['AccessType']))
    
    return URL
    
 
  def _getShellyData(self):
    URL = self._getShellyStatusUrl()
    meter_r = requests.get(url = URL)
    
    # check for response
    if not meter_r:
        raise ConnectionError("No response from Shelly 3EM - %s" % (URL))
    
    meter_data = meter_r.json()     
    
    # check for Json
    if not meter_data:
        raise ValueError("Converting response to JSON failed")
    
    
    return meter_data
  
  def _getGoeChargerStatusUrl(self):
    config = self._getConfig()
    accessType = config['DEFAULT']['AccessType']
    
    if accessType == 'OnPremise': 
        URL = "http://%s/status" % (config['ONPREMISE']['GoEHost'])
    else:
        raise ValueError("AccessType %s is not supported" % (config['DEFAULT']['AccessType']))
    
    return URL
  
  
  def _getGoeChargerData(self):
    URL = self._getGoeChargerStatusUrl()
    request_data = requests.get(url = URL)
    
    # check for response
    if not request_data:
        raise ConnectionError("No response from go-eCharger - %s" % (URL))
    
    json_data = request_data.json()     
    
    # check for Json
    if not json_data:
        raise ValueError("Converting response to JSON failed")
      
    return json_data
  
 
  def _signOfLife(self):
    logging.info("--- Start: sign of life ---")
    logging.info("Last _update() call: %s" % (self._lastUpdate))
    logging.info("Last '/Ac/Power': %s" % (self._dbusservice['/Ac/Power']))
    logging.info("--- End: sign of life ---")
    return True
 
 
  def _getBalancingMeterResult(self, phaseSumIn, phaseSumOut):
    resultIn = 0
    resultOut = 0
    debugMsg = "(%s) input sum %s/%s" % (self._lastMeterCounter,phaseSumIn,phaseSumOut)
    
    #check config
    config = self._getConfig()
    useBalancing = int(config['DEFAULT']['GridMeterBalancing'])
    
    if self._lastMeterCounter != 0 and useBalancing == 1:
        #calc diff values from last value to now
        #new values should be higher than current
        diffIn = phaseSumIn - self._lastMeterSumIn
        diffOut = phaseSumOut - self._lastMeterSumOut        
        debugMsg += " - last-meter-diff %s/%s" % (diffIn, diffOut)
        
        #resultIn
        ###resultIn = self._lastMeterSumIn + diffIn
        resultIn = self._lastResultIn + diffIn
        if diffOut >= diffIn:
            resultIn -= diffIn
            
        #resultOut
        ###resultOut = self._lastMeterSumOut + diffOut
        resultOut = self._lastResultOut + diffOut
        if diffOut >= diffIn:
            resultOut -= diffIn
            
        #next run values
        #on next run we the result of this run - so the base is the same
        ###self._lastMeterSumIn = resultIn
        ###self._lastMeterSumOut = resultOut
        
        #debug, debug, debug....puhh
        diffResultIn = resultIn - self._lastResultIn
        diffResultOut = resultOut - self._lastResultOut
        debugMsg += " - result %s/%s (last-result-diff %s/%s)" %(resultIn, resultOut, diffResultIn, diffResultOut)        
    else:
        #default case - lets take input for output
        #will help for next run
        resultIn = phaseSumIn
        resultOut = phaseSumOut
        
        #next run values
        #on first run we take the input - so we use this for next run as a basis
        #but on next run me must take the result - so on "3rd" run we take "result" as basis
        ###self._lastMeterSumIn = phaseSumIn
        ###self._lastMeterSumOut = phaseSumOut
        
        #debug output - identify else case
        debugMsg += " - else-case result=input"
    
    # set last values
    self._lastMeterSumIn = phaseSumIn
    self._lastMeterSumOut = phaseSumOut
    self._lastResultIn = resultIn
    self._lastResultOut = resultOut
    
    # raise counter
    self._lastMeterCounter += 1
        
    # logging
    logging.info(debugMsg)
    
    return resultIn, resultOut, debugMsg
    
    
  def _update(self):   
    try:
       #get data from Shelly 3em
       meter_data = self._getShellyData()
       goe_data = self._getGoeChargerData()
       
       #send data to DBus
       self._dbusservice['/Ac/Power'] = meter_data['total_power'] # positive: consumption, negative: feed into grid
       self._dbusservice['/Ac/L1/Voltage'] = meter_data['emeters'][0]['voltage']
       self._dbusservice['/Ac/L2/Voltage'] = meter_data['emeters'][1]['voltage']
       self._dbusservice['/Ac/L3/Voltage'] = meter_data['emeters'][2]['voltage']
       self._dbusservice['/Ac/L1/Current'] = meter_data['emeters'][0]['current'] + int(goe_data['nrg'][4] * 0.1)
       self._dbusservice['/Ac/L2/Current'] = meter_data['emeters'][1]['current'] + int(goe_data['nrg'][5] * 0.1)
       self._dbusservice['/Ac/L3/Current'] = meter_data['emeters'][2]['current'] + int(goe_data['nrg'][6] * 0.1)
       self._dbusservice['/Ac/L1/Power'] = meter_data['emeters'][0]['power'] + int(goe_data['nrg'][7] * 0.1 * 1000)
       self._dbusservice['/Ac/L2/Power'] = meter_data['emeters'][1]['power'] + int(goe_data['nrg'][8] * 0.1 * 1000)
       self._dbusservice['/Ac/L3/Power'] = meter_data['emeters'][2]['power'] + int(goe_data['nrg'][9] * 0.1 * 1000)
       self._dbusservice['/Ac/L1/Energy/Forward'] = (meter_data['emeters'][0]['total']/1000)
       self._dbusservice['/Ac/L2/Energy/Forward'] = (meter_data['emeters'][1]['total']/1000)
       self._dbusservice['/Ac/L3/Energy/Forward'] = (meter_data['emeters'][2]['total']/1000)
       self._dbusservice['/Ac/L1/Energy/Reverse'] = (meter_data['emeters'][0]['total_returned']/1000)
       self._dbusservice['/Ac/L2/Energy/Reverse'] = (meter_data['emeters'][1]['total_returned']/1000)
       self._dbusservice['/Ac/L3/Energy/Reverse'] = (meter_data['emeters'][2]['total_returned']/1000)
       #self._dbusservice['/Ac/Energy/Forward'] = self._dbusservice['/Ac/L1/Energy/Forward'] + self._dbusservice['/Ac/L2/Energy/Forward'] + self._dbusservice['/Ac/L3/Energy/Forward']
       #self._dbusservice['/Ac/Energy/Reverse'] = self._dbusservice['/Ac/L1/Energy/Reverse'] + self._dbusservice['/Ac/L2/Energy/Reverse'] + self._dbusservice['/Ac/L3/Energy/Reverse'] 
       
       # only update Energy on first run or every 3min.
       if (self._lastMeterCounter == 0) or ((time.time() - self._lastMeterTime) > 180):
            _totalForward = self._dbusservice['/Ac/L1/Energy/Forward'] + self._dbusservice['/Ac/L2/Energy/Forward'] + self._dbusservice['/Ac/L3/Energy/Forward'] + int(float(goe_data['eto']) / 10.0)
            _totalReverse = self._dbusservice['/Ac/L1/Energy/Reverse'] + self._dbusservice['/Ac/L2/Energy/Reverse'] + self._dbusservice['/Ac/L3/Energy/Reverse']
            self._dbusservice['/Ac/Energy/Forward'], self._dbusservice['/Ac/Energy/Reverse'], self._dbusservice['/Debug'] = self._getBalancingMeterResult(_totalForward, _totalReverse)
            self._lastMeterTime = time.time()
            
       
       #logging
       logging.debug("House Consumption (/Ac/Power): %s" % (self._dbusservice['/Ac/Power']))
       logging.debug("House Forward (/Ac/Energy/Forward): %s" % (self._dbusservice['/Ac/Energy/Forward']))
       logging.debug("House Reverse (/Ac/Energy/Revers): %s" % (self._dbusservice['/Ac/Energy/Reverse']))
       logging.debug("---");
       
       # increment UpdateIndex - to show that new data is available
       index = self._dbusservice['/UpdateIndex'] + 1  # increment index
       if index > 255:   # maximum value of the index
         index = 0       # overflow from 255 to 0
       self._dbusservice['/UpdateIndex'] = index

       #update lastupdate vars
       self._lastUpdate = time.time()              
    except Exception as e:
       logging.critical('Error at %s', '_update', exc_info=e)
       
    # return true, otherwise add_timeout will be removed from GObject - see docs http://library.isr.ist.utl.pt/docs/pygtk2reference/gobject-functions.html#function-gobject--timeout-add
    return True
 
  def _handlechangedvalue(self, path, value):
    logging.debug("someone else updated %s to %s" % (path, value))
    return True # accept the change
 


def main():
  #configure logging
  logging.basicConfig(      format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S',
                            level=logging.INFO,
                            handlers=[
                                logging.FileHandler("%s/current.log" % (os.path.dirname(os.path.realpath(__file__)))),
                                logging.StreamHandler()
                            ])
 
  try:
      logging.info("Start");
  
      from dbus.mainloop.glib import DBusGMainLoop
      # Have a mainloop, so we can send/receive asynchronous calls to and from dbus
      DBusGMainLoop(set_as_default=True)
     
      #formatting 
      _kwh = lambda p, v: (str(round(v, 2)) + ' KWh')
      _a = lambda p, v: (str(round(v, 1)) + ' A')
      _w = lambda p, v: (str(round(v, 1)) + ' W')
      _v = lambda p, v: (str(round(v, 1)) + ' V')   
     
      #start our main-service
      pvac_output = DbusShelly3emService(
        servicename='com.victronenergy.grid',
        deviceinstance=40,
        paths={
          '/Ac/Energy/Forward': {'initial': 0, 'textformat': _kwh}, # energy bought from the grid
          '/Ac/Energy/Reverse': {'initial': 0, 'textformat': _kwh}, # energy sold to the grid
          '/Ac/Power': {'initial': 0, 'textformat': _w},
          
          '/Ac/Current': {'initial': 0, 'textformat': _a},
          '/Ac/Voltage': {'initial': 0, 'textformat': _v},
          
          '/Ac/L1/Voltage': {'initial': 0, 'textformat': _v},
          '/Ac/L2/Voltage': {'initial': 0, 'textformat': _v},
          '/Ac/L3/Voltage': {'initial': 0, 'textformat': _v},
          '/Ac/L1/Current': {'initial': 0, 'textformat': _a},
          '/Ac/L2/Current': {'initial': 0, 'textformat': _a},
          '/Ac/L3/Current': {'initial': 0, 'textformat': _a},
          '/Ac/L1/Power': {'initial': 0, 'textformat': _w},
          '/Ac/L2/Power': {'initial': 0, 'textformat': _w},
          '/Ac/L3/Power': {'initial': 0, 'textformat': _w},
          '/Ac/L1/Energy/Forward': {'initial': 0, 'textformat': _kwh},
          '/Ac/L2/Energy/Forward': {'initial': 0, 'textformat': _kwh},
          '/Ac/L3/Energy/Forward': {'initial': 0, 'textformat': _kwh},
          '/Ac/L1/Energy/Reverse': {'initial': 0, 'textformat': _kwh},
          '/Ac/L2/Energy/Reverse': {'initial': 0, 'textformat': _kwh},
          '/Ac/L3/Energy/Reverse': {'initial': 0, 'textformat': _kwh},
        })
     
      logging.info('Connected to dbus, and switching over to gobject.MainLoop() (= event based)')
      mainloop = gobject.MainLoop()
      mainloop.run()            
  except Exception as e:
    logging.critical('Error at %s', 'main', exc_info=e)
if __name__ == "__main__":
  main()
