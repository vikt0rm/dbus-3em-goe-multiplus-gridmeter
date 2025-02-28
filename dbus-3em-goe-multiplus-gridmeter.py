#!/usr/bin/env python
# vim: ts=2 sw=2 et

# import normal packages
import platform 
import logging
import logging.handlers
from logging.handlers import RotatingFileHandler
import os
import sys
if sys.version_info.major == 2:
    import gobject
else:
    from gi.repository import GLib as gobject
import time
import requests # for http GET
import configparser # for config/ini file
from dbus import SessionBus, SystemBus, DBusException
 
# our own packages from victron
sys.path.insert(1, os.path.join(os.path.dirname(__file__), '/opt/victronenergy/dbus-systemcalc-py/ext/velib_python'))
from vedbus import VeDbusService, VeDbusItemImport


class DbusShelly3emService:
  def __init__(self, paths, productname='Gridmeter 3em+goe+mp2', connection='Shelly 3EM HTTP JSON service'):
    config = self._getConfig()
    deviceinstance = int(config['DEFAULT']['DeviceInstance'])
    customname = config['DEFAULT']['CustomName']
    role = config['DEFAULT']['Role']
    self.goeVoltage = None
    self.multiplusVoltage = None

    allowed_roles = ['pvinverter','grid']
    if role in allowed_roles:
        servicename = 'com.victronenergy.' + role
    else:
        logging.error("Configured Role: %s is not in the allowed list")
        exit()

    if role == 'pvinverter':
        productid = 0xA144
    else:
        productid = 45069

    # Connect to the sessionbus. Note that on ccgx we use systembus instead.
    self._dbusConn = SessionBus() if 'DBUS_SESSION_BUS_ADDRESS' in os.environ else SystemBus()

    self._dbusServiceName = "{}.http_{:02d}".format(servicename, deviceinstance)
    self._dbusservice = VeDbusService(self._dbusServiceName, register=False)
    self._paths = paths
 
    logging.debug("%s /DeviceInstance = %d" % (servicename, deviceinstance))
 
    # Create the management objects, as specified in the ccgx dbus-api document
    self._dbusservice.add_path('/Mgmt/ProcessName', __file__)
    self._dbusservice.add_path('/Mgmt/ProcessVersion', 'Unkown version, and running on Python ' + platform.python_version())
    self._dbusservice.add_path('/Mgmt/Connection', connection)
 
    # Create the mandatory objects
    self._dbusservice.add_path('/DeviceInstance', deviceinstance)
    self._dbusservice.add_path('/ProductId', productid)
    self._dbusservice.add_path('/DeviceType', 345) # found on https://www.sascha-curth.de/projekte/005_Color_Control_GX.html#experiment - should be an ET340 Engerie Meter
    self._dbusservice.add_path('/ProductName', productname)
    self._dbusservice.add_path('/CustomName', customname)
    self._dbusservice.add_path('/Latency', None)
    self._dbusservice.add_path('/FirmwareVersion', 0.3)
    self._dbusservice.add_path('/HardwareVersion', 0)
    self._dbusservice.add_path('/Connected', 1)
    self._dbusservice.add_path('/Role', role)
    self._dbusservice.add_path('/Position', self._getShellyPosition()) # normaly only needed for pvinverter
    self._dbusservice.add_path('/Serial', self._getShellySerial())
    self._dbusservice.add_path('/UpdateIndex', 0)
 
    # add path values to dbus
    for path, settings in self._paths.items():
      self._dbusservice.add_path(
        path, settings['initial'], gettextcallback=settings['textformat'], writeable=True, onchangecallback=self._handlechangedvalue)
    
    # register the dbus service
    self._dbusservice.register()

    # last update
    self._lastUpdate = 0
 
    # add _update function 'timer'
    gobject.timeout_add(500, self._update) # pause 500ms before the next request
    
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
 
 
  def _getShellyPosition(self):
    config = self._getConfig()
    value = config['DEFAULT']['Position']
    
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
    meter_r = None
    try:
       meter_r = requests.get(url = URL, timeout=None)
    except:
       logging.warning("Something went wrong during Shelly Status request.")
    
    # check for response
    if not meter_r:
        raise ConnectionError("No response from Shelly 3EM - %s" % (URL))
    
    meter_data = meter_r.json()     
    
    # check for Json
    if not meter_data:
        raise ValueError("Converting response to JSON failed")
    
    
    return meter_data
 
 
  def _signOfLife(self):
    #dbusservice = VeDbusService(self._dbusServiceName)
    logging.info("--- Start: sign of life ---")
    logging.info("Last _update() call: %s" % (self._lastUpdate))
    logging.info("Last '/Ac/Power': %s" % (self._dbusservice['/Ac/Power']))
    logging.info("--- End: sign of life ---")
    return True
  
  def _getCombinedPower(self, shellyPower, goePowerPath, dbusPaths, considerMp2 = False):
    goeDbusPath = "com.victronenergy.evcharger.http_43"
    goePower = 0.0
    goeConnected = goeDbusPath in dbusPaths
    
    if goeConnected:
       goePower = VeDbusItemImport(self._dbusConn, "com.victronenergy.evcharger.http_43", goePowerPath, None, False).get_value()
       if goePower is None:
          goePower = 0.0
          logging.error("goePower is invalid")
    else:
       logging.debug("go-eCharger not connected")
    
    mp2DbusPath = "com.victronenergy.vebus.ttyUSB0"
    multiplusPower = 0.0
    mp2Ready = mp2DbusPath in dbusPaths
    if mp2Ready:
       multiplusPower = VeDbusItemImport(self._dbusConn, "com.victronenergy.vebus.ttyUSB0", '/Ac/ActiveIn/P', None, False).get_value()
       if multiplusPower is None:
          multiplusPower = 0.0
          logging.error("mp2Power is invalid")
    else:
       logging.error("mp2 is not ready yet")

    if not considerMp2:
       multiplusPower = 0
       logging.debug("mp2Power not considered")

    logging.debug("_getCombinedPower (shellyPower): %s (goePower): %s (multiplusPower): %s" % (shellyPower, goePower, multiplusPower))

    return shellyPower + goePower + multiplusPower

  def _getCombinedAmps(self, shellyAmps, goePowerPath, dbusPaths, considerMp2 = False):

    goeAmps = 0.0
    goeDbusPath = "com.victronenergy.evcharger.http_43"
    goeConnected = goeDbusPath in dbusPaths

    if goeConnected:
       goePower = VeDbusItemImport(self._dbusConn, goeDbusPath, goePowerPath, None, False).get_value()
       if not self.goeVoltage:
          self.goeVoltage = VeDbusItemImport(self._dbusConn, goeDbusPath, '/Ac/Voltage', None, False).get_value()
          logging.info("goeVoltage imported: %s" % (self.goeVoltage))
       goeVoltage = self.goeVoltage
       if goePower is None or goeVoltage is None:
          goePower = 0
          logging.error("goe values are not plausible: goePower = %s goeVoltage = %s" % (goePower, goeVoltage))
       elif goeVoltage > 0:
          goeAmps = goePower / goeVoltage
    else:
       logging.debug("go-eCharger not connected")
   
    mp2DbusPath = "com.victronenergy.vebus.ttyUSB0"
    multiplusAmps = 0.0
    mp2Ready = mp2DbusPath in dbusPaths
    if mp2Ready:
       multiplusPower = VeDbusItemImport(self._dbusConn, mp2DbusPath, '/Ac/ActiveIn/L1/P', None, False).get_value()
       if not self.multiplusVoltage:
          self.multiplusVoltage = VeDbusItemImport(self._dbusConn, mp2DbusPath, '/Ac/ActiveIn/L1/V', None, False).get_value()
          logging.info("multiplusVoltage imported: %s" % (self.multiplusVoltage))
       multiplusVoltage = self.multiplusVoltage
       
       if multiplusPower is None or multiplusVoltage is None:
          multiplusAmps = 0
          logging.error("mp2 path does not exist")
       elif multiplusVoltage > 0:
          multiplusAmps = multiplusPower / multiplusVoltage
    else:
       logging.error("mp2 not ready yet")
          
    if not considerMp2:
       multiplusAmps = 0
       logging.debug("mp2Power path not considered")

    logging.debug("_getCombinedAmps (shellyAmps): %s (goeAmps): %s (multiplusAmps): %s" % (shellyAmps, goeAmps, multiplusAmps))

    return shellyAmps + goeAmps + multiplusAmps

  def _getEnergyFromPower(self, power):
     # New Version - from xris99
     #Calc = 60min * 60 sec / 0.500 (refresh interval of 500ms) * 1000
     return power/(60*60/0.5*1000)

  def _update(self): 
    try:
      #get data from Shelly 3em
      meter_data = self._getShellyData()
      config = self._getConfig()

      try:
        remapL1 = int(config['ONPREMISE']['L1Position'])
      except KeyError:
        remapL1 = 1

      if remapL1 > 1:
        old_l1 = meter_data['emeters'][0]
        meter_data['emeters'][0] = meter_data['emeters'][remapL1-1]
        meter_data['emeters'][remapL1-1] = old_l1
 
      dbusPaths = self._dbusConn.list_names()
      #send data to DBus
      self._dbusservice['/Ac/L1/Voltage'] = meter_data['emeters'][0]['voltage']
      self._dbusservice['/Ac/L2/Voltage'] = meter_data['emeters'][1]['voltage']
      self._dbusservice['/Ac/L3/Voltage'] = meter_data['emeters'][2]['voltage']
      self._dbusservice['/Ac/L1/Current'] = self._getCombinedAmps(meter_data['emeters'][0]['current'] * meter_data['emeters'][0]['pf'],  '/Ac/L1/Power', dbusPaths, True)
      self._dbusservice['/Ac/L2/Current'] = self._getCombinedAmps(meter_data['emeters'][1]['current'] * meter_data['emeters'][1]['pf'], '/Ac/L2/Power', dbusPaths)
      self._dbusservice['/Ac/L3/Current'] = self._getCombinedAmps(meter_data['emeters'][2]['current'] * meter_data['emeters'][2]['pf'], '/Ac/L3/Power', dbusPaths)
      self._dbusservice['/Ac/L1/Power'] = self._getCombinedPower(meter_data['emeters'][0]['power'], '/Ac/L1/Power', dbusPaths, True)
      self._dbusservice['/Ac/L2/Power'] = self._getCombinedPower(meter_data['emeters'][1]['power'], '/Ac/L2/Power', dbusPaths)
      self._dbusservice['/Ac/L3/Power'] = self._getCombinedPower(meter_data['emeters'][2]['power'], '/Ac/L3/Power', dbusPaths)
      self._dbusservice['/Ac/Power'] = self._dbusservice['/Ac/L1/Power'] + self._dbusservice['/Ac/L2/Power'] + self._dbusservice['/Ac/L3/Power']
      
      if (self._dbusservice['/Ac/L1/Power'] > 0):
         self._dbusservice['/Ac/L1/Energy/Forward'] = self._dbusservice['/Ac/L1/Energy/Forward'] + self._getEnergyFromPower(self._dbusservice['/Ac/L1/Power']) 
      else:
         self._dbusservice['/Ac/L1/Energy/Reverse'] = self._dbusservice['/Ac/L1/Energy/Reverse'] + self._getEnergyFromPower(self._dbusservice['/Ac/L1/Power']*-1) 
         
      self._dbusservice['/Ac/L2/Energy/Forward'] = (meter_data['emeters'][1]['total']/1000)
      self._dbusservice['/Ac/L3/Energy/Forward'] = (meter_data['emeters'][2]['total']/1000)
      self._dbusservice['/Ac/L2/Energy/Reverse'] = (meter_data['emeters'][1]['total_returned']/1000) 
      self._dbusservice['/Ac/L3/Energy/Reverse'] = (meter_data['emeters'][2]['total_returned']/1000) 
      
      # Old version
      #dbusservice['/Ac/Energy/Forward'] = dbusservice['/Ac/L1/Energy/Forward'] + dbusservice['/Ac/L2/Energy/Forward'] + dbusservice['/Ac/L3/Energy/Forward']
      #dbusservice['/Ac/Energy/Reverse'] = dbusservice['/Ac/L1/Energy/Reverse'] + dbusservice['/Ac/L2/Energy/Reverse'] + dbusservice['/Ac/L3/Energy/Reverse'] 
      
      # New Version - from xris99
      #Calc = 60min * 60 sec / 0.500 (refresh interval of 500ms) * 1000
      if (self._dbusservice['/Ac/Power'] > 0):
           self._dbusservice['/Ac/Energy/Forward'] = self._dbusservice['/Ac/Energy/Forward'] + self._getEnergyFromPower(self._dbusservice['/Ac/Power'])            
      if (self._dbusservice['/Ac/Power'] < 0):
           self._dbusservice['/Ac/Energy/Reverse'] = self._dbusservice['/Ac/Energy/Reverse'] + self._getEnergyFromPower(self._dbusservice['/Ac/Power']*-1)

      
      #logging
      logging.debug("House Consumption (/Ac/Power): %s" % (self._dbusservice['/Ac/Power']))
      logging.debug("House Forward (/Ac/Energy/Forward): %s" % (self._dbusservice['/Ac/Energy/Forward']))
      logging.debug("House Reverse (/Ac/Energy/Revers): %s" % (self._dbusservice['/Ac/Energy/Reverse']))
      logging.debug("---")
      
      # increment UpdateIndex - to show that new data is available an wrap
      self._dbusservice['/UpdateIndex'] = (self._dbusservice['/UpdateIndex'] + 1 ) % 256

      #update lastupdate vars
      self._lastUpdate = time.time()
    except (ValueError, requests.exceptions.ConnectionError, requests.exceptions.Timeout, ConnectionError) as e:
       logging.critical('Error getting data from Shelly - check network or Shelly status. Setting power values to 0. Details: %s', e, exc_info=e)       
       self._dbusservice['/Ac/L1/Power'] = 0                                       
       self._dbusservice['/Ac/L2/Power'] = 0                                       
       self._dbusservice['/Ac/L3/Power'] = 0
       self._dbusservice['/Ac/Power'] = 0
       self._dbusservice['/UpdateIndex'] = (self._dbusservice['/UpdateIndex'] + 1 ) % 256        
    except Exception as e:
       logging.critical('Error at %s', '_update', exc_info=e)
       
    # return true, otherwise add_timeout will be removed from GObject - see docs http://library.isr.ist.utl.pt/docs/pygtk2reference/gobject-functions.html#function-gobject--timeout-add
    return True
 
  def _handlechangedvalue(self, path, value):
    logging.debug("someone else updated %s to %s" % (path, value))
    return True # accept the change




def getLogLevel():
  config = configparser.ConfigParser()
  config.read("%s/config.ini" % (os.path.dirname(os.path.realpath(__file__))))
  logLevelString = config['DEFAULT']['LogLevel']
  
  if logLevelString:
    level = logging.getLevelName(logLevelString)
  else:
    level = logging.INFO
    
  return level


def main():
  #configure logging
  logging.basicConfig(      format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S',
                            level=getLogLevel(),
                            handlers=[
                                RotatingFileHandler("%s/current.log" % (os.path.dirname(os.path.realpath(__file__))), maxBytes=10000),
                                logging.StreamHandler()
                            ])
 
  try:
      logging.info("Start")
  
      from dbus.mainloop.glib import DBusGMainLoop
      # Have a mainloop, so we can send/receive asynchronous calls to and from dbus
      DBusGMainLoop(set_as_default=True)
     
      #formatting 
      _kwh = lambda p, v: (str(round(v, 2)) + ' kWh')
      _a = lambda p, v: (str(round(v, 1)) + ' A')
      _w = lambda p, v: (str(round(v, 1)) + ' W')
      _v = lambda p, v: (str(round(v, 1)) + ' V')   
     
      #start our main-service
      pvac_output = DbusShelly3emService(
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
  except (ValueError, requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
    logging.critical('Error in main type %s', str(e))
  except Exception as e:
    logging.critical('Error at %s', 'main', exc_info=e)
if __name__ == "__main__":
  main()
