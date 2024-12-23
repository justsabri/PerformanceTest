from utils.Utils import Params
from utils.Utils import DeviceInfo
import serial.tools.list_ports

class DeviceManagerImpl():
    def __init__(self):
        pass

    def getDeviceInfo(self):
        deviceInfo = DeviceInfo()

        #获取所有串口
        deviceInfo.ports = serial.tools.list_ports.comports()

        #获取网口

        
        #获取其它软件采集端口能力
        #deviceInfo.xxx = 

        return deviceInfo

    def config(self, param: Params):
        pass

    def startDataFlow(self):
        pass