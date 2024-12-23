from DeviceManager import DeviceManagerImpl
from utils.Utils import Params

        
class ParamManagerImpl():
    def __init__(self):
        self.deviceManager = DeviceManagerImpl()

    def getDeviceInfo(self):
        return self.deviceManager.getDeviceInfo()

    def config(self, param: Params):
        self.deviceManager.config(param)
    
    def startDataFlow(self):
        self.deviceManager.startDataFlow()
    