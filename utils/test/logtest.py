
import sys,os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import Logger

 
logger = Logger.GetLogger(__name__)

class LOGTest:
    def __init__(self) -> None:
        self.llogger = Logger.GetLogger(__name__)

    def printLog(self,i):
        self.llogger.debug(i)

logger.debug('This is a debug message')
logg = LOGTest()
logg.printLog(1)

for i in range(100):
    logg.printLog(i)
