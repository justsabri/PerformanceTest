
from utils import Logger
from enum import Enum, auto
from abc import ABC, abstractmethod
from DataAnalyzer import DataAnalyzer
import sys,os

class ReportType(Enum):
    CS_REPORT = auto()


def getReportGen(type: ReportType):
    if not isinstance(type, ReportType):
        raise TypeError(f"Expected an int, but got {type(type).__name__}")
    
    if type == ReportType.CS_REPORT:
        from CSReportGen import CSReportGenImpl
        return CSReportGenImpl()
    
    return RepoortGenBase()
    

class RepoortGenBase(ABC):
    def __init__(self):
        self.logger = Logger.GetLogger(__name__)
        self.REPORT_PATH = f"{sys.path[0]}\\reports\\"
        if not os.path.exists(self.REPORT_PATH):
            os.mkdir(self.REPORT_PATH)
        self.logger.debug('report path is ' + self.REPORT_PATH)
        #print(os.path.dirname(os.path.abspath(__file__)))
        self.analyzer = DataAnalyzer()

    @abstractmethod
    def genReport(self, dataPath):
        self.logger.info("default method, should not run this func")
    