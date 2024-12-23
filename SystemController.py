from ParamManager import ParamManagerImpl
import ReportGen
from utils.Utils import singleton

@singleton
class SystemControllerImpl():
    def __init__(self):
        self.paramManager = ParamManagerImpl()
        
        self.cb = {}
        self.reportGens = {}

    def setCallbacck(self, cbtype, cb):
        self.cb[cbtype] = cb

    def genReport(self, type: ReportGen.ReportType, dataPath: str, **kwargs):
        if type not in self.reportGens:
            self.reportGens[type] = ReportGen.getReportGen(type)
        reportGen = self.reportGens[type]
        return reportGen.genReport(dataPath, **kwargs)
    
    def updateReport(self, type: ReportGen.ReportType, index: int, **kwargs):
        if type not in self.reportGens:
            self.reportGens[type] = ReportGen.getReportGen(type)
        reportGen = self.reportGens[type]
        return reportGen.updateReport(index, **kwargs)
    
    def saveReport(self, type: ReportGen.ReportType):
        if type not in self.reportGens:
            self.reportGens[type] = ReportGen.getReportGen(type)
        reportGen = self.reportGens[type]
        return reportGen.saveReport()
