import ReportGen
from utils.Utils import singleton

@singleton
class SystemControllerImpl():
    def __init__(self):
        self.cb = {}

    def setCallbacck(self, cbtype, cb):
        self.cb[cbtype] = cb

    def genReport(self, type: ReportGen.ReportType, dataPath: str):
        reportGen = ReportGen.getReportGen(type)
        reportGen.genReport(dataPath)
