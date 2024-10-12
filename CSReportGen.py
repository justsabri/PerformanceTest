
from ReportGen import RepoortGenBase
class CSReportGenImpl(RepoortGenBase):
    def __init__(self):
        super().__init__()
    
    def genReport(self, dataPath):
        self.logger.debug("gen cs report")
        self.analyzer.handleCsData(dataPath)
        reportFile = f"{self.REPORT_PATH}csReport.doc"
        self.logger.debug("savr cs report to " + reportFile)
