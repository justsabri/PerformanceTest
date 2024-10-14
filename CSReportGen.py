
from ReportGen import RepoortGenBase
class CSReportGenImpl(RepoortGenBase):
    def __init__(self):
        super().__init__()
    
    def genReport(self, dataPath:str, **kwargs):
        self.logger.debug(f"gen cs report from: {dataPath}")
        for field, value in kwargs.items():
            self.logger.info(f"Field: {field}, Value: {value}")
        x,y = self.analyzer.handleCsData(dataPath)
        # 完成图表数据处理，返回图表数据


    def updateReport(self, index: int, **kwargs):
        self.logger.info(f"update report: {index} {kwargs}")
        for field, value in kwargs.items():
            self.logger.info(f"Field: {field}, Value: {value}")

    def saveReport(self):
        reportFile = f"{self.REPORT_PATH}csReport.doc"
        self.logger.debug("savr cs report to " + reportFile)
        return reportFile
