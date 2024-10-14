import SystemController
from SystemController import SystemControllerImpl
import ReportGen

ctl = SystemControllerImpl()
ctl.genReport(ReportGen.ReportType.CS_REPORT, 'd:\\')
ctl.updateReport(ReportGen.ReportType.CS_REPORT, 0, xLimit = "10", yLimit = '10')
ctl.saveReport(ReportGen.ReportType.CS_REPORT)
