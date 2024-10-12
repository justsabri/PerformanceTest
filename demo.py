import SystemController
from SystemController import SystemControllerImpl
import ReportGen

ctl = SystemControllerImpl()
ctl.genReport(ReportGen.ReportType.CS_REPORT, 'd:\\')

