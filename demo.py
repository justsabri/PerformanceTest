import SystemController
from SystemController import SystemControllerImpl
import ReportGen

voyageData = [
    {'MELoad': '100%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_192347_100%-1\testdata.csv'},
    {'MELoad': '100%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_195154_100%-2\testdata.csv'},
    {'MELoad': '75%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_211152_75%-1\testdata.csv'},
    {'MELoad': '75%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_213744_75%-2\testdata.csv'},
    {'MELoad': '75%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_215955_75%-3\testdata.csv'},
    {'MELoad': '75%', 'windSpeed': '3', 'windOri': '16', 'depth': '50', 'dataPath': r'E:\深海\实船性能测试\试验数据\试验数据\2024_0427_222314_75%-4\testdata.csv'}
]

shipAndEnvirnmentParam = {
    'voyageData': voyageData,
    'vesselName': 'Celsius Essen', 'fwdDraught': 3.364, 'weather': 'CLOUDY', 'airTemperature': 10,
    'hullNo': 'JL0311(C)', 'midDraught': 5.6, 'waveScale': 3, 'airPressure': 1014,
    'date': '2024.04.27~28', 'aftDraught':8.26, 'windDirection': 'NE', 'waterTemperature':8,
    'seaArea': 'Yellow China Sea', 'displacement': 22166, 'windScale': 3, 'waterDensity': 1.023
}


ctl = SystemControllerImpl()
ctl.genReport(ReportGen.ReportType.CS_REPORT, 'd:\\', **shipAndEnvirnmentParam)
ctl.updateReport(ReportGen.ReportType.CS_REPORT, 0, xLimit = "10", yLimit = '10')
ctl.saveReport(ReportGen.ReportType.CS_REPORT)
