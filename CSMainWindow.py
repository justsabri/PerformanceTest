from UI import Ui_cs_enter
from UI.VoyageItem import VoyageItem
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QSizePolicy
from SystemController import SystemControllerImpl
import ReportGen
import sys

class CSMainWindow(QMainWindow, Ui_cs_enter.Ui_MainWindow):
    def __init__(self, parent = None):
        super(CSMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.vIndex = 0
        self.controller = SystemControllerImpl()
        self.figManager = None

        self.addButton.clicked.connect(self.onAddButtonClicked)
        self.deleteButton.clicked.connect(self.onDeleteButtonClicked)
        self.startAnalyse.clicked.connect(self.startCSAnalyse)
        self.genReport.clicked.connect(self.genCSReport)

        self.vesselName.setPlaceholderText('船舶名称')
        self.fwdDraught.setPlaceholderText('单位: m')
        self.weather.setPlaceholderText('e.g. CLOUDY')
        self.airTemperature.setPlaceholderText('单位：℃')
        self.hullNo.setPlaceholderText('船舶编号')
        self.midDraught.setPlaceholderText('单位: m')
        self.waveScale.setPlaceholderText('单位: Class')
        self.airPressure.setPlaceholderText('单位: mbar')
        self.date.setPlaceholderText('e.g. 2024.04.27~28')
        self.aftDraught.setPlaceholderText('单位: m')
        self.windDirection.setPlaceholderText('单位：°')
        self.waterTemperature.setPlaceholderText('单位：℃')
        self.seaArea.setPlaceholderText('e.g. Yellow Sea')
        self.displacement.setPlaceholderText('单位: m3')
        self.windScale.setPlaceholderText('单位: beaufoot')
        self.waterDensity.setPlaceholderText('单位: t/m3')


    def onAddButtonClicked(self):
        self.vIndex += 1
        v_item = VoyageItem()
        v_item.setIndex(self.vIndex)
        v_item.updateTitle()
        v_item.setCb(self.showFigs)
        self.boxVLayout.addWidget(v_item)
        self.updateVoyageCount()
    
    def onDeleteButtonClicked(self):
        # 删除 QGroupBox 中的最后一个widget（如果有的话）
        if self.boxVLayout.count() > 0:
            last_widget = self.boxVLayout.itemAt(self.boxVLayout.count() - 1).widget()
            self.boxVLayout.removeWidget(last_widget)
            self.vIndex -= 1
            last_widget.deleteLater()  # 删除控件
            self.updateVoyageCount()

    def packParams(self):
        voyageData = []
        for i in range(0, self.boxVLayout.count()):
            item = self.boxVLayout.itemAt(i).widget()
            param = item.getParams()
            if param is None:
                return None
            voyageData.append(item.getParams())
        
        vesselName = self.checkParam(self.vesselName.text())
        fwdDraught = self.checkParam(self.fwdDraught.text())
        weather = self.checkParam(self.weather.text())
        airTemperature = self.checkParam(self.airTemperature.text())
        hullNo = self.checkParam(self.hullNo.text())
        midDraught = self.checkParam(self.midDraught.text())
        waveScale = self.checkParam(self.waveScale.text())
        airPressure = self.checkParam(self.airPressure.text())
        date = self.checkParam(self.date.text())
        aftDraught = self.checkParam(self.aftDraught.text())
        windDirection = self.checkParam(self.windDirection.text())
        waterTemperature = self.checkParam(self.waterTemperature.text())
        seaArea = self.checkParam(self.seaArea.text())
        displacement = self.checkParam(self.displacement.text())
        windScale = self.checkParam(self.windScale.text())
        waterDensity = self.checkParam(self.waterDensity.text())
        shipAndEnvirnmentParam = {
            'voyageData': voyageData,
            'vesselName': vesselName, 'fwdDraught': fwdDraught, 'weather': weather, 'airTemperature': airTemperature,
            'hullNo': hullNo, 'midDraught': midDraught, 'waveScale': waveScale, 'airPressure': airPressure,
            'date': date, 'aftDraught':aftDraught, 'windDirection': windDirection, 'waterTemperature':waterTemperature,
            'seaArea': seaArea, 'displacement': displacement, 'windScale': windScale, 'waterDensity': waterDensity
        }
        # print(shipAndEnvirnmentParam)
        return shipAndEnvirnmentParam
    
    def checkParam(self, param):
        if not param:
            param = 'unknown'
        return param
    
    def updateVoyageCount(self):
        # print(self.vIndex)
        """更新显示的 Widget 数量。"""
        self.count_label.setText(f"当前工况数: {self.vIndex}")

    def startCSAnalyse(self):
        shipAndEnvirnmentParam = self.packParams()
        if shipAndEnvirnmentParam is None:
            return
        self.figManager = self.controller.genReport(ReportGen.ReportType.CS_REPORT, 'd:\\', **shipAndEnvirnmentParam)
        for i in range(self.boxVLayout.count()):
            voyage = self.boxVLayout.itemAt(i).widget()
            voyage.setResultOptional()

    def showFigs(self, index):
        """
        清空所有子小部件
        """
        # print(index, self.figLayout.count())
        if self.figLayout is not None:
            while self.figLayout.count():
                child = self.figLayout.takeAt(0)
                widget = child.widget()
                if widget:
                    widget.setParent(None)  # 从布局中移除，但不销毁部件
                # if child.widget():
                #     child.widget().deleteLater()  # 删除小部件
        
        # print('222', self.figLayout.count())
        if not self.figManager:
            print('do not analyse first')
            # 弹窗提示先分析

            return
        
        canvas = self.figManager.get_canvas_by_index(index - 1)
        for canva in canvas:
            canva.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.figLayout.addWidget(canva)

        # print('333', self.figLayout.count())

    def genCSReport(self):
        self.controller.saveReport(ReportGen.ReportType.CS_REPORT)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = CSMainWindow()
    myWin.show()
    sys.exit(app.exec_()) 