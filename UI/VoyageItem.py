from .Ui_voyage_item import Ui_Form
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QSizePolicy, QFileDialog, QMessageBox


class VoyageItem(Ui_Form, QtWidgets.QWidget):
    def __init__(self,parent =None):
        super(VoyageItem, self).__init__(parent)
        self.setupUi(self)

        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # 设置高度为固定
        self.resultButton.setVisible(False)
        self.resultButton.clicked.connect(self.showResult)
        self.pushButton.clicked.connect(self.openFiles)

        self.filePath.setPlaceholderText("请输入文件路径")
        self.load.setPlaceholderText("e.g. 75%")
        self.windDirection.setPlaceholderText('单位：°')
        self.windSpeed.setPlaceholderText('单位: m/s')
        self.waterDepth.setPlaceholderText('单位: m')

        self.title = '工况'
        self.index = 1
        self.cb = None
    
    def updateTitle(self):
        self.label.setText(self.title + str(self.index))

    def setIndex(self, i):
        self.index = i
    
    def setCb(self, func):
        self.cb = func

    def openFiles(self):
        file_path, _ = QFileDialog.getOpenFileName(None, "选择工况文件", "", "All Files (*);;Text Files (*.txt)")
        self.filePath.setText(file_path)

    def getParams(self):
        load = self.load.text()
        if not load:
            self.showErrorBox(self.title + str(self.index) + 'MELoad不能为空')
            return None
        if load[-1:] != '%':
            load = load + '%'
        params = {
            'MELoad': load,
            'windSpeed': self.windSpeed.text(),
            'windOri': self.windDirection.text(),
            'depth': self.waterDepth.text(),            
            'dataPath': self.filePath.text()
        }
        return params
    
    def showErrorBox(self, msg):
        # 创建错误框
        error_box = QMessageBox()
        error_box.setIcon(QMessageBox.Critical)  # 设置为错误类型
        error_box.setWindowTitle("错误")
        error_box.setText(msg)
        error_box.setInformativeText("更多详细信息：")
        error_box.setStandardButtons(QMessageBox.Ok)  # 设置按钮
        error_box.exec_()

    def setResultOptional(self):
        self.resultButton.setVisible(True)
    
    def showResult(self):
        if not self.cb:
            return None
        return self.cb(self.index)
        