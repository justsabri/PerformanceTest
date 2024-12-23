from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

def singleton(cls):
    instances = {}

    def get_instance(*args, **kwargs):
        if cls not in instances:
            instances[cls] = cls(*args, **kwargs)
        return instances[cls]

    return get_instance

class Params():
    def __init__(self):
        pass

class DeviceInfo():
    def __init__(self):
        self.ports = {}

class CSData():
    __slots__ = ['trialNo', 'startTime', 'endTime', 'MELoad', 'windDirection', 'windSpeed','depth','rpm','shaftPower','heading', 'initialHeading', 'finalHeading','period', 'seq', 'distance','speed', 'initialSpeed', 'finalSpeed', 'df', 'traceFigName', 'speedFigName']
    def __init__(self):
        self.trialNo = ''
        self.startTime = 0
        self.MELoad = 0
        self.windDirection = 0
        self.windSpeed = 0
        self.depth = 0
        self.rpm = 0
        self.shaftPower = 0
        self.heading = 0
        self.period = 0
        self.distance = 0
        self.speed = 0

class FigureManager:
    """管理多个 Figure 和 FigureCanvas 的类"""
    def __init__(self):
        self.figures = {}  # 存储 {name: (Figure, FigureCanvas)} 的字典
        self.seqToName = {} # 存储工况数据顺序和图名字 {seq： [name1, name2]}

    def add_figure(self, seq, name, figure, canvas):
        """添加新的 Figure 和 Canvas"""
        print(seq, name)
        self.figures[name] = (figure, canvas)
        if seq not in self.seqToName:
            self.seqToName[seq] = []
        self.seqToName[seq].append(name)

    def get_canvas(self, name):
        """根据名称获取 FigureCanvas"""
        return self.figures[name][1] if name in self.figures else None
    
    def get_canvas_by_index(self, index):
        """根据索引获取 FigureCanvas"""
        if index > len(self.figures) - 1:
            return None

        figNames = self.seqToName[index]
        print('by id', figNames)
        figs = []
        for name in figNames:
            figs.append(self.get_canvas(name))

        print(len(figs))
        return figs

    def get_figure(self, name):
        """根据名称获取 Figure"""
        return self.figures[name][0] if name in self.figures else None

    def update_figure(self, name, **kwargs):
        figure = self.get_figure(name)
        if figure:
            ax = figure.axes[0]  # 获取第一个子图
            # 设置坐标轴
            left = kwargs.get('left')
            right = kwargs.get('right')
            top = kwargs.get('top')
            bottom = kwargs.get('bottom')
            ax.set_xlim(left, right)
            ax.set_ylim(bottom, top)

            # 更新其他参数

            # 刷新显示
            canvas = self.get_canvas(name)
            canvas.draw()

    def remove_figure(self, name):
        """移除指定的 Figure 和 Canvas"""
        if name in self.figures:
            del self.figures[name]
        for seq, names in enumerate(self.seqToName):
            if name in names:
                names.remove(name)
                break

    def save_figure(self, name, path):
        fig = self.get_figure(name)
        if fig:
            fig.savefig(path, bbox_inches='tight')

import wmi

def check_network_adapter_status():
    c = wmi.WMI()
    for adapter in c.Win32_NetworkAdapter():
        # 打印网络适配器的名称和连接状态
        print(f"名称: {adapter.Name}")
        print(f"MAC 地址: {adapter.MACAddress}")
        print(f"连接状态: {adapter.NetConnectionStatus}")
        print(f"网线是否已插入: {'是' if adapter.NetConnectionStatus == 2 else '否'}")
        print('-' * 40)