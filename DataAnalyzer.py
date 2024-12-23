from utils import Logger
from utils.Utils import singleton
import pandas as pd
from utils.Utils import CSData
import numpy as np
from scipy import stats

@singleton
class DataAnalyzer():
    def __init__(self):
        self.logger = Logger.GetLogger(__name__)

    def handleCsData(self, dataPath: str):
        self.logger.debug("datapath is " + dataPath)
        csData = CSData()
        df = pd.read_csv(dataPath)
        csData.startTime = df['BTC时间'].iloc[0][-8:]
        csData.endTime = df['BTC时间'].iloc[-1][-8:]
        csData.rpm = round(df['左轴转速(r/min)'].mean())
        csData.shaftPower = round(df['左轴功率(KW)'].mean(), 1)
        csData.heading = round(df['船艏向(°)'].mean(), 1)
        csData.initialHeading = df['船艏向(°)'].iloc[0]
        csData.finalHeading = df['船艏向(°)'].iloc[-1]
        csData.distance = round(df['位移(m)'].iloc[-1] - df['位移(m)'].iloc[0])
        csData.speed = round(df['实时航速(kn)'].mean(), 2)
        csData.initialSpeed = df['实时航速(kn)'].iloc[0]
        csData.finalSpeed = df['实时航速(kn)'].iloc[0]

        # 轨迹图数据处理
        csData.df = df[['高斯坐标X(m)','高斯坐标Y(m)', '船艏向(°)']].copy()
        x0 = csData.df['高斯坐标X(m)'][0]
        y0 = csData.df['高斯坐标Y(m)'][0]
        csData.df['x_0'] = csData.df['高斯坐标X(m)'] - x0
        csData.df['y_0'] = csData.df['高斯坐标Y(m)'] - y0

        # 最小二乘法拟合直线，直线角度作为旋转依据
        xy_array = csData.df[['x_0', 'y_0']].to_numpy()
        slope, intercept, r_value, p_value, std_err = stats.linregress(xy_array[:, 0], xy_array[:, 1])
        # 计算拟合直线与y轴的夹角(弧度)
        angle = np.arctan(slope)
        # 旋转矩阵，计算需要旋转的角度
        if csData.df['x_0'].iloc[-1] > 0 and csData.df['y_0'].iloc[-1] > 0:
            angle = np.pi / 2 - angle
        elif csData.df['x_0'].iloc[-1] < 0 and csData.df['y_0'].iloc[-1] > 0:
            angle = np.pi / 2 + angle
        elif csData.df['x_0'].iloc[-1] < 0 and csData.df['y_0'].iloc[-1] < 0:
            angle = -np.pi / 2 - angle
        elif csData.df['x_0'].iloc[-1] > 0 and csData.df['y_0'].iloc[-1] < 0:
            angle = -np.pi / 2 + angle
        #angle += np.pi / 2
        # print('111', slope, angle, angle*180/np.pi, csData.df['船艏向(°)'][0]* np.pi / 180, csData.df['船艏向(°)'][0])
        rotation_matrix = np.array([[np.cos(angle), -np.sin(angle)], [np.sin(angle), np.cos(angle)]])
        # 旋转折线
        rotated_points = np.dot(xy_array, rotation_matrix)
        csData.df.loc[:, 'x_0_r'] = rotated_points[:, 0]
        csData.df.loc[:, 'y_0_r'] = rotated_points[:, 1]

        # 航速-距离-时间
        time = df['BTC时间'].apply(lambda t : (int(t[-8:-6]) * 60 + int(t[-5:-3])) * 60 + int(t[-2:]))
        csData.df['time'] = time - time.iloc[0]
        csData.df['speed'] = df['实时航速(kn)'].copy()
        csData.df['distance'] = df['迹程(m)'] - df['迹程(m)'].iloc[0]
        return csData
