import os
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

class TimeStampedRotatingFileHandler(RotatingFileHandler):
    def __init__(self, filename, maxBytes=0, backupCount=0, encoding=None, delay=False):
        # 记录日志文件的初始创建时间
        self.start_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + f".{datetime.now().microsecond // 1000:03d}"
        # 记录初始日志文件的文件名
        self.current_filename = filename
        super().__init__(filename, maxBytes=maxBytes, backupCount=backupCount, encoding=encoding, delay=delay)
        print(self.baseFilename)

    def doRollover(self):
        """
        重写日志切分的功能，使用时间戳命名旧日志文件。
        """
        if self.stream:
            self.stream.close()
            self.stream = None

        # 获取当前的日志文件名
        current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + f".{datetime.now().microsecond // 1000:03d}"
        new_filename = f"{LOG_PATH}{self.start_time} - {current_time}.log"

        # 如果新文件存在，删除它
        if os.path.exists(new_filename):
            os.remove(new_filename)
        
        # 重命名当前日志文件为带有时间戳的文件
        os.rename(self.current_filename, new_filename)
        
        # 更新开始时间为新的日志文件开始时间
        self.start_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + f".{datetime.now().microsecond // 1000:03d}"

        # 更新切分后的新日志文件名
        self.current_filename = f"{LOG_PATH}{self.start_time}.log"
        self.baseFilename = self.current_filename

        # 调用父类方法，继续处理切分后的文件
        if not self.delay:
            self.stream = self._open()


current_logtime = datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + f".{datetime.now().microsecond // 1000:03d}"
if not os.path.exists('log') or not os.path.isdir('log'):
    os.mkdir('log')
LOG_PATH = f"log/{current_logtime}/"
os.mkdir(LOG_PATH)
origin_filename = f"{LOG_PATH}{current_logtime}.log"
#print(origin_filename)
handler = TimeStampedRotatingFileHandler(origin_filename, maxBytes=1024*1024*10, backupCount=10)
#handler.namer = lambda name: name.split('.')[0] + '.' + name.split('.')[1]  # 修改文件名的命名规则，去除后缀中的日期和时间
handler.setLevel(logging.DEBUG)

# formatter = logging.Formatter('%(asctime)s - %(name)s -  %(filename)s - %(module)s - %(funcName)s - %(lineno)d - %(levelname)s - %(message)s')
formatter = logging.Formatter('%(asctime)s %(process)d %(thread)d %(levelname)s %(module)s:%(funcName)s:%(lineno)d %(message)s')
handler.setFormatter(formatter)


def GetLogger(name):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    logger.addHandler(handler)
    return logger