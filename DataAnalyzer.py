from utils import Logger
from utils.Utils import singleton

@singleton
class DataAnalyzer():
    def __init__(self):
        self.logger = Logger.GetLogger(__name__)

    def handleCsData(self, dataPath: str):
        self.logger.debug("datapath is " + dataPath)
