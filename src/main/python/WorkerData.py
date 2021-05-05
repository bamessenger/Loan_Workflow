import uuid
import pathlib

from PyQt5.QtCore import QObject, pyqtSignal, QThreadPool, QRunnable
from ExcelFiles import XLFile


class WorkerSignals(QObject):
    # Create worker signals
    started = pyqtSignal(str)
    currentStatus = pyqtSignal(str)
    completed = pyqtSignal()


class WorkerManager(QObject):
    _workers = {}

    def __init__(self):
        super().__init__()

        # Create a threadpool for workers.
        self.threadpool = QThreadPool()
        self.signals = WorkerSignals()

    def enqueue(self, worker):
        self.threadpool.start(worker)
        self._workers[worker.jobID] = worker

    def notifyCompletion(self, jobID):
        pass


class DataWorker(QRunnable):
    # Worker for the data transfer of Encompass data
    def __init__(self, ePath, wdPath, wrPath):
        super().__init__()
        # create unique identifier for each worker
        self.jobID = str(uuid.uuid4().hex)
        self.signals = WorkerSignals()
        self.efile = ePath
        self.wdfile = wdPath
        self.wrfile = wrPath
        self.data = XLFile()

    def run(self):
        self.signals.currentStatus.emit('Starting data transfer......Done')
        self.data.fileRead(encompPath=self.efile)
        efileName = pathlib.Path(self.efile).stem
        self.signals.currentStatus.emit(efileName + ' file read......Done')
        self.data.excelWrite(wrkflwDataPath=self.wdfile)
        wdfileName = pathlib.Path(self.wdfile).stem
        self.signals.currentStatus.emit(efileName + ' data written to '
                                        + wdfileName + '......Done')
        self.signals.currentStatus.emit(wdfileName + ' saved and '
                                                    'closed......Done')
        self.data.dashData(wrkflwDataPath=self.wdfile, wrkflwRptPath=self.wrfile)
        self.signals.completed.emit()

class EmailWorker(QRunnable):
    # Worker for the Emailing of recipients of the Dashboard
    def __init__(self):
        super().__init__()

    def run(self):
        pass
