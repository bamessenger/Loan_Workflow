import uuid
from time import sleep

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
    def __init__(self, ePath, wPath):
        super().__init__()
        # create unique identifier for each worker
        self.jobID = str(uuid.uuid4().hex)
        self.signals = WorkerSignals()
        self.efile = ePath
        self.wfile = wPath
        self.data = XLFile()

    def run(self):
        self.signals.currentStatus.emit('Starting File Update......Done')
        self.data.fileRead(encompPath=self.efile)
        self.signals.currentStatus.emit('Encompass data file read......Done')
        self.data.excelWrite(wrkflwPath=self.wfile)
        self.signals.currentStatus.emit('Write Encompass Data to '
                                        'Daily Workflow workbook......Done')
        sleep(3)
        self.signals.currentStatus.emit('Save and close Daily Workflow '
                                        'workbook......Done')
        self.signals.completed.emit()


class EmailWorker(QRunnable):
    # Worker for the Emailing of recipients of the Dashboard
    def __init__(self):
        super().__init__()

    def run(self):
        pass
