from PyQt5.QtCore import QObject, pyqtSignal, QThreadPool


class WorkerSignals(QObject):
    # Create worker signals
    started = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    completed = pyqtSignal()


class WorkerManager(QObject):
    _workers = {}

    def __init__(self):
        super().__init__()

        # Create a threadpool for workers.
        self.threadpool = QThreadPool()
        self.signals = WorkerSignals()

    def enqueue(self, worker):
        worker.signals.finished.connect(self.notifyCompletion)
        self.threadpool.start(worker)
        self._workers[worker.jobID] = worker

    def notifyCompletion(self, jobID):
        pass


class Worker(QObject):

    def __init__(self):
        super().__init__()

    def run(self):
        pass
