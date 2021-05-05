from src.main.python.AdminGui import Ui_MainWindow
from PyQt5 import QtWidgets
from WorkerData import DataWorker, WorkerManager
from PyQt5.QtWidgets import QMessageBox


class MainWindowUI(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindowUI, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.msgBox = QMessageBox()
        self.worker = WorkerManager()
        # Connect buttons with Methods
        self.ui.btnMstFileSlct.clicked.connect(self.browseEncompFile)
        self.ui.btnDlyWrkflwDataSlct.clicked.connect(self.browseWrkFlwDataFile)
        self.ui.btnDlyWrkflwRptSlct.clicked.connect(self.browseWrkFlwRptingFile)
        self.ui.btnDataUpdte.clicked.connect(self.startProc)

    def browseEncompFile(self):
        # Browse and select Encompass data extract within file explorer
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.encompFile, _ = QtWidgets.QFileDialog.getOpenFileName(None, "Open",
                                                                   "", "Excel "
                                                                       "Files ("
                                                                       "*.xl"
                                                                       "*);;All "
                                                                       "Files "
                                                                       "(*)",
                                                                   options=options)
        if self.encompFile:
            self.ui.lneMstFile.setText(self.encompFile)

    def browseWrkFlwDataFile(self):
        # Browse and select Daily Workflow Data file within file explorer
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.wrkflwDataFile, _ = QtWidgets.QFileDialog.getOpenFileName(None,
                                                                    "Open",
                                                                   "", "Excel "
                                                                       "Files ("
                                                                       "*.xl"
                                                                       "*);;All "
                                                                       "Files "
                                                                       "(*)",
                                                                   options=options)
        if self.wrkflwDataFile:
            self.ui.lneDlyWrkflwDataFile.setText(self.wrkflwDataFile)

    def browseWrkFlwRptingFile(self):
        # Browse and select Daily Workflow Data file within file explorer
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.wrkflwRptFile, _ = QtWidgets.QFileDialog.getOpenFileName(None,
                                                                    "Open",
                                                                   "", "Excel "
                                                                       "Files ("
                                                                       "*.xl"
                                                                       "*);;All "
                                                                       "Files "
                                                                       "(*)",
                                                                   options=options)
        if self.wrkflwRptFile:
            self.ui.lneDlyWrkflwRptFile.setText(self.wrkflwRptFile)

    def startProc(self):
        # Clear Status Dialogue in case user reruns the application
        self.ui.teDataStatOut.clear()
        # Do error checking
        try:
            self.startDataWorker()
        except AttributeError:
            self.msgBox.setIcon(QMessageBox.Critical)
            self.msgBox.setText("Missing Encompass and/or Workflow File")
            self.msgBox.setWindowTitle("Missing Data")
            self.msgBox.setStandardButtons(QMessageBox.Ok)
            self.msgBox.exec()

    def progressDialogue(self, text):
        self.ui.teDataStatOut.append(text)

    def completedProc(self):
        self.msgBox.setIcon(QMessageBox.Information)
        self.msgBox.setText("Encompass Data Transfer Complete")
        self.msgBox.setWindowTitle("Program Status")
        self.msgBox.setStandardButtons(QMessageBox.Ok)
        self.msgBox.exec()

    def startDataWorker(self):
        w = DataWorker(ePath=self.encompFile, wdPath=self.wrkflwDataFile,
                       wrPath=self.wrkflwRptFile)
        w.signals.currentStatus.connect(self.progressDialogue)
        w.signals.completed.connect(self.completedProc)
        self.worker.enqueue(w)

    def startEmailWorker(self):
        pass
