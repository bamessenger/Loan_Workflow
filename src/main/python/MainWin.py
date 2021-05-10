import psutil

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
        self.task = 0
        # Connect buttons with Methods
        self.ui.btnMstFileSlct.clicked.connect(self.browseEncompFile)
        self.ui.btnDlyWrkflwDataSlct.clicked.connect(self.browseWrkFlwDataFile)
        self.ui.btnDlyWrkflwRptSlct.clicked.connect(self.browseWrkFlwRptingFile)
        self.ui.btnDataUpdte.clicked.connect(self.startProc)

    def browseEncompFile(self):
        # Browse and select Encompass data extract within file explorer
        # Option is only given to open an Excel file
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.encompFile, _ = \
            QtWidgets.QFileDialog.getOpenFileName(None,"Open","","Excel Files "
                                                                 "(*.xl*)",
                                                                 options=options)
        if self.encompFile:
            self.ui.lneMstFile.setText(self.encompFile)

    def browseWrkFlwDataFile(self):
        # Browse and select Daily Workflow Data file within file explorer
        # Option is only given to open an Excel file
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.wrkflwDataFile, _ = \
            QtWidgets.QFileDialog.getOpenFileName(None,"Open","","Excel Files "
                                                                 "(*.xl*)",
                                                                 options=options)
        if self.wrkflwDataFile:
            self.ui.lneDlyWrkflwDataFile.setText(self.wrkflwDataFile)

    def browseWrkFlwRptingFile(self):
        # Browse and select Daily Workflow Data file within file explorer
        # Option is only given to open an Excel file
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.wrkflwRptFile, _ = \
            QtWidgets.QFileDialog.getOpenFileName(None,"Open","","Excel Files "
                                                                 "(*.xl*)",
                                                                 options=options)
        if self.wrkflwRptFile:
            self.ui.lneDlyWrkflwRptFile.setText(self.wrkflwRptFile)

    def startProc(self):
        # Clear Status Dialogue in case user reruns the application
        self.ui.teDataStatOut.clear()
        self.ui.lneTasksCmpl.clear()
        self.ui.lneCurrStatus.setText('Starting')
        self.ui.lneTasksCmpl.setText('Loading 15 Tasks...')
        self.task = 0
        self.excptHandlerStart()

    def excptHandlerFiles(self):
        self.msgBox.setIcon(QMessageBox.Critical)
        self.msgBox.setText('File Path or Name is incorrect.')
        self.msgBox.setWindowTitle('File Error')
        self.msgBox.setStandardButtons(QMessageBox.Ok)
        self.msgBox.exec()

    def excptHandlerStart(self):
        # Do error checking
        # Check to see if Excel is running, if so, produce error message
        if "EXCEL.EXE" in (i.name() for i in psutil.process_iter()):
            self.msgBox.setIcon(QMessageBox.Critical)
            self.msgBox.setText("Please close all Excel workbooks.\n\n\n"
                                "NOTE: If workbooks are closed, run\n"
                                "Task Manager (crtl+atl+delete), and\n"
                                "close background instance of Excel.\n"
                                )
            self.msgBox.setWindowTitle("Excel is Running")
            self.msgBox.setStandardButtons(QMessageBox.Ok)
            self.msgBox.exec()
            return
        # Error will occur if user hasn't selected all three files
        try:
            self.startDataWorker()
        except AttributeError as error:
            self.msgBox.setIcon(QMessageBox.Critical)
            self.msgBox.setText(str(error))
            self.msgBox.setWindowTitle("Missing Data")
            self.msgBox.setStandardButtons(QMessageBox.Ok)
            self.msgBox.exec()

    def progressDialogue(self, text):
        self.ui.teDataStatOut.append(text)

    def tasksCompleted(self, num):
        # self.task will serve as a counter for signals as they emit 1 each
        # completed task
        self.task += num
        self.ui.lneTasksCmpl.setText(str(self.task) + ' out of 15 completed')

    def currStatDialogue(self, text):
        if text == 'Running':
            self.ui.lneCurrStatus.setText(text)
            self.ui.lneCurrStatus.setStyleSheet("color: rgb(16, 88, 82);\n"
                                                "font-weight: bold;")
        else:
            self.ui.lneCurrStatus.setText(text)
            self.ui.lneCurrStatus.setStyleSheet("color: rgb(135, 135, 135);\n"
                                                "font-weight: regular;")

    def completedProc(self):
        self.msgBox.setIcon(QMessageBox.Information)
        self.msgBox.setText("Encompass Data Transfer Complete")
        self.msgBox.setWindowTitle("Program Status")
        self.msgBox.setStandardButtons(QMessageBox.Ok)
        self.msgBox.exec()

    def startDataWorker(self):
        #
        w = DataWorker(ePath=self.encompFile, wdPath=self.wrkflwDataFile,
                       wrPath=self.wrkflwRptFile)
        w.signals.output.connect(self.progressDialogue)
        w.signals.completed.connect(self.completedProc)
        w.signals.currentStatus.connect(self.currStatDialogue)
        w.signals.tskComplete.connect(self.tasksCompleted)
        self.worker.enqueue(w)

    def startEmailWorker(self):
        pass
