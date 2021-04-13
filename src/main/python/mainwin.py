from PyQt5.QtCore import QThreadPool
from admingui import Ui_MainWindow
from PyQt5 import QtWidgets


class MainWindowUI(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindowUI, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.threadpool = QThreadPool()
        self.ui.btnMstFileSlct.clicked.connect(self.browseEncompFile)
        self.ui.btnDlyWrkflwSlct.clicked.connect(self.browseWrkFlwFile)

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

    def browseWrkFlwFile(self):
        # Browse and select Daily Workflow file within file explorer
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.wrkflwFile, _ = QtWidgets.QFileDialog.getOpenFileName(None, "Open",
                                                                   "", "Excel "
                                                                       "Files ("
                                                                       "*.xl"
                                                                       "*);;All "
                                                                       "Files "
                                                                       "(*)",
                                                                   options=options)
        if self.wrkflwFile:
            self.ui.lneDlyWrkflwFile.setText(self.wrkflwFile)
