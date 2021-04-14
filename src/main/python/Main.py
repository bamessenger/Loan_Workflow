import sys

from fbs_runtime.application_context.PyQt5 import ApplicationContext
from MainWin import MainWindowUI


class AppContext(ApplicationContext):
    def run(self):
        window = MainWindowUI()
        version = self.build_settings["version"]
        window.setWindowTitle("Loan Workflow Admin v" + version)
        window.show()
        return self.app.exec_()


if __name__ == '__main__':
    appctxt = AppContext()  # 1. Instantiate ApplicationContext
    exit_code = appctxt.run()  # 2. Invoke run()
    sys.exit(exit_code)
