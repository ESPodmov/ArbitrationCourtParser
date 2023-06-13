import os

from PyQt5.QtWidgets import QApplication
import interface
import sys

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = interface.MainWindow(os.getcwd())
    mainWindow.show()
    sys.exit(app.exec_())
