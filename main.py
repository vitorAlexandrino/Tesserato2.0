import sys
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
import time

from main_ui import Ui_MainWindow
from SplashScreen_ui import Ui_SplashScreen

counter = 0
class SplashScreen (QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_SplashScreen()
        self.ui.setupUi(self)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)
        self.timer.start(25)

        self.show()
    
    def update(self):
        global counter
        self.ui.progressBar.setValue(counter)
        if counter >= 100:
            self.timer.stop()
            self.main = MainWindow()
            self.main.show()

        counter += 1

class MainWindow (QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SplashScreen()
    sys.exit(app.exec())
