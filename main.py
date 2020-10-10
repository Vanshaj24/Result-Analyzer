from PyQt5 import QtWidgets
from IGCSE import Ui_IGCSE
from IB import Ui_IB
from home import Ui_Dialog
class HomePage(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super(HomePage, self).__init__(parent)
        self.setupUi(self)
        self.goToIB.clicked.connect(self.hide)
        self.goToIGCSE.clicked.connect(self.hide)
        self.setWindowTitle("Result Converter")

class IB(QtWidgets.QMainWindow, Ui_IB):
    def __init__(self, parent=None):
        super(IB, self).__init__(parent)
        self.setupUi(self)
        self.back.clicked.connect(self.hide)
        self.setWindowTitle("IB")

class IGCSE(QtWidgets.QMainWindow, Ui_IGCSE):
    def __init__(self, parent=None):
        super(IGCSE, self).__init__(parent)
        self.setupUi(self)
        self.back.clicked.connect(self.hide)
        self.setWindowTitle("IGCSE")
class ScreenManager:
    def __init__(self):
        self.home = HomePage()
        self.ib = IB()
        self.igcse = IGCSE()
        self.home.show()
        self.home.goToIGCSE.clicked.connect(self.igcse.show)
        self.home.goToIB.clicked.connect(self.ib.show)
        self.ib.back.clicked.connect(self.home.show)
        self.igcse.back.clicked.connect(self.home.show)
if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    manager = ScreenManager()
    sys.exit(app.exec_())