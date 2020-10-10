from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(718, 567)
        Dialog.setStyleSheet("background-color: rgb(255, 255, 255);")
        Dialog.setWindowIcon(QtGui.QIcon('jpis.ico'))
        self.gridLayout = QtWidgets.QGridLayout(Dialog)
        self.gridLayout.setObjectName("gridLayout")
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem, 5, 0, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem1, 1, 0, 1, 1)
        self.goToIGCSE = QtWidgets.QPushButton(Dialog)
        self.goToIGCSE.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
"background-color: rgb(143, 143, 143);")
        self.goToIGCSE.setObjectName("goToIGCSE")
        self.gridLayout.addWidget(self.goToIGCSE, 4, 0, 1, 1, QtCore.Qt.AlignHCenter)
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem2, 3, 0, 1, 1)
        self.goToIB = QtWidgets.QPushButton(Dialog)
        self.goToIB.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
"background-color: rgb(143, 143, 143);")
        self.goToIB.setObjectName("goToIB")
        self.gridLayout.addWidget(self.goToIB, 2, 0, 1, 1, QtCore.Qt.AlignHCenter)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setText("")
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1, QtCore.Qt.AlignHCenter)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        jpisLogo = QtGui.QPixmap('jpis.jfif')
        self.label.setPixmap(jpisLogo)
        self.label.show()

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.goToIGCSE.setText(_translate("Dialog", "IGCSE"))
        self.goToIB.setText(_translate("Dialog", "IB"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
