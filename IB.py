import os
from statistics import mean
import pdfplumber
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import pyrebase

firebaseConfig = {
    "apiKey": "AIzaSyASpoPMWaMARGs7ZkzolN01mmCipXd69Qc",
    "authDomain": "resultconverter.firebaseapp.com",
    "databaseURL": "https://resultconverter.firebaseio.com",
    "projectId": "resultconverter",
    "storageBucket": "resultconverter.appspot.com",
    "messagingSenderId": "747529888790",
    "appId": "1:747529888790:web:edb2185bc08cf6812dc137",
    "measurementId": "G-PE09ZR81H9"
}
firebase = pyrebase.initialize_app(firebaseConfig)
storage = firebase.storage()
database = firebase.database()


class Ui_IB(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(718, 567)
        MainWindow.setStyleSheet("background-color: rgb(255, 255, 255);")
        MainWindow.setWindowIcon(QtGui.QIcon('jpis.ico'))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_3.setObjectName("gridLayout_3")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem, 0, 0, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem1, 0, 2, 1, 1)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.gridLayout.addItem(spacerItem2, 3, 2, 1, 1)
        self.browseOld = QtWidgets.QComboBox(self.centralwidget)
        self.browseOld.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                     "background-color: rgb(143, 143, 143);font: 75 10pt \"Microsoft Tai Le\";\n"
                                     "")
        self.browseOld.setObjectName("browseOld")

        self.gridLayout.addWidget(self.browseOld, 4, 3, 1, 1)
        spacerItem3 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem3, 5, 0, 1, 1)
        self.browseNew = QtWidgets.QPushButton(self.centralwidget)
        self.browseNew.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                     "background-color: rgb(143, 143, 143);")
        self.browseNew.setObjectName("browseNew")
        self.gridLayout.addWidget(self.browseNew, 4, 0, 1, 1)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setStyleSheet("font: 75 22pt \"Microsoft Tai Le\";")
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 1, 1, 3, QtCore.Qt.AlignHCenter)
        self.back = QtWidgets.QPushButton(self.centralwidget)
        self.back.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "background-color: rgb(143, 143, 143);font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "")
        self.back.setText('Back')
        self.gridLayout.addWidget(self.back,0, 0, 1, 1, QtCore.Qt.AlignHCenter)
        spacerItem4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem4, 8, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "")
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 2, 0, 1, 1, QtCore.Qt.AlignHCenter)
        self.saveOld = QtWidgets.QPushButton(self.centralwidget)
        self.saveOld.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "background-color: rgb(143, 143, 143);font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "")
        self.saveOld.setObjectName("saveOld")
        self.gridLayout.addWidget(self.saveOld, 6, 3, 1, 1)
        self.saveNew = QtWidgets.QPushButton(self.centralwidget)
        self.saveNew.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "background-color: rgb(143, 143, 143);font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "")
        self.saveNew.setObjectName("saveNew")
        self.gridLayout.addWidget(self.saveNew, 6, 0, 1, 1)
        spacerItem5 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem5, 1, 0, 1, 1)
        spacerItem6 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem6, 3, 0, 1, 1)
        spacerItem7 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem7, 7, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                   "")
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 2, 3, 1, 1, QtCore.Qt.AlignHCenter)
        spacerItem8 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem8, 5, 2, 1, 1)
        spacerItem9 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem9, 4, 1, 1, 1)
        self.gridLayout_3.addLayout(self.gridLayout, 0, 1, 1, 1)
        self.fileName = QtWidgets.QLabel(self.centralwidget)
        self.fileName.setStyleSheet("font: 75 10pt \"Microsoft Tai Le\";\n"
                                    "")
        self.fileName.setText("")
        self.fileName.setObjectName("fileName")
        self.gridLayout_3.addWidget(self.fileName, 1, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 878, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.browseNew.setMaximumSize(QtCore.QSize(200, 16777215))
        self.saveNew.setMaximumSize(QtCore.QSize(200, 16777215))
        self.saveOld.setMaximumSize(QtCore.QSize(200, 16777215))
        self.browseOld.setMaximumSize(QtCore.QSize(200, 16777215))
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.activeFile = ""
        self.browseNew.clicked.connect(self.browseFileNew)
        self.saveNew.clicked.connect(self.saveFileNew)
        try:
            data = list(database.child('Excel').child('IB').get().val().keys())
        except:
            data = 0
        if data == 0:
            self.browseOld.addItem("No File Saved")
        else:
            for i in range(len(data)):
                file = database.child("Excel").child('IB').child(data[i]).get().val()
                self.browseOld.addItem(file)
        self.saveOld.clicked.connect(self.saveFileOld)

    def saveFileOld(self):
        if self.browseOld.itemText(0) == "No File Saved":
            self.msg = QtWidgets.QMessageBox()
            self.msg.setWindowTitle("Error")
            self.msg.setText("There are no files saved Online")
            self.msg.setIcon(QtWidgets.QMessageBox.Warning)
            self.msg.exec()
        else:
            self.filesaveOld = QtWidgets.QFileDialog.getSaveFileName()[0]
            filename = str(self.filesaveOld).split('/')[-1]
            filename.replace("/", "\\")
            print(self.browseOld.currentText())
            if self.filesaveOld != "":
                storage.child("Excel/IB/" + self.browseOld.currentText()).download(filename + ".xlsx")
                self.msg = QtWidgets.QMessageBox()
                self.msg.setWindowTitle("Saved")
                self.msg.setText("File retrieved and downloaded")
                self.msg.setIcon(QtWidgets.QMessageBox.Information)
                self.msg.exec()

    def saveFileNew(self):
        if self.activeFile == "":
            self.msg = QtWidgets.QMessageBox()
            self.msg.setWindowTitle("Error")
            self.msg.setText("Please Select a file first")
            self.msg.setIcon(QtWidgets.QMessageBox.Warning)
            self.msg.exec()
        else:
            msgBox = QtWidgets.QMessageBox()
            msgBox.setText('Where do you want to save this file?')
            msgBox.addButton(QtWidgets.QPushButton('Local'), QtWidgets.QMessageBox.YesRole)
            msgBox.addButton(QtWidgets.QPushButton('Local and Online'), QtWidgets.QMessageBox.NoRole)
            msgBox.addButton(QtWidgets.QPushButton('Cancel'), QtWidgets.QMessageBox.RejectRole)
            result = msgBox.exec()
            if int(result) == 0:
                self.savePath = QFileDialog.getSaveFileName(directory=self.activeFile, filter="Excel Files (*.xlsx)")
                filesave = str(self.savePath[0]).split('/')[-1]
                self.generate(filesave, self.filepath[0], 0)
            elif int(result) == 1:
                self.savePath = QFileDialog.getSaveFileName(directory=self.activeFile, filter="Excel Files (*.xlsx)")
                filesave = str(self.savePath[0]).split('/')[-1]
                self.generate(filesave, self.filepath[0], 1)

    def browseFileNew(self):
        self.filepath = QFileDialog.getOpenFileName(filter="PDF Files (*.pdf)")

        self.fileName.setText(self.filepath[0])
        self.activeFile = str(self.filepath[0]).split('/')[-1].split('.')[0]
        print(self.activeFile)

    def generate(self, filesave, filename, type):
        pdf = pdfplumber.open(filename)
        page = pdf.pages[0]
        text = page.extract_text()
        text.encode('utf-8')
        table_list = text.split('\n')
        candidate = table_list[3].split(' ')
        session = candidate[1] + " " + candidate[2]
        schoolCode = candidate[4]
        subjectDict = {}
        candName = []
        candidateNumber = []
        subjectDict['ECONOMICS SL'] = {}
        subjectDict['ECONOMICS HL'] = {}
        subjectDict['ECONOMICS EE'] = {}
        subjectDict['ENGLISH A: Lang and Literature SL'] = {}
        subjectDict['ENGLISH A: Lang and Literature HL'] = {}
        subjectDict['ENGLISH A EE'] = {}
        subjectDict['FRENCH AB. SL'] = {}
        subjectDict['FRENCH B SL'] = {}
        subjectDict['MATHEMATICS SL'] = {}
        subjectDict['MATHEMATICS HL'] = {}
        subjectDict['MATH.STUDIES SL'] = {}
        subjectDict['MATHEMATICS EE'] = {}
        subjectDict['PHYSICS SL'] = {}
        subjectDict['PHYSICS HL'] = {}
        subjectDict['PHYSICS EE'] = {}
        subjectDict['BIOLOGY SL'] = {}
        subjectDict['BIOLOGY HL'] = {}
        subjectDict['BIOLOGY EE'] = {}
        subjectDict['COMPUTER SC. SL'] = {}
        subjectDict['COMPUTER SC. HL'] = {}
        subjectDict['COMPUTER SC. EE'] = {}
        subjectDict['HINDI B SL'] = {}
        subjectDict['HINDI B HL'] = {}
        subjectDict['HINDI B EE'] = {}
        subjectDict['CHEMISTRY SL'] = {}
        subjectDict['CHEMISTRY HL'] = {}
        subjectDict['CHEMISTRY EE'] = {}
        subjectDict['ENV. AND SOC. SL'] = {}
        subjectDict['ENV. AND SOC. EE'] = {}
        subjectDict['ENV. AND SOC. HL'] = {}
        subjectDict['BUSINESS MANAGEMENT SL'] = {}
        subjectDict['BUSINESS MANAGEMENT HL'] = {}
        subjectDict['BUSINESS MANAGEMENT EE'] = {}
        subjectDict['SPANISH AB. SL'] = {}
        subjectDict['SPANISH AB. HL'] = {}
        subjectDict['SPANISH AB. EE'] = {}
        subjectDict['THEORY KNOWL. TK'] = {}
        subjectDict['VISUAL ARTS HL'] = {}
        subjectDict['VISUAL ARTS SL'] = {}
        subjectDict['VISUAL ARTS EE'] = {}
        subjectDict['PSYCHOLOGY HL'] = {}
        subjectDict['PSYCHOLOGY SL'] = {}
        subjectDict['PSYCHOLOGY EE'] = {}
        subjectDict['HISTORY HL'] = {}
        subjectDict['HISTORY SL'] = {}
        subjectDict['HISTORY EE'] = {}
        subjectDict['HISTORY AMERICAS HL'] = {}
        subjectDict['WLD. STUDIES HEALTH AND DEV EE'] = {}
        subjectDict['FRENCH B EE'] = {}
        total = []
        for pageNo in range(len(pdf.pages)):
            page = pdf.pages[pageNo]
            text = page.extract_text()
            text.encode('utf-8')
            table_list = text.split('\n')
            candidate = table_list[3].split(' ')
            candidateNumber.append(candidate[5])
            candName.append(table_list[4].split('Name ')[1])
            candNo = candidate[5]
            d1 = {candNo: ""}
            subjectDict['ECONOMICS SL'].update(d1)
            subjectDict['ECONOMICS HL'].update(d1)
            subjectDict['ECONOMICS EE'].update(d1)
            subjectDict['ENGLISH A: Lang and Literature SL'].update(d1)
            subjectDict['ENGLISH A: Lang and Literature HL'].update(d1)
            subjectDict['ENGLISH A EE'].update(d1)
            subjectDict['FRENCH AB. SL'].update(d1)
            subjectDict['FRENCH B SL'].update(d1)
            subjectDict['MATHEMATICS SL'].update(d1)
            subjectDict['MATHEMATICS HL'].update(d1)
            subjectDict['MATH.STUDIES SL'].update(d1)
            subjectDict['MATHEMATICS EE'].update(d1)
            subjectDict['PHYSICS SL'].update(d1)
            subjectDict['PHYSICS HL'].update(d1)
            subjectDict['PHYSICS EE'].update(d1)
            subjectDict['BIOLOGY SL'].update(d1)
            subjectDict['BIOLOGY HL'].update(d1)
            subjectDict['BIOLOGY EE'].update(d1)
            subjectDict['COMPUTER SC. SL'].update(d1)
            subjectDict['COMPUTER SC. HL'].update(d1)
            subjectDict['COMPUTER SC. EE'].update(d1)
            subjectDict['HINDI B SL'].update(d1)
            subjectDict['HINDI B HL'].update(d1)
            subjectDict['HINDI B EE'].update(d1)
            subjectDict['CHEMISTRY SL'].update(d1)
            subjectDict['CHEMISTRY HL'].update(d1)
            subjectDict['CHEMISTRY EE'].update(d1)
            subjectDict['ENV. AND SOC. SL'].update(d1)
            subjectDict['ENV. AND SOC. EE'].update(d1)
            subjectDict['ENV. AND SOC. HL'].update(d1)
            subjectDict['BUSINESS MANAGEMENT SL'].update(d1)
            subjectDict['BUSINESS MANAGEMENT HL'].update(d1)
            subjectDict['BUSINESS MANAGEMENT EE'].update(d1)
            subjectDict['SPANISH AB. SL'].update(d1)
            subjectDict['SPANISH AB. HL'].update(d1)
            subjectDict['SPANISH AB. EE'].update(d1)
            subjectDict['THEORY KNOWL. TK'].update(d1)
            subjectDict['VISUAL ARTS HL'].update(d1)
            subjectDict['VISUAL ARTS SL'].update(d1)
            subjectDict['VISUAL ARTS EE'].update(d1)
            subjectDict['PSYCHOLOGY HL'].update(d1)
            subjectDict['PSYCHOLOGY SL'].update(d1)
            subjectDict['PSYCHOLOGY EE'].update(d1)
            subjectDict['HISTORY HL'].update(d1)
            subjectDict['HISTORY SL'].update(d1)
            subjectDict['HISTORY EE'].update(d1)
            subjectDict['HISTORY AMERICAS HL'].update(d1)
            subjectDict['WLD. STUDIES HEALTH AND DEV EE'].update(d1)
            subjectDict['FRENCH B EE'].update(d1)
            for i in range(9, len(table_list) - 4):
                b = table_list[i].split(" ")
                if table_list[i] != 'Additional subjects':
                    ses = b[1] + " " + b[2]
                    grade = table_list[i].split(" " + ses + " - ")
                    subject = grade[1].split(' in')[0]
                    marks = grade[0]
                    try:
                        d1 = {candNo: int(marks)}
                    except:
                        d1 = {candNo: marks}
                    subjectDict[subject].update(d1)
            t = table_list[len(table_list) - 3].split('Total Points: ')[1]
            total.append(int(t))
        subjectNames = subjectDict.keys()
        df = pd.DataFrame()
        df['Candidate Number'] = candidateNumber
        df['Candidate Name'] = candName
        s = sorted(zip(total, candidateNumber, candName), reverse=True)[:3]
        top3marks = [s[0][0], s[1][0], s[2][0]]
        top3can = [s[0][1], s[1][1], s[2][1]]
        top3name = [s[0][2], s[1][2], s[2][2]]
        average = []
        average.append(int(mean(total)))
        average.append('')
        average.append('')
        df2 = pd.DataFrame(
            {'Top Candidates Name: ': top3name, 'Top Candidates Number': top3can, ' Top Candidates Marks': top3marks,
             'Average Total Points': average})

        for subject in subjectNames:
            grades = list(subjectDict[subject].values())
            df[subject] = grades

        df['Total Points'] = total
        dfs = {'All Marks': df, 'Analysis': df2}
        filesavename = filesave
        writer = pd.ExcelWriter(filesavename, engine='xlsxwriter')
        for sheet in dfs.keys():
            dfs[sheet].to_excel(writer, sheet_name=sheet, index=None, header=True)
        writer.save()
        if type == 0:
            self.msg = QtWidgets.QMessageBox()
            self.msg.setWindowTitle("Saved")
            self.msg.setText("File saved locally")
            self.msg.setIcon(QtWidgets.QMessageBox.Information)
            self.msg.exec()
        if type == 1:
            storage.child('Excel/IB/'+ filesavename).put(filesavename)
            database.child("Excel").child('IB').push(filesavename)
            self.msg = QtWidgets.QMessageBox()
            self.msg.setWindowTitle("Saved")
            self.msg.setText("File saved online and locally")
            self.msg.setIcon(QtWidgets.QMessageBox.Information)
            self.msg.exec()
            if self.browseOld.itemText(0) == "No File Saved":
                self.browseOld.removeItem(0)
                self.browseOld.addItem(filesavename)
            else:
                self.browseOld.addItem(filesavename)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Result Converter"))
        self.browseNew.setText(_translate("MainWindow", "Browse"))
        self.label.setText(_translate("MainWindow", "Result Converter for IB"))
        self.label_3.setText(_translate("MainWindow", "Generate New File"))
        self.saveOld.setText(_translate("MainWindow", "Save"))
        self.saveNew.setText(_translate("MainWindow", "Save"))
        self.label_2.setText(_translate("MainWindow", "Browse Old File"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_IB()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())