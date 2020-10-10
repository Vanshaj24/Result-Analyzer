import os
from statistics import mean

import PyPDF2
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


class Ui_IGCSE(object):
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
        self.gridLayout.addWidget(self.back, 0, 0, 1, 1, QtCore.Qt.AlignHCenter)
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
            data = list(database.child('Excel').child('IGCSE').get().val().keys())
        except:
            data = 0
        if data == 0:
            self.browseOld.addItem("No File Saved")
        else:
            for i in range(len(data)):
                file = database.child("Excel").child('IGCSE').child(data[i]).get().val()
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
                storage.child("Excel/" + self.browseOld.currentText()).download(filename + ".xlsx")
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

    def sorting(self, subjects, name, cn, sub):
        highestMarks,candName,candNumber,subName,lowestMarks,lowestIDnumber,average,lowestIDname = [],[],\
                                                                                                   [],[],[],\
                                                                                                   [],[],[]
        count = 0
        for subject in subjects:
            high = -100
            lowest = 200
            highestID = ""
            lowestID = ""
            for marks in range(len(subject)):
                try:
                    if int(subject[marks]) > high:
                        high = int(subject[marks])
                        highestID = marks
                    if int(subject[marks]) < lowest:
                        low = int(subject[marks])
                        lowestID = marks
                except:
                    continue
            if high>0:
                highestMarks.append(high)
                candName.append(name[highestID])
                candNumber.append(cn[highestID])
                subName.append(sub[count])
                lowestMarks.append(low)
                lowestIDname.append(name[lowestID])
                lowestIDnumber.append(cn[lowestID])

            f = [str(g) for g in subject]
            f.sort()
            s = list(filter(("UN").__ne__, f))
            g = list(filter(("NO").__ne__, s))
            try:
                sorted = [int(g) for g in list(filter((" ").__ne__, g))]
                average.append(int(mean(sorted)))
            except:
                continue
            count += 1
        return highestMarks, candName, candNumber, subName, lowestMarks, lowestIDname, lowestIDnumber, average

    def generate(self, filesave, filename, type):
        df = pd.DataFrame()
        name = []
        dob = []
        cn = []
        eco, addmath, lit, lang, fren, bs, hin, evm, math, phy, chem, it, \
        his, acc, art, thai, bio, cs, span,intmath = [],[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []
        read_pdf = PyPDF2.PdfFileReader(
            filename)
        number_of_pages = read_pdf.getNumPages()
        for i in range(0, number_of_pages):
            eco.append(" ")
            addmath.append(" ")
            math.append(" ")
            lit.append(" ")
            lang.append(" ")
            fren.append(" ")
            hin.append(" ")
            bs.append(" ")
            evm.append(" ")
            phy.append(" ")
            chem.append(" ")
            it.append(" ")
            his.append(" ")
            acc.append(" ")
            art.append(" ")
            thai.append(" ")
            bio.append(" ")
            cs.append(" ")
            span.append(" ")
            intmath.append(" ")

        # making the array 'name' with all the names of the candidate
        for i in range(0, number_of_pages):
            page = read_pdf.getPage(i)
            page_content = page.extractText()
            b = 1
            while b < 650:
                if page_content[b:b + 3] == "No.":
                    break
                b = b + 1
            c = b + 3
            while b < 650:
                if page_content[b] == "0" or page_content[b] == "1" or page_content[b] == "2" or page_content[b] == "3" \
                        or page_content[b] == "4" or page_content[b] == "5" or page_content[b] == "6" \
                        or page_content[b] == "7" or page_content[b] == "8" or page_content[b] == "9":
                    break
                b = b + 1
            name.append(page_content[c:b])

            name[i] = str(name[i]).replace('\n', '')
            c = b
            while b < 650:
                if page_content[b] == "I":
                    break
                b = b + 1
            dob.append(page_content[c:b])
            dob[i] = str(dob[i]).replace('\n', '')
            while b < 1000:
                if page_content[b] == "/":
                    break
                b = b + 1
            c = b
            while b < 1000:
                if page_content[b] == "C":
                    break
                b = b + 1
            cn.append(page_content[c:b])
            cn[i] = str(cn[i]).replace('\n', '').replace('/', '')
            while b < 1000:
                if page_content[b] == "0" or page_content[b] == "1" or page_content[b] == "2" or page_content[b] == "3" \
                        or page_content[b] == "4" or page_content[b] == "5" or page_content[b] == "6" \
                        or page_content[b] == "7" or page_content[b] == "8" or page_content[b] == "9":
                    break
                b = b + 1
            while b < 1000:
                if page_content[b] == "I":
                    break
                b = b + 1
            try:
                for k in range(0, 30):
                    while b < 1000:
                        if page_content[b] == "0" or page_content[b] == "1" or page_content[b] == "2" or page_content[
                            b] == "3" \
                                or page_content[b] == "4" or page_content[b] == "5" or page_content[b] == "6" \
                                or page_content[b] == "7" or page_content[b] == "8" or page_content[b] == "9":
                            break
                        b = b + 1
                    code = page_content[b:b + 4]
                    b += 4
                    while b < 1500:
                        if page_content[b] == "0" or page_content[b] == "1" or page_content[b] == "2" or page_content[
                            b] == "3" \
                                or page_content[b] == "4" or page_content[b] == "5" or page_content[b] == "6" \
                                or page_content[b] == "7" or page_content[b] == "8" or page_content[b] == "9" \
                                or page_content[b:b + 9] == "NO RESULT" or page_content[b:b + 8] == "UNGRADED":
                            break
                        b = b + 1
                    marks = page_content[b:b + 2]
                    if code == "0455":
                        eco[i] = marks
                    elif code == "0486":
                        lit[i] = marks
                    elif code == "0549":
                        hin[i] = marks
                    elif code == "0606":
                        addmath[i] = marks
                    elif code == "0450":
                        bs[i] = marks
                    elif code == "0680":
                        evm[i] = marks
                    elif code == "0520":
                        fren[i] = marks
                    elif code == "0500":
                        lang[i] = marks
                    elif code == "0580":
                        math[i] = marks
                    elif code == "0625":
                        phy[i] = marks
                    elif code == "0620":
                        chem[i] = marks
                    elif code == "0610":
                        bio[i] = marks
                    elif code == "0400":
                        art[i] = marks
                    elif code == "0452":
                        acc[i] = marks
                    elif code == "0518":
                        thai[i] = marks
                    elif code == "0530":
                        span[i] = marks
                    elif code == "0417":
                        it[i] = marks
                    elif code == "0478":
                        cs[i] = marks
                    elif code == "0470":
                        his[i] = marks
                    elif code == "0607":
                        intmath[i] = marks
            except:
                continue
        avg = []
        nos = []
        dict = {0: eco, 1: fren, 2: addmath, 3: hin, 4: bs, 5: lang, 6: lit, 7: evm, 8: math, 9: phy, 10: chem, 11: bio,
        12: it, 13: thai, 14: span, 15: his, 16: art, 17: acc, 18: cs, 19: intmath}
        for i in range(0, number_of_pages):
            for k in range(len(dict)):
                a = dict[k]
                if a[i] != " " and a[i] != "UN" and a[i] != "NO":
                    a[i] = int(a[i])

        for k in range(0, len(dict)):
            a = dict.get(k)
            for i in range(0, number_of_pages):
                if a[i] != " " and a[i] != "UN" and a[i] != "NO":
                    a[i] = int(a[i])

        for i in range(0, number_of_pages):
            list1 = []
            avg1 = 0
            for k in range(0, len(dict)):
                a = dict.get(k)
                if a[i] != " " and a[i] != "UN" and a[i] != "NO":
                    list1.append(a[i])
            for k in range(0, 5):
                max1 = 0
                for j in range(len(list1)):
                    if list1[j] > max1:
                        max1 = list1[j]
                        index = j
                avg1 += max1
                try:
                    list1.pop(index)
                except:
                    continue
            avg.append(avg1 / 5)

        for i in range(0, len(dict)):
            nos.append(0)

        for k in range(0, len(dict)):
            a = dict.get(k)
            for i in range(0, number_of_pages):
                if a[i] != " ":
                    nos[k] += 1

        df['Candidate Name'] = name
        df['Date of Birth'] = dob
        df['Candidate No.'] = cn
        df['Economics'] = eco
        df['French'] = fren
        df['Add Math'] = addmath
        df['Hindi'] = hin
        df['Business Studies'] = bs
        df['English Language'] = lang
        df['English Literature'] = lit
        df['EVM'] = evm
        df['Mathematics (W/O Coursework)'] = math
        df['Physics'] = phy
        df['Chemistry'] = chem
        df['Biology'] = bio
        df['Information and Communication Technology'] = it
        df['First Language Thai'] = thai
        df['Spanish'] = span
        df['History'] = his
        df['Art and Design'] = art
        df['Accounts'] = acc
        df['Computer Science'] = cs
        df['International Math'] = intmath
        df['%Best5'] = avg
        subjects = [eco, fren, addmath, hin, bs, lang, lit, evm, math, phy, chem, bio, it, thai, span, his, art, acc, cs, intmath]
        sub = ['Economics', 'French', 'Add Math', 'Hindi', 'Business Studies', 'English Language', 'English Literature',
             'EVM', 'Mathematics(W/O Coursework)', 'Physics', 'Chemistry', 'Biology',
             'Information and Communication Technology', 'First Language Thai',
             'Spanish', 'History', 'Art and Design', 'Accounts', 'Computer Science','International Math']
        highestMarks, candName, candNumber, subName, lowestMarks, lowestIDname, lowestIDnumber, average = self.sorting(
            subjects,name, cn, sub)
        df2 = pd.DataFrame({'Subject Name': subName, 'H Candidate Number': candNumber, 'H Candidate Name': candName,
                            'Highest Marks': highestMarks,
                            'L Candidate Number': lowestIDnumber, 'L Candidate Name': lowestIDname, 'Lowest Marks': lowestMarks,
                            'Average Marks': average})
        dfs = {'All Marks': df, 'Subject Analysis': df2}

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
            storage.child('Excel/IGCSE/' + filesavename).put(filesavename)
            database.child("Excel").child('IGCSE').push(filesavename)
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
        self.label.setText(_translate("MainWindow", "Result Converter for IGCSE"))
        self.label_3.setText(_translate("MainWindow", "Generate New File"))
        self.saveOld.setText(_translate("MainWindow", "Save"))
        self.saveNew.setText(_translate("MainWindow", "Save"))
        self.label_2.setText(_translate("MainWindow", "Browse Old File"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_IGCSE()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
