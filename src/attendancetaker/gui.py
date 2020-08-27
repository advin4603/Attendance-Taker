# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'gui.ui'
##
## Created by: Qt User Interface Compiler version 5.15.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide2.QtCore import (QCoreApplication, QDate, QDateTime, QMetaObject,
                            QObject, QPoint, QRect, QSize, QTime, QUrl, Qt)
from PySide2.QtGui import (QBrush, QColor, QConicalGradient, QCursor, QFont,
                           QFontDatabase, QIcon, QKeySequence, QLinearGradient, QPalette, QPainter,
                           QPixmap, QRadialGradient)
from PySide2.QtWidgets import *
import os
import attendancetaker.logic
import datetime
import pyperclip
from pathlib import Path
import sys


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not os.path.isdir("Logs"):
            os.mkdir("Logs")

        if not os.path.isdir("Attendances"):
            os.mkdir("Attendances")
        attendancetaker.logic.load_settings()
        self.settings = attendancetaker.logic.settings
        if not os.path.isfile(self.settings["dataPath"]):
            qfd = QFileDialog()
            defaultPath = os.getcwd()
            fltr = "Excel Files (*.xlsx)"
            path = QFileDialog.getOpenFileName(qfd, "Select Students List", defaultPath, fltr)
            if not path[0]:
                msg = QMessageBox()
                msg.setWindowTitle("Error")
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Error")
                msg.setInformativeText("You must select a valid student list.")
                msg.exec_()
                sys.exit()
            attendancetaker.logic.change_settings("dataPath", path[0])

        attendancetaker.logic.load_data(self.settings["dataPath"])
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(779, 690)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.headingLabel = QLabel(self.centralwidget)
        self.headingLabel.setObjectName(u"headingLabel")
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.headingLabel.sizePolicy().hasHeightForWidth())
        self.headingLabel.setSizePolicy(sizePolicy)
        self.headingLabel.setMinimumSize(QSize(0, 60))
        font = QFont()
        font.setFamily(u"Century Gothic")
        font.setPointSize(30)
        font.setBold(True)
        font.setWeight(75)
        self.headingLabel.setFont(font)
        self.headingLabel.setAlignment(Qt.AlignCenter)

        self.verticalLayout.addWidget(self.headingLabel)

        self.contentBox = QHBoxLayout()
        self.contentBox.setSpacing(0)
        self.contentBox.setObjectName(u"contentBox")
        self.outputBox = QTextBrowser(self.centralwidget)
        self.outputBox.setObjectName(u"outputBox")

        self.contentBox.addWidget(self.outputBox)

        self.verticalLayout_3 = QVBoxLayout()
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.verticalLayout_3.setContentsMargins(30, -1, 30, -1)
        self.goButton = QPushButton(self.centralwidget)
        self.goButton.setObjectName(u"goButton")
        self.goButton.setMinimumSize(QSize(0, 0))
        font1 = QFont()
        font1.setFamily(u"Century Gothic")
        font1.setBold(True)
        font1.setWeight(75)
        self.goButton.setFont(font1)
        self.goButton.clicked.connect(self.go)

        self.verticalLayout_3.addWidget(self.goButton)

        self.dataChooser = QPushButton(self.centralwidget)
        self.dataChooser.setObjectName(u"dataChooser")
        self.dataChooser.setMinimumSize(QSize(0, 0))
        self.dataChooser.setFont(font1)
        self.dataChooser.clicked.connect(self.choose_data)

        self.verticalLayout_3.addWidget(self.dataChooser)

        self.sheetChoice = QComboBox(self.centralwidget)
        for sheet in attendancetaker.logic.data:
            self.sheetChoice.addItem(sheet)
        self.sheetChoice.setObjectName(u"sheetChoice")
        self.sheetChoice.setFont(font1)
        self.sheetChoice.setContextMenuPolicy(Qt.DefaultContextMenu)
        self.sheetChoice.setLayoutDirection(Qt.LeftToRight)
        self.sheetChoice.setDuplicatesEnabled(True)

        self.verticalLayout_3.addWidget(self.sheetChoice)

        self.saveButton = QPushButton(self.centralwidget)
        self.saveButton.setObjectName(u"saveButton")
        self.saveButton.setFont(font1)
        self.saveButton.clicked.connect(self.save)

        self.verticalLayout_3.addWidget(self.saveButton)

        self.downloadButton = QPushButton(self.centralwidget)
        self.downloadButton.setObjectName(u"downloadButton")
        self.downloadButton.setFont(font1)
        self.downloadButton.clicked.connect(self.download)

        self.verticalLayout_3.addWidget(self.downloadButton)

        self.contentBox.addLayout(self.verticalLayout_3)

        self.verticalLayout.addLayout(self.contentBox)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
        self.absentees: List[str] = []

    def choose_data(self):
        qfd = QFileDialog()
        defaultPath = self.settings["dataPath"]
        fltr = "Excel Files (*.xlsx)"
        path = QFileDialog.getOpenFileName(qfd, "Select Students List", defaultPath, fltr)
        if path[0]:
            attendancetaker.logic.change_settings("dataPath", path[0])
            attendancetaker.logic.load_data(self.settings["dataPath"])
            self.sheetChoice.clear()
            for sheet in attendancetaker.logic.data:
                self.sheetChoice.addItem(sheet)

    def download(self, *args, **kwargs):
        path = QFileDialog.getSaveFileName(caption="Download Attendance to",
                                           dir=str(Path(self.settings[
                                                            "downloadPath"]) / Path(
                                               f"Attendance {datetime.datetime.now().date()}.txt")),
                                           filter="*.txt")
        if not path[0]:
            return
        attendancetaker.logic.change_settings("downloadPath", str(Path(path[0]).parent))

        file_name = f"Attendances\\Attendance {datetime.datetime.now().date()}.txt"
        if not os.path.isfile(file_name):
            # If file does not exist then add heading
            with open(file_name, "w") as f:
                heading = f"Absentees - {datetime.datetime.now().date()}"
                print(heading, file=f)
                content = heading
        else:
            with open(file_name, "r") as f:
                content = f.read()

        with open(path[0], "w") as new_file:
            print(content, file=new_file)

    def save(self, *args, **kwargs):
        subject, ok = QInputDialog.getText(self, 'Save Attendance', 'Enter Subject Name:')
        if not ok:
            return
        file_name = f"Attendances\\Attendance {datetime.datetime.now().date()}.txt"
        # Log the Absentees in Attendance{date}.txt under the Attendances directory
        if not os.path.isfile(file_name):
            # If file does not exist then add heading
            with open(file_name, "w") as f:
                print(f"Absentees - {datetime.datetime.now().date()}", file=f)

        with open(file_name, "a") as f:
            if len(self.absentees) == 0:
                # No Absentees.
                print(f"{subject} : No Absentees", file=f)
            elif len(self.absentees) == 1:
                # Just 1 Absentee.
                print(f"{subject} : {self.absentees[0].strip()}", file=f)
            else:
                # Multiple Absentees.
                print(
                    f"{subject} : {', '.join([i.strip() for i in self.absentees[:-1]])}, and {self.absentees[-1].strip()}",
                    file=f)

    def go(self, *args, **kwargs):
        student_data = attendancetaker.logic.data[self.sheetChoice.currentText()].copy()
        name = f"Logs\\Attendance;{datetime.datetime.now()}.txt".replace(
            " ", ";").replace(":", "-")
        with open(name, "w") as f:
            # Get present students from clipboard.
            s = pyperclip.paste()
            print(s, file=f)

        # Open the file of all present students to read all present students.
        with open(name, "r") as file:
            # Separate all students in a list line by line.
            lines = [i.strip("\n") for i in file.readlines()]
            # Make a list to store the indices of the lines containing names of recognized students that need to be deleted from the list of lines so that only unrecognized names remain.
            remove_line = []

            # Loop over every line.
            for line in lines:
                # Start looking for the name in the database of students.
                for key, val in student_data.items():
                    # Check if name matches the one in database.
                    if val in line:
                        # If match found then remove the name from the database so that it is not checked for again and add the line to remove lines.
                        del student_data[key]
                        remove_line.append(line)
                        break

        # Remove all the lines with recognized names.
        for line in remove_line:
            lines.remove(line)
        output_text = ""

        # Print out all the Absentees that remain in the database.
        if student_data:
            output_text += "Absentees:\n" + "\n".join([f"{i} : {student_data[i].strip()}" for i in student_data])
        else:
            output_text += ("No Absentees.")

        output_text += ("\n" + "-" * 20 + "\n")

        # Print out all the unrecognized names.
        output_text += "Unrecognized Students:\n" + "\n".join([i.strip() for i in list(filter(lambda n: n, lines))])

        output_text += ("\n" + "-" * 20 + "\n")
        self.outputBox.setText(output_text)
        self.absentees = list(student_data.values())

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.headingLabel.setText(QCoreApplication.translate("MainWindow", u"Attendance Taker", None))
        # if QT_CONFIG(whatsthis)
        self.outputBox.setWhatsThis(
            QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Output Box</p></body></html>", None))
        # endif // QT_CONFIG(whatsthis)
        # if QT_CONFIG(whatsthis)
        self.goButton.setWhatsThis(
            QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Start Attendance</p></body></html>", None))
        # endif // QT_CONFIG(whatsthis)
        self.goButton.setText(QCoreApplication.translate("MainWindow", u"Go", None))

        self.dataChooser.setWhatsThis(
            QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Select student list</p></body></html>",
                                       None))
        self.dataChooser.setText(QCoreApplication.translate("MainWindow", u"Change List", None))

        # if QT_CONFIG(whatsthis)
        self.sheetChoice.setWhatsThis(QCoreApplication.translate("MainWindow",
                                                                 u"<html><head/><body><p>Choose the sheet containing all students</p></body></html>",
                                                                 None))
        # endif // QT_CONFIG(whatsthis)
        # if QT_CONFIG(whatsthis)
        self.saveButton.setWhatsThis(QCoreApplication.translate("MainWindow",
                                                                u"<html><head/><body><p>Save the current attendance.</p></body></html>",
                                                                None))
        # endif // QT_CONFIG(whatsthis)
        self.saveButton.setText(QCoreApplication.translate("MainWindow", u"Save", None))
        # if QT_CONFIG(whatsthis)
        self.downloadButton.setWhatsThis(QCoreApplication.translate("MainWindow",
                                                                    u"<html><head/><body><p>Download the saved attendances. <span style=\" font-weight:600;\">Attention - clears the saved attendances</span></p></body></html>",
                                                                    None))
        # endif // QT_CONFIG(whatsthis)
        self.downloadButton.setText(QCoreApplication.translate("MainWindow", u"Download", None))
    # retranslateUi
