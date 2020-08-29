from PySide2.QtCore import (QCoreApplication, QDate, QDateTime, QMetaObject,
                            QObject, QPoint, QRect, QSize, QTime, QUrl, Qt)
from PySide2.QtGui import (QBrush, QColor, QConicalGradient, QCursor, QFont,
                           QFontDatabase, QIcon, QKeySequence, QLinearGradient, QPalette, QPainter,
                           QPixmap, QRadialGradient)
from PySide2.QtWidgets import *
import os
import attendancetaker.dataHandler
import datetime
import pyperclip
from pathlib import Path
import sys
import shutil
import csv


class DragDropButton(QPushButton):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
            self.setStyleSheet("background-color : lightgreen;")
        else:
            e.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):

        # If no Logs directory found at current working directory then create one.
        if not os.path.isdir("Logs"):
            os.mkdir("Logs")

        # If no Attendances directory found at current working directory then create one.
        if not os.path.isdir("Attendances"):
            os.mkdir("Attendances")

        # Load settings
        attendancetaker.dataHandler.load_settings()

        # Store settings as an attribute for easy access.
        self.settings = attendancetaker.dataHandler.settings

        # If the path containing excel workbook doesnt exist, then ask user to select one.
        if not os.path.isfile(self.settings["dataPath"]):
            qfd = QFileDialog()
            defaultPath = os.getcwd()
            fltr = "Excel Files (*.xlsx)"
            path = QFileDialog.getOpenFileName(qfd, "Select Students List", defaultPath, fltr)
            if not path[0]:
                # If user does not select excel workbook containing student list then display warning and quit.
                msg = QMessageBox()
                msg.setWindowTitle("Error")
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Error")
                msg.setInformativeText("You must select a valid student list.")
                msg.exec_()
                sys.exit()

            # Remember the excel file chosen by user.
            attendancetaker.dataHandler.change_settings("dataPath", path[0])

        # load the excel workbook containing lists of students
        attendancetaker.dataHandler.load_data(self.settings["dataPath"])

        # Building the user interface
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
        self.goButton = DragDropButton(self.centralwidget)
        self.goButton.setObjectName(u"goButton")
        self.goButton.setMinimumSize(QSize(0, 0))
        font1 = QFont()
        font1.setFamily(u"Century Gothic")
        font1.setBold(True)
        font1.setWeight(75)
        self.goButton.setFont(font1)
        self.goButton.clicked.connect(self.go)
        self.goButton.dropEvent = self.goButtonDrag

        self.verticalLayout_3.addWidget(self.goButton)

        self.dataChooser = QPushButton(self.centralwidget)
        self.dataChooser.setObjectName(u"dataChooser")
        self.dataChooser.setMinimumSize(QSize(0, 0))
        self.dataChooser.setFont(font1)
        self.dataChooser.clicked.connect(self.choose_data)

        self.verticalLayout_3.addWidget(self.dataChooser)

        self.sheetChoice = QComboBox(self.centralwidget)
        for sheet in attendancetaker.dataHandler.data:
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

        self.clearLogsButton = QPushButton(self.centralwidget)
        self.clearLogsButton.setObjectName(u"clearLogsButton")
        self.clearLogsButton.setFont(font1)
        self.clearLogsButton.clicked.connect(self.clearLogs)

        self.verticalLayout_3.addWidget(self.clearLogsButton)

        self.clearAttendancesButton = QPushButton(self.centralwidget)
        self.clearAttendancesButton.setObjectName(u"clearAttendancesButton")
        self.clearAttendancesButton.setFont(font1)
        self.clearAttendancesButton.clicked.connect(self.clearAttendances)

        self.verticalLayout_3.addWidget(self.clearAttendancesButton)

        self.contentBox.addLayout(self.verticalLayout_3)

        self.verticalLayout.addLayout(self.contentBox)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
        self.absentees: List[str] = []

    def goButtonDrag(self, e):
        if e.mimeData().urls()[0].isLocalFile:
            draggedPath = str(e.mimeData().urls()[0].toLocalFile())
            if Path(draggedPath).suffix in (".csv",):
                with open(draggedPath, "r") as csvFile:
                    content = [i for j in csv.reader(csvFile) for i in j]
                self.takeAttendance("\n".join(content))
        self.goButton.setStyleSheet("")

    def clearLogs(self):
        if os.path.isdir("Logs"):
            shutil.rmtree("Logs")
        os.mkdir("Logs")

    def clearAttendances(self):
        if os.path.isdir("Attendances"):
            shutil.rmtree("Attendances")
        os.mkdir("Attendances")

    def choose_data(self):
        """Function called to prompt user to select excel workbook containing lists of students"""
        qfd = QFileDialog()
        defaultPath = self.settings["dataPath"]
        fltr = "Excel File (*.xlsx)"
        path = QFileDialog.getOpenFileName(qfd, "Select Students List", defaultPath, fltr)
        if path[0]:
            attendancetaker.dataHandler.change_settings("dataPath", path[0])
            attendancetaker.dataHandler.load_data(self.settings["dataPath"])
            self.sheetChoice.clear()
            for sheet in attendancetaker.dataHandler.data:
                self.sheetChoice.addItem(sheet)

    def download(self):
        """Function called to create a file containing all saved attendances and save it to a location on the pc."""
        path = QFileDialog.getSaveFileName(caption="Download Attendance to",
                                           dir=str(Path(self.settings[
                                                            "downloadPath"]) / Path(
                                               f"Attendance {datetime.datetime.now().date()}.txt")),
                                           filter="Text File (*.txt)")
        if not path[0]:
            return
        attendancetaker.dataHandler.change_settings("downloadPath", str(Path(path[0]).parent))

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
        """Function to prompt user for a subject name to save the attendance under."""
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

    def takeAttendance(self, studentList: str):
        """Function to get list of present people from clipboard and determine absentees."""
        student_data = attendancetaker.dataHandler.data[self.sheetChoice.currentText()].copy()
        name = f"Logs\\Attendance;{datetime.datetime.now()}.txt".replace(
            " ", ";").replace(":", "-")
        with open(name, "w") as f:
            # Get present students from clipboard.
            s = studentList
            print(s, file=f)

        # Open the file of all present students to read all present students.
        with open(name, "r") as file:
            # Separate all students in a list line by line.
            lines = [i.strip("\n") for i in file.readlines()]
            # Make a list to store the indices of the lines containing names of recognized students that need to be
            # deleted from the list of lines so that only unrecognized names remain.
            remove_line = []

            # Loop over every line.
            for line in lines:
                # Start looking for the name in the database of students.
                for key, val in student_data.items():
                    # Check if name matches the one in database.
                    if val in line:
                        # If match found then remove the name from the database so that it is not checked for again
                        # and add the line to remove lines.
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
            output_text += "No Absentees."

        output_text += ("\n" + "-" * 20 + "\n")

        # Print out all the unrecognized names.
        unrec_names = [i.strip() for i in filter(lambda n: n, lines)]
        if len(unrec_names) > 0:
            output_text += "Unrecognized Students:\n" + "\n".join(unrec_names)
        else:
            output_text += "Unrecognized Students:\n" + "No Unrecognized students."

        output_text += ("\n" + "-" * 20 + "\n")
        self.outputBox.setText(output_text)
        self.absentees = list(student_data.values())

    def go(self, *args, **kwargs):
        """Function when go button is clicked."""
        self.takeAttendance(pyperclip.paste())

    def retranslateUi(self, MainWindow):
        """Set English text in ui."""
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

        self.clearLogsButton.setWhatsThis(QCoreApplication.translate("MainWindow",
                                                                     u"<html><head/><body><p>Clear all Log Files.</p></body></html>",
                                                                     None))
        self.clearLogsButton.setText(QCoreApplication.translate("MainWindow", u"Clear Logs", None))

        self.clearAttendancesButton.setWhatsThis(QCoreApplication.translate("MainWindow",
                                                                            u"<html><head/><body><p>Clear all Attendances.</p></body></html>",
                                                                            None))
        self.clearAttendancesButton.setText(QCoreApplication.translate("MainWindow", u"Clear Attendances", None))
