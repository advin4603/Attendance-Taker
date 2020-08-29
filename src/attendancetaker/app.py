"""
Determines absentees and unrecognized people from a list of students.
"""
import sys
from PySide2 import QtWidgets
from PySide2.QtWidgets import QMessageBox
from PySide2.QtGui import QIcon
from attendancetaker.gui import Ui_MainWindow
import traceback
import os
import datetime


class AttendanceTaker(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        # Load GUI
        try:
            self.init_ui()
        except SystemExit:
            # If sys.exit() called while loading, then quit.
            sys.exit()
        except:
            # If Error Encountered while loading, display error message and quit.
            msg = QMessageBox()
            msg.setWindowTitle("Critical Error")
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            error_file_name = f"Tracebacks\\Traceback{datetime.datetime.now()}".replace(":", ";").replace(".",
                                                                                                          ",") + ".txt"
            msg.setInformativeText(f"- Something went wrong -\n\nCheck {os.getcwd()}\\{error_file_name} for details.")

            with open(error_file_name, "w") as error_file:
                # Save error in traceback file for debugging
                print(traceback.format_exc(), file=error_file)
            msg.exec_()
            sys.exit()

    def init_ui(self):
        self.setupUi(self)
        self.setWindowTitle('Attendance Taker')
        appIcon = QIcon("Resources/Vexels-Office-Clipboard.ico")
        self.setWindowIcon(appIcon)
        self.show()


def main():
    app = QtWidgets.QApplication(sys.argv)
    main_window = AttendanceTaker()
    sys.exit(app.exec_())
