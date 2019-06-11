from PySide2 import QtCore, QtGui, QtWidgets
from ui import Ui_Dialog
import sys

ui = None
Dialog = None
app = None

def pre_init():
    global ui
    global app
    global Dialog
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)


def init():
    Dialog.show()
    app.exec_()


def append_log(message):
    ui.append_log(message)

