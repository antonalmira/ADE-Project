import sys
import os
import ctypes
from PyQt5 import QtWidgets, QtCore, QtGui
from app import DocuApp
from utils import get_resource_path


def main():

    if os.name == 'nt':
        myappid = 'powerintegrations.tardis.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

    if hasattr(QtCore.Qt, 'HighDpiScaleFactorRoundingPolicy'):
        QtWidgets.QApplication.setHighDpiScaleFactorRoundingPolicy(
            QtCore.Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )

    app = QtWidgets.QApplication(sys.argv)

    app_icon = QtGui.QIcon(get_resource_path(os.path.join('resources', 'icons', 'tardis_icon.ico')))
    app.setWindowIcon(app_icon)

    window = DocuApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()