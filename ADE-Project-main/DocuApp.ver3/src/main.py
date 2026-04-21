# main.py
import sys
import os
import ctypes
from PyQt5 import QtWidgets, QtCore
from app import DocuApp

def main():
    # 1. FIX: Taskbar Icon (Windows only)
    # This prevents Windows from grouping your app under the Python icon.
    if os.name == 'nt':
        myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    # 2. DPI SCALING SETTINGS
    # These must be set BEFORE the QApplication is created.
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
    # Safely set rounding policy (only available in Qt 5.14+)
    if hasattr(QtCore.Qt, 'HighDpiScaleFactorRoundingPolicy'):
        QtWidgets.QApplication.setHighDpiScaleFactorRoundingPolicy(
            QtCore.Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )

    # 3. START APPLICATION
    app = QtWidgets.QApplication(sys.argv)
    
    # Instantiate your window
    window = DocuApp()
    
    # Show the window
    window.show()
    
    # Run the event loop
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()