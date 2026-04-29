# main.py
import sys
import os
import ctypes
from PyQt5 import QtWidgets, QtCore
from app import DocuApp

def main():
  
    if os.name == 'nt':
        myappid = 'powerintegrations.tardis.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


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