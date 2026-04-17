# src/main.py
import sys
import os
from PyQt5 import QtWidgets
from app import DocuApp


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = DocuApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()