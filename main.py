import sys
from MainWindow import MainIntroWindow
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtCore
QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

def main():
    ### Create Applicatin Window ###
    app = QApplication(sys.argv)

    # Create filebrowser and day selector.
    browser = MainIntroWindow()
    #window = browser.window('demo', icon=icon_image).Layout(layout)

    # Close application.
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()