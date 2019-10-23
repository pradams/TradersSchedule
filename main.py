import sys
from MainWindow import MainIntroWindow
from PyQt5.QtWidgets import QApplication


### Create Applicatin Window ###
app = QApplication(sys.argv)

# Create filebrowser and day selector.
browser = MainIntroWindow()

# Close application.

sys.exit(app.exec_())



























