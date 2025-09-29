import sys
import os
os.environ['QT_QPA_PLATFORM'] = 'xcb'
from PyQt6.QtWidgets import QApplication
from windows import SpreadsheetEditor

if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setApplicationName("Test Case Editor")
    app.setApplicationVersion("0.1")

    editor = SpreadsheetEditor()
    editor.show()

    sys.exit(app.exec())
