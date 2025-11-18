import sys
import os
from ui.window import MainWindow
from PyQt6.QtWidgets import QApplication
from ui.styles import get_stylesheet, DarkTheme
from PyQt6.QtGui import QIcon

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("KDD IDE")
    app.setApplicationVersion("1.0.1")
    app.setWindowIcon(QIcon(resource_path("resources/testeditor_logo.PNG")))
    style_sheet = get_stylesheet(DarkTheme)
    app.setStyleSheet(style_sheet)
    editor = MainWindow()
    editor.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
