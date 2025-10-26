import sys
from ui.window import MainWindow
from PyQt6.QtWidgets import QApplication
from ui.styles import app_stylesheet

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Test Case Editor")
    app.setApplicationVersion("0.1")
    app.setStyleSheet(app_stylesheet)
    editor = MainWindow()
    editor.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
