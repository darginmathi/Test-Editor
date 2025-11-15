from PyQt6.QtWidgets import (QDockWidget, QTextEdit, QVBoxLayout, QWidget,
                             QHBoxLayout, QPushButton, QLabel)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QTextCursor

class ConsoleDock(QDockWidget):
    clear_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__("Console", parent)
        self.setup_ui()

    def setup_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        header_layout = QHBoxLayout()

        self.status_label = QLabel("Ready")
        header_layout.addWidget(self.status_label)

        header_layout.addStretch()

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_requested.emit)
        self.clear_btn.setMaximumWidth(60)
        header_layout.addWidget(self.clear_btn)

        layout.addLayout(header_layout)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setPlaceholderText("Test output will appear here...")

        self.output_text.setAcceptRichText(False)

        layout.addWidget(self.output_text)
        self.setWidget(widget)

    def log(self, message):
        self.output_text.append(message)
        self._scroll_to_bottom()

    def log_test_start(self, project, module):
        self.log(f'[INFO] Press "Ctrl+Shift+R" to stop run execution')

    def log_test_result(self, exit_code, message=""):
        if exit_code == 0:
            self.log(f"✓ Test completed {message}")
        else:
            self.log(f"✗ Test failed with code {exit_code} {message}")

    def log_maven_output(self, line):
        if line.strip():
            self.output_text.append(line.rstrip())

    def log_config(self, message):
        self.log(message)

    def _scroll_to_bottom(self):
        cursor = self.output_text.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.output_text.setTextCursor(cursor)
        self.output_text.ensureCursorVisible()

    def clear_output(self):
        self.output_text.clear()
        self.status_label.setText("Ready")

    def set_status(self, status, is_error=False):
        self.status_label.setText(status)
