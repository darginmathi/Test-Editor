from PyQt6.QtWidgets import (QDockWidget, QTextEdit, QVBoxLayout, QWidget,
                             QHBoxLayout, QPushButton, QLabel)
from PyQt6.QtCore import Qt, pyqtSignal

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

        layout.addWidget(self.output_text)
        self.setWidget(widget)

    def append_output(self, text):
        self.output_text.append(text)
        cursor = self.output_text.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.output_text.setTextCursor(cursor)

    def clear_output(self):
        self.output_text.clear()
        self.status_label.setText("Ready")

    def set_status(self, status, is_error=False):
        color = "red" if is_error else "green"
        self.status_label.setText(f"<span style='color: {color};'>{status}</span>")
