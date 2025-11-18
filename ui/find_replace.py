from PyQt6.QtWidgets import (QDialog, QLineEdit, QPushButton, QHBoxLayout,
                             QGridLayout, QCheckBox, QLabel)
from PyQt6.QtCore import pyqtSignal

class FindReplace(QDialog):
    findNextClicked = pyqtSignal(str, bool)
    findPrevClicked = pyqtSignal(str, bool)
    replaceClicked = pyqtSignal(str, str, bool)
    replaceAllClicked = pyqtSignal(str, str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Find and Replace")

        layout = QGridLayout(self)

        layout.addWidget(QLabel("Find:"), 0 , 0)
        self.find_input = QLineEdit(self)
        layout.addWidget(self.find_input, 0 , 1)

        layout.addWidget(QLabel("Replace:"), 1 , 0)
        self.replace_input = QLineEdit(self)
        layout.addWidget(self.replace_input, 1 , 1)

        self.match_cell_checkbox = QCheckBox("Match Entire Cell", self)
        layout.addWidget(self.match_cell_checkbox, 2, 1)

        self.find_next_btn = QPushButton("Find Next", self)
        self.find_prev_btn = QPushButton("Find Prev", self)
        self.replace_btn = QPushButton("Replace", self)
        self.replace_all_btn = QPushButton("Replace All", self)
        self.close_btn = QPushButton("Close", self)

        button_grid = QGridLayout()
        button_grid.addWidget(self.find_prev_btn, 0, 0)
        button_grid.addWidget(self.find_next_btn, 0, 1)

        button_grid.addWidget(self.replace_btn, 1, 0)
        button_grid.addWidget(self.replace_all_btn, 1, 1)

        button_grid.addWidget(self.close_btn, 2, 1)

        button_wrapper = QHBoxLayout()
        button_wrapper.addLayout(button_grid)
        #button_wrapper.addStretch()

        layout.addLayout(button_wrapper, 3, 0, 1, 2)

        self.find_prev_btn.clicked.connect(self.on_find_prev)
        self.find_next_btn.clicked.connect(self.on_find_next)
        self.replace_btn.clicked.connect(self.on_replace)
        self.replace_all_btn.clicked.connect(self.on_replace_all)
        self.close_btn.clicked.connect(self.close)

        self.last_find_text = ""

    def on_find_prev(self):
        find_text = self.find_input.text()
        match_cell = self.match_cell_checkbox.isChecked()
        self.last_find_text = find_text
        self.findPrevClicked.emit(find_text, match_cell)

    def on_find_next(self):
        find_text = self.find_input.text()
        match_cell = self.match_cell_checkbox.isChecked()
        self.last_find_text = find_text
        self.findNextClicked.emit(find_text, match_cell)

    def on_replace(self):
        find_text = self.find_input.text()
        replace_text = self.replace_input.text()
        match_cell = self.match_cell_checkbox.isChecked()
        self.last_find_text = find_text
        self.replaceClicked.emit(find_text, replace_text, match_cell)

    def on_replace_all(self):
        find_text = self.find_input.text()
        replace_text = self.replace_input.text()
        match_cell = self.match_cell_checkbox.isChecked()
        self.last_find_text = find_text
        self.replaceAllClicked.emit(find_text, replace_text, match_cell)

    def showEvent(self, event):
        super().showEvent(event)
        self.find_input.selectAll()
        self.find_input.setFocus()
