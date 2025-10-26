from PyQt6.QtCore import QObject, Qt
from PyQt6.QtWidgets import QApplication

from .undo_commands import EditCellCommand, InsertRowsCommand, DeleteRowsCommand

class TableController(QObject):

    def __init__(self, parent, undo_stack):
        super().__init__(parent)
        self.undo_stack = undo_stack

    def insert_rows(self, model, position, count):
        command = InsertRowsCommand(model, position, count)
        self.undo_stack.push(command)

    def delete_rows(self, model, rows):
        if not rows:
            return
        self.undo_stack.beginMacro("Delete Rows")

        for row_index in sorted(rows, reverse=True):
            command = DeleteRowsCommand(model, row_index, 1)
            self.undo_stack.push(command)
        self.undo_stack.endMacro()

    def copy(self, model, selection):
        indexes = selection.indexes()
        if not indexes:
            return

        rows = sorted(set(index.row() for index in indexes))
        cols = sorted(set(index.column() for index in indexes))

        text_lines = []
        for row in rows:
            row_data = []
            for col in cols:
                index = model.index(row, col)
                value = model.data(index, Qt.ItemDataRole.DisplayRole)
                row_data.append(str(value) if value is not None else "")
            text_lines.append('\t'.join(row_data))

        QApplication.clipboard().setText('\n'.join(text_lines))

    def paste(self, model, start_index):
        if not start_index.isValid():
            return

        clipboard_text = QApplication.clipboard().text().strip()
        if not clipboard_text:
            return

        data = []
        for line in clipboard_text.split('\n'):
            line = line.strip()
            if line:
                if '\t' in line:
                    cells = line.split('\t')
                else:
                    cells = [line]
                data.append(cells)
        if not data:
            return

        self.undo_stack.beginMacro("Paste")
        start_row = start_index.row()
        start_col = start_index.column()

        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                target_row = start_row + i
                target_col = start_col + j
                if (target_row < model.rowCount() and
                    target_col < model.columnCount()):
                    index = model.index(target_row, target_col)
                    old_value = model.data(index, Qt.ItemDataRole.EditRole)

                    if str(value) != str(old_value):
                        command = EditCellCommand(model, index, value, old_value)
                        self.undo_stack.push(command)

        self.undo_stack.endMacro()

    def cut(self, model, selection):
        self.copy(model, selection)
        self.clear(model, selection)

    def clear(self, model, selection):
        indexes = selection.indexes()
        if not indexes:
            return

        self.undo_stack.beginMacro("Clear Contents")
        for index in indexes:
            if index.isValid():
                old_value = model.data(index, Qt.ItemDataRole.EditRole)
                if str(old_value) != "":
                    command = EditCellCommand(model, index, "", old_value)
                    self.undo_stack.push(command)
        self.undo_stack.endMacro()
