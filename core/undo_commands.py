from PyQt6.QtCore import QModelIndex
from PyQt6.QtGui import QUndoCommand

class EditCellCommand(QUndoCommand):
    def __init__(self, model, index, new_value, old_value, parent=None):
        super().__init__(parent)
        self.model = model
        self.index = QModelIndex(index)
        self.new_value = new_value
        self.old_value = old_value
        self.setText(f"Edit Cell ({index.row() + 1}, {index.column() + 1})")

    def undo(self):
        try:
            self.model.setData(self.index, self.old_value)
        except Exception as e:
            print(f"Error in undo: {e}")

    def redo(self):
        try:
            self.model.setData(self.index, self.new_value)
        except Exception as e:
            print(f"Error in redo: {e}")

class InsertRowsCommand(QUndoCommand):
    def __init__(self, model, position, count, parent=None):
        super().__init__(parent)
        self.model = model
        self.position = position
        self.count = count
        self.setText(f"Insert {count} Row(s)")

    def undo(self):
        self.model.removeRows(self.position, self.count)

    def redo(self):
        self.model.insertRows(self.position, self.count)

class DeleteRowsCommand(QUndoCommand):
    def __init__(self, model, position, count, parent=None):
        super().__init__(parent)
        self.model = model
        self.position = position
        self.count = count
        self.deleted_data = model.df.iloc[position : position + count].copy()
        self.setText(f"Delete {count} Row(s)")

    def undo(self):
        self.model._reinsert_dataframe(self.position, self.deleted_data)

    def redo(self):
        self.model.removeRows(self.position, self.count)
