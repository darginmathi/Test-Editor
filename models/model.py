import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex
from typing import Optional, Any

class TableModel(QAbstractTableModel):
    def __init__(self, df: Optional[pd.DataFrame] = None) -> None:
        super().__init__()

        if df is None:
            self.df = pd.DataFrame()
        else:
            self.df = df.copy()
        self.has_unsaved_changes = False

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self.df) if not self.df.empty else 0

    def columnCount(self, parent=QModelIndex()) -> int:
        return len(self.df.columns) if not self.df.empty else 0

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.df.empty:
            return None

        row, col = index.row(), index.column()

        if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
            value = self.df.iat[row, col]
            if pd.isna(value):
                return ""
            return str(value)
        return None

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if not index.isValid() or role != Qt.ItemDataRole.EditRole or self.df.empty:
            return False
        row, col = index.row(), index.column()
        old_value = str(self.df.iat[row, col])
        if str(value) != old_value:
            self.df.iat[row, col] = value
            self.dataChanged.emit(index, index)
            self.has_unsaved_changes = True
            return True
        return False

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return " "
            if orientation == Qt.Orientation.Vertical:
                return ""
        return None

    def flags(self, index):
        if not index.isValid() or self.df.empty:
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsDropEnabled

    def loadData(self, df):
        self.beginResetModel()
        self.df = df.copy() if not df.empty else pd.DataFrame()
        self.has_unsaved_changes = False
        self.endResetModel()

    def insertRows(self, position, rows, parent=QModelIndex()):
        if self.df.empty:
            return False
        try:
            self.beginInsertRows(parent, position, position + rows - 1)
            new_rows = pd.DataFrame([[""]*len(self.df.columns)] * rows, columns=self.df.columns)
            self.df = pd.concat([self.df.iloc[:position], new_rows, self.df.iloc[position:]]).reset_index(drop=True)
            self.endInsertRows()
            self.has_unsaved_changes = True
        except Exception as e:
            print(f"Error in insertRows: {e}")
            return False

        return True

    def removeRows(self, position, rows, parent=QModelIndex()):
        if self.df.empty or position >= len(self.df):
            return False
        try:
            self.beginRemoveRows(parent, position, position + rows - 1)
            self.df = self.df.drop(self.df.index[position:position+rows]).reset_index(drop=True)
            self.endRemoveRows()

            self.has_unsaved_changes = True
            return True
        except Exception as e:
            print(f"Error in removeRows: {e}")
            return False

    def _reinsert_dataframe(self, position, df_to_insert):
        if df_to_insert.empty:
            return False

        try:
            self.beginInsertRows(QModelIndex(), position, position + len(df_to_insert) - 1)
            self.df = pd.concat([self.df.iloc[:position], df_to_insert, self.df.iloc[position:]]).reset_index(drop=True)
            self.endInsertRows()
            return True
        except Exception as e:
            print(f"Error re-inserting rows: {e}")
            self.endInsertRows()
            return False

    def mark_saved(self):
        self.has_unsaved_changes = False

    def _check_cell_match(self, r, c, find_text_lower, match_cell):
        cell_value = self.data(self.index(r, c), Qt.ItemDataRole.DisplayRole)
        if cell_value:
            cell_value_lower = str(cell_value).lower()
            if match_cell:
                if cell_value_lower == find_text_lower:
                    return True
            else:
                if find_text_lower in cell_value_lower:
                    return True
        return False

    def find_next(self, text, start_index, match_cell=False):
        if self.df.empty or not text:
            return QModelIndex(), False

        find_text_lower = text.lower()
        start_row = start_index.row()
        start_col = start_index.column()

        if start_col < self.columnCount() - 1:
            start_col += 1
        else:
            if start_row < self.rowCount() - 1:
                 start_row += 1
                 start_col = 0
            else:
                 start_row = 0
                 start_col = 0

        current_col = start_col
        for r in range(start_row, self.rowCount()):
            for c in range(current_col, self.columnCount()):
                if self._check_cell_match(r, c, find_text_lower, match_cell):
                    return self.index(r, c), False
            current_col = 0

        for r in range(0, start_row + 1):
            end_col = start_index.column() + 1 if r == start_index.row() else self.columnCount()
            for c in range(0, end_col):
                if self._check_cell_match(r, c, find_text_lower, match_cell):
                    return self.index(r, c), True

        return QModelIndex(), False

    def find_prev(self, text, start_index, match_cell=False):
        if self.df.empty or not text:
            return QModelIndex(), False

        find_text_lower = text.lower()
        start_row = start_index.row()
        start_col = start_index.column()

        if start_col > 0:
            start_col -= 1
        else:
            if start_row > 0:
                start_row -= 1
                start_col = self.columnCount() - 1
            else:
                start_row = self.rowCount() - 1
                start_col = self.columnCount() - 1

        current_col = start_col
        for r in range(start_row, -1, -1):
            for c in range(current_col, -1, -1):
                if self._check_cell_match(r, c, find_text_lower, match_cell):
                    return self.index(r, c), False
            current_col = self.columnCount() - 1

        for r in range(self.rowCount() - 1, start_index.row() - 1, -1):
            start_c = start_index.column() - 1 if r == start_index.row() else self.columnCount() - 1
            for c in range(start_c, -1, -1):
                if self._check_cell_match(r, c, find_text_lower, match_cell):
                    return self.index(r, c), True

        return QModelIndex(), False

    def find_all(self, text, match_cell=False):
        matches = []
        if self.df.empty or not text:
            return matches

        find_text = text.lower()

        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                cell_value = self.data(self.index(i, j), Qt.ItemDataRole.DisplayRole)
                if cell_value:
                    cell_value_lower = str(cell_value).lower()
                    if match_cell:
                        if cell_value_lower == find_text:
                            matches.append(self.index(i, j))
                    else:
                        if find_text in cell_value_lower:
                            matches.append(self.index(i, j))
        return matches







