import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex

class TableModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame = None):
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

    def headerData(self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return chr(65 + section) if section <  26 else f"Col{section+1}"
            if orientation == Qt.Orientation.Vertical:
                return str(section + 1)
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


