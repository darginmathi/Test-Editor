import pandas as pd
from .model import TableModel

from PyQt6.QtCore import QModelIndex, Qt
from typing import Optional, Any

class ObjRepoModel(TableModel):

    TYPE_COL = 0
    NAME_COL = 1
    XPATH_COL = 2
    LOCATOR_COL = 3

    ABBRIVATION_MAP = {
        "cel": "Cell",
        "lnk": "Link",
        "lbl": "Label",
        "tgl": "Toggle",
        "cir": "Circle",
        "btn": "Button",
        "txt": "TextBox",
        "ddl": "DropDown",
        "chk": "Checkbox"
    }

    def __init__(self, df: Optional[pd.DataFrame] = None) -> None:
        super().__init__(df=df)

    @classmethod
    def create_preset(cls):
        columns = [str(i) for i in range(4)]
        data = [
            ["Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"],
            ["", "", "XPATH", ""],
            ["END", "", "", ""]
        ]
        return pd.DataFrame(data, columns=columns)

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if not index.isValid():
            return False

        result = super().setData(index, value, role)

        if result and role == Qt.ItemDataRole.EditRole and index.column() == self.NAME_COL:
            for abbr, full_type in self.ABBRIVATION_MAP.items():
                if value.startswith(abbr):
                    type_index = self.index(index.row(), self.TYPE_COL)
                    super().setData(type_index, full_type, role)
                    break
        return result

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Vertical:
            return str(section + 1)
        return super().headerData(section, orientation, role)

    def get_duplicate_names(self):
        if self.df.empty:
            return set()
        # Ensure we're working with string types, and handle potential float NaNs
        names = self.df.iloc[:, self.NAME_COL].astype(str)
        # Filter out empty strings and 'END'
        names = names[names.str.strip() != '']
        names = names[names != 'END']
        names = names[names != 'User friendly name of Object']
        
        # Find duplicates
        duplicates = names[names.duplicated()].unique()
        return set(duplicates)

    def insertRows(self, position, rows, parent=QModelIndex()):
        success = super().insertRows(position, rows, parent)
        if success:
            try:
                for i in range(rows):
                    row_index = position + i
                    xpath_index = self.index(row_index, self.XPATH_COL)

                    if row_index < len(self.df):
                        self.df.iat[row_index, self.XPATH_COL] = "XPATH"
                        self.dataChanged.emit(xpath_index, xpath_index, [Qt.ItemDataRole.EditRole])
                self.layoutChanged.emit()
            except Exception as e:
                    print(f"Error setting type or updating IDs after insert: {e}")
                    return False
        return success
