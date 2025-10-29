import pandas as pd
from .model import TableModel

from PyQt6.QtCore import QModelIndex, Qt

class ObjRepoModel(TableModel):

    TYPE_COL = 0
    NAME_COL = 1
    XPATH_COL = 2
    LOCATOR_COL = 3

    def __init__(self, df: pd.DataFrame = None):
        super().__init__(df=df)

    @classmethod
    def create_preset(cls):
        columns = [str(i) for i in range(4)]
        data = [
            ["Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"],
            ["Link", "lnkAdmin", "XPATH", "//*[@id=\"page-admin\"]"],
            ["END", "", "", ""]
        ]
        return pd.DataFrame(data, columns=columns)

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
                self.layoutChangedSignal.emit()
            except Exception as e:
                    print(f"Error setting type or updating IDs after insert: {e}")
                    return False
        return success
