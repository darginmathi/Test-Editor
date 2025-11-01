import pandas as pd
from .model import TableModel

from PyQt6.QtCore import QModelIndex, Qt, pyqtSignal

class TestScenarioModel(TableModel):

    TYPE_COL = 0
    ID_COL = 1
    SKIP_COL = 2
    DESC_COL = 3
    STEPS_COL = 4
    EXPECTED_COL = 5
    COMMAND_COL = 6
    DATA1_COL = 7
    DATA5_COL = 11

    def __init__(self, df: pd.DataFrame = None):
        super().__init__(df=df)

    @classmethod
    def create_preset(cls, module_name, module_abbr):
        columns = [str(i) for i in range(12)]
        abbr = module_abbr.upper()[:2]
        data = [
            ["Type", "ID", "Skip", "Description", "Steps Performed", "Expected Results", "Command", "Data1", "Data2", "Data3", "Data4", "Data5"],
            ["UCB", module_name, "", f"SynOption {module_name} Test Scripts", "", "", "", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT1", "", "SynOption Test Scripts", "Launch Application And Login", "A successful login should happen.", "StartAppWithLogin", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT2", "", "", "Scenario Started", "", "StartScenario", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT3", "", "", "", "", "", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT4", "", "", "Scenario Ended", "", "EndScenario", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT5", "", "SynOption Test Scripts", "Close Application", "", "StopApp", "", "", "", "", ""],
            ["UCF", module_name, "", f"SynOption {module_name} Test Results", "", "", "", "", "", "", "", ""],
            ["END", "", "", "", "", "", "", "", "", "", "", ""]
        ]
        return pd.DataFrame(data, columns=columns)

    def insertRows(self, position, rows, parent=QModelIndex()):
        success = super().insertRows(position, rows, parent)
        if success:
            try:
                for i in range(rows):
                    row_index = position + i
                    type_index = self.index(row_index, self.TYPE_COL)

                    if row_index < len(self.df):
                        self.df.iat[row_index, self.TYPE_COL] = "TC"
                        self.dataChanged.emit(type_index, type_index, [Qt.ItemDataRole.EditRole])

                self._update_test_case_ids()
            except Exception as e:
                    print(f"Error setting type or updating IDs after insert: {e}")
                    return False
        return success

    def removeRows(self, position, rows, parent=QModelIndex()):
        success = super().removeRows(position, rows, parent)
        if success:
            self._update_test_case_ids()
        return success

    def _reinsert_dataframe(self, position, df_to_insert):
         success = super()._reinsert_dataframe(position, df_to_insert)
         if success:
             self._update_test_case_ids()
         return success

    def _update_test_case_ids(self):
        if self.df.empty:
            return
        abbr = ""
        tc_counter = 1
        indices_to_update = []

        for i in range(len(self.df)):
            cell_type = str(self.df.iat[i, self.TYPE_COL])
            if cell_type == "TC":
                cell_id = str(self.df.iat[i, self.ID_COL])
                if cell_id.startswith("TC-") and len(cell_id) > 5:
                    abbr = cell_id[3:5]
                    break

        if not abbr:
            abbr = "UN"

        for i in range(len(self.df)):
            cell_type = str(self.df.iat[i, self.TYPE_COL])
            if cell_type == "TC":
                old_id = str(self.df.iat[i, self.ID_COL])
                new_id = f"TC-{abbr}_AUT{tc_counter}"
                if new_id != old_id:
                    self.df.iat[i, self.ID_COL] = new_id
                    index = self.index(i, self.ID_COL)
                    indices_to_update.append(index)
                tc_counter += 1

        if indices_to_update:
            first = indices_to_update[0]
            last = indices_to_update[-1]

            self.dataChanged.emit(first, last, [Qt.ItemDataRole.DisplayRole])




