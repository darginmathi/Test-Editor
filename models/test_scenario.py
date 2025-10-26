import pandas as pd
from .model import TableModel

class TestScenarioModel(TableModel):

    TYPE_COL = 0
    ID_COL = 1
    COMMAND_COL = 6
    DATA1_COL = 7

    def __init__(self, df: pd.DataFrame = None):
        super().__init__(df=df)

    @classmethod
    def create_preset(cls, module_name, module_abbr):
        columns = [str(i) for i in range(12)]
        abbr = module_abbr.upper()[:2]
        data = [
            ["Type", "ID", "Skip", "Description", "Steps Performed", "Expected Results", "Command", "Data1", "Data2", "Data3", "Data4", "Data5"],
            ["UCB", module_name, "", f"SynOption {module_name} Test Scripts", "", "", "", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT1", "", "SynOption Test Scripts", "Launch Application And Login", "A successful login should happen.", "StartAppWithLogin", "e2etest", "Synergy1!", "6SDLAUWYJWZUYWT6OEFEMHDPOYJLNPY7", "", ""],
            ["TC", f"TC-{abbr}_AUT2", "", "", "Scenario Started", "", "StartScenario", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT3", "", "", "", "", "", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT4", "", "", "Scenario Ended", "", "EndScenario", "", "", "", "", ""],
            ["TC", f"TC-{abbr}_AUT5", "", "SynOption Test Scripts", "Close Application", "", "StopApp", "", "", "", "", ""],
            ["UCF", module_name, "", f"SynOption {module_name} Test Results", "", "", "", "", "", "", "", ""],
            ["END", "", "", "", "", "", "", "", "", "", "", ""]
        ]
        return pd.DataFrame(data, columns=columns)

    '''def insertRows(self, position, rows, parent=QModelIndex()):
        if self.df.empty:
            return False

        try:
            self.beginInsertRows(parent, position, position + rows - 1)
            new_rows = pd.DataFrame([[""]*len(self.df.columns)] * rows, columns=self.df.columns)
            self.df = pd.concat([self.df.iloc[:position], new_rows, self.df.iloc[position:]]).reset_index(drop=True)
            self.endInsertRows()

            # abbr = ""
            tc_counter = 1

            for i in range(len(self.df)):
                cell_id = str(self.df.iat[i, self.ID_COL])
                if cell_id.startswith("TC-"):
                    abbr = cell_id[3:5] # Grabs 'QS' from 'TC-QS_AUT1'
                    break

            if abbr:
                indices_to_update = []
                for i in range(len(self.df)):
                    cell_type = str(self.df.iat[i, self.TYPE_COL])
                    if cell_type == "TC":
                        new_id = f"TC-{abbr}_AUT{tc_counter}"
                        self.df.iat[i, self.ID_COL] = new_id
                        tc_counter += 1

                        # Tell the view to redraw this cell
                        index = self.index(i, self.ID_COL)
                        indices_to_update.append(index)

                if indices_to_update:
                    first = indices_to_update[0]
                    last = indices_to_update[-1]
                    self.dataChanged.emit(first, last, [Qt.ItemDataRole.DisplayRole])

            self.has_unsaved_changes = True
            return True

        except Exception as e:
            print(f"Error in insertRows: {e}")
            self.endInsertRows()
            return False'''

    '''def removeRows(self, position, rows, parent=QModelIndex()):
        success = super().removeRows(position, rows, parent)

        if success:
            self.update_test_case_ids()

        return success

    def update_test_case_ids(self):
        pass'''

