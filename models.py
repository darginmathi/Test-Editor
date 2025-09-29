import re
import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PyQt6.QtGui import QColor
from typing import List
from command_manager import command_manager

class SpreadsheetModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame = None, is_test_scenario=True):
        super().__init__()
        self.is_test_scenario = is_test_scenario
        self.undo_stack = []
        self.redo_stack = []
        if df is None:
            self.df = pd.DataFrame()
        else:
            self.df = df.copy()

        # Update command manager with object names if this is Objects model
        if not self.is_test_scenario and not self.df.empty:
            command_manager.update_object_names(self.df)

    def rowCount(self, parent=None):
        return len(self.df) if not self.df.empty else 0

    def columnCount(self, parent=None):
        return len(self.df.columns) if not self.df.empty else 0

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.df.empty:
            return None

        row, col = index.row(), index.column()

        if role == Qt.ItemDataRole.BackgroundRole:
            # Color code autocomplete cells
            if self.is_test_scenario and col in [6, 7]:  # Command and Data1 columns
                if row < len(self.df):
                    row_type = str(self.df.iat[row, 0]) if len(self.df.columns) > 0 else ""
                    if row_type == "TC":
                        return QColor(240, 248, 255)  # Light blue background

        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
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

        self.undo_stack.append(('edit', row, col, old_value, value))
        self.redo_stack.clear()

        self.df.iat[row, col] = value
        self.dataChanged.emit(index, index)

        return True

    def flags(self, index):
        if not index.isValid() or self.df.empty:
            return Qt.ItemFlag.NoItemFlags

        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return chr(65 + section) if section < 26 else f"Col{section+1}"
            else:
                return str(section + 1)
        return None

    def get_autocomplete_options(self, index: QModelIndex, search_text: str = "") -> List[str]:
        """Get autocomplete options for specific cells"""
        if self.df.empty or not index.isValid():
            return []

        row, col = index.row(), index.column()

        # For Test Scenario - Command column (column 6)
        if self.is_test_scenario and col == 6:  # Command column
            if row < len(self.df) and len(self.df.columns) > 0:
                current_type = str(self.df.iat[row, 0])
                if current_type == "TC":
                    return command_manager.search_commands(search_text)

        # For Test Scenario - Data1 column (column 7) - Object names from Objects file
        elif self.is_test_scenario and col == 7:  # Data1 column
            if row < len(self.df) and len(self.df.columns) > 0:
                current_type = str(self.df.iat[row, 0])
                if current_type == "TC":
                    return command_manager.search_objects(search_text)

        return []

    def is_autocomplete_cell(self, index: QModelIndex) -> bool:
        """Check if a cell should have autocomplete"""
        if self.df.empty or not index.isValid():
            return False

        row, col = index.row(), index.column()

        if self.is_test_scenario and col in [6, 7]:  # Command and Data1 columns
            if row < len(self.df) and len(self.df.columns) > 0:
                current_type = str(self.df.iat[row, 0])
                return current_type == "TC"

        return False
    def insertRow(self, row):
        if self.df.empty:
            if self.df.columns.empty:
                if self.is_test_scenario:
                    columns = ["Type", "ID", "Skip", "Description", "Steps Performed",
                              "Expected Results", "Command", "Data1", "Data2", "Data3"]
                else:
                    columns = ["Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"]
                self.df = pd.DataFrame(columns=columns)

            self.beginInsertRows(QModelIndex(), 0, 0)
            self.df.loc[0] = [""] * len(self.df.columns)
            self.endInsertRows()

            self.undo_stack.append(('insert', 0, self.df.iloc[0].copy()))
            self.redo_stack.clear()
            return True

        if row < 0 or row > len(self.df):
            return False

        new_row = {col: "" for col in self.df.columns}

        if self.is_test_scenario:
            if len(self.df.columns) > 2:
                new_row[self.df.columns[0]] = "TC"
        else:
            if len(self.df.columns) > 2:
                new_row[self.df.columns[2]] = "XPATH"

        self.beginInsertRows(QModelIndex(), row, row)
        if row == 0:
            self.df = pd.concat([pd.DataFrame([new_row]), self.df], ignore_index=True)
        elif row == len(self.df):
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            top_part = self.df.iloc[:row]
            bottom_part = self.df.iloc[row:]
            self.df = pd.concat([top_part, pd.DataFrame([new_row]), bottom_part], ignore_index=True)
        self.endInsertRows()

        self.undo_stack.append(('insert', row, pd.Series(new_row)))
        self.redo_stack.clear()

        if self.is_test_scenario:
            self._updateID()

        return True

    def deleteRow(self, row):
        if self.df.empty or row < 0 or row >= len(self.df):
            return False

        deleted_row = self.df.iloc[row].copy()
        self.undo_stack.append(('delete', row, deleted_row))
        self.redo_stack.clear()

        self.beginRemoveRows(QModelIndex(), row, row)
        self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
        self.endRemoveRows()
        if self.is_test_scenario:
            self._updateID()

        return True

    def _updateID(self):
        if not self.is_test_scenario or self.df.empty or len(self.df.columns) < 2:
            return

        tc_rows = []
        pattern = r"^TC-([A-Z]+)_AUT(\d+)$"

        # Collect all TC rows and parse their IDs
        for i in range(len(self.df)):
            current_type = str(self.df.iat[i, 0])
            current_id = str(self.df.iat[i, 1])

            if current_type == "TC":
                match = re.match(pattern, current_id)
                if match:
                    module_abbr = match.group(1)
                    tc_num = int(match.group(2))
                    tc_rows.append((i, tc_num, current_id, module_abbr))
                else:
                    tc_rows.append((i, -1, current_id, None))

        if not tc_rows:
            return

        # Find the module abbreviation to use
        module_abbr = "UM"
        for _, _, _, abbr in tc_rows:
            if abbr is not None:
                module_abbr = abbr
                break

        # Sort by row index to maintain order
        tc_rows.sort(key=lambda x: x[0])

        changed_rows = []
        current_number = 1  # Always start from 1

        # Update all TC rows with sequential numbering
        for row_idx, current_num, current_id, _ in tc_rows:
            new_id = f"TC-{module_abbr}_AUT{current_number}"

            # Update if the ID has changed
            if current_id != new_id:
                self.df.iat[row_idx, 1] = new_id
                changed_rows.append(row_idx)

            current_number += 1

        # Emit data changed signals
        if changed_rows:
            for row_idx in changed_rows:
                id_index = self.index(row_idx, 1)
                self.dataChanged.emit(id_index, id_index)

    def undo(self):
        if not self.undo_stack:
            return False
        action = self.undo_stack.pop()
        action_type = action[0]

        if action_type == 'edit':
            row, col, old_value, new_value = action[1], action[2], action[3], action[4]
            current_value = str(self.df.iat[row, col])
            self.df.iat[row, col] = old_value
            self.redo_stack.append(('edit', row, col, current_value, old_value))
            index = self.index(row, col)
            self.dataChanged.emit(index, index)

        elif action_type == 'insert':
            row, inserted_data = action[1], action[2]
            deleted_row = self.df.iloc[row].copy()
            self.beginRemoveRows(QModelIndex(), row, row)
            self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
            self.endRemoveRows()
            self.redo_stack.append(('delete', row, deleted_row))

        elif action_type == 'delete':
            row, deleted_data = action[1], action[2]
            self.beginInsertRows(QModelIndex(), row, row)
            self.df = pd.concat([self.df.iloc[:row], deleted_data.to_frame().T, self.df.iloc[row:]]).reset_index(drop=True)
            self.endInsertRows()
            self.redo_stack.append(('insert', row, deleted_data))

        # Update IDs after undo for TestScenario
        if self.is_test_scenario:
            self._updateID()

        return True

    def redo(self):
        """Redo last undone action"""
        if not self.redo_stack:
            return False

        action = self.redo_stack.pop()
        action_type = action[0]

        if action_type == 'edit':
            row, col, old_value, new_value = action[1], action[2], action[3], action[4]
            current_value = str(self.df.iat[row, col])
            self.df.iat[row, col] = new_value
            self.undo_stack.append(('edit', row, col, current_value, new_value))
            index = self.index(row, col)
            self.dataChanged.emit(index, index)

        elif action_type == 'insert':
            row, inserted_data = action[1], action[2]
            self.beginInsertRows(QModelIndex(), row, row)
            self.df = pd.concat([self.df.iloc[:row], inserted_data.to_frame().T, self.df.iloc[row:]]).reset_index(drop=True)
            self.endInsertRows()
            self.undo_stack.append(('delete', row, inserted_data))

        elif action_type == 'delete':
            row, deleted_data = action[1], action[2]
            deleted_row = self.df.iloc[row].copy()
            self.beginRemoveRows(QModelIndex(), row, row)
            self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
            self.endRemoveRows()
            self.undo_stack.append(('insert', row, deleted_row))

        # Update IDs after redo for TestScenario
        if self.is_test_scenario:
            self._updateID()

        return True

    def loadData(self, df):
        """Load new data into the model"""
        self.beginResetModel()
        self.df = df.copy() if not df.empty else pd.DataFrame()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.endResetModel()

        # Update IDs after loading data for TestScenario
        #if self.is_test_scenario:
            #self._updateID()
