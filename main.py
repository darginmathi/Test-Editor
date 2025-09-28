import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QHBoxLayout, QHeaderView,
    QMenu, QTabWidget
)
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PyQt6.QtGui import QAction, QKeySequence, QShortcut


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

    def rowCount(self, parent=None):
        return len(self.df) if not self.df.empty else 0

    def columnCount(self, parent=None):
        return len(self.df.columns) if not self.df.empty else 0

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.df.empty:
            return None

        row, col = index.row(), index.column()
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

        # Save to undo stack
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

    def insertRow(self, row):
        """Insert a row at the specified position with appropriate defaults"""
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

        # Create new row with appropriate defaults
        new_row = {col: "" for col in self.df.columns}

        if self.is_test_scenario:
            # TestScenario defaults
            new_row["A"] = "TC"
            # Don't set ID here - let the update function handle it
        else:
            # Objects defaults - XPATH in By-Type column (index 2)
            if len(self.df.columns) > 2:
                new_row[self.df.columns[2]] = "XPATH"  # By-Type column

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

        # Save to undo stack
        self.undo_stack.append(('insert', row, pd.Series(new_row)))
        self.redo_stack.clear()

        # Auto-update IDs only for TestScenario tab
        if self.is_test_scenario:
            self._updateID()

        return True

    def deleteRow(self, row):
        """Delete a row at the specified position"""
        if self.df.empty or row < 0 or row >= len(self.df):
            return False

        # Save to undo stack
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
        type_col = self.df.columns[0]
        id_col = self.df.columns[1]
        tc_rows = []
        for i in range(len(self.df)):
            current_type = str(self.df.iat[i, 0])
            current_id = str(self.df.iat[i, 1])
            if current_type == "TC":
                if current_id.startswith("TC-UN_AUT"):
                    try:
                        tc_num = int(current_id[9:])
                        tc_rows.append((i, tc_num, current_id))
                    except ValueError:
                        tc_rows.append((i, -1, current_id))
                else:
                    tc_rows.append((i, -1, current_id))

        if not tc_rows:
            return
        tc_rows.sort(key=lambda x: x[0])
        current_numbers = [num for _, num, _ in tc_rows if num != -1]

        if current_numbers:
            max_num = max(current_numbers)
            next_available = max_num + 1
        else:
            next_available = 101
        current_number = 101
        changed_rows = []

        for row_idx, current_num, current_id in tc_rows:
            new_id = f"TC-UM_AUT{current_number}"
            if current_num != current_number or current_num == -1:
                self.df.iat[row_idx, 1] = new_id
                changed_rows.append(row_idx)
            current_number += 1
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
        if self.is_test_scenario:
            self._updateID()


class SpreadsheetEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Spreadsheet Editor")
        self.resize(1400, 800)

        # Create tab widget
        self.tab_widget = QTabWidget()

        # Create two tabs
        self.tab1 = QWidget()
        self.tab2 = QWidget()

        # Setup tabs
        self.setupTab1()
        self.setupTab2()

        # Add tabs to widget
        self.tab_widget.addTab(self.tab1, "TestScenario")
        self.tab_widget.addTab(self.tab2, "Objects")

        # Create models for both tabs with appropriate types
        self.model1 = SpreadsheetModel(is_test_scenario=True)  # TestScenario
        self.model2 = SpreadsheetModel(is_test_scenario=False) # Objects

        # Setup tables
        self.setupTable1()
        self.setupTable2()

        # Create top menu
        self.createTopMenu()

        # Set central widget
        self.setCentralWidget(self.tab_widget)

        # Create shortcuts
        self.createShortcuts()

    def createShortcuts(self):
        """Create keyboard shortcuts"""
        self.save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self)
        self.save_shortcut.activated.connect(self.saveFile)

        self.undo_shortcut = QShortcut(QKeySequence("Ctrl+Z"), self)
        self.undo_shortcut.activated.connect(self.undo)

        self.redo_shortcut = QShortcut(QKeySequence("Ctrl+Y"), self)
        self.redo_shortcut.activated.connect(self.redo)

    def setupTab1(self):
        """Setup TestScenario tab"""
        layout = QVBoxLayout(self.tab1)
        self.table1 = QTableView()
        layout.addWidget(self.table1)
        self.table1.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table1.customContextMenuRequested.connect(self.showContextMenu1)

    def setupTab2(self):
        """Setup Objects tab"""
        layout = QVBoxLayout(self.tab2)
        self.table2 = QTableView()
        layout.addWidget(self.table2)
        self.table2.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table2.customContextMenuRequested.connect(self.showContextMenu2)

    def setupTable1(self):
        """Configure table1 appearance"""
        self.table1.setModel(self.model1)
        self.table1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table1.setAlternatingRowColors(True)
        self.table1.setSortingEnabled(False)
        self.setColumnWidths(self.table1, self.model1)

    def setupTable2(self):
        """Configure table2 appearance"""
        self.table2.setModel(self.model2)
        self.table2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table2.setAlternatingRowColors(True)
        self.table2.setSortingEnabled(False)
        self.setColumnWidths(self.table2, self.model2)

    def createTopMenu(self):
        """Create top menu bar"""
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")

        new_file_action = QAction("NewFile", self)
        new_file_action.triggered.connect(self.newFile)
        file_menu.addAction(new_file_action)

        load_preset_action = QAction("LoadPreset", self)
        load_preset_action.triggered.connect(self.loadPreset)
        file_menu.addAction(load_preset_action)

        load_file_action = QAction("LoadFile", self)
        load_file_action.triggered.connect(self.loadFile)
        file_menu.addAction(load_file_action)

        save_file_action = QAction("SaveFile", self)
        save_file_action.triggered.connect(self.saveFile)
        file_menu.addAction(save_file_action)

        edit_menu = menubar.addMenu("Edit")

        undo_action = QAction("Undo", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.undo)
        edit_menu.addAction(undo_action)

        redo_action = QAction("Redo", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.redo)
        edit_menu.addAction(redo_action)

    def setColumnWidths(self, table, model):
        """Set column widths"""
        if model.df.empty:
            return

        for col_index in range(model.columnCount()):
            if col_index == 0:
                table.setColumnWidth(col_index, 100)
            elif col_index == 1:
                table.setColumnWidth(col_index, 150)
            elif col_index == 2:
                table.setColumnWidth(col_index, 100)
            elif col_index == 3:
                table.setColumnWidth(col_index, 250)
            else:
                table.setColumnWidth(col_index, 120)

    def showContextMenu1(self, position):
        self.showContextMenu(position, self.table1, self.model1)

    def showContextMenu2(self, position):
        self.showContextMenu(position, self.table2, self.model2)

    def showContextMenu(self, position, table, model):
        menu = QMenu(self)

        insert_action = QAction("Insert Row", self)
        insert_action.triggered.connect(lambda: self.insertRowContext(table, model))
        menu.addAction(insert_action)

        delete_action = QAction("Delete Row", self)
        delete_action.triggered.connect(lambda: self.deleteRowContext(table, model))
        menu.addAction(delete_action)

        menu.exec(table.viewport().mapToGlobal(position))

    def insertRowContext(self, table, model):
        current_index = table.currentIndex()

        if not current_index.isValid():
            target_row = len(model.df) if not model.df.empty else 0
        else:
            target_row = current_index.row()

        if model.insertRow(target_row):
            if target_row < len(model.df):
                new_index = model.index(target_row, 0)
                table.setCurrentIndex(new_index)
                table.edit(new_index)

    def deleteRowContext(self, table, model):
        current_index = table.currentIndex()
        if not current_index.isValid():
            return
        target_row = current_index.row()
        model.deleteRow(target_row)

    def createTestScenarioPreset(self):
        columns = [str(i) for i in range(12)]

        data = [
            ["Type", "ID", "Skip", "Description", "Steps Performed", "Expected Results", "Command", "Data1", "Data2", "Data3", "Data4", "Data5"],
            ["UCB", "Graphs", "", "SynOption Graphs Test Scripts", "", "", "", "", "", "", "", ""],
            ["TC", "TC-UM_AUT101", "", "SynOption Test Scripts", "Launch Application And Login", "A successful login should happen.", "StartAppWithLogin", "e2etest", "Synergy1!", "6SDLAUWYJWZUYWT6OEFEMHDPOYJLNPY7", "", ""],
            ["TC", "TC-UM_AUT102", "", "", "Scenario Started", "", "StartScenario", "", "", "", "", ""],
            ["TC", "TC-UM_AUT103", "", "", "", "", "", "", "", "", "", ""],
            ["TC", "TC-UM_AUT104", "", "", "Scenario Ended", "", "EndScenario", "", "", "", "", ""],
            ["TC", "TC-UM_AUT105", "", "SynOption Test Scripts", "Close Application", "", "StopApp", "", "", "", "", ""],
            ["UCF", "Graphs", "", "SynOption Graphs Test Results", "", "", "", "", "", "", "", ""],
            ["END", "", "", "", "", "", "", "", "", "", "", ""]
        ]

        return pd.DataFrame(data, columns=columns)

    def createObjectsPreset(self):
        columns = [str(i) for i in range(4)]

        data = [
            ["Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"],
            ["Link", "lnkAdmin", "XPATH", "//*[@id=\"page-admin\"]"],
            ["END", "", "", ""]
        ]

        return pd.DataFrame(data, columns=columns)

    def newFile(self):
        if not self.model1.df.empty or not self.model2.df.empty:
            reply = QMessageBox.question(
                self, "New File",
                "Create new files? This will lose any unsaved changes.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        self.model1.loadData(pd.DataFrame())
        self.model2.loadData(pd.DataFrame())
        self.setColumnWidths(self.table1, self.model1)
        self.setColumnWidths(self.table2, self.model2)

    def loadPreset(self):
        if not self.model1.df.empty or not self.model2.df.empty:
            reply = QMessageBox.question(
                self, "Load Preset",
                "This will replace current data. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        test_scenario_df = self.createTestScenarioPreset()
        self.model1.loadData(test_scenario_df)

        objects_df = self.createObjectsPreset()
        self.model2.loadData(objects_df)

        self.setColumnWidths(self.table1, self.model1)
        self.setColumnWidths(self.table2, self.model2)

    def loadFile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open TestScenario Excel File", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            test_scenario_df = pd.read_excel(path, sheet_name="TestScenario")
            objects_df = pd.read_excel(path, sheet_name="Objects")

            self.model1.loadData(test_scenario_df)
            self.model2.loadData(objects_df)

            self.setColumnWidths(self.table1, self.model1)
            self.setColumnWidths(self.table2, self.model2)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")

    def saveFile(self):
        if self.model1.df.empty and self.model2.df.empty:
            QMessageBox.warning(self, "Empty Sheets", "No data to save")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "test_cases.xlsx", "Excel Files (*.xlsx)"
        )
        if path:
            try:
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    if not self.model1.df.empty:
                        self.model1.df.to_excel(writer, sheet_name='TestScenario', index=False)
                    if not self.model2.df.empty:
                        self.model2.df.to_excel(writer, sheet_name='Objects', index=False)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file:\n{str(e)}")

    def undo(self):
        current_tab_index = self.tab_widget.currentIndex()
        if current_tab_index == 0:
            self.model1.undo()
        else:
            self.model2.undo()

    def redo(self):
        current_tab_index = self.tab_widget.currentIndex()
        if current_tab_index == 0:
            self.model1.redo()
        else:
            self.model2.redo()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setApplicationName("Test Case Editor")
    app.setApplicationVersion("1.0")

    editor = SpreadsheetEditor()
    editor.show()

    sys.exit(app.exec())
