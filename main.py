import sys
import re
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QHBoxLayout, QHeaderView,
    QMenu, QTabWidget, QInputDialog, QDialog, QHBoxLayout, QListWidget, QLabel, QPushButton, QListWidgetItem
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
        if self.is_test_scenario:
            self._updateID()


class SpreadsheetEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Spreadsheet Editor")
        self.resize(1400, 800)

        self.current_module_name = None
        self.current_module_abbr = None
        self.current_scenario_path = None
        self.current_obj_path = None

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

        create_preset_action = QAction("CreatePreset", self)
        create_preset_action.triggered.connect(self.createPreset)
        file_menu.addAction(create_preset_action)

        load_file_action = QAction("LoadFile", self)
        load_file_action.triggered.connect(self.loadFile)
        file_menu.addAction(load_file_action)

        save_file_action = QAction("SaveFile", self)
        save_file_action.triggered.connect(self.saveFile)
        file_menu.addAction(save_file_action)

        save_preset_as_action = QAction("SaveFileAs", self)
        save_preset_as_action.triggered.connect(self.saveFileAs)
        file_menu.addAction(save_preset_as_action)

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

    def createTestScenarioPreset(self, module_name, module_abbr):
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

    def createPreset(self):
        if not self.model1.df.empty or not self.model2.df.empty:
            reply = QMessageBox.question(
                self, "Create Preset",
                "This will replace current data. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        # Prompt for module name
        module_name, ok = QInputDialog.getText(
            self, "Module Name", "Enter module name (e.g., QuickSanity):"
        )

        if not ok or not module_name.strip():
            return

        module_name = module_name.strip()

        # Prompt for module abbreviation
        module_abbr, ok = QInputDialog.getText(
            self, "Module Abbreviation",
            f"Enter 2-letter abbreviation for {module_name} (e.g., QS):",
            text=module_name[:2].upper()
        )

        if not ok or not module_abbr.strip():
            return

        module_abbr = module_abbr.strip().upper()[:2]

        # Store module info for later use
        self.current_module_name = module_name
        self.current_module_abbr = module_abbr

        # Create presets with module name
        test_scenario_df = self.createTestScenarioPreset(module_name, module_abbr)
        self.model1.loadData(test_scenario_df)

        objects_df = self.createObjectsPreset()
        self.model2.loadData(objects_df)

        self.setColumnWidths(self.table1, self.model1)
        self.setColumnWidths(self.table2, self.model2)

        QMessageBox.information(
            self, "Preset Created",
            f"Preset created for module: {module_name}\n\n"
            f"When saving for the first time, use 'Save Preset As' to create files:\n"
            f"• Automation_Module_{module_name}.xlsx\n"
            f"• ObjRep_Module_{module_name}_Test.xlsx"
        )

    def saveFileAs(self):
        if self.model1.df.empty and self.model2.df.empty:
            QMessageBox.warning(self, "Empty Sheets", "No data to save")
            return

        dialog = CustomFileDialog(self, mode="save")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            scenario_path, obj_path = dialog.getSelectedFiles()

            if scenario_path and obj_path:
                try:
                    # Create directories if they don't exist
                    os.makedirs(os.path.dirname(scenario_path), exist_ok=True)
                    os.makedirs(os.path.dirname(obj_path), exist_ok=True)

                    with pd.ExcelWriter(scenario_path, engine='openpyxl') as writer:
                        if not self.model1.df.empty:
                            self.model1.df.to_excel(writer, sheet_name='TestScenario', index=False)

                    with pd.ExcelWriter(obj_path, engine='openpyxl') as writer:
                        if not self.model2.df.empty:
                            self.model2.df.to_excel(writer, sheet_name='Objects', index=False)

                    # Store current paths for future saves
                    self.current_scenario_path = scenario_path
                    self.current_obj_path = obj_path

                    QMessageBox.information(self, "Success", "Files saved successfully!")

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save files:\n{str(e)}")

    def loadFile(self):
        dialog = CustomFileDialog(self, mode="open")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            scenario_path, obj_path = dialog.getSelectedFiles()

            if scenario_path and obj_path:
                try:
                    test_scenario_df = pd.read_excel(scenario_path, sheet_name="TestScenario")
                    objects_df = pd.read_excel(obj_path, sheet_name="Objects")

                    self.model1.loadData(test_scenario_df)
                    self.model2.loadData(objects_df)

                    self.setColumnWidths(self.table1, self.model1)
                    self.setColumnWidths(self.table2, self.model2)

                    # Store current paths for saving
                    self.current_scenario_path = scenario_path
                    self.current_obj_path = obj_path

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to load files:\n{str(e)}")

    def saveFile(self):
        if self.model1.df.empty and self.model2.df.empty:
            QMessageBox.warning(self, "Empty Sheets", "No data to save")
            return

        # If we already have paths (from loading), use them
        if hasattr(self, 'current_scenario_path') and hasattr(self, 'current_obj_path'):
            try:
                with pd.ExcelWriter(self.current_scenario_path, engine='openpyxl') as writer:
                    if not self.model1.df.empty:
                        self.model1.df.to_excel(writer, sheet_name='TestScenario', index=False)

                with pd.ExcelWriter(self.current_obj_path, engine='openpyxl') as writer:
                    if not self.model2.df.empty:
                        self.model2.df.to_excel(writer, sheet_name='Objects', index=False)

                QMessageBox.information(self, "Success", "Files saved successfully!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save files:\n{str(e)}")
        else:
            # Use custom dialog for new saves
            self.saveFileAs()

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

class CustomFileDialog(QDialog):
    def __init__(self, parent=None, mode="open"):
        super().__init__(parent)
        self.mode = mode
        self.selected_module = None
        self.setWindowTitle("Select Module and Files" if mode == "open" else "Save Module Files")
        self.setModal(True)
        self.resize(600, 400)

        self.setupUI()

    def setupUI(self):
        layout = QVBoxLayout()

        # Data directory selection
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(QLabel("Data Directory:"))
        self.dir_label = QLabel("Not selected")
        dir_layout.addWidget(self.dir_label)
        self.select_dir_btn = QPushButton("Select Directory")
        self.select_dir_btn.clicked.connect(self.selectDataDirectory)
        dir_layout.addWidget(self.select_dir_btn)
        layout.addLayout(dir_layout)

        # Module selection
        layout.addWidget(QLabel("Select Module:"))
        self.module_list = QListWidget()
        self.module_list.itemSelectionChanged.connect(self.onModuleSelected)
        layout.addWidget(self.module_list)

        # File selection
        file_layout = QHBoxLayout()

        # Test Scenario files
        scenario_layout = QVBoxLayout()
        scenario_layout.addWidget(QLabel("Test Scenario Files:"))
        self.scenario_list = QListWidget()
        scenario_layout.addWidget(self.scenario_list)
        file_layout.addLayout(scenario_layout)

        # Object Repository files
        obj_layout = QVBoxLayout()
        obj_layout.addWidget(QLabel("Object Repository Files:"))
        self.obj_list = QListWidget()
        obj_layout.addWidget(self.obj_list)
        file_layout.addLayout(obj_layout)

        layout.addLayout(file_layout)

        # Buttons
        button_layout = QHBoxLayout()
        if self.mode == "open":
            self.open_btn = QPushButton("Open Files")
            self.open_btn.clicked.connect(self.accept)
            self.open_btn.setEnabled(False)
            button_layout.addWidget(self.open_btn)
        else:
            self.save_btn = QPushButton("Save Files")
            self.save_btn.clicked.connect(self.accept)
            button_layout.addWidget(self.save_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def selectDataDirectory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Data Directory")
        if directory:
            self.data_directory = directory
            self.dir_label.setText(directory)
            self.populateModules()

    def populateModules(self):
        self.module_list.clear()
        self.scenario_list.clear()
        self.obj_list.clear()

        # Look for testSuites and objectRepositories directories
        test_suites_path = os.path.join(self.data_directory, "testSuites")
        obj_repos_path = os.path.join(self.data_directory, "objectRepositories")

        if not os.path.exists(test_suites_path) or not os.path.exists(obj_repos_path):
            QMessageBox.warning(self, "Error", "Required directories not found in selected data directory")
            return

        # Get all module directories (qaoptimus, qatitan, etc.)
        modules = set()

        if os.path.exists(test_suites_path):
            for item in os.listdir(test_suites_path):
                if os.path.isdir(os.path.join(test_suites_path, item)):
                    modules.add(item)

        if os.path.exists(obj_repos_path):
            for item in os.listdir(obj_repos_path):
                if os.path.isdir(os.path.join(obj_repos_path, item)):
                    modules.add(item)

        if not modules:
            QMessageBox.information(self, "No Modules", "No modules found in the selected directory")
            return

        for module in sorted(modules):
            self.module_list.addItem(module)

    def onModuleSelected(self):
        self.scenario_list.clear()
        self.obj_list.clear()

        if not self.module_list.selectedItems():
            return

        self.selected_module = self.module_list.selectedItems()[0].text()

        # Populate test scenario files
        test_suites_path = os.path.join(self.data_directory, "testSuites", self.selected_module)
        if os.path.exists(test_suites_path):
            for file in os.listdir(test_suites_path):
                if file.endswith('.xlsx'):
                    # Remove "Automation_Module_" prefix and ".xlsx" suffix for display
                    display_name = file.replace('Automation_Module_', '').replace('.xlsx', '')
                    item = QListWidgetItem(display_name)
                    item.file_path = os.path.join(test_suites_path, file)
                    self.scenario_list.addItem(item)

        # Populate object repository files
        obj_repos_path = os.path.join(self.data_directory, "objectRepositories", self.selected_module)
        if os.path.exists(obj_repos_path):
            for file in os.listdir(obj_repos_path):
                if file.endswith('.xlsx'):
                    # Remove "ObjRep_Module_" prefix and "_Test.xlsx" suffix for display
                    display_name = file.replace('ObjRep_Module_', '').replace('_Test.xlsx', '')
                    item = QListWidgetItem(display_name)
                    item.file_path = os.path.join(obj_repos_path, file)
                    self.obj_list.addItem(item)

        # Enable open button if we have files to open
        if self.mode == "open":
            has_scenario_files = self.scenario_list.count() > 0
            has_obj_files = self.obj_list.count() > 0
            self.open_btn.setEnabled(has_scenario_files and has_obj_files)

    def getSelectedFiles(self):
        if not hasattr(self, 'data_directory') or not self.selected_module:
            return None, None

        scenario_file = None
        obj_file = None

        if self.scenario_list.selectedItems():
            scenario_file = self.scenario_list.selectedItems()[0].file_path
        elif self.mode == "save":  # For save, we can create new files
            scenario_file = os.path.join(self.data_directory, "testSuites", self.selected_module,
                                       f"Automation_Module_{self.selected_module}.xlsx")

        if self.obj_list.selectedItems():
            obj_file = self.obj_list.selectedItems()[0].file_path
        elif self.mode == "save":  # For save, we can create new files
            obj_file = os.path.join(self.data_directory, "objectRepositories", self.selected_module,
                                  f"ObjRep_Module_{self.selected_module}_Test.xlsx")

        return scenario_file, obj_file


if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setApplicationName("Test Case Editor")
    app.setApplicationVersion("1.0")

    editor = SpreadsheetEditor()
    editor.show()

    sys.exit(app.exec())
