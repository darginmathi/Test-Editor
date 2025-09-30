import os
import re
import pandas as pd
from PyQt6.QtWidgets import (
    QMainWindow, QTableView, QVBoxLayout, QWidget, QTabWidget,
    QFileDialog, QMessageBox, QHBoxLayout, QHeaderView, QMenu,
    QInputDialog, QDialog, QApplication, QStatusBar, QLabel
)
from PyQt6.QtCore import Qt, QModelIndex, QTimer
from PyQt6.QtGui import QAction

from models.models import SpreadsheetModel
from ui import CustomFileDialog
from command_manager import command_manager


class SpreadsheetEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Spreadsheet Editor")
        self.resize(1600, 900)

        app_font = QApplication.font()
        app_font.setPointSize(int(app_font.pointSize() * 1.2))
        QApplication.setFont(app_font)

        self.current_module_name = None
        self.current_module_abbr = None
        self.current_scenario_path = None
        self.current_obj_path = None
        self.data_directory = None

        self.setupStatusBar()

        self.tab_widget = QTabWidget()

        self.tab1 = QWidget()
        self.tab2 = QWidget()

        self.setupTab1()
        self.setupTab2()

        self.tab_widget.addTab(self.tab1, "TestScenario")
        self.tab_widget.addTab(self.tab2, "Objects")

        self.model1 = SpreadsheetModel(is_test_scenario=True)  # TestScenario
        self.model2 = SpreadsheetModel(is_test_scenario=False) # Objects

        self.setupTable1()
        self.setupTable2()

        self.createTopMenu()

        self.setCentralWidget(self.tab_widget)

    def setupStatusBar(self):
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)

        # Permanent message label
        self.status_message = QLabel("Ready")
        status_bar.addPermanentWidget(self.status_message)

        # Temporary message label (clears after timeout)
        self.temp_message = QLabel("")
        status_bar.addWidget(self.temp_message)

    def showStatusMessage(self, message: str, timeout=3000):
        """Show temporary status message that clears after timeout"""
        self.temp_message.setText(message)
        if timeout > 0:
            QTimer.singleShot(timeout, self.clearTempMessage)

    def clearTempMessage(self):
        """Clear temporary status message"""
        self.temp_message.setText("")

    def showPermanentMessage(self, message: str):
        """Show permanent status message"""
        self.status_message.setText(message)

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
        """Configure table1 appearance with auto-adjusting columns and rows"""
        self.table1.setModel(self.model1)

        # Configure header for auto-adjustment
        header = self.table1.horizontalHeader()
        vertical_header = self.table1.verticalHeader()

        # Auto-adjust columns to content
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        # Auto-adjust rows to content
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        # Allow manual resizing as well
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        # Set minimum column widths for better usability
        self.table1.setColumnWidth(0, 50)   # Type
        self.table1.setColumnWidth(1, 100)  # ID
        self.table1.setColumnWidth(2, 40)   # Skip
        self.table1.setColumnWidth(3, 200)  # Description
        self.table1.setColumnWidth(4, 200)  # Steps Performed
        self.table1.setColumnWidth(5, 100)  # Expected Results
        self.table1.setColumnWidth(6, 150)  # Command
        self.table1.setColumnWidth(7, 200)  # Data1
        self.table1.setColumnWidth(8, 100)  # Data2
        self.table1.setColumnWidth(9, 100)  # Data3

        # Enable word wrap for better text display
        self.table1.setWordWrap(True)

        self.table1.setAlternatingRowColors(True)
        self.table1.setSortingEnabled(False)
        self.table1.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table1.customContextMenuRequested.connect(self.showContextMenu1)

    def setupTable2(self):
        """Configure table2 appearance with auto-adjusting columns and rows"""
        self.table2.setModel(self.model2)

        # Configure header for auto-adjustment
        header = self.table2.horizontalHeader()
        vertical_header = self.table2.verticalHeader()

        # Auto-adjust columns to content
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        # Auto-adjust rows to content
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        # Allow manual resizing as well
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        vertical_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        # Set minimum column widths for better usability
        self.table2.setColumnWidth(0, 100)   # Type
        self.table2.setColumnWidth(1, 250)  # User friendly name of Object
        self.table2.setColumnWidth(2, 80)  # By-Type
        self.table2.setColumnWidth(3, 400)  # Webdriver friendly name (reduced from 900)

        # Enable word wrap for better text display
        self.table2.setWordWrap(True)

        self.table2.setAlternatingRowColors(True)
        self.table2.setSortingEnabled(False)
        self.table2.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table2.customContextMenuRequested.connect(self.showContextMenu2)

    def createTopMenu(self):
        """Create top menu bar"""
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")

        new_file_action = QAction("New File", self)
        new_file_action.triggered.connect(self.newFile)
        file_menu.addAction(new_file_action)

        load_file_action = QAction("Load File", self)
        load_file_action.triggered.connect(self.loadFile)
        file_menu.addAction(load_file_action)

        save_file_action = QAction("Save File", self)
        save_file_action.triggered.connect(self.saveFile)
        file_menu.addAction(save_file_action)

        save_as_action = QAction("Save File As", self)
        save_as_action.triggered.connect(self.saveFileAs)
        file_menu.addAction(save_as_action)

        # Add Auto-adjust action
        edit_menu = menubar.addMenu("Edit")

        undo_action = QAction("Undo", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.undo)
        edit_menu.addAction(undo_action)

        redo_action = QAction("Redo", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.redo)
        edit_menu.addAction(redo_action)

        # Add auto-adjust action
        auto_adjust_action = QAction("Auto-adjust Columns", self)
        auto_adjust_action.setShortcut("Ctrl+R")
        auto_adjust_action.triggered.connect(self.autoAdjustColumns)
        edit_menu.addAction(auto_adjust_action)

    def autoAdjustColumns(self):
        """Auto-adjust all columns in the current tab"""
        current_tab_index = self.tab_widget.currentIndex()
        if current_tab_index == 0:
            self.table1.resizeColumnsToContents()
            self.table1.resizeRowsToContents()
        else:
            self.table2.resizeColumnsToContents()
            self.table2.resizeRowsToContents()
        self.showStatusMessage("Columns and rows auto-adjusted", 2000)

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

        # Add auto-adjust to context menu
        auto_adjust_action = QAction("Auto-adjust Columns", self)
        auto_adjust_action.triggered.connect(self.autoAdjustColumns)
        menu.addAction(auto_adjust_action)

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
        # Prompt for module name and abbreviation
        module_name, ok = QInputDialog.getText(
            self, "New File", "Enter module name (e.g., QuickSanity):"
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

        # Check if we have unsaved changes
        if not self.model1.df.empty or not self.model2.df.empty:
            reply = QMessageBox.question(
                self, "New File",
                "Create new file? This will lose any unsaved changes.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        # Store module info for later use
        self.current_module_name = module_name
        self.current_module_abbr = module_abbr

        # Create presets with module name
        test_scenario_df = self.createTestScenarioPreset(module_name, module_abbr)
        self.model1.loadData(test_scenario_df)

        objects_df = self.createObjectsPreset()
        self.model2.loadData(objects_df)

        # Update command manager with new object names
        command_manager.update_object_names(objects_df)

        # Clear current file paths since this is a new file
        self.current_scenario_path = None
        self.current_obj_path = None

        # Auto-adjust columns after loading new data
        QTimer.singleShot(100, self.autoAdjustColumns)

        # Show status message instead of popup
        self.showStatusMessage(f"New file created for module: {module_name}", 5000)

    def loadFile(self):
        dialog = CustomFileDialog(self, mode="open")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            scenario_path, obj_path = dialog.getSelectedFiles()

            if scenario_path and obj_path:
                try:
                    # Extract module name from filename
                    scenario_filename = os.path.basename(scenario_path)
                    module_match = re.search(r'Automation_Module_([^.]+)\.xlsx', scenario_filename)
                    if module_match:
                        self.current_module_name = module_match.group(1)

                    # Set data directory from the loaded file path
                    self.data_directory = os.path.dirname(os.path.dirname(scenario_path))

                    # Update command manager with data directory
                    command_manager.set_data_directory(self.data_directory)

                    # Read Excel files exactly as they are
                    test_scenario_df = pd.read_excel(
                        scenario_path,
                        sheet_name="TestScenario",
                        header=None,
                        dtype=object,
                        keep_default_na=False
                    )
                    objects_df = pd.read_excel(
                        obj_path,
                        sheet_name="Objects",
                        header=None,
                        dtype=object,
                        keep_default_na=False
                    )

                    # Load the data exactly as read from Excel
                    self.model1.loadData(test_scenario_df)
                    self.model2.loadData(objects_df)

                    self.current_scenario_path = scenario_path
                    self.current_obj_path = obj_path

                    # Update command manager with loaded object names
                    command_manager.update_object_names(objects_df)

                    # Auto-adjust columns after loading data
                    QTimer.singleShot(100, self.autoAdjustColumns)

                    # Show status message instead of popup
                    self.showStatusMessage(f"Files loaded successfully - Module: {self.current_module_name}", 5000)
                    self.showPermanentMessage(f"Loaded: {self.current_module_name}")

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to load files:\n{str(e)}")

    def saveFile(self):
        if self.model1.df.empty and self.model2.df.empty:
            self.showStatusMessage("No data to save", 3000)
            return

        # If we already have paths (from loading or previous save), use them
        if hasattr(self, 'current_scenario_path') and hasattr(self, 'current_obj_path') and self.current_scenario_path and self.current_obj_path:
            try:
                with pd.ExcelWriter(self.current_scenario_path, engine='openpyxl') as writer:
                    if not self.model1.df.empty:
                        self.model1.df.to_excel(writer, sheet_name='TestScenario', index=False, header=False)

                with pd.ExcelWriter(self.current_obj_path, engine='openpyxl') as writer:
                    if not self.model2.df.empty:
                        self.model2.df.to_excel(writer, sheet_name='Objects', index=False, header=False)

                # Show status message instead of popup
                self.showStatusMessage("Files saved successfully", 3000)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save files:\n{str(e)}")
        else:
            # Use custom dialog for new saves
            self.saveFileAs()

    def saveFileAs(self):
        if self.model1.df.empty and self.model2.df.empty:
            self.showStatusMessage("No data to save", 3000)
            return

        dialog = CustomFileDialog(self, mode="save")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            scenario_path, obj_path = dialog.getSelectedFiles()

            if scenario_path and obj_path:
                try:
                    # Extract data directory from save path
                    self.data_directory = os.path.dirname(os.path.dirname(scenario_path))

                    # Update command manager with data directory
                    command_manager.set_data_directory(self.data_directory)

                    # Create directories if they don't exist
                    os.makedirs(os.path.dirname(scenario_path), exist_ok=True)
                    os.makedirs(os.path.dirname(obj_path), exist_ok=True)

                    with pd.ExcelWriter(scenario_path, engine='openpyxl') as writer:
                        if not self.model1.df.empty:
                            self.model1.df.to_excel(writer, sheet_name='TestScenario', index=False, header=False)

                    with pd.ExcelWriter(obj_path, engine='openpyxl') as writer:
                        if not self.model2.df.empty:
                            self.model2.df.to_excel(writer, sheet_name='Objects', index=False, header=False)

                    # Store current paths for future saves
                    self.current_scenario_path = scenario_path
                    self.current_obj_path = obj_path

                    # Show status message instead of popup
                    self.showStatusMessage("Files saved successfully", 3000)
                    if hasattr(self, 'current_module_name') and self.current_module_name:
                        self.showPermanentMessage(f"Loaded: {self.current_module_name}")

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save files:\n{str(e)}")

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
