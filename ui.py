import os
import re
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QLabel,
    QPushButton, QListWidgetItem, QFileDialog, QMessageBox,
    QSplitter, QComboBox, QWidget
)
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QFont


class CustomFileDialog(QDialog):
    def __init__(self, parent=None, mode="open"):
        super().__init__(parent)
        self.mode = mode
        self.selected_module = None
        self.setWindowTitle("Select Module and Files" if mode == "open" else "Save Module Files")
        self.setModal(True)
        self.resize(900, 600)

        self.setFont(QFont("Arial", 12))
        self.setupUI()
        self.loadLastDirectory()

    def setupUI(self):
        layout = QVBoxLayout()

        # Data directory selection
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(QLabel("Data Directory:"))
        self.dir_label = QLabel("Not selected")
        self.dir_label.setFont(QFont("Arial", 11))
        dir_layout.addWidget(self.dir_label)
        self.select_dir_btn = QPushButton("Select Directory")
        self.select_dir_btn.clicked.connect(self.selectDataDirectory)
        dir_layout.addWidget(self.select_dir_btn)
        layout.addLayout(dir_layout)

        # Sorting options
        sort_layout = QHBoxLayout()
        sort_layout.addWidget(QLabel("Sort by:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Alphabetical", "Modification Date", "Size"])
        self.sort_combo.currentTextChanged.connect(self.populateModules)
        sort_layout.addWidget(self.sort_combo)
        sort_layout.addStretch()
        layout.addLayout(sort_layout)

        # Horizontal split
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left: Modules
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("Modules:"))
        self.module_list = QListWidget()
        self.module_list.itemSelectionChanged.connect(self.onModuleSelected)
        left_layout.addWidget(self.module_list)
        splitter.addWidget(left_widget)

        # Right: Only Test Scenario files list with status
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # Test Scenario files
        scenario_layout = QVBoxLayout()
        scenario_layout.addWidget(QLabel("Test Scenario Files:"))
        self.scenario_list = QListWidget()
        self.scenario_list.itemSelectionChanged.connect(self.onScenarioSelected)
        scenario_layout.addWidget(self.scenario_list)
        right_layout.addLayout(scenario_layout)

        # Status label for matching object file
        self.status_label = QLabel("")  # Start with empty message
        self.status_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border: 1px solid #ccc;")
        self.status_label.setMinimumHeight(0)  # Ensure consistent height
        right_layout.addWidget(self.status_label)

        splitter.addWidget(right_widget)
        splitter.setSizes([300, 600])
        layout.addWidget(splitter)

        # Buttons - FIX: Create the correct button based on mode
        button_layout = QHBoxLayout()

        if self.mode == "open":
            self.open_btn = QPushButton("Open Files")
            self.open_btn.clicked.connect(self.openFiles)
            self.open_btn.setEnabled(False)
            button_layout.addWidget(self.open_btn)
        else:
            self.save_btn = QPushButton("Save Files")  # FIX: Create save_btn for save mode
            self.save_btn.clicked.connect(self.saveFiles)
            button_layout.addWidget(self.save_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def loadLastDirectory(self):
        settings = QSettings("TestEditor", "Settings")
        last_dir = settings.value("last_directory", "")
        if last_dir and os.path.exists(last_dir):
            self.data_directory = last_dir
            self.dir_label.setText(last_dir)
            self.populateModules()

    def saveLastDirectory(self):
        if hasattr(self, 'data_directory'):
            settings = QSettings("TestEditor", "Settings")
            settings.setValue("last_directory", self.data_directory)

    def selectDataDirectory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Data Directory")
        if directory:
            self.data_directory = directory
            self.dir_label.setText(directory)
            self.saveLastDirectory()
            self.populateModules()

    def populateModules(self):
        self.module_list.clear()
        self.scenario_list.clear()
        self.status_label.setText("Select a test scenario file to see matching object file status")

        if not hasattr(self, 'data_directory'):
            return

        test_suites_path = os.path.join(self.data_directory, "testSuites")
        obj_repos_path = os.path.join(self.data_directory, "objectRepositories")

        if not os.path.exists(test_suites_path) or not os.path.exists(obj_repos_path):
            QMessageBox.warning(self, "Error", "Required directories not found")
            return

        # Get modules and sort them
        modules = set()
        for path in [test_suites_path, obj_repos_path]:
            if os.path.exists(path):
                for item in os.listdir(path):
                    if os.path.isdir(os.path.join(path, item)):
                        modules.add(item)

        modules_list = sorted(list(modules), key=lambda x: x.lower())
        for module in modules_list:
            self.module_list.addItem(module)

    def getDataDirectory(self):
        return self.data_directory if hasattr(self, 'data_directory') else None

    def onModuleSelected(self):
        self.scenario_list.clear()
        self.status_label.setText("")  # Clear the status message initially
        self.status_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border: 1px solid #ccc;")

        # FIX: Check which button exists before trying to disable it
        if self.mode == "open" and hasattr(self, 'open_btn'):
            self.open_btn.setEnabled(False)
        elif self.mode == "save" and hasattr(self, 'save_btn'):
            # For save mode, we can always enable save button if a module is selected
            self.save_btn.setEnabled(True)

        if not self.module_list.selectedItems():
            return

        self.selected_module = self.module_list.selectedItems()[0].text()

        # Populate test scenario files
        test_suites_path = os.path.join(self.data_directory, "testSuites", self.selected_module)
        if os.path.exists(test_suites_path):
            for file in os.listdir(test_suites_path):
                if file.endswith('.xlsx'):
                    item = QListWidgetItem(file)
                    item.file_path = os.path.join(test_suites_path, file)
                    self.scenario_list.addItem(item)

    def onScenarioSelected(self):
        if not self.scenario_list.selectedItems():
            self.status_label.setText("Select a test scenario file to see matching object file status")
            self.status_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border: 1px solid #ccc;")

            # FIX: Check which button exists before trying to disable it
            if self.mode == "open" and hasattr(self, 'open_btn'):
                self.open_btn.setEnabled(False)
            return

        selected_scenario = self.scenario_list.selectedItems()[0]
        scenario_file = selected_scenario.file_path
        scenario_filename = os.path.basename(scenario_file)

        # Find matching object repository file
        obj_repos_path = os.path.join(self.data_directory, "objectRepositories", self.selected_module)
        matching_obj_file = self.findMatchingObjectFile(scenario_filename, obj_repos_path)

        if matching_obj_file:
            # Only show success message if we're in open mode and button exists
            if self.mode == "open" and hasattr(self, 'open_btn'):
                self.status_label.setText("")
                self.status_label.setStyleSheet("padding: 10px; background-color: #e8f5e8; border: 1px solid #4caf50; color: #2e7d32;")
                self.open_btn.setEnabled(True)
            else:
                # In save mode or when open_btn doesn't exist, show minimal info or nothing
                self.status_label.setText("")
                self.status_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border: 1px solid #ccc;")

            self.selected_scenario_file = scenario_file
            self.selected_obj_file = matching_obj_file

            # For save mode, ensure save button is enabled when files are available
            if self.mode == "save" and hasattr(self, 'save_btn'):
                self.save_btn.setEnabled(True)
        else:
            # Show warning only when object file is NOT available
            self.status_label.setText(
                f"Ensure the object repository {scenario_filename} exists."
            )
            self.status_label.setStyleSheet("padding: 10px; background-color: #fff3cd; border: 1px solid #ffc107; color: #856404;")

            # Disable open button if object file is missing
            if self.mode == "open" and hasattr(self, 'open_btn'):
                self.open_btn.setEnabled(False)

    def findMatchingObjectFile(self, scenario_filename, obj_repos_path):
        """Find the matching object repository file for a given test scenario file"""
        if not os.path.exists(obj_repos_path):
            return None

        # Extract module name from scenario filename
        # Pattern: Automation_Module_ModuleName.xlsx or similar
        module_match = re.search(r'Automation_Module_([^.]+)', scenario_filename)
        if module_match:
            module_name = module_match.group(1)
            expected_obj_file = f"ObjRep_Module_{module_name}_Test.xlsx"
            expected_path = os.path.join(obj_repos_path, expected_obj_file)

            if os.path.exists(expected_path):
                return expected_path

        # Alternative: Look for any object file that might match
        # Try to find files with similar names
        for file in os.listdir(obj_repos_path):
            if file.endswith('.xlsx'):
                # Check if the base names are similar (without prefixes/suffixes)
                scenario_base = scenario_filename.replace('Automation_Module_', '').replace('.xlsx', '')
                obj_base = file.replace('ObjRep_Module_', '').replace('_Test.xlsx', '')

                if scenario_base == obj_base:
                    return os.path.join(obj_repos_path, file)

        return None

    def openFiles(self):
        if not hasattr(self, 'selected_scenario_file') or not hasattr(self, 'selected_obj_file'):
            QMessageBox.warning(self, "Error", "No files selected or matching object file not found")
            return

        if not os.path.exists(self.selected_scenario_file) or not os.path.exists(self.selected_obj_file):
            QMessageBox.warning(self, "Error", "One or both files no longer exist")
            return

        self.accept()

    def saveFiles(self):
        if not self.module_list.selectedItems():
            QMessageBox.warning(self, "Error", "Please select a module to save the files into.")
            return

        self.selected_module = self.module_list.selectedItems()[0].text()

        # Use the module name from the parent if available, otherwise use selected module
        parent = self.parent()
        if hasattr(parent, 'current_module_name') and parent.current_module_name:
            module_name = parent.current_module_name
        else:
            module_name = self.selected_module

        scenario_filename = f"Automation_Module_{module_name}.xlsx"
        obj_filename = f"ObjRep_Module_{module_name}_Test.xlsx"

        self.selected_scenario_file = os.path.join(self.data_directory, "testSuites", self.selected_module, scenario_filename)
        self.selected_obj_file = os.path.join(self.data_directory, "objectRepositories", self.selected_module, obj_filename)

        self.accept()

    def getSelectedFiles(self):
        if hasattr(self, 'selected_scenario_file') and hasattr(self, 'selected_obj_file'):
            return self.selected_scenario_file, self.selected_obj_file
        return None, None

    def closeEvent(self, event):
        self.saveLastDirectory()
        super().closeEvent(event)
