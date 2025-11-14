import os
import re
from pathlib import Path
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QLabel,
    QPushButton, QListWidgetItem, QFileDialog, QMessageBox,
    QSplitter, QWidget
)
from PyQt6.QtCore import Qt, QSettings

class FileUI(QDialog):
    def __init__(self,main = None, mode="open"):
        super().__init__(main)
        self.mode = mode
        self.main = main
        self.e2e_dir = None
        self.data_directory = None
        self.selected_project = None
        self.selected_scenario_path = None
        self.selected_obj_path = None

        self._setup_window()
        self._setup_ui()
        self._load_last_directory()

    def _setup_window(self):
        title = "Select Files" if self.mode == "open" else "Save Files"
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(900, 600)

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.addLayout(self._create_directory_layout())
        layout.addWidget(self._create_files_explorer())
        layout.addLayout(self._create_buttons_section())
        self.setLayout(layout)

    def _create_directory_layout(self):
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(QLabel("E2E Directory: "))

        self.dir_label = QLabel("Not Selected")
        dir_layout.addWidget(self.dir_label)

        self.select_dir = QPushButton("Select Dir")
        self.select_dir.clicked.connect(self._select_e2e_directory)
        dir_layout.addWidget(self.select_dir)

        return dir_layout

    def _create_files_explorer(self):
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._create_projects_panel())
        splitter.addWidget(self._create_modules_panel())
        splitter.setSizes([300, 600])
        return splitter

    def _create_projects_panel(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("Projects"))

        self.projects_list = QListWidget()
        self.projects_list.itemSelectionChanged.connect(self._on_project_selected)
        layout.addWidget(self.projects_list)

        return widget

    def _create_modules_panel(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("Modules"))

        self.modules_list = QListWidget()
        self.modules_list.itemSelectionChanged.connect(self._on_module_selected)
        layout.addWidget(self.modules_list)

        self.status = QLabel("")
        self.status.setStyleSheet(
            "padding: 10px; border-radius: 4px; border: 1px solid #ccc;"
        )
        self.status.setMinimumHeight(50)
        self.status.setMaximumHeight(50)
        self.status.setMaximumWidth(600)
        layout.addWidget(self.status)

        return widget

    def _create_buttons_section(self):
        button_layout = QHBoxLayout()
        if self.mode == "open":
            self.open_btn = QPushButton("Open Files")
            self.open_btn.clicked.connect(self._open_files)
            self.open_btn.setEnabled(False)
            button_layout.addWidget(self.open_btn)
        else:
            self.save_files = QPushButton("Save Files")
            self.save_files.clicked.connect(self._save_files)
            self.save_files.setEnabled(False)
            button_layout.addWidget(self.save_files)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)

        return button_layout

    # Directory Management

    def _load_last_directory(self):
        settings = QSettings("TestEditor", "Settings")
        e2e_dir = settings.value("e2e_dir", "")

        if e2e_dir and os.path.exists(e2e_dir):
            self.e2e_dir = e2e_dir
            self.data_directory = os.path.join(e2e_dir, "data")
            self.dir_label.setText(e2e_dir)
            self._populate_projects()
            self._load_last_project()

    def _save_last_directory(self):
        if self.data_directory:
            settings = QSettings("TestEditor", "Settings")
            settings.setValue("e2e_dir", self.e2e_dir)

    def _select_e2e_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            e2e_dir = directory
            data_dir = os.path.join(directory, "data")

            if self._validate_directory(e2e_dir, data_dir):
                self.e2e_dir = directory
                self.data_directory = data_dir

                self.dir_label.setText(self.data_directory)
                self._save_last_directory()
                self._populate_projects()
                self._load_last_project()
            else:
                QMessageBox.warning(self, "Invalid Directory",
                "Selected directory should be the E2E application root with 'data/' subdirectory")

    def _validate_directory(self, e2e_dir, data_dir):
        app_path = Path(e2e_dir)
        data_path = Path(data_dir)

        has_pom = (app_path / "pom.xml").exists()
        has_src = (app_path / "src").exists()

        has_test_suites = (data_path / "testSuites").exists()
        has_obj_repos = (data_path / "objectRepositories").exists()

        return has_pom and has_src and has_test_suites and has_obj_repos

    def _load_last_project(self):
        if self.mode != "open":
            return
        settings = QSettings("TestEditor", "Settings")
        last_project = settings.value("last_project", "")
        if not last_project:
            return

        for i in range(self.projects_list.count()):
            item = self.projects_list.item(i)
            if item.text() == last_project:
                self.projects_list.setCurrentItem(item)
                break

    # Dir Data Population

    def _populate_projects(self):
        self.projects_list.clear()
        self.modules_list.clear()
        self.status.setText("Select a test scenario")

        if not self.data_directory:
            return

        test_suit_path = os.path.join(self.data_directory, "testSuites")
        obj_repo_path = os.path.join(self.data_directory, "objectRepositories")

        if not os.path.exists(test_suit_path):
            QMessageBox.warning(self, "Error", "Required directory testSuites not found")
            return
        elif not os.path.exists(obj_repo_path):
            QMessageBox.warning(self, "Error", "Required directory objectRepositories not found")
            return

        projects = set()
        for path in [test_suit_path, obj_repo_path]:
            if os.path.exists(path):
                for item in os.listdir(path):
                    if os.path.isdir(os.path.join(path, item)):
                        projects.add(item)

        projects_list = sorted(list(projects))
        for project in projects_list:
            self.projects_list.addItem(project)

    def _on_project_selected(self):
        self.modules_list.clear()
        self._reset_status()

        self._set_action_button_state(False)

        if not self.projects_list.selectedItems():
            return

        self.selected_project = self.projects_list.selectedItems()[0].text()

        if self.mode == "save":
            self._set_action_button_state(True)

        test_suit_path = os.path.join(self.data_directory, "testSuites", self.selected_project)

        if os.path.exists(test_suit_path):
            for file in os.listdir(test_suit_path):
                if file.endswith(".xlsx"):
                    item = QListWidgetItem(file)
                    item.file_path = os.path.join(test_suit_path, file)
                    self.modules_list.addItem(item)

    def _on_module_selected(self):
        if not self.modules_list.selectedItems():
            self._reset_status()
            self._set_action_button_state(False)
            return

        selected_scenario = self.modules_list.selectedItems()[0]
        scenario_path = selected_scenario.file_path

        obj_path = os.path.join(self.data_directory, "objectRepositories", self.selected_project)
        matching_obj_path = self._find_matching_object_file(scenario_path, obj_path)

        if matching_obj_path:
            self._handle_matching_file_found(scenario_path, matching_obj_path)
        else:
            self._handle_matching_file_not_found(os.path.basename(scenario_path))

    def _find_matching_object_file(self, scenario_path, obj_repo_path):
        if not os.path.exists(obj_repo_path):
            return None

        scenario_file = os.path.basename(scenario_path)
        module_match = re.search(r'Automation_Module_([^.]+)\.xlsx', scenario_file)

        if module_match:
            module_name = module_match.group(1)
            expected_obj_file = f"ObjRep_Module_{module_name}_Test.xlsx"
            expected_path = os.path.join(obj_repo_path, expected_obj_file)

            if os.path.exists(expected_path):
                return expected_path

        return None

    # Status

    def _reset_status(self):
        self.status.setText("Select a test scenario")
        self.status.setStyleSheet("padding: 10px; border-radius: 4px; border: 1px solid #ccc;")

    def _set_action_button_state(self, enabled):
        if self.mode == "open" and hasattr(self, "open_btn"):
            self.open_btn.setEnabled(enabled)
        if self.mode == "save" and hasattr(self, "save_files"):
            self.save_files.setEnabled(enabled)

    def _handle_matching_file_found(self, scenario_path, obj_path):
        self.selected_scenario_path = scenario_path
        self.selected_obj_path = obj_path
        if self.mode == "open":
            self.status.setText("Matching object file found")
            self.status.setStyleSheet(
                "padding: 10px; border-radius: 4px; border: 1px solid #4caf50; color: #2e7d32;"
            )
        else:
            self._reset_status()

        self._set_action_button_state(True)

    def _handle_matching_file_not_found(self, scenario_filename):
        module_name = scenario_filename.replace('Automation_Module_', '').replace('.xlsx', '')
        self.status.setText(
            f"Expected: ObjRep_Module_{module_name}_Test.xlsx not found"
        )
        self.status.setStyleSheet(
            "padding: 10px; border-radius: 4px; border: 1px solid #ffc107; color: #856404;"
        )

        if self.mode == "open":
            self._set_action_button_state(False)

    # File operations

    def _open_files(self):
        if not self.selected_scenario_path or not self.selected_obj_path:
            QMessageBox.warning(self, "Error", "No files selected or objrepo not found")
            return

        if not os.path.exists(self.selected_scenario_path) or not os.path.exists(self.selected_obj_path):
            QMessageBox.warning(self, "Error", "One or more files no longer exist")
            return

        if self.projects_list.selectedItems():
            settings = QSettings("TestEditor", "Settings")
            last_project = self.projects_list.selectedItems()[0].text()
            settings.setValue("last_project", last_project)

        self.accept()

    def _save_files(self):
        if not self.projects_list.selectedItems():
            QMessageBox.warning(self, "Error", "Select the project to save files into")
            return

        self.selected_project = self.projects_list.selectedItems()[0].text()

        tab = self.main.get_current_tab()
        module_name = tab.module_name


        scenario_filename = f"Automation_Module_{module_name}.xlsx"
        obj_filename = f"ObjRep_Module_{module_name}_Test.xlsx"

        self.selected_scenario_path = os.path.join(
            self.data_directory, "testSuites",
            self.selected_project, scenario_filename
        )
        self.selected_obj_path = os.path.join(
            self.data_directory, "objectRepositories",
            self.selected_project, obj_filename
        )

        self.accept()

    def get_selected_files_path(self):
        if self.selected_scenario_path and self.selected_obj_path:
            return self.selected_scenario_path, self.selected_obj_path
        return None, None

    def get_data_directory(self):
        return self.data_directory

    # Event Handlers

    def closeEvent(self, event):
        self._save_last_directory()
        super().closeEvent(event)
