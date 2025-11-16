from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtWidgets import (
    QMenuBar,
    QMessageBox,
    QToolBar,
    QToolButton,
    QWidget,
    QHBoxLayout,
    QPushButton,
    QStyle,
)
from core.commands import COMMANDS
from models.test_scenario import TestScenarioModel
from typing import TYPE_CHECKING

from ui.tab import TabWidget
if TYPE_CHECKING:
    from ui.window import MainWindow

class MenuBar(QMenuBar):
    def __init__(self, main: "MainWindow") -> None:
        super().__init__(main)
        self.main = main
        self.file_ops = main.file_ops
        self.create_menu_bar()
        self.create_corner_run_toolbar()

    def create_menu_bar(self) -> None:

        menubar = self.main.menuBar()
        if menubar is None:
            return

        self._create_file_menu(menubar)
        self._create_edit_menu(menubar)

        auto_adjust_action = QAction("Auto Adjust", self.main)
        auto_adjust_action.triggered.connect(self.auto_adjust_cells)
        auto_adjust_action.setShortcut("Ctrl+A")
        auto_adjust_action.setToolTip("Auto adjust columns")
        menubar.addAction(auto_adjust_action)

        fixed_width_action = QAction("Preset Width", self.main)
        fixed_width_action.triggered.connect(self.apply_fixed_widths)
        fixed_width_action.setShortcut("Ctrl+Shift+A")
        fixed_width_action.setToolTip("Apply default fixed widths")
        menubar.addAction(fixed_width_action)

        generate_action = QAction("Generate Test Cases", self.main)
        generate_action.triggered.connect(self.generate_test_cases_action)
        generate_action.setShortcut("Ctrl+G")
        generate_action.setToolTip("Generate Test Cases")
        menubar.addAction(generate_action)


    def _create_file_menu(self, menubar: QMenuBar) -> None:
        file_menu = menubar.addMenu("File")
        if file_menu is None:
            return

        new_action = QAction("New File", self.main)
        new_action.triggered.connect(self.file_ops.new_file)
        new_action.setShortcut("Ctrl+N")
        file_menu.addAction(new_action)

        load_action = QAction("Open File", self.main)
        load_action.triggered.connect(self.file_ops.open_file)
        load_action.setShortcut("Ctrl+E")
        file_menu.addAction(load_action)

        file_menu.addSeparator()

        save_action = QAction("Save File", self.main)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.file_ops.save_file)
        file_menu.addAction(save_action)

        save_as_action = QAction("Save File As", self.main)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.file_ops.save_file_as)
        file_menu.addAction(save_as_action)

        file_menu.addSeparator()

    def _create_edit_menu(self, menubar: QMenuBar) -> None:
        edit_menu = menubar.addMenu("Edit")
        if edit_menu is None:
            return

        edit_menu.addAction(self.main.undo_action)
        edit_menu.addAction(self.main.redo_action)

        edit_menu.addSeparator()
        find_action = QAction("Find/Replace", self.main)
        find_action.setShortcut(QKeySequence.StandardKey.Find)
        find_action.triggered.connect(self.main._show_find_dialog)
        edit_menu.addAction(find_action)

    def generate_test_cases_action(self) -> None:
        current_tab = self.main.get_current_tab()
        if not current_tab and hasattr(current_tab, 'controller'):
            return
        current_table = self.main.get_current_table()
        if not current_table:
            return

        model = current_table.model

        for row in range(model.rowCount()):
            command = model.data(model.index(row, TestScenarioModel.COMMAND_COL), 0)
            if command in COMMANDS:
                steps_data = model.data(model.index(row, TestScenarioModel.STEPS_COL), 0)
                expected_data = model.data(model.index(row, TestScenarioModel.EXPECTED_COL), 0)
                if steps_data or expected_data:
                    overwrite_needed = True
                    break

        should_proceed = True
        should_overwrite = False
        if overwrite_needed:
            msg_box = QMessageBox(self.main)
            msg_box.setWindowTitle("Generate Test Cases")
            msg_box.setText("Some test cases already have data.")
            msg_box.setInformativeText("How would you like to proceed?")

            reset_button = msg_box.addButton("Reset All", QMessageBox.ButtonRole.DestructiveRole)
            fill_empty_button = msg_box.addButton("Fill Empty", QMessageBox.ButtonRole.AcceptRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)

            msg_box.setDefaultButton(fill_empty_button)
            msg_box.exec()

            clicked_button = msg_box.clickedButton()
            if clicked_button == reset_button:
                should_overwrite = True
            elif clicked_button == fill_empty_button:
                should_overwrite = False
            else:
                should_proceed = False
                self.main.show_status_message("Operation cancelled.", "info", 3000)

            if should_proceed:
                num_generated = current_tab.controller.generate_test_cases(model, overwrite=should_overwrite)
                if num_generated > 0:
                    message = f"Successfully generated {num_generated} test case(s)."
                    self.main.show_status_message(message, "success", 5000)
                else:
                    message = "No test cases were generated."
                    self.main.show_status_message(message, "info", 5000)

    def auto_adjust_cells(self) -> None:
        table = self.main.get_current_table()
        if table:
            table.auto_adjust_cells()

    def apply_fixed_widths(self) -> None:
        tab = self.main.get_current_tab()
        if isinstance(tab, TabWidget) and tab.inner_tabs.currentIndex() == 0:
            scenario_table_view = tab.table1.table

            scenario_table_view.setColumnWidth(0, 50)  # Type
            scenario_table_view.setColumnWidth(1, 130) # ID
            scenario_table_view.setColumnWidth(2, 50)  # Skip
            scenario_table_view.setColumnWidth(3, 300) # Description
            scenario_table_view.setColumnWidth(4, 350) # Steps Performed
            scenario_table_view.setColumnWidth(5, 50) # Expected Results
            scenario_table_view.setColumnWidth(6, 250) # Command
            for col_index in range(7, 12):
                scenario_table_view.setColumnWidth(col_index, 250)
            self.main.status_bar.clearMessage()

        else:
            self.main.show_status_message(
                "Fixed widths only apply to the active TestScenario table.",
                message_type="warning",
                timeout=5000
            )

    def create_corner_run_toolbar(self):
        toolbar_widget = QWidget()
        toolbar_widget.setObjectName
        toolbar_layout = QHBoxLayout(toolbar_widget)
        toolbar_layout.setContentsMargins(5, 0, 5, 0)
        toolbar_layout.setSpacing(1)

        self.run_button = QPushButton("▶")
        self.run_button.setToolTip("Run Current Module (Ctrl+R)")
        self.run_button.clicked.connect(self.run_current_module)
        self.run_button.setFlat(True)
        self.run_button.setObjectName("RunButton")
        toolbar_layout.addWidget(self.run_button)

        self.stop_button = QPushButton("■")
        self.stop_button.setToolTip("Stop Execution (Ctrl+Shift+R)")
        self.stop_button.clicked.connect(self.stop_test)
        self.stop_button.setFlat(True)
        self.stop_button.setObjectName("StopButton")
        self.stop_button.setEnabled(False)
        toolbar_layout.addWidget(self.stop_button)

        self.config_button = QPushButton("⚙")
        self.config_button.setToolTip("Run Configuration")
        self.config_button.clicked.connect(self.show_run_config)
        self.config_button.setFlat(True)
        self.config_button.setObjectName("ConfigButton")
        toolbar_layout.addWidget(self.config_button)

        toolbar_layout.addStretch()

        self.main.menuBar().setCornerWidget(toolbar_widget, Qt.Corner.TopRightCorner)

    def run_current_module(self):
        self.main.run_current_module()

    def stop_test(self):
        self.main.stop_test()

    def show_run_config(self):
        self.main.show_run_config()
