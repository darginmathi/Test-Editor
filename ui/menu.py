from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import QMenuBar
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
        load_action.setShortcut("Ctrl+O")
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
        generate_action = QAction("Generate Test Cases", self.main)
        generate_action.triggered.connect(self.generate_test_cases_action)
        edit_menu.addAction(generate_action)

        edit_menu.addSeparator()
        find_action = QAction("Find/Replace", self.main)
        find_action.setShortcut(QKeySequence.StandardKey.Find)
        find_action.triggered.connect(self.main._show_find_dialog)
        edit_menu.addAction(find_action)

    def generate_test_cases_action(self) -> None:
        current_tab = self.main.get_current_tab()
        if current_tab and hasattr(current_tab, 'controller'):
            current_table = self.main.get_current_table()
            if current_table:
                num_generated = current_tab.controller.generate_test_cases(current_table.model)
                if num_generated > 0:
                    message = f"Generated {num_generated} test case(s)."
                    self.main.show_status_message(message, "success", 5000)
                else:
                    message = f"No empty fields found to generate."
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
