from PyQt6.QtWidgets import QTabWidget, QMessageBox, QVBoxLayout, QWidget, QAbstractItemView
from PyQt6.QtCore import pyqtSignal, Qt, QModelIndex
from PyQt6.QtGui import QUndoStack

from models import TestScenarioModel, ObjRepoModel
from .table import Table

from .delegates import ComboBoxDelegate
from core import TableController

class TabWidget(QWidget):
    closeRequested = pyqtSignal(object)

    def __init__(self, module_name, main, project_name=None):
        super().__init__()
        self.main = main
        self.module_name = module_name
        self.project_name = project_name
        self.scenario_path = None
        self.obj_path = None

        self.undo_stack = QUndoStack(self)

        self.model1 = TestScenarioModel()
        self.model2 = ObjRepoModel()

        commands = self.main.command_manager.get_command_names()
        colors = self.main.get_current_colors()

        self.delegate = ComboBoxDelegate(self, self.undo_stack, commands, colors)

        self.table1 = Table(model=self.model1, undo_stack=self.undo_stack, delegate=self.delegate, parent_tab=self)
        self.table2 = Table(model=self.model2, undo_stack=self.undo_stack, delegate=self.delegate)

        self.controller = TableController(main, self.undo_stack)

        self.setup_ui()
        self.connect_signals()

        self.apply_zoom(self.main.current_font_size)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.inner_tabs = QTabWidget()
        self.inner_tabs.addTab(self.table1, "TestScenario")
        self.inner_tabs.addTab(self.table2, "Objects")

        layout.addWidget(self.inner_tabs)

    def connect_signals(self):
        self.table1.insertRowsRequested.connect(lambda p, c: self.controller.insert_rows(self.model1, p, c))
        self.table1.deleteRowsRequested.connect(lambda r: self.controller.delete_rows(self.model1, r))
        self.table1.copyRequested.connect(lambda s: self.controller.copy(self.model1, s))
        self.table1.pasteRequested.connect(lambda i: self.controller.paste(self.model1, i))
        self.table1.cutRequested.connect(lambda s: self.controller.cut(self.model1, s))
        self.table1.clearRequested.connect(lambda s: self.controller.clear(self.model1, s))

        self.table2.insertRowsRequested.connect(lambda p, c: self.controller.insert_rows(self.model2, p, c))
        self.table2.deleteRowsRequested.connect(lambda r: self.controller.delete_rows(self.model2, r))
        self.table2.copyRequested.connect(lambda s: self.controller.copy(self.model2, s))
        self.table2.pasteRequested.connect(lambda i: self.controller.paste(self.model2, i))
        self.table2.cutRequested.connect(lambda s: self.controller.cut(self.model2, s))
        self.table2.clearRequested.connect(lambda s: self.controller.clear(self.model2, s))
        self.model1.dataChanged.connect(self.update_tab_text)
        self.model2.dataChanged.connect(self.update_tab_text)
        self.model2.dataChanged.connect(self.on_object_model_changed)
        self.table1.skipRequested.connect(lambda r: self.controller.skip(self.model1, r))
        self.table1.unskipRequested.connect(lambda r: self.controller.unskip(self.model1, r))

        self.undo_stack.cleanChanged.connect(self.update_tab_text)

    def mark_saved(self):
        self.undo_stack.setClean()

    def update_tab_text(self):
        tab = self.main.main_tab
        if tab:
            index = tab.indexOf(self)
            if self.project_name:
                base_name = f"{self.project_name} | {self.module_name}"
            else:
                base_name = self.module_name or "Untitled"

            if not self.undo_stack.isClean():
                tab.setTabText(index, f"{base_name} •")
            else:
                tab.setTabText(index, base_name)

    def close_tab(self):
        if not self.undo_stack.isClean():
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                f"Save changes to {self.module_name} before closing?",
                QMessageBox.StandardButton.Save |
                QMessageBox.StandardButton.Cancel |
                QMessageBox.StandardButton.Discard
            )
            if reply == QMessageBox.StandardButton.Save:
                if not self.main.file_ops.save_file():
                    return False
            elif reply == QMessageBox.StandardButton.Cancel:
                return False
        return True

    def apply_zoom(self, font_size):

        style = f"font-size: {font_size}px;"

        self.table1.table.setStyleSheet(style)
        self.table2.table.setStyleSheet(style)

        self.table1.table.horizontalHeader().setStyleSheet(style)
        self.table1.table.verticalHeader().setStyleSheet(style)
        self.table2.table.horizontalHeader().setStyleSheet(style)
        self.table2.table.verticalHeader().setStyleSheet(style)

        self.table1.auto_adjust_cells()
        self.table2.auto_adjust_cells()

    def update_delegate_commands(self, commands):
        if hasattr(self.delegate, 'update_command_list'):
             self.delegate.update_command_list(commands)
        elif hasattr(self.delegate, 'commands'):
            self.delegate.commands = commands

    def on_object_model_changed(self):
        self.delegate.invalidate_objects_cache()

    def is_object_in_repository(self, object_name):
        if not object_name:
            return False

        object_name_column_in_obj_repo = 1 # Assuming 'User friendly name of Object' is the 2nd column (0-indexed)
        found_indexes = self.model2.match(
            self.model2.index(0, object_name_column_in_obj_repo),
            Qt.ItemDataRole.DisplayRole,
            object_name,
            1, # Max 1 match
            Qt.MatchFlag.MatchExactly
        )
        return bool(found_indexes)

    def go_to_object_definition(self, scenario_index: QModelIndex):
        object_name = self.model1.data(scenario_index, Qt.ItemDataRole.DisplayRole)
        if not object_name:
            self.main.show_status_message("Selected cell does not contain an object name.", "warning", 3000)
            return

        # Search for the object in the object repository model (self.model2)
        # Assuming 'Object Name' is the first column in ObjRepoModel
        object_name_column_in_obj_repo = 1 # Assuming 'User friendly name of Object' is the 2nd column (0-indexed)
        found_indexes = self.model2.match(
            self.model2.index(0, object_name_column_in_obj_repo),
            Qt.ItemDataRole.DisplayRole,
            object_name,
            1, # Max 1 match
            Qt.MatchFlag.MatchExactly
        )

        if found_indexes:
            found_index = found_indexes[0]
            # Switch to the Objects tab
            self.inner_tabs.setCurrentIndex(1) # Assuming Objects tab is at index 1

            # Select and scroll to the found object
            self.table2.table.selectionModel().clearSelection()
            self.table2.table.setCurrentIndex(found_index)
            self.table2.table.scrollTo(found_index, QAbstractItemView.ScrollHint.PositionAtCenter)
        else:
            self.main.show_status_message(f"Object '{object_name}' not found in the Object Repository.", "warning", 3000)
