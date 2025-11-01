from PyQt6.QtWidgets import (QMainWindow, QWidget, QTabWidget, QVBoxLayout, QMessageBox, QLabel, QStatusBar, QPushButton, QFrame, QTableView)
from PyQt6.QtCore import Qt, QTimer, QSettings, QModelIndex
from PyQt6.QtGui import QAction, QKeySequence

import re
from .menu import MenuBar
from .tab import TabWidget
from .find_replace import FindReplace
from .styles import THEMES
from models import TableModel
from core import FileController
from core.command_manager import CommandManager
from core.undo_commands import EditCellCommand


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Editor")
        self.resize(1600, 900)
        self.showMaximized()

        self.current_font_size = 12
        self.base_font_size = self.current_font_size
        self.current_connected_stack = None
        self.last_find_index = QModelIndex()

        self.main_tab = QTabWidget()
        self.main_tab.setTabsClosable(True)
        self.main_tab.tabCloseRequested.connect(self.close_tab)
        self.main_tab.currentChanged.connect(self.on_tab_changed)

        self.welcome_widget = self._create_welcome_widget()
        self.file_ops = FileController(self)
        self.create_undo_redo()
        self.command_manager = CommandManager()
        settings = QSettings("TestEditor", "Settings")
        saved_path = settings.value(self.command_manager.SETTINGS_KEY, "")
        self.command_manager.load_commands(saved_path)
        self.command_manager.commandsReloaded.connect(self._update_delegates_command_list)

        self.menu_bar = MenuBar(self)
        self.setup_status_bar()
        self.center_status_label.setObjectName("StatusLabel")
        self.center_status_label.hide()
        self._setup_find_and_replace()

        self.setCentralWidget(self.main_tab)
        self.show_welcome_screen()

    def create_new_tab(self, module_name=None, scenario_path=None, obj_path=None, project_name=None):
        self.hide_welcome_screen()

        existing_tab = self.find_existing_tab(project_name, module_name)
        if existing_tab:
            index = self.main_tab.indexOf(existing_tab)
            self.main_tab.setCurrentIndex(index)
            QMessageBox.information(
                self,
                "Duplicate File",
                "File already open, Switching to existing tab"
            )
            return existing_tab

        tab = TabWidget(module_name, self, project_name=project_name)

        if scenario_path and obj_path:
            tab.scenario_path = scenario_path
            tab.obj_path = obj_path

        index = self.main_tab.addTab(tab, "")
        self.main_tab.setCurrentIndex(index)
        tab.update_tab_text()
        '''if project_name:
            self.setWindowTitle(f"Test Case Editor - {project_name} | {module_name}")
        else:
            self.setWindowTitle(f"Test Case Editor - {module_name}")'''
        self.setWindowTitle("Test Case Editor")
        self.on_tab_changed(index)
        return tab

    def find_existing_tab(self, project_name, module_name):
        if project_name and module_name:
            for i in range(self.main_tab.count()):
                tab = self.main_tab.widget(i)
                if (tab.project_name == project_name and tab.module_name == module_name):
                    return tab
        return None

    def close_tab(self, index):
        tab = self.main_tab.widget(index)

        if tab.close_tab():
            self.main_tab.removeTab(index)
            tab.deleteLater()

            if self.main_tab.count() == 0:
                self.show_welcome_screen()

    def on_tab_changed(self, index):
        if self.current_connected_stack:
            try:
                self.undo_action.triggered.disconnect(self.current_connected_stack.undo)
            except TypeError:
                pass
            try:
                self.redo_action.triggered.disconnect(self.current_connected_stack.redo)
            except TypeError:
                pass
            try:
                self.current_connected_stack.canUndoChanged.disconnect(self.undo_action.setEnabled)
            except TypeError:
                pass
            try:
                self.current_connected_stack.canRedoChanged.disconnect(self.redo_action.setEnabled)
            except TypeError:
                pass
            try:
                self.current_connected_stack.indexChanged.disconnect(self.update_undo_status)
            except TypeError:
                pass

        new_stack = None

        if index >= 0:
            tab = self.main_tab.widget(index)
            if isinstance(tab, TabWidget):
                stack = tab.undo_stack

                self.undo_action.triggered.connect(stack.undo)
                self.redo_action.triggered.connect(stack.redo)
                stack.canUndoChanged.connect(self.undo_action.setEnabled)
                stack.canRedoChanged.connect(self.redo_action.setEnabled)
                self.undo_action.setEnabled(stack.canUndo())
                self.redo_action.setEnabled(stack.canRedo())
                stack.indexChanged.connect(self.update_undo_status)

                new_stack = stack
            else:
                self.undo_action.setEnabled(False)
                self.redo_action.setEnabled(False)
                self.status_bar.clearMessage()
        else:
            self.undo_action.setEnabled(False)
            self.redo_action.setEnabled(False)
            self.status_bar.clearMessage()

        self.current_connected_stack = new_stack
        self.update_undo_status()


    def setup_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.undo_status_label = QLabel("")
        self.undo_status_label.setMinimumWidth(250)
        self.undo_status_label.setMaximumWidth(250)
        self.undo_status_label.setStyleSheet("padding-left: 10px; padding-right: 10px;")

        self.center_status_label = QLabel("")
        self.center_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.center_status_label.setStyleSheet("padding-left: 10px; padding-right: 10px;")
        self.center_status_label.hide()

        self.zoom_out_button = QPushButton("-")
        self.zoom_out_button.setObjectName("StatusBarButton")
        self.zoom_out_button.setToolTip("Zoom Out (Ctrl+-)")
        self.zoom_out_button.clicked.connect(self.zoom_out)
        self.zoom_out_button.setShortcut("Ctrl+-")

        self.zoom_label = QLabel("100%")
        self.zoom_label.setToolTip("Current Zoom Level")
        self.zoom_label.setStyleSheet("padding-left: 5px; padding-right: 5px;")

        self.zoom_in_button = QPushButton("+")
        self.zoom_in_button.setObjectName("StatusBarButton")
        self.zoom_in_button.setToolTip("Zoom In (Ctrl+=)")
        self.zoom_in_button.clicked.connect(self.zoom_in)
        self.zoom_in_button.setShortcut("Ctrl+=")

        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.VLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)

        self.reset_zoom_button = QPushButton("Reset")
        self.reset_zoom_button.setObjectName("StatusBarButton")
        self.reset_zoom_button.setToolTip("Reset Zoom (Ctrl+0)")
        self.reset_zoom_button.clicked.connect(self.reset_zoom)
        self.reset_zoom_button.setShortcut("Ctrl+0")

        self.status_bar.insertWidget(0, self.undo_status_label)

        self.status_bar.addPermanentWidget(self.zoom_out_button)
        self.status_bar.addPermanentWidget(self.zoom_label)
        self.status_bar.addPermanentWidget(self.zoom_in_button)
        self.status_bar.addPermanentWidget(separator)
        self.status_bar.addPermanentWidget(self.reset_zoom_button)

        self.status_bar.insertWidget(1, self.center_status_label, 1)

        self.update_zoom_label()

    def update_undo_status(self):
        if self.current_connected_stack and self.current_connected_stack.canUndo():
            undo_text = self.current_connected_stack.undoText()
            self.undo_status_label.setText(f"Undo: {undo_text}")
        else:
            self.undo_status_label.clear()

    def show_status_message(self, message, message_type="info", timeout=5000):
        self.center_status_label.setProperty("message_type", message_type)

        self.center_status_label.style().unpolish(self.center_status_label)
        self.center_status_label.style().polish(self.center_status_label)

        self.center_status_label.setText(message)
        self.center_status_label.show()
        QTimer.singleShot(timeout, self.center_status_label.hide)

    def select_command_file(self):
        if self.command_manager.command_file_prompt(self):
            self.show_status_message("Command file reloaded successfully.", "success", 3000)
        else:
             self.show_status_message("Command file selection cancelled or failed.", "warning", 3000)

    def zoom_in(self):
        self.change_zoom(2)

    def zoom_out(self):
        self.change_zoom(-2)

    def reset_zoom(self):
        self.change_zoom(new_size=self.base_font_size)

    def change_zoom(self, delta=0, new_size=None):
        if new_size is not None:
            target_size = new_size
        else:
            target_size = self.current_font_size + delta

        min_size = 6
        max_size = 48
        target_size = max(min_size, min(target_size, max_size))

        if target_size == self.current_font_size:
            return

        self.current_font_size = target_size

        for i in range(self.main_tab.count()):
            widget = self.main_tab.widget(i)
            if isinstance(widget, TabWidget):
                widget.apply_zoom(self.current_font_size)
                self.update_zoom_label()

    def update_zoom_label(self):
        percentage = int((self.current_font_size / self.base_font_size) * 100)
        self.zoom_label.setText(f"{percentage}%")

    def get_current_tab(self):
        current_index = self.main_tab.currentIndex()
        if current_index >= 0:
            return self.main_tab.widget(current_index)
        return None

    def get_current_models(self):
        tab = self.get_current_tab()
        if tab:
            return tab.model1, tab.model2
        return None, None

    def get_current_table(self):
        tab = self.get_current_tab()
        if tab and isinstance(tab, TabWidget):
            current_inner_index = tab.inner_tabs.currentIndex()
            return tab.table1 if current_inner_index == 0 else tab.table2
        return None

    def create_undo_redo(self):
        self.undo_action = QAction("Undo", self)
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.setEnabled(False)

        self.redo_action = QAction("Redo", self)
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        self.redo_action.setEnabled(False)

    def _create_welcome_widget(self):
        welcome_widget = QWidget()
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        logo = QLabel()
        logo.setText("Test Case Studio")
        logo.setProperty("class", "welcome-logo")
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)

        instructions = QLabel("Create a new file or open an existing one to get started")
        instructions.setProperty("class", "welcome-instructions")
        instructions.setAlignment(Qt.AlignmentFlag.AlignCenter)

        tips = QLabel("• Use Ctrl+N to create a new test case\n• Use Ctrl+O to open existing test cases\n• Use Ctrl+R to auto-adjust cell sizes")
        tips.setProperty("class", "welcome-tips")
        tips.setAlignment(Qt.AlignmentFlag.AlignLeft)

        layout.addWidget(logo)
        layout.addWidget(instructions)
        layout.addWidget(tips)
        welcome_widget.setLayout(layout)
        return welcome_widget

    def show_welcome_screen(self):
        if self.main_tab.count() == 0:
            self.main_tab.addTab(self.welcome_widget, "Home")
            self.main_tab.setTabsClosable(False)
            self.setWindowTitle("Test Case Editor")

    def hide_welcome_screen(self):
        for i in range(self.main_tab.count()):
            if self.main_tab.widget(i) == self.welcome_widget:
                self.main_tab.removeTab(i)
                self.main_tab.setTabsClosable(True)
                break

    def _update_delegates_command_list(self, new_command_list):
        print("Updating command lists in open tabs...")
        for i in range(self.main_tab.count()):
            widget = self.main_tab.widget(i)
            if isinstance(widget, TabWidget):
                widget.update_delegate_commands(new_command_list)

    def _setup_find_and_replace(self):
        self.find_dialog = FindReplace(self)
        self.find_dialog.hide()

        self.find_dialog.findNextClicked.connect(self.on_find_next)
        self.find_dialog.findPrevClicked.connect(self.on_find_prev)
        self.find_dialog.replaceClicked.connect(self.on_replace)
        self.find_dialog.replaceAllClicked.connect(self.on_replace_all)

    def _show_find_dialog(self):
        self.last_find_index = QModelIndex()

        main_geo = self.geometry()

        dialog_size = self.find_dialog.sizeHint()
        status_bar_height = self.status_bar.height()

        margin = 10
        new_x = main_geo.left() + margin
        new_y = main_geo.bottom() - dialog_size.height() - status_bar_height - margin
        self.find_dialog.move(new_x, new_y)

        self.find_dialog.show()
        self.find_dialog.activateWindow()
        self.find_dialog.find_input.setFocus()
        self.find_dialog.find_input.selectAll()

    def on_find_prev(self, find_text, match_cell):
        current_table = self.get_current_table()
        if not current_table or not find_text:
            return

        model = current_table.model

        start_index = self.last_find_index
        if not start_index.isValid():
            start_index = model.index(model.rowCount(), 0)

        prev_match, wrapped  = model.find_prev(find_text, start_index, match_cell)

        if prev_match.isValid():
            self.last_find_index = prev_match
            current_table.table.scrollTo(prev_match, QTableView.ScrollHint.PositionAtCenter)
            current_table.table.setCurrentIndex(prev_match)
        else:
            self.last_find_index = QModelIndex()
            self.show_status_message(f"No Instance of '{find_text}' found.", "info", 5000)

    def on_find_next(self, find_text, match_cell):
        current_table = self.get_current_table()
        if not current_table or not find_text:
            return

        model = current_table.model

        start_index = self.last_find_index
        if not start_index.isValid():
            start_index = model.index(-1, -1)

        next_match, wrapped = model.find_next(find_text, start_index, match_cell)

        if next_match.isValid():
            self.last_find_index = next_match
            current_table.table.scrollTo(next_match, QTableView.ScrollHint.PositionAtCenter)
            current_table.table.setCurrentIndex(next_match)
        else:
            self.last_find_index = QModelIndex()
            self.show_status_message(f"No Instance of '{find_text}' found.", "info", 5000)

    def on_replace(self, find_text, replace_text, match_cell):
        current_table = self.get_current_table()
        current_tab = self.get_current_tab()

        if not current_table  or not find_text:
            return

        current_index = current_table.table.currentIndex()
        if not current_index.isValid():
            self.on_find_next(find_text, match_cell)
            return

        model = current_table.model
        current_text = model.data(current_index, Qt.ItemDataRole.DisplayRole) or ""

        matches = False
        if match_cell:
            matches = (find_text.lower() == current_text.lower())
        else:
            matches = (find_text.lower() in current_text.lower())

        if matches:
            old_value = current_text
            if match_cell:
                new_value = replace_text
            else:
                new_value = re.sub(find_text, replace_text, old_value, flags=re.IGNORECASE)

            cmd = EditCellCommand(model, current_index, new_value, old_value)
            current_tab.undo_stack.push(cmd)

            self.last_find_index = current_index
            self.on_find_next(find_text, match_cell)


    def on_replace_all(self, find_text, replace_text, match_cell):
        current_table = self.get_current_table()
        current_tab = self.get_current_tab()
        if not current_table or not current_tab or not find_text:
            return

        model = current_table.model
        matches = model.find_all(find_text, match_cell)

        if not matches:
            self.show_status_message(f"No occurrences of '{find_text}' found.", "info", 5000)
            return

        current_tab.undo_stack.beginMacro(f"Replace All '{find_text}'")

        for index in matches:
            old_value = model.data(index, Qt.ItemDataRole.DisplayRole) or ""

            if match_cell:
                new_value = replace_text
            else:
                new_value = re.sub(find_text, replace_text, old_value, flags=re.IGNORECASE)

            if new_value != old_value:
                cmd = EditCellCommand(model, index, new_value, old_value)
                current_tab.undo_stack.push(cmd)

        current_tab.undo_stack.endMacro()
        self.show_status_message(f"Replaced {len(matches)} occurrences.", "success", 3000)

    def closeEvent(self, event):
        unsaved_tabs = []
        for i in range(self.main_tab.count()):
            widget = self.main_tab.widget(i)
            if isinstance(widget, TabWidget) and not widget.undo_stack.isClean():
                unsaved_tabs.append(widget)

        if unsaved_tabs:
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setText("You have unsaved changes.")
            msg_box.setInformativeText("Do you want to discard your changes and exit?")
            discard_button = msg_box.addButton("Discard", QMessageBox.ButtonRole.DestructiveRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            msg_box.setDefaultButton(cancel_button)
            msg_box.exec()

            if msg_box.clickedButton() == discard_button:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def get_current_colors(self):
        return THEMES.get("dark", THEMES["dark"])







