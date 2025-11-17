import os
from PyQt6.QtWidgets import (QMainWindow, QWidget, QTabWidget, QVBoxLayout, QMessageBox, QLabel, QStatusBar, QPushButton, QFrame, QTableView, QFileDialog)
from PyQt6.QtCore import Qt, QTimer, QModelIndex, QSettings, QDateTime
from PyQt6.QtGui import QAction, QKeySequence, QShortcut
from core.log_finder import LogFinder
from PyQt6.QtWebEngineWidgets import QWebEngineView

import re
from .menu import MenuBar
from .tab import TabWidget
from .find_replace import FindReplace
from .styles import DarkTheme, LightTheme
from models import TableModel
from core import FileController
from core.command_manager import CommandManager
from core.undo_commands import EditCellCommand
from models.run_config import RunConfig, get_base_url
from .output_dock import OutputDock
from core.test_runner import TestRunner
from .run_config import RunConfigDialog
from PyQt6.QtWidgets import QDialog



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Editor")
        self.resize(1600, 900)
        self.showMaximized()

        self.current_font_size = 16
        self.base_font_size = self.current_font_size
        self.current_connected_stack = None
        self.last_find_index = QModelIndex()

        self.settings = QSettings("TestEditor", "Settings")
        self.e2e_dir = self.settings.value("e2e_dir", "")

        self.main_tab = QTabWidget()
        self.main_tab.setTabsClosable(True)
        self.main_tab.tabBar().setMovable(True)
        self.main_tab.tabCloseRequested.connect(self.close_tab)
        self.main_tab.currentChanged.connect(self.on_tab_changed)

        self.welcome_widget = self._create_welcome_widget()
        self.file_ops = FileController(self)
        self.create_undo_redo()
        self.command_manager = CommandManager()
        self.command_manager.commandsReloaded.connect(self._update_delegates_command_list)

        self.menu_bar = MenuBar(self)
        self.setup_status_bar()
        self.center_status_label.setObjectName("StatusLabel")
        self.center_status_label.hide()
        self._setup_find_and_replace()
        self.test_runner = None
        self.output_dock = OutputDock(self)
        # self.output_dock.setFixedHeight(int(self.height() * 0.3))
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.output_dock)
        self.output_dock.open_log_button.clicked.connect(self._on_open_log_manually)
        self.run_settings = RunConfig()
        self.run_settings.load_from_settings()

        self.setup_run_toolbar()

        self.create_shortcuts()
        self.setCentralWidget(self.main_tab)
        self.show_welcome_screen()

    def create_shortcuts(self):
        for i in range(1, 10):
            shortcut = QShortcut(QKeySequence(f"Ctrl+{i}"), self)
            shortcut.activated.connect(lambda idx=i-1: self.switch_main_tab(idx))

        self.shortcut_ctrl_tab = QShortcut(QKeySequence("Ctrl+Tab"), self)
        self.shortcut_ctrl_tab.activated.connect(self.toggle_table)

        self.shortcut_ctrl_w = QShortcut(QKeySequence("Ctrl+W"), self)
        self.shortcut_ctrl_w.activated.connect(lambda: self.close_tab(self.main_tab.currentIndex()))

        shortcut_maximize_dock = QShortcut(QKeySequence("Ctrl+M"), self)
        shortcut_maximize_dock.activated.connect(self.output_dock.toggle_maximize)

        shortcut_next_dock_tab = QShortcut(QKeySequence("Ctrl+PgDown"), self)
        shortcut_next_dock_tab.activated.connect(self.navigate_dock_tabs_next)
        shortcut_prev_dock_tab = QShortcut(QKeySequence("Ctrl+PgUp"), self)
        shortcut_prev_dock_tab.activated.connect(self.navigate_dock_tabs_prev)

        shortcut_log_back = QShortcut(QKeySequence("Alt+Left"), self)
        shortcut_log_back.activated.connect(self.navigate_log_back)
        shortcut_log_forward = QShortcut(QKeySequence("Alt+Right"), self)
        shortcut_log_forward.activated.connect(self.navigate_log_forward)

    def switch_main_tab(self, index):
        if index < self.main_tab.count():
            self.main_tab.setCurrentIndex(index)

    def toggle_table(self):
        current_tab = self.get_current_tab()
        if current_tab and hasattr(current_tab, 'inner_tabs'):
            current_index = current_tab.inner_tabs.currentIndex()
            new_index = 1 - current_index
            current_tab.inner_tabs.setCurrentIndex(new_index)

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
        self.setWindowTitle("Test Case Editor")
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

        if isinstance(tab, TabWidget):
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
                self.current_connected_stack.canUndoChanged.disconnect(self.redo_action.setEnabled)
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
        if self.test_runner and self.test_runner.is_test_running():
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setText("E2E Test running.......")
            msg_box.setInformativeText("Do you want to stop the run and exit?")
            discard_button = msg_box.addButton("Stop and Exit", QMessageBox.ButtonRole.DestructiveRole)
            cancel_button = msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            msg_box.setDefaultButton(cancel_button)
            msg_box.exec()
            if msg_box.clickedButton() == discard_button:
                self.test_runner.stop_test()
                self.test_runner.cleanup()
                event.accept()
            else:
                event.ignore()
                return

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
                if self.test_runner:
                    self.test_runner.stop_test()
                    self.test_runner.cleanup()
                event.accept()
            else:
                event.ignore()
        else:
            if self.test_runner:
                self.test_runner.stop_test()
                self.test_runner.cleanup()
            event.accept()

    def get_current_colors(self):
        return DarkTheme()

    def setup_run_toolbar(self):
        self.output_dock.clear_requested.connect(self.clear_output)
        self.console_shortcut = QShortcut(QKeySequence("Ctrl+`"), self)
        self.console_shortcut.activated.connect(self.toggle_console)
        self.output_dock.hide()

    def toggle_console(self):
        if self.output_dock.isVisible():
            self.output_dock.hide()
            self.centralWidget().show()
        else:
            self.output_dock.show()
            self.output_dock.raise_()

            if self.output_dock.is_dock_maximized():
                self.centralWidget().hide()

    def clear_output(self):
        console = self.output_dock.get_console()
        console.clear()

    def run_current_module(self):
        tab = self.get_current_tab()

        if not tab or not hasattr(tab, 'project_name'):
            self.show_status_message("No test file open", "warning", 3000)
            return

        if not tab.project_name or not tab.module_name:
            self.show_status_message("Current file is not a valid test module", "warning", 3000)
            return

        console = self.output_dock.get_console()

        if self.output_dock.isHidden():
            self.output_dock.show()

        if not self.test_runner:
            if not self.e2e_dir:
                QMessageBox.warning(self, "E2E directory not set", "Set dir in open file menu")
                return
            self.test_runner = TestRunner(self.e2e_dir)
            self.test_runner.output_received.connect(console.append)
            self.test_runner.error_received.connect(console.append)
            self.test_runner.process_finished.connect(self.on_test_finished)
            self.test_runner.process_started.connect(self.on_test_started)

        console.clear()
        self.output_dock.tabs.setCurrentWidget(console)

        url = getattr(self, 'current_run_url', None) or get_base_url(tab.project_name)

        config = RunConfig(
            project_name=tab.project_name,
            module_name=tab.module_name,
            base_url=url,
            browser=self.run_settings.browser,
            video_option=self.run_settings.video_option,
            wait_time=self.run_settings.wait_time
        )

        start_time = QDateTime.currentDateTime()
        project = config.project_name
        module = config.module_name
        logs_dir = os.path.join(self.e2e_dir, "logs")

        self.log_finder = LogFinder(start_time, project, module, logs_dir)
        self.log_finder.log_found.connect(self._on_log_found)
        self.test_runner.run_test(config)


    def stop_test(self):
        if self.test_runner and self.test_runner.is_test_running():
            self.test_runner.stop_test()
        else:
            QMessageBox.information(self, "No Test Running", "No test is currently running")

    def show_run_config(self):
        tab = self.get_current_tab()

        if not tab or not hasattr(tab, 'project_name'):
            self.show_status_message("No test file open", "warning", 3000)
            return

        if not tab.project_name or not tab.module_name:
            self.show_status_message("Current file is not a valid test module", "warning", 3000)
            return

        url = getattr(self, 'current_run_url', None) or get_base_url(tab.project_name)

        current_config = RunConfig(
            project_name=tab.project_name,
            module_name=tab.module_name,
            base_url=url,
            browser=self.run_settings.browser,
            video_option=self.run_settings.video_option,
            wait_time=self.run_settings.wait_time
        )

        dialog = RunConfigDialog(current_config, self)
        self._position_dialog_near_run_button(dialog)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_config = dialog.get_config()
            self.run_settings.browser = updated_config.browser
            self.run_settings.video_option = updated_config.video_option
            self.run_settings.wait_time = updated_config.wait_time
            self.run_settings.save_to_settings()
            self.current_url = updated_config.base_url
            self.show_status_message("Run configuration saved", "info", 3000)

    def _position_dialog_near_run_button(self, dialog):
        if hasattr(self.menu_bar, 'config_button'):
            config_button = self.menu_bar.config_button
            button_global_pos = config_button.mapToGlobal(config_button.rect().bottomLeft())

            dialog_size = dialog.sizeHint()
            screen_geometry = self.screen().availableGeometry()

            x = screen_geometry.right() - dialog_size.width()
            y = button_global_pos.y() + 5

            if y + dialog_size.height() > screen_geometry.bottom():
                y = screen_geometry.bottom() - dialog_size.height()

            dialog.move(x, y)
        else:
            screen_geometry = self.screen().availableGeometry()
            dialog_size = dialog.sizeHint()
            x = screen_geometry.right() - dialog_size.width()
            y = screen_geometry.top() + 50
            dialog.move(x, y)

    def _on_log_found(self, filepath):
        parts = os.path.basename(filepath).split('_')
        title = f"{parts[0]}/{parts[2]}"
        self.output_dock.open_log_tab(filepath, title)

    def _on_open_log_manually(self):
        if not self.e2e_dir:
            return
        logs_dir = os.path.join(self.e2e_dir, "logs")
        filepath, _ = QFileDialog.getOpenFileName(self, "Open Log File", logs_dir, "HTML Files (*.html)")
        if filepath:
            parts = os.path.basename(filepath).split('_')
            title = f"{parts[0]}/{parts[2]}"
            self.output_dock.open_log_tab(filepath, title)

    def on_test_started(self):
        if hasattr(self.menu_bar, 'run_button'):
            self.menu_bar.run_button.setEnabled(False)
            self.menu_bar.stop_button.setEnabled(True)

        self.output_dock.start_log_polling()

    def on_test_finished(self, exit_code):
        self.output_dock.stop_log_polling()

        console = self.output_dock.get_console()
        if exit_code == 0:
            console.append("\n[INFO] Run finished successfully.")
            self.show_status_message("Run Completed", "success", 5000)
        if exit_code != 0:
            console.append(f"\n[ERROR] Run finished with exit code: {exit_code}.")
            self.show_status_message("Run Failed", "error", 5000)

        if hasattr(self.menu_bar, 'run_button'):
            self.menu_bar.run_button.setEnabled(True)
            self.menu_bar.stop_button.setEnabled(False)

    # Shortcuts for dock

    def navigate_dock_tabs_next(self):
         current_index = self.output_dock.tabs.currentIndex()
         count = self.output_dock.tabs.count()
         new_index = (current_index + 1) % count
         self.output_dock.tabs.setCurrentIndex(new_index)

    def navigate_dock_tabs_prev(self):
        current_index = self.output_dock.tabs.currentIndex()
        count = self.output_dock.tabs.count()
        new_index = (current_index - 1 + count) % count
        self.output_dock.tabs.setCurrentIndex(new_index)

    def navigate_log_back(self):
        current_widget = self.output_dock.tabs.currentWidget()
        if isinstance(current_widget, QWebEngineView):
            current_widget.back()

    def navigate_log_forward(self):
        current_widget = self.output_dock.tabs.currentWidget()
        if isinstance(current_widget, QWebEngineView):
            current_widget.forward()






