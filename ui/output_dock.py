from PyQt6.QtWidgets import (QDockWidget, QTabWidget,  QTextEdit, QTabBar, QToolBar, QVBoxLayout, QWidget,
                             QHBoxLayout, QPushButton, QLabel)
from PyQt6.QtCore import Qt, pyqtSignal, QUrl, QFileSystemWatcher
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage, QWebEngineHistory

class OutputDock(QDockWidget):
    clear_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__("OUTPUT", parent)
        self.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea)

        title_bar = QWidget()
        title_bar_layout = QHBoxLayout(title_bar)
        title_bar_layout.setContentsMargins(5, 0, 5, 0)
        title_bar_layout.setSpacing(5)

        title_label = QLabel("Output:")
        title_bar_layout.addWidget(title_label)
        title_bar_layout.addStretch()

        self.is_maximized = False
        self.restore_height = 0

        self.back_button = QPushButton("Back")
        self.forward_button = QPushButton("Forward")
        self.open_log_button = QPushButton("Open Log")
        self.maximize_button = QPushButton("Maximize")

        self.back_button.hide()
        self.forward_button.hide()

        title_bar_layout.addWidget(self.back_button)
        title_bar_layout.addWidget(self.forward_button)
        title_bar_layout.addWidget(self.open_log_button)
        title_bar_layout.addWidget(self.maximize_button)

        self.setTitleBarWidget(title_bar)
        self.maximize_button.clicked.connect(self.toggle_maximize)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        layout.addWidget(self.tabs)

        self.setWidget(container)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.tabs.addTab(self.console, "Console")
        self.tabs.tabBar().setTabButton(0, QTabBar.ButtonPosition.RightSide, None)
        self.tabs.currentChanged.connect(self.on_tab_changed)

    def get_console(self):
        return self.console

    def open_log_tab(self, filepath, title):
        report_view = QWebEngineView()
        report_view.setUrl(QUrl.fromLocalFile(filepath))

        index = self.tabs.addTab(report_view, title)
        self.tabs.setCurrentIndex(index)

        watcher = QFileSystemWatcher([filepath], self)
        watcher.fileChanged.connect(report_view.reload)

    def close_tab(self, index):
        if index == 0:
            return
        widget_to_remove = self.tabs.widget(index)
        self.tabs.removeTab(index)
        widget_to_remove.deleteLater()

    def on_tab_changed(self, index):
        previous_widget = self.property("current_log_widget")
        if previous_widget:
            try:
                page = previous_widget.page()
                page.action(QWebEnginePage.WebAction.Back).changed.disconnect(self.back_button.setEnabled)
                page.action(QWebEnginePage.WebAction.Forward).changed.disconnect(self.forward_button.setEnabled)
                self.back_button.clicked.disconnect(previous_widget.back)
                self.forward_button.clicked.disconnect(previous_widget.forward)
            except (TypeError, RuntimeError):
                pass

        current_widget = self.tabs.widget(index)

        if isinstance(current_widget, QWebEngineView):
            self.back_button.show()
            self.forward_button.show()

            page = current_widget.page()

            self.back_button.clicked.connect(current_widget.back)
            self.forward_button.clicked.connect(current_widget.forward)

            page.action(QWebEnginePage.WebAction.Back).changed.connect(self._update_back_button_status)
            page.action(QWebEnginePage.WebAction.Forward).changed.connect(self._update_forward_button_status)

            self._update_back_button_status()
            self._update_forward_button_status()

        else:
            self.back_button.hide()
            self.forward_button.hide()

    def toggle_maximize(self):
        main_window = self.parent()
        if not main_window:
            print("skipped toggle maximize")
            return

        if self.is_maximized:
            # main_window.resizeDocks([self], [self.previous_height], Qt.Orientation.Vertical)
            # self.setMaximumHeight(16777215) # Reset max height limit (this is a Qt constant for 'no limit')
            # self.setFixedHeight(self.restore_height)
            main_window.centralWidget().show()
            self.is_maximized = False
            self.maximize_button.setText("Maximize")
        else:
            if self.restore_height == 0:
                self.restore_height = self.height()

            # maximized_height = int((main_window.centralWidget().height() + self.height()) * 0.8)
            # maximized_height = int(main_window.height())
            # # # main_window.resizeDocks([self], [maximized_height], Qt.Orientation.Vertical)
            # self.setFixedHeight(maximized_height)
            main_window.centralWidget().hide()
            self.is_maximized = True
            self.maximize_button.setText("Restore")

    def _update_back_button_status(self):
        current_widget = self.property("current_log_widget")
        if current_widget:
            page = current_widget.page()
            self.back_button.setEnabled(page.action(QWebEnginePage.WebAction.Back).isEnabled())

    def _update_forward_button_status(self):
        current_widget = self.property("current_log_widget")
        if current_widget:
            page = current_widget.page()
            self.forward_button.setEnabled(page.action(QWebEnginePage.WebAction.Forward).isEnabled())

    def is_dock_maximized(self):
        return self.is_maximized
