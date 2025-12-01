from PyQt6.QtWidgets import (QDockWidget, QTabWidget,  QTextEdit, QTabBar, QVBoxLayout, QWidget,
                             QHBoxLayout, QPushButton)
from PyQt6.QtCore import Qt, pyqtSignal, QUrl
from PyQt6.QtGui import QShortcut, QKeySequence
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage
from PyQt6.QtWidgets import QStyle


class OutputDock(QDockWidget):
    clear_requested = pyqtSignal()
    toggle_visibility_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__("OUTPUT", parent)
        self.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea)

        title_bar = QWidget()
        title_bar_layout = QHBoxLayout(title_bar)
        title_bar_layout.setContentsMargins(0, 5, 0, 0)
        title_bar_layout.setSpacing(5)

        self.is_maximized = False
        self.restore_height = 0

        style = self.style()
        # back_icon = style.standardIcon(QStyle.StandardPixmap.SP_ArrowBack)
        # forward_icon = style.standardIcon(QStyle.StandardPixmap.SP_ArrowForward)
        open_log_icon = style.standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        self.maximize_icon = style.standardIcon(QStyle.StandardPixmap.SP_TitleBarMaxButton)
        self.restore_icon = style.standardIcon(QStyle.StandardPixmap.SP_TitleBarNormalButton)
        self.minimize_icon = style.standardIcon(QStyle.StandardPixmap.SP_TitleBarMinButton)

        self.back_button = QPushButton("◀")
        self.forward_button = QPushButton("▶")
        self.refresh_button = QPushButton("⟳")
        self.maximize_button = QPushButton(icon=self.maximize_icon)
        self.open_log_button = QPushButton(icon=open_log_icon)
        self.minimize_button = QPushButton(icon=self.minimize_icon)

        self.back_button.setObjectName("DockTitleBarButton")
        self.forward_button.setObjectName("DockTitleBarButton")
        self.refresh_button.setObjectName("DockTitleBarButton")
        self.maximize_button.setObjectName("DockTitleBarButton")
        self.open_log_button.setObjectName("DockTitleBarButton")
        self.minimize_button.setObjectName("DockTitleBarButton")

        self.back_button.hide()
        self.forward_button.hide()
        self.refresh_button.hide()

        nav_container = QWidget()
        nav_container.setObjectName("DockButtonContainerL")
        nav_layout = QHBoxLayout(nav_container)
        nav_layout.setContentsMargins(4, 4, 4, 4)
        nav_layout.setSpacing(2)
        nav_layout.addWidget(self.back_button)
        nav_layout.addWidget(self.forward_button)
        nav_layout.addWidget(self.refresh_button)

        action_container = QWidget()
        action_container.setObjectName("DockButtonContainerR")
        action_layout = QHBoxLayout(action_container)
        action_layout.setContentsMargins(4, 4, 4, 4)
        action_layout.setSpacing(2)
        action_layout.addWidget(self.open_log_button)
        action_layout.addWidget(self.minimize_button)
        action_layout.addWidget(self.maximize_button)

        title_bar_layout.addWidget(nav_container)
        title_bar_layout.addStretch()
        title_bar_layout.addWidget(action_container)

        self.back_button.setFlat(True)
        self.forward_button.setFlat(True)
        self.refresh_button.setFlat(True)
        self.open_log_button.setFlat(True)
        self.maximize_button.setFlat(True)
        self.minimize_button.setFlat(True)

        self.back_button.setToolTip("Back (Alt+Left)")
        self.forward_button.setToolTip("Forward (Alt+Right)")
        self.refresh_button.setToolTip("Refresh (F5)")
        self.open_log_button.setToolTip("Open Log File")
        self.maximize_button.setToolTip("Maximize (Ctrl+M)")
        self.minimize_button.setToolTip("Minimize (Ctrl+`)")

        self.setTitleBarWidget(title_bar)
        self.minimize_button.clicked.connect(self.toggle_visibility_requested.emit)
        self.maximize_button.clicked.connect(self.toggle_maximize)
        self.refresh_button.clicked.connect(self.refresh_current_view)

        self.refresh_shortcut = QShortcut(QKeySequence(Qt.Key.Key_F5), self)
        self.refresh_shortcut.activated.connect(self.refresh_current_view)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabBar().setMovable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        layout.addWidget(self.tabs)

        self.setWidget(container)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.tabs.addTab(self.console, "Console")
        self.tabs.tabBar().setTabButton(0, QTabBar.ButtonPosition.RightSide, None)
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self._warmup_web_engine()

    def get_console(self):
        return self.console

    def open_log_tab(self, filepath, title):
        report_view = QWebEngineView()
        report_view.setUrl(QUrl.fromLocalFile(filepath))

        index = self.tabs.addTab(report_view, title)
        self.tabs.setCurrentIndex(index)

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
            self.refresh_button.show()

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
            self.refresh_button.hide()

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
            # self.maximize_button.setText("Maximize")
            self.maximize_button.setIcon(self.maximize_icon)
        else:
            if self.restore_height == 0:
                self.restore_height = self.height()

            # maximized_height = int((main_window.centralWidget().height() + self.height()) * 0.8)
            # maximized_height = int(main_window.height())
            # # # main_window.resizeDocks([self], [maximized_height], Qt.Orientation.Vertical)
            # self.setFixedHeight(maximized_height)
            main_window.centralWidget().hide()
            self.is_maximized = True
            self.maximize_button.setIcon(self.restore_icon)
            # self.maximize_button.setText("Restore")

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

    def refresh_current_view(self):
        current_widget = self.tabs.currentWidget()
        if isinstance(current_widget, QWebEngineView):
            def get_scroll_and_refresh(scroll_y):
                url = current_widget.url().toLocalFile()
                if not url:
                    return

                if not (url.lower().endswith(".html") or url.lower().endswith(".htm")):
                    current_widget.reload()
                    return

                try:
                    with open(url, 'r', encoding='utf-8') as f:
                        new_html = f.read()

                    js_html = new_html.replace('\\', '\\\\').replace('`', '\\`').replace('$', '\\$')

                    js_code = f"""
                        document.documentElement.innerHTML = `{js_html}`;
                        window.scrollTo(0, {scroll_y});
                    """
                    current_widget.page().runJavaScript(js_code)

                except FileNotFoundError:
                    current_widget.reload() # Fallback

            current_widget.page().runJavaScript("window.scrollY;", 0, get_scroll_and_refresh)

    def is_dock_maximized(self):
        return self.is_maximized


    def _warmup_web_engine(self):
        temp_view = QWebEngineView()
        temp_view.deleteLater()



