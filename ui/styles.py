class DarkTheme:
    BG_DARK = "hsl(0, 0%, 0%)"
    BG = "hsl(0, 0%, 10%)"
    BG_LIGHT = "hsl(0, 0%, 20%)"
    TEXT = "hsl(0, 0%, 95%)"
    TEXT_MUTED = "hsl(0, 0%, 80%)"
    HIGHLIGHT = "hsl(0, 0%, 40%)"
    BORDER = "hsl(0, 0%, 30%)"
    BORDER_MUTED = "hsl(0, 0%, 20%)"
    SHADOW = "hsl(0, 0%, 30%)"
    PRIMARY = "hsl(210, 0%, 10%)"
    SECONDARY = "hsl(270, 30%, 60%)"
    DANGER = "hsl(0, 70%, 60%)"
    WARNING = "hsl(35, 90%, 60%)"
    SUCCESS = "hsl(140, 50%, 55%)"
    INFO = "hsl(190, 70%, 60%)"

class LightTheme:
    BG_DARK = "hsl(0, 0%, 100%)"
    BG = "hsl(0, 0%, 95%)"
    BG_LIGHT = "hsl(0, 0%, 90%)"
    TEXT = "hsl(0, 0%, 0%)"
    TEXT_MUTED = "hsl(0, 0%, 10%)"
    HIGHLIGHT = "hsl(0, 0%, 60%)"
    BORDER = "hsl(0, 0%, 70%)"
    BORDER_MUTED = "hsl(0, 0%, 80%)"
    SHADOW = "hsl(0, 0%, 30%)"
    PRIMARY = "hsl(210, 80%, 60%)"
    SECONDARY = "hsl(270, 30%, 60%)"
    DANGER = "hsl(0, 70%, 60%)"
    WARNING = "hsl(35, 90%, 60%)"
    SUCCESS = "hsl(140, 50%, 55%)"
    INFO = "hsl(190, 70%, 60%)"


def get_stylesheet(theme) -> str:
    return f"""
    /* Global */
    QWidget {{
        font-family: "Segoe UI";
        font-size: 16px;
        color: {theme.TEXT_MUTED};
        background-color: {theme.BG_DARK};
    }}

    /* Menu Bar */
    QMenuBar {{
    }}

    QMenuBar::item {{
        border: none;
        background-color: transparent;
        padding: 2px 8px;
        color: {theme.TEXT};
    }}

    QMenuBar::item:selected {{
        background-color: {theme.BG_LIGHT};
        border-radius: 4px;
    }}

    QMenuBar::item:pressed {{
        background-color: {theme.BG};
    }}

    /* Dropdown Menus */
    QMenu {{
        background-color: {theme.BG_DARK};
        border: 1px solid {theme.BORDER};
        padding: 1px;
        border-radius: 4px;
    }}

    QMenu::item {{
        padding: 4px 24px 4px 28px;
        background-color: transparent;
        border-radius: 4px;
        color: {theme.TEXT};
    }}

    QMenu::item:selected {{
        background-color: {theme.BG};
    }}

    QMenu::item:pressed {{
        background-color: {theme.BG_DARK};
    }}

    QMenu::separator {{
        height: 1px;
        background-color: {theme.BORDER};
        margin: 2px 8px;
    }}

    /* Table View */
    QTableView {{
        background-color: {theme.BG_DARK};
        gridline-color: {theme.BORDER};
        outline: 0;
    }}

    QTableView::item:selected {{
        background-color: {theme.PRIMARY};
        color: {theme.TEXT};
    }}

    QTableView QLineEdit {{
        color: {theme.TEXT};
        background-color: {theme.BG_DARK};
        border: 1px solid {theme.BORDER};
        border-radius: 4px;
        selection-color: {theme.TEXT};
        selection-background-color: {theme.BORDER_MUTED};
    }}

    QTableView::item:focus {{
        outline: 1px dotted {theme.TEXT};
        outline-offset: -1px;
    }}

    /* Headers */
    QHeaderView::section {{
        background-color: {theme.BG_DARK};
        padding: -2px -1px;
        border: none;
        color: {theme.TEXT};
    }}

    QHeaderView::section:horizontal {{
    border-bottom: 1px solid {theme.BORDER};
    /* border-right: 1px solid {theme.BORDER}; */
    }}

    QHeaderView::section:vertical {{
    border-right: 1px solid {theme.BORDER};
    /* border-bottom: 1px solid {theme.BORDER}; */
    }}

    /* Buttons */
    QPushButton {{
        background-color: {theme.BG_DARK};
        border: 1px solid {theme.BORDER};
        border-radius: 4px;
        padding: 3px 8px;
        min-width: 60px;
        min-height: 20px;
        color: {theme.TEXT};
    }}

    QPushButton:hover {{
        background-color: {theme.BG_LIGHT};
    }}

    QPushButton:pressed {{
        background-color: {theme.BG};
    }}

    QPushButton:disabled {{
        background-color: {theme.BG_DARK};
        color: {theme.TEXT_MUTED};
    }}

   /* Toolbar Run Control Buttons - Icon Colors Only */
       /* Container for dock button groups */
    QWidget#RunToolbar {{
        background-color: transparent;
        border-top: none;
        border-right: 2px solid {theme.BG};
        border-left: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-top-left-radius: 0px;
        border-bottom-left-radius: 6px;
        border-top-right-radius: 0px;
        border-bottom-right-radius: 6px;
        padding: 0px;
    }}

    QPushButton#RunButton {{
        border: none;
        background-color: transparent;
        padding: 2px;
        border-radius: 6px;
        min-width: 24px;
        min-height: 24px;
        color: {theme.SUCCESS};
    }}

    QPushButton#RunButton:hover {{
        background-color: {theme.BG_LIGHT};
        border-left: 1px solid {theme.BORDER};
        border-top: 1px solid {theme.BORDER};
        border-bottom: 1px solid {theme.BG};
        border-right: 1px solid {theme.BG};
    }}

    QPushButton#RunButton:pressed {{
        background-color: {theme.BG};
        border-left: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-right: 1px solid {theme.BORDER};
    }}

    QPushButton#RunButton:disabled {{
        background-color: transparent;
        color: {theme.TEXT_MUTED};
        opacity: 0.3;
    }}

    QPushButton#StopButton {{
        border: none;
        background-color: transparent;
        padding: 2px;
        border-radius: 6px;
        min-width: 24px;
        min-height: 24px;
        color: {theme.DANGER};
    }}

    QPushButton#StopButton:hover {{
        background-color: {theme.BG_LIGHT};
        border-left: 1px solid {theme.BORDER};
        border-top: 1px solid {theme.BORDER};
        border-bottom: 1px solid {theme.BG};
        border-right: 1px solid {theme.BG};
    }}

    QPushButton#StopButton:pressed {{
        background-color: {theme.BG};
        border-left: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-right: 1px solid {theme.BORDER};
    }}

    QPushButton#StopButton:disabled {{
        background-color: transparent;
        color: {theme.TEXT_MUTED};
        opacity: 0.3;
    }}

    QPushButton#ConfigButton {{
        border: none;
        background-color: transparent;
        padding: 2px;
        border-radius: 6px;
        min-width: 24px;
        min-height: 24px;
        color: {theme.INFO};
    }}

    QPushButton#ConfigButton:hover {{
        background-color: {theme.BG_LIGHT};
        border-left: 1px solid {theme.BORDER};
        border-top: 1px solid {theme.BORDER};
        border-bottom: 1px solid {theme.BG};
        border-right: 1px solid {theme.BG};
    }}

    QPushButton#ConfigButton:pressed {{
        background-color: {theme.BG};
        border-left: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-right: 1px solid {theme.BORDER};
    }}

    QPushButton#ConfigButton:disabled {{
        background-color: transparent;
        color: {theme.TEXT_MUTED};
        opacity: 0.3;
    }}

    /* Status Bar Buttons Buttons zoom control */

    QPushButton#StatusBarButton {{
        border: none;
        padding: 2px 4px;
        background-color: transparent;
        border-radius: 6px;
        min-width: 15px;
        min-height: 15px;
    }}

    QPushButton#StatusBarButton:hover {{
        background-color: {theme.BG_LIGHT};
        border-left: 1px solid {theme.BORDER};
        border-top: 1px solid {theme.BORDER};
        border-bottom: 1px solid {theme.BG};
        border-right: 1px solid {theme.BG};
    }}

    QPushButton#StatusBarButton:pressed {{
        background-color: {theme.BG};
        border-left: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-right: 1px solid {theme.BORDER};
    }}

    QStatusBar::item {{
    border: none;
    }}

    /* Input Fields */
    QLineEdit, QTextEdit {{
        border: 1px solid {theme.BORDER};
        border-radius: 4px;
        padding: 2px 4px;
        selection-background-color: {theme.BG_LIGHT};
        background-color: {theme.BG_DARK};
        color: {theme.TEXT};
    }}

    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {theme.BG_LIGHT};
    }}

        /* Scrollbars */
    QScrollBar:vertical {{
        border: none;
        border-left: 1px solid {theme.BORDER};
        background-color: {theme.BG_DARK};
        width: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {theme.HIGHLIGHT};
        min-height: 20px;
        margin: 2px 4px;
        border-radius: 6px;
    }}

    QScrollBar::handle:vertical:hover {{
        background-color: {theme.BORDER};
    }}

    QScrollBar::handle:vertical:pressed {{
        background-color: {theme.BORDER_MUTED};
    }}

    QScrollBar::sub-page:vertical, QScrollBar::add-page:vertical {{
        background-color: {theme.BG_DARK};
    }}

    /* Hide the top/bottom arrow buttons */
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
        border: none;
        background: none;
    }}

    QScrollBar:horizontal {{
        border: none;
        border-top: 1px solid {theme.BORDER};
        background-color: {theme.BG_DARK};
        height: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {theme.HIGHLIGHT};
        min-width: 20px;
        margin: 4px 2px;
        border-radius: 6px;
    }}

    QScrollBar::handle:horizontal:hover {{
        background-color: {theme.BORDER};
    }}

    QScrollBar::handle:horizontal:pressed {{
        background-color: {theme.BORDER_MUTED};
    }}

    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
        background-color: {theme.BG_DARK};
    }}

    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
        border: none;
        background: none;
    }}

    /* Corner between scrollbars */
    QScrollBar::corner {{
        background-color: {theme.BG_DARK};
        border-top: 1px solid {theme.BORDER};
        border-left: 1px solid {theme.BORDER};
    }}

    /* Tab Widget */
    QTabWidget {{
        background-color: {theme.BG_DARK};
    }}

    QTabWidget::pane {{
        border: 1px solid {theme.BORDER};
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {theme.BG};
        border: 1px solid {theme.BORDER};
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        padding: 3px 8px;
        margin-right: 0px;
        min-height: 20px;
        color: {theme.TEXT_MUTED};
    }}

    QTabBar::tab:selected {{
        background-color: {theme.BG_DARK};
        border-bottom: 1px solid {theme.BG_DARK};
        color: {theme.TEXT};
    }}

    QTabBar::tab:hover:!selected {{
        background-color: {theme.BG_LIGHT};
    }}

    /* Tooltips */
    QToolTip {{
        background-color: {theme.BG_DARK};
        border: 1px solid {theme.BORDER};
        padding: 2px 4px;
        color: {theme.TEXT};
    }}

    /* Status Message Label */
    QLabel#StatusLabel {{
        padding-left: 10px;
        padding-right: 10px;
        border-radius: 4px;
    }}

    QLabel#StatusLabel[message_type="info"] {{
        color: {theme.INFO};
    }}

    QLabel#StatusLabel[message_type="success"] {{
        color: {theme.SUCCESS};
    }}

    QLabel#StatusLabel[message_type="warning"] {{
        color: {theme.WARNING};
    }}

    QLabel#StatusLabel[message_type="error"] {{
        color: {theme.DANGER};
    }}

    /* Welcome Screen Styles */
    QLabel[class="welcome-logo"] {{
        font-size: 64px;
        font-weight: bold;
        padding: 40px;
    }}

    QLabel[class="welcome-instructions"] {{
        font-size: 22px;
        margin-top: 30px;
        line-height: 1.5;
    }}

    QLabel[class="welcome-tips"] {{
        font-size: 22px;
        margin-top: 30px;
        line-height: 1.5;
        padding-left: 80px;
    }}

    /* List Widgets */
    QListWidget {{
        background-color: {theme.BG_DARK};
        border: 1px solid {theme.BORDER};
        border-radius: 4px;
        outline: none;
    }}

    QListWidget::item:selected {{
        background-color: {theme.BG_LIGHT};
        color: {theme.TEXT};
    }}

    QListWidget::item:hover:!selected {{
         background-color: {theme.BORDER};
         color: {theme.TEXT};
    }}

    /* Make the splitter handle invisible */
    QSplitter::handle {{
        border: none;
        width: 0px;
        image: none;
        background-color: transparent;
    }}

    /* Container for dock button groups */
    QWidget#DockButtonContainerR {{
        background-color: transparent;
        border: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-right: none;
        border-left: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-top-left-radius: 6px;
        border-bottom-left-radius: 6px;
        border-top-right-radius: 0px;
        border-bottom-right-radius: 0px;
        padding: 0px;
    }}

    /* Container for dock button groups */
    QWidget#DockButtonContainerL {{
        background-color: transparent;
        border: none;
        border: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-left: none;
        border-right: 1px solid {theme.BORDER};
        border-top-left-radius: 0px;
        border-bottom-left-radius: 0px;
        border-top-right-radius: 6px;
        border-bottom-right-radius: 6px;
        padding: 0px;
    }}

    /* Output toolbar buttons (specific to output dock) */
    QPushButton#DockTitleBarButton {{
        border: none;
        background-color: transparent;
        padding: 2px;
        border-radius: 6px;
        min-width: 24px;
        min-height: 24px;
    }}

    QPushButton#DockTitleBarButton:hover {{
        background-color: {theme.BG_LIGHT};
        border-left: 1px solid {theme.BORDER};
        border-top: 1px solid {theme.BORDER};
        border-bottom: 1px solid {theme.BG};
        border-right: 1px solid {theme.BG};
    }}

    QPushButton#DockTitleBarButton:pressed {{
        background-color: {theme.BG};
        border-left: 2px solid {theme.BG};
        border-top: 2px solid {theme.BG};
        border-bottom: 1px solid {theme.BORDER};
        border-right: 1px solid {theme.BORDER};
    }}
"""
