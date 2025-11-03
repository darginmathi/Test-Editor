#    Muted Red (`#E57373`): hsl(0, 58%, 68%)
#    Muted Blue (`#64B5F6`): hsl(207, 86%, 68%)
THEMES = {
    "dark": {
        "BG_DARK": "hsl(0, 0%, 0%)",
        "BG": "hsl(0, 0%, 10%)",
        "BG_LIGHT": "hsl(0, 0%, 20%)",
        "TEXT": "hsl(0, 0%, 95%)",
        "TEXT_MUTED": "hsl(0, 0%, 80%)",
        "HIGHLIGHT": "hsl(0, 0%, 40%)",
        "BORDER": "hsl(0, 0%, 30%)",
        "BORDER_MUTED": "hsl(0, 0%, 20%)",
        "SHADOW": "hsl(0, 0%, 30%)",
        "PRIMARY": "hsl(210, 0%, 10%)",
        "SECONDARY": "hsl(270, 30%, 60%)",
        "DANGER": "hsl(0, 70%, 60%)",
        "WARNING": "hsl(35, 90%, 60%)",
        "SUCCESS": "hsl(140, 50%, 55%)",
        "INFO": "hsl(190, 70%, 60%)"
    },
    "mdark": {
        "BG_DARK": "hsl(0, 0%, 10%)",
        "BG": "hsl(0, 0%, 0%)",
        "BG_LIGHT": "hsl(0, 0%, 20%)",
        "TEXT": "hsl(0, 0%, 95%)",
        "TEXT_MUTED": "hsl(0, 0%, 80%)",
        "HIGHLIGHT": "hsl(0, 0%, 0%)",
        "BORDER": "hsl(0, 0%, 30%)",
        "BORDER_MUTED": "hsl(0, 0%, 20%)",
        "SHADOW": "hsl(0, 0%, 30%)",
        "PRIMARY": "hsl(210, 0%, 10%)",
        "SECONDARY": "hsl(270, 30%, 60%)",
        "DANGER": "hsl(0, 70%, 60%)",
        "WARNING": "hsl(35, 90%, 60%)",
        "SUCCESS": "hsl(140, 50%, 55%)",
        "INFO": "hsl(190, 70%, 60%)"
    },
    "light": {
        "BG_DARK": "hsl(0, 0%, 100%)",
        "BG": "hsl(0, 0%, 95%)",
        "BG_LIGHT": "hsl(0, 0%, 90%)",
        "TEXT": "hsl(0, 0%, 0%)",
        "TEXT_MUTED": "hsl(0, 0%, 10%)",
        "HIGHLIGHT": "hsl(0, 0%, 60%)",
        "BORDER": "hsl(0, 0%, 70%)",
        "BORDER_MUTED": "hsl(0, 0%, 80%)",
        "SHADOW": "hsl(0, 0%, 30%)",
        "PRIMARY": "hsla(210, 80%, 60%, 0.15)",
        "SECONDARY": "hsla(270, 30%, 60%, 0.15)",
        "DANGER": "hsla(0, 70%, 60%, 0.15)",
        "WARNING": "hsla(35, 90%, 60%, 0.15)",
        "SUCCESS": "hsla(140, 50%, 55%, 0.15)",
        "INFO": "hsla(190, 70%, 60%, 0.15)"
    }
}

def get_stylesheet(theme_name="dark"):
    colors = THEMES.get(theme_name, THEMES["dark"])

    return f"""
    /* Global */
    QWidget {{
        font-family: "Segoe UI";
        font-size: 16px;
        color: {colors["TEXT_MUTED"]};
        background-color: {colors["BG_DARK"]};
    }}

    /* Menu Bar */
    QMenuBar {{
    }}

    QMenuBar::item {{
        border: none;
        background-color: transparent;
        padding: 2px 8px;
        color: {colors["TEXT"]};
    }}

    QMenuBar::item:selected {{
        background-color: {colors["BG_LIGHT"]};
        border-radius: 4px;
    }}

    QMenuBar::item:pressed {{
        background-color: {colors["BG"]};
    }}

    /* Dropdown Menus */
    QMenu {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        padding: 1px;
        border-radius: 4px;
    }}

    QMenu::item {{
        padding: 4px 24px 4px 28px;
        background-color: transparent;
        border-radius: 4px;
        color: {colors["TEXT"]};
    }}

    QMenu::item:selected {{
        background-color: {colors["BG"]};
    }}

    QMenu::item:pressed {{
        background-color: {colors["BG_DARK"]};
    }}

    QMenu::separator {{
        height: 1px;
        background-color: {colors["BORDER"]};
        margin: 2px 8px;
    }}

    /* Table View */
    QTableView {{
        background-color: {colors["BG_DARK"]};
        gridline-color: {colors["BORDER"]};
        outline: 0;
    }}

    QTableView::item:selected {{
        background-color: {colors["PRIMARY"]};
        color: {colors["TEXT"]};
    }}

    QTableView QLineEdit {{
        color: {colors["TEXT"]};
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        border-radius: 4px;
        selection-color: {colors["TEXT"]};
        selection-background-color: {colors["PRIMARY"]};
    }}

    QTableView::item:focus {{
        outline: 1px dotted {colors["TEXT"]};
        outline-offset: -1px;
    }}

    /* Headers */
    QHeaderView::section {{
        background-color: {colors["BG_DARK"]};
        padding: -2px -1px;
        border: none;
        color: {colors["TEXT"]};
    }}

    QHeaderView::section:horizontal {{
    border-bottom: 1px solid {colors["BORDER"]};
    /* border-right: 1px solid {colors["BORDER"]}; */
    }}

    QHeaderView::section:vertical {{
    border-right: 1px solid {colors["BORDER"]};
    /* border-bottom: 1px solid {colors["BORDER"]}; */
    }}

    /* Buttons */
    QPushButton {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        border-radius: 4px;
        padding: 3px 8px;
        min-width: 60px;
        min-height: 20px;
        color: {colors["TEXT"]};
    }}

    QPushButton:hover {{
        background-color: {colors["BG_LIGHT"]};
    }}

    QPushButton:pressed {{
        background-color: {colors["BG"]};
    }}

    QPushButton:disabled {{
        background-color: {colors["BG_DARK"]};
        color: {colors["TEXT_MUTED"]};
    }}

    QPushButton#StatusBarButton {{
        border: none;
        padding: 2px 4px;
        background-color: transparent;
        min-width: 15px;
        min-height: 15px;
    }}

    QPushButton#StatusBarButton:hover {{
        background-color: {colors["BG_LIGHT"]};
    }}

    QPushButton#StatusBarButton:pressed {{
        background-color: {colors["BG"]};
    }}

    QStatusBar::item {{
    border: none;
    }}

    /* Input Fields */
    QLineEdit, QTextEdit {{
        border: 1px solid {colors["BORDER"]};
        border-radius: 4px;
        padding: 2px 4px;
        selection-background-color: {colors["BG_LIGHT"]};
        background-color: {colors["BG_DARK"]};
        color: {colors["TEXT"]};
    }}

    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {colors["BG_LIGHT"]};
    }}

        /* Scrollbars */
    QScrollBar:vertical {{
        border: none;
        border-left: 1px solid {colors["BORDER"]};
        background-color: {colors["BG_DARK"]};
        width: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {colors["HIGHLIGHT"]};
        min-height: 20px;
        margin: 2px 4px;
        border-radius: 6px;
    }}

    QScrollBar::handle:vertical:hover {{
        background-color: {colors["BORDER"]};
    }}

    QScrollBar::handle:vertical:pressed {{
        background-color: {colors["BORDER_MUTED"]};
    }}

    QScrollBar::sub-page:vertical, QScrollBar::add-page:vertical {{
        background-color: {colors["BG_DARK"]};
    }}

    /* Hide the top/bottom arrow buttons */
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
        border: none;
        background: none;
    }}

    QScrollBar:horizontal {{
        border: none;
        border-top: 1px solid {colors["BORDER"]};
        background-color: {colors["BG_DARK"]};
        height: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {colors["HIGHLIGHT"]};
        min-width: 20px;
        margin: 4px 2px;
        border-radius: 6px;
    }}

    QScrollBar::handle:horizontal:hover {{
        background-color: {colors["BORDER"]};
    }}

    QScrollBar::handle:horizontal:pressed {{
        background-color: {colors["BORDER_MUTED"]};
    }}

    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
        background-color: {colors["BG_DARK"]};
    }}

    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
        border: none;
        background: none;
    }}

    /* Corner between scrollbars */
    QScrollBar::corner {{
        background-color: {colors["BG_DARK"]};
        border-top: 1px solid {colors["BORDER"]};
        border-left: 1px solid {colors["BORDER"]};
    }}

    /* Tab Widget */
    QTabWidget {{
        background-color: {colors["BG_DARK"]};
    }}

    QTabWidget::pane {{
        border: 1px solid {colors["BORDER"]};
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {colors["BG"]};
        border: 1px solid {colors["BORDER"]};
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        padding: 3px 8px;
        margin-right: 0px;
        min-height: 20px;
        color: {colors["TEXT_MUTED"]};
    }}

    QTabBar::tab:selected {{
        background-color: {colors["BG_DARK"]};
        border-bottom: 1px solid {colors["BG_DARK"]};
        color: {colors["TEXT"]};
    }}

    QTabBar::tab:hover:!selected {{
        background-color: {colors["BG_LIGHT"]};
    }}

    /* Tooltips */
    QToolTip {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        padding: 2px 4px;
        color: {colors["TEXT"]};
    }}

    /* Status Message Label */
    QLabel#StatusLabel {{
        padding-left: 10px;
        padding-right: 10px;
        border-radius: 4px;
    }}

    QLabel#StatusLabel[message_type="info"] {{
        color: {colors["INFO"]};
    }}

    QLabel#StatusLabel[message_type="success"] {{
        color: {colors["SUCCESS"]};
    }}

    QLabel#StatusLabel[message_type="warning"] {{
        color: {colors["WARNING"]};
    }}

    QLabel#StatusLabel[message_type="error"] {{
        color: {colors["DANGER"]};
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
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        border-radius: 4px;
        outline: none;
    }}

    QListWidget::item:selected {{
        background-color: {colors["BG_LIGHT"]};
        color: {colors["TEXT"]};
    }}

    QListWidget::item:hover:!selected {{
         background-color: {colors["BORDER"]};
         color: {colors["TEXT"]};
    }}

    /* Make the splitter handle invisible */
    QSplitter::handle {{
        border: none;
        width: 0px;
        image: none;
        background-color: transparent;
    }}
"""
