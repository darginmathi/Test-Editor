THEMES = {
    "dark": {
        "BG_DARK": "hsl(0, 0%, 0%)",
        "BG": "hsl(0, 0%, 5%)",
        "BG_LIGHT": "hsl(0, 0%, 10%)",
        "TEXT": "hsl(0, 0%, 95%)",
        "TEXT_MUTED": "hsl(0, 0%, 70%)",
        "HIGHLIGHT": "hsl(0, 0%, 40%)",
        "BORDER": "hsl(0, 0%, 30%)",
        "BORDER_MUTED": "hsl(0, 0%, 20%)",
        "PRIMARY": "hsl(206, 100%, 65%)",
        "SECONDARY": "hsl(40, 100%, 37%)",
        "DANGER": "hsl(7, 100%, 66%)",
        "WARNING": "hsl(53, 100%, 70%)",
        "SUCCESS": "hsl(162, 100%, 22%)",
        "INFO": "hsl(217, 100%, 70%)",
    },
    "light": {
        "BG_DARK": "hsl(0, 0%, 95%)",
        "BG": "hsl(0, 0%, 100%)",
        "BG_LIGHT": "hsl(0, 0%, 100%)",
        "TEXT": "hsl(0, 0%, 5%)",
        "TEXT_MUTED": "hsl(0, 0%, 30%)",
        "HIGHLIGHT": "hsl(0, 0%, 60%)",
        "BORDER": "hsl(0, 0%, 700%)",
        "BORDER_MUTED": "hsl(0, 0%, 80%)",
        "PRIMARY": "hsl(206, 100%, 65%)",
        "SECONDARY": "hsl(40, 100%, 37%)",
        "DANGER": "hsl(7, 100%, 66%)",
        "WARNING": "hsl(53, 100%, 70%)",
        "SUCCESS": "hsl(162, 100%, 22%)",
        "INFO": "hsl(217, 100%, 70%)",
    }
}

def get_stylesheet(theme_name="dark"):
    colors = THEMES.get(theme_name, THEMES["dark"])

    return f"""
    /* Global */
    QWidget {{
        font-family: "Segoe UI", "Calibri", "Arial";
        font-size: 16px;
        color: {colors["TEXT_MUTED"]};
        background-color: {colors["BG"]};
    }}

    /* Menu Bar */
    QMenuBar {{
        background-color: {colors["BG_DARK"]};
        border-bottom: 1px solid {colors["BORDER"]};
        padding: 2px;
    }}

    QMenuBar::item {{
        background-color: transparent;
        padding: 4px 12px;
        border: 1px solid transparent;
        color: {colors["TEXT"]};
    }}

    QMenuBar::item:selected {{
        background-color: {colors["BG_LIGHT"]};
        border: 1px solid {colors["BORDER"]};
    }}

    /* Dropdown Menus */
    QMenu {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        padding: 1px;
        border-radius: 0px;
    }}

    QMenu::item {{
        padding: 4px 24px 4px 28px;
        background-color: transparent;
        color: {colors["TEXT"]};
    }}

    QMenu::item:selected {{
        background-color: {colors["BG_LIGHT"]};
    }}

    QMenu::separator {{
        height: 1px;
        background-color: {colors["BORDER"]};
        margin: 2px 8px;
    }}

    /* Table View */
    QTableView {{
        gridline-color: {colors["BORDER"]};
        selection-background-color: {colors["BG_LIGHT"]};
        border: 1px solid {colors["BORDER"]};
        outline: 0;
    }}

    QTableView::item:selected {{
        background-color: {colors["BG_LIGHT"]};
        color: {colors["TEXT"]};
    }}

    QTableView QLineEdit {{
        color: {colors["TEXT"]};
        background-color: {colors["BG"]};
        border: 1px solid {colors["BG_LIGHT"]};
        selection-color: {colors["TEXT"]};
        selection-background-color: {colors["BG_LIGHT"]};
    }}

    /* Headers */
    QHeaderView::section {{
        background-color: {colors["BG_DARK"]};
        padding: 4px;
        border: 1px solid {colors["BORDER"]};
        border-top: none;
        border-left: none;
        font-weight: normal;
        color: {colors["TEXT"]};
    }}

    QHeaderView::section:last {{
        border-right: 1px solid {colors["BORDER"]};
    }}

    /* Buttons */
    QPushButton {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        padding: 4px 12px;
        min-width: 75px;
        min-height: 24px;
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

    /* Input Fields */
    QLineEdit, QTextEdit {{
        border: 1px solid {colors["BORDER"]};
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
        background-color: {colors["BG"]};
        width: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {colors["BG_DARK"]};
        min-height: 20px;
        margin: 2px;
    }}

    QScrollBar:horizontal {{
        background-color: {colors["BG"]};
        height: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {colors["BG_DARK"]};
        min-width: 20px;
        margin: 2px;
    }}

    /* Tab Widget */
    QTabWidget::pane {{
        border: 1px solid {colors["BORDER"]};
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {colors["BG_DARK"]};
        border: 1px solid {colors["BORDER"]};
        border-bottom: none;
        padding: 4px 12px;
        margin-right: 1px;
        color: {colors["TEXT_MUTED"]};
    }}

    QTabBar::tab:selected {{
        background-color: {colors["BG"]};
        border-bottom: 1px solid {colors["BG"]};
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
        border-radius: 3px;
    }}

    QLabel#StatusLabel[message_type="info"] {{
        background-color: {colors["INFO"]};
        color: {colors["BG"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="success"] {{
        background-color: {colors["SUCCESS"]};
        color: {colors["BG"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="warning"] {{
        background-color: {colors["WARNING"]};
        color: {colors["BG"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="error"] {{
        background-color: {colors["DANGER"]};
        color: {colors["BG"]};
        border: 1px solid {colors["BORDER"]};
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
"""
