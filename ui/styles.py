# ui/styles.py

# Define color palettes for different themes
THEMES = {
    "dark": {
        "BG_DARK": "hsl(336 0% 1%)",
        "BG": "hsl(300 0% 4%)",
        "BG_LIGHT": "hsl(0 0% 9%)",
        "TEXT": "hsl(300 0% 95%)",
        "TEXT_MUTED": "hsl(300 0% 69%)",
        "HIGHLIGHT": "hsl(330 0% 39%)",
        "BORDER": "hsl(0 0% 28%)",
        "BORDER_MUTED": "hsl(300 0% 18%)",
        "PRIMARY": "hsl(210 77% 72%)",
        "SECONDARY": "hsl(32 61% 63%)",
        "DANGER": "hsl(9 26% 64%)",
        "WARNING": "hsl(52 19% 57%)",
        "SUCCESS": "hsl(146 17% 59%)",
        "INFO": "hsl(217 28% 65%)",
    },
    "light": {
        "BG_DARK": "hsl(0 0% 90%)",
        "BG": "hsl(300 0% 95%)",
        "BG_LIGHT": "hsl(300 50% 100%)",
        "TEXT": "hsl(300 0% 4%)",
        "TEXT_MUTED": "hsl(0 0% 28%)",
        "HIGHLIGHT": "hsl(300 50% 100%)",
        "BORDER": "hsl(0 0% 50%)",
        "BORDER_MUTED": "hsl(340 0% 62%)",
        "PRIMARY": "hsl(207 78% 27%)",
        "SECONDARY": "hsl(38 100% 17%)",
        "DANGER": "hsl(9 21% 41%)",
        "WARNING": "hsl(52 23% 34%)",
        "SUCCESS": "hsl(147 19% 36%)",
        "INFO": "hsl(217 22% 41%)",
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
        color: {colors["TEXT"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="success"] {{
        background-color: {colors["SUCCESS"]};
        color: {colors["TEXT"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="warning"] {{
        background-color: {colors["WARNING"]};
        color: {colors["TEXT"]};
        border: 1px solid {colors["BORDER"]};
    }}

    QLabel#StatusLabel[message_type="error"] {{
        background-color: {colors["DANGER"]};
        color: {colors["TEXT"]};
        border: 1px solid {colors["BORDER"]};
    }}
"""
