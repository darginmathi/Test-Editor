# ui/styles.py

# Excel-like Color Scheme
COLOR_EXCEL_HEADER_BG = "#f2f2f2"
COLOR_EXCEL_HEADER_BORDER = "#d0d0d0"
COLOR_EXCEL_GRID_LINES = "#e1e1e1"
COLOR_EXCEL_BG_WHITE = "#ffffff"
COLOR_EXCEL_SELECTION_BLUE = "#cce8ff"
COLOR_EXCEL_SELECTION_BORDER = "#0078d4"
COLOR_EXCEL_MENU_BG = "#fafafa"
COLOR_EXCEL_MENU_BORDER = "#c8c8c8"
COLOR_EXCEL_MENU_HOVER = "#e5f3ff"
COLOR_EXCEL_MENU_SELECTED = "#cde6ff"
COLOR_EXCEL_TEXT_DARK = "#000000"
COLOR_EXCEL_TEXT_MEDIUM = "#323232"
COLOR_EXCEL_TEXT_LIGHT = "#666666"
COLOR_EXCEL_BUTTON_NORMAL = "#f0f0f0"
COLOR_EXCEL_BUTTON_HOVER = "#e1e1e1"
COLOR_EXCEL_BUTTON_ACTIVE = "#d0d0d0"
COLOR_EXCEL_SCROLLBAR = "#e6e6e6"
COLOR_EXCEL_SCROLLBAR_HANDLE = "#c8c8c8"
COLOR_EXCEL_ACTIVE_CELL_BORDER = "#ff8000" # Orange for active cell border, like Excel

app_stylesheet = f"""
    /* ===== GLOBAL BASE STYLES ===== */
    /* These apply to ALL widgets through inheritance */
    QWidget {{
        font-family: "Segoe UI", "Calibri", "Arial";
        font-size: 16px;
        color: {COLOR_EXCEL_TEXT_DARK};
        background-color: {COLOR_EXCEL_BG_WHITE};
    }}

    /* ===== SPECIFIC COMPONENT STYLES ===== */
    /* Only define what's different from base styles */

    /* Menu Bar */
    QMenuBar {{
        background-color: {COLOR_EXCEL_MENU_BG};
        border-bottom: 1px solid {COLOR_EXCEL_MENU_BORDER};
        padding: 2px;
    }}

    QMenuBar::item {{
        background-color: transparent;
        padding: 4px 12px;
        border: 1px solid transparent;
    }}

    QMenuBar::item:selected {{
        background-color: {COLOR_EXCEL_MENU_HOVER};
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
    }}

    /* Dropdown Menus */
    QMenu {{
        background-color: {COLOR_EXCEL_MENU_BG};
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        padding: 1px;
        border-radius: 0px;
    }}

    QMenu::item {{
        padding: 4px 24px 4px 28px;
        background-color: transparent;
    }}

    QMenu::item:selected {{
        background-color: {COLOR_EXCEL_SELECTION_BLUE};
    }}

    QMenu::separator {{
        height: 1px;
        background-color: {COLOR_EXCEL_MENU_BORDER};
        margin: 2px 8px;
    }}

    /* Table View - Excel Style */
    QTableView {{
        gridline-color: {COLOR_EXCEL_GRID_LINES};
        selection-background-color: {COLOR_EXCEL_SELECTION_BLUE};
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        outline: 0;
    }}

    QTableView::item:selected {{
        background-color: {COLOR_EXCEL_SELECTION_BLUE};
        color: {COLOR_EXCEL_TEXT_DARK}; /* Ensure text is dark */
    }}

    QTableView QLineEdit {{
        color: {COLOR_EXCEL_TEXT_DARK};
        background-color: {COLOR_EXCEL_BG_WHITE};
        border: 1px solid {COLOR_EXCEL_ACTIVE_CELL_BORDER};
        selection-color: {COLOR_EXCEL_TEXT_DARK};
        selection-background-color: {COLOR_EXCEL_SELECTION_BLUE};
    }}

    /* Headers */
    QHeaderView::section {{
        background-color: {COLOR_EXCEL_HEADER_BG};
        padding: 4px;
        border: 1px solid {COLOR_EXCEL_HEADER_BORDER};
        border-top: none;
        border-left: none;
        font-weight: normal;
    }}

    QHeaderView::section:last {{
        border-right: 1px solid {COLOR_EXCEL_HEADER_BORDER};
    }}

    /* Buttons */
    QPushButton {{
        background-color: {COLOR_EXCEL_BUTTON_NORMAL};
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        padding: 4px 12px;
        min-width: 75px;
        min-height: 24px;
    }}

    QPushButton:hover {{
        background-color: {COLOR_EXCEL_BUTTON_HOVER};
    }}

    QPushButton:pressed {{
        background-color: {COLOR_EXCEL_BUTTON_ACTIVE};
    }}

    QPushButton:disabled {{
        background-color: {COLOR_EXCEL_MENU_BG};
        color: {COLOR_EXCEL_TEXT_LIGHT};
    }}

    QPushButton#StatusBarButton {{
        border: none; /* Remove border */
        padding: 2px 4px; /* Smaller padding */
        background-color: transparent; /* Flat background */
        min-width: 15px; /* Allow smaller width */
        min-height: 15px; /* Allow smaller height */
        /* You can set a fixed size here too if preferred */
        /* fixed-width: 25px; */
        /* fixed-height: 25px; */
    }}

    QPushButton#StatusBarButton:hover {{
        background-color: {COLOR_EXCEL_BUTTON_HOVER}; /* Subtle hover */
    }}

    QPushButton#StatusBarButton:pressed {{
        background-color: {COLOR_EXCEL_BUTTON_ACTIVE}; /* Subtle press */
    }}

    /* Input Fields */
    QLineEdit, QTextEdit {{
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        padding: 2px 4px;
        selection-background-color: {COLOR_EXCEL_SELECTION_BLUE};
    }}

    QLineEdit:focus, QTextEdit:focus {{
        border: 1px solid {COLOR_EXCEL_SELECTION_BORDER};
    }}

    /* Scrollbars */
    QScrollBar:vertical {{
        background-color: {COLOR_EXCEL_SCROLLBAR};
        width: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {COLOR_EXCEL_SCROLLBAR_HANDLE};
        min-height: 20px;
        margin: 2px;
    }}

    QScrollBar:horizontal {{
        background-color: {COLOR_EXCEL_SCROLLBAR};
        height: 16px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {COLOR_EXCEL_SCROLLBAR_HANDLE};
        min-width: 20px;
        margin: 2px;
    }}

    /* Tab Widget */
    QTabWidget::pane {{
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {COLOR_EXCEL_MENU_BG};
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        border-bottom: none;
        padding: 4px 12px;
        margin-right: 1px;
    }}

    QTabBar::tab:selected {{
        background-color: {COLOR_EXCEL_BG_WHITE};
        border-bottom: 1px solid {COLOR_EXCEL_BG_WHITE};
    }}

    QTabBar::tab:hover:!selected {{
        background-color: {COLOR_EXCEL_BUTTON_HOVER};
    }}

    /* Tooltips */
    QToolTip {{
        background-color: #ffffe1;
        border: 1px solid {COLOR_EXCEL_MENU_BORDER};
        padding: 2px 4px;
    }}
"""
