from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTableView, QHeaderView, QMenu, QInputDialog
from PyQt6.QtCore import Qt, pyqtSignal, QItemSelection, QModelIndex
from PyQt6.QtGui import QAction, QKeySequence, QShortcut

class Table(QWidget):
    insertRowsRequested = pyqtSignal(int, int)
    deleteRowsRequested = pyqtSignal(list)
    copyRequested = pyqtSignal(QItemSelection)
    pasteRequested = pyqtSignal(QModelIndex)
    cutRequested = pyqtSignal(QItemSelection)
    clearRequested = pyqtSignal(QItemSelection)

    def __init__(self, model, undo_stack=None, delegate=None):
        super().__init__()
        self.model = model
        self.undo_stack = undo_stack
        self.delegate = delegate

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        self.table = QTableView()
        layout.addWidget(self.table)
        self.setLayout(layout)

        self.setup_table()
        self.setup_context_menu()
        self.setup_shortcuts()

    def setup_table(self):
        self.table.setModel(self.model)
        if self.delegate:
            self.table.setItemDelegate(self.delegate)
        h_header = self.table.horizontalHeader()
        v_header = self.table.verticalHeader()
        h_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        v_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.auto_adjust_cells()
        self.table.setWordWrap(True)
        self.table.setAlternatingRowColors(False)
        self.table.setSortingEnabled(False)

    def auto_adjust_cells(self):
        self.table.resizeColumnsToContents()

    def auto_adjust_rows(self):
        self.table.resizeRowsToContents()

    def setup_context_menu(self):
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

    def show_context_menu(self, position):
        menu = QMenu(self)

        insert_row_action = QAction("Insert Rows", self)
        insert_row_action.triggered.connect(self.insert_rows)
        insert_row_action.setShortcut("Ctrl+I")
        menu.addAction(insert_row_action)

        delete_row_action = QAction("Delete Rows", self)
        delete_row_action.triggered.connect(self.delete_rows)
        delete_row_action.setShortcut("Ctrl+D")
        menu.addAction(delete_row_action)

        menu.addSeparator()

        copy_action = QAction("Copy", self)
        copy_action.triggered.connect(self.copy_cells)
        copy_action.setShortcut("Ctrl+C")
        menu.addAction(copy_action)

        paste_action = QAction("Paste", self)
        paste_action.triggered.connect(self.paste_cells)
        paste_action.setShortcut("Ctrl+V")
        menu.addAction(paste_action)

        cut_action = QAction("Cut", self)
        cut_action.triggered.connect(self.cut_cells)
        cut_action.setShortcut("Ctrl+X")
        menu.addAction(cut_action)

        menu.addSeparator()

        delete_action = QAction("Clear Contents", self)
        delete_action.triggered.connect(self.clear_cells)
        delete_action.setShortcut("Delete")
        menu.addAction(delete_action)

        menu.exec(self.table.mapToGlobal(position))

    def setup_shortcuts(self):
        self.insert_shortcut = QShortcut(QKeySequence("Ctrl+I"), self)
        self.insert_shortcut.activated.connect(self.insert_rows)

        self.delete_shortcut = QShortcut(QKeySequence("Ctrl+D"), self)
        self.delete_shortcut.activated.connect(self.delete_rows)

        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.activated.connect(self.copy_cells)

        self.paste_shortcut = QShortcut(QKeySequence("Ctrl+V"), self)
        self.paste_shortcut.activated.connect(self.paste_cells)

        self.cut_shortcut = QShortcut(QKeySequence("Ctrl+X"), self)
        self.cut_shortcut.activated.connect(self.cut_cells)

        self.clear_shortcut = QShortcut(QKeySequence("Delete"), self)
        self.clear_shortcut.activated.connect(self.clear_cells)


    def copy_cells(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            self.copyRequested.emit(selection)

    def paste_cells(self):
        self.pasteRequested.emit(self.table.currentIndex())

    def cut_cells(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            self.cutRequested.emit(selection)

    def clear_cells(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            self.clearRequested.emit(selection)

    def insert_rows(self):
        selection = self.table.selectionModel().selection()

        rows = sorted(set(index.row() for index in selection.indexes()))
        num_rows = len(rows)
        if num_rows < 1:
            return
        position = rows[0]

        if num_rows > 1:
            self.insertRowsRequested.emit(position, num_rows)
        if num_rows == 1:
            num, ok = QInputDialog.getInt(
                self,
                "Insert Rows",
                "No. of Rows:",
                1,
                1,
                100
            )
            if ok and num > 0:
                rows_to_insert = num
                self.insertRowsRequested.emit(position, rows_to_insert)
            else:
                return

    def delete_rows(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            rows = sorted(list(set(index.row() for index in selection.indexes() if index.isValid())))
        elif self.table.currentIndex().isValid():
            rows = [self.table.currentIndex().row()]
        else:
            return
        if rows:
            self.deleteRowsRequested.emit(rows)
