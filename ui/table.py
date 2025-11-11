from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTableView, QHeaderView, QMenu, QInputDialog, QMessageBox
from PyQt6.QtCore import Qt, pyqtSignal, QItemSelection, QModelIndex
from PyQt6.QtGui import QAction, QKeySequence, QShortcut

class Table(QWidget):
    insertRowsRequested = pyqtSignal(int, int)
    deleteRowsRequested = pyqtSignal(list)
    copyRequested = pyqtSignal(QItemSelection)
    pasteRequested = pyqtSignal(QModelIndex)
    cutRequested = pyqtSignal(QItemSelection)
    clearRequested = pyqtSignal(QItemSelection)
    skipRequested = pyqtSignal(list)
    unskipRequested = pyqtSignal(list)

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

        self.model.rowsInserted.connect(self.merge_cells)
        self.model.rowsRemoved.connect(self.merge_cells)
        self.model.dataChanged.connect(self._on_data_changed_commands)
        self.model.modelReset.connect(self.merge_cells)

        self.merge_cells()

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

        remove_row_action = QAction("Delete Rows", self)
        remove_row_action.triggered.connect(self.remove_rows)
        remove_row_action.setShortcut("Ctrl+D")
        menu.addAction(remove_row_action)

        menu.addSeparator()

        mark_skip_action = QAction("Mark Skip", self)
        mark_skip_action.triggered.connect(self.skip)
        mark_skip_action.setShortcut("Ctrl+R")
        menu.addAction(mark_skip_action)

        mark_unskip_action = QAction("Mark Unskip", self)
        mark_unskip_action.triggered.connect(self.unskip)
        mark_unskip_action.setShortcut("Ctrl+T")
        menu.addAction(mark_unskip_action)

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

        self.remove_shortcut = QShortcut(QKeySequence("Ctrl+D"), self)
        self.remove_shortcut.activated.connect(self.remove_rows)

        self.skip_shortcut = QShortcut(QKeySequence("Ctrl+R"), self)
        self.skip_shortcut.activated.connect(self.skip)

        self.unskip_shortcut = QShortcut(QKeySequence("Ctrl+T"), self)
        self.unskip_shortcut.activated.connect(self.unskip)

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
                1000
            )
            if ok and num > 0:
                rows_to_insert = num
                self.insertRowsRequested.emit(position, rows_to_insert)
            else:
                return

    def remove_rows(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            rows = sorted(list(set(index.row() for index in selection.indexes() if index.isValid())))
        elif self.table.currentIndex().isValid():
            rows = [self.table.currentIndex().row()]
        else:
            return
        if rows:
            reply = QMessageBox.question(
                self,
                "Confirm Delete",
                f"Are you sure you want to delete {len(rows)} row(s)",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.deleteRowsRequested.emit(rows)
            else:
                return



    def merge_cells(self):
        self.table.clearSpans()

        model = self.model
        command_col = 6
        desc_col = 3

        blocks = []
        start_row_stack = []

        start_row = 3
        end_row = model.rowCount() - 3
        if end_row < start_row:
            return

        row = 3
        for row in range(row, end_row):
            command_index = model.index(row, command_col)
            command = model.data(command_index, Qt.ItemDataRole.DisplayRole)

            if command == "StartScenario":
                start_row_stack.append(row)

            elif command == "EndScenario":
                if start_row_stack:
                    start_row = start_row_stack.pop()
                    blocks.append((start_row, row))

        for start, end in blocks:
            span_size = end - start + 1
            if span_size > 1:
                self.table.setSpan(start, desc_col, span_size, 1)

    def skip(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            rows = sorted(list(set(index.row() for index in selection.indexes() if index.isValid())))
            if rows:
                self.skipRequested.emit(rows)

    def unskip(self):
        selection = self.table.selectionModel().selection()
        if not selection.isEmpty():
            rows = sorted(list(set(index.row() for index in selection.indexes() if index.isValid())))
            if rows:
                self.unskipRequested.emit(rows)

    def _on_data_changed_commands(self, topLeft, bottomRight, roles):
        command_col = 6

        if topLeft.column() <= command_col <= bottomRight.column():
            self.merge_cells()
