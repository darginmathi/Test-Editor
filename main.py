import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget,
    QPushButton, QFileDialog, QMessageBox, QHeaderView, QHBoxLayout
)
from PyQt6.QtCore import QAbstractTableModel, Qt, QModelIndex

class SpreadSheetModel(QAbstractTableModel):
    def __init__(self, df=None):
        super().__init__()
        if df is None or df.empty:
            self.df = self.createEmptyStructure()
        else:
            self.df = df.copy()

    def createEmptyStructure(self):
        columns = ["A", "B", "C", "D", "E",
                  "F", "G", "H", "I", "J"]
        data = [
            ["Type", "ID", "Skip", "Description", "Steps Performed","Expected Results", "Command", "Data1", "Data2", "Data3"],
            ["UCB", "Graphs", "", "SynOption Graphs Test Scripts", "", "", "", "", "", ""],
            ["TC", "TC-LN_AUT101", "", "", "", "", "", "", "", ""],
            ["UCF", "Graphs", "", "SynOption Graphs Test Scripts", "", "", "", "", "", ""],
            ["END", "", "", "", "", "", "", "", "", ""]
        ]

        return pd.DataFrame(data, columns=columns)

    def rowCount(self, parent=QModelIndex()):
        return len(self.df)

    def columnCount(self, parent=QModelIndex()):
        return len(self.df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            value = self.df.iat[row, col]
            if pd.isna(value):
                return str("")
            return str(value)
        return None

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if not index.isValid() or role != Qt.ItemDataRole.EditRole:
            return False
        row, col = index.row(), index.column()
        if row < 1 or row >= len(self.df) - 1:  # Allow editing between first and last rows
            return False
        self.df.iat[row, col] = value
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        row = index.row()
        if row < 1 or row >= len(self.df) - 1:  # First and last rows are read-only
            return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self.df.columns[section]
        return None

    def insertRow(self, row):
        if row < 1 or row >= len(self.df) - 1:
            return False

        new_row = {col: "" for col in self.df.columns}
        new_row["Type"] = "TC"

        self.beginInsertRows(QModelIndex(), row, row)
        top_part = self.df.iloc[:row]
        bottom_part = self.df.iloc[row:]
        self.df = pd.concat([top_part, pd.DataFrame([new_row]), bottom_part], ignore_index=True)
        self.endInsertRows()
        self._updateID()
        return True

    def deleteRow(self, row):
        if row < 1 or row >= len(self.df) - 1:
            return False

        self.beginRemoveRows(QModelIndex(), row, row)
        self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
        self.endRemoveRows()
        self._updateID()
        return True

    def _updateID(self):
        tc_counter = 101
        for i in range(1, len(self.df) - 1):
            if i < len(self.df) and self.df.iloc[i, 0] == "TC":
                current_id = self.df.iat[i, 1]
                if not current_id:
                    new_id = f"TC-LN_AUT{tc_counter}"
                    self.df.iat[i, 1] = new_id
                    tc_counter += 1

                    id_index = self.index(i, 1)
                    self.dataChanged.emit(id_index, id_index)


class SpreadSheetEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Case Editor")
        self.resize(1400, 800)
        self.model = SpreadSheetModel()
        self.setupUI()

    def setupUI(self):
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setSortingEnabled(False)
        self.setColumnWidths()
        self.createButtons()
        self.setupLayout()

    def setColumnWidths(self):
        column_widths = {
            "Type": 80,
            "ID": 150,
            "Skip": 60,
            "Description": 200,
            "Steps Performed": 100,
            "Expected Results": 100,
            "Command": 150,
            "Data1": 120,
            "Data2": 120,
            "Data3": 120
        }

        for col, width in column_widths.items():
            if col in self.model.df.columns:
                col_index = self.model.df.columns.get_loc(col)
                self.table.setColumnWidth(col_index, width)

    def createButtons(self):
        self.add_row_btn = QPushButton("Add Test Case")
        self.add_row_btn.clicked.connect(self.addTestcase)

        self.remove_row_btn = QPushButton("Remove Selected Test Case")
        self.remove_row_btn.clicked.connect(self.removeTestcase)

        self.save_btn = QPushButton("Save Excel")
        self.save_btn.clicked.connect(self.save_excel)

        self.load_btn = QPushButton("Load Excel")
        self.load_btn.clicked.connect(self.load_excel)

        self.new_btn = QPushButton("New Sheet")
        self.new_btn.clicked.connect(self.newSheet)

    def setupLayout(self):
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_row_btn)
        button_layout.addWidget(self.remove_row_btn)
        button_layout.addWidget(self.new_btn)
        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.save_btn)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.table)
        main_layout.addLayout(button_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def addTestcase(self):
        selected_indexes = self.table.selectionModel().selectedRows()
        if selected_indexes:
            row = selected_indexes[0].row() + 1
        else:
            row = len(self.model.df) - 1  # Insert before the last row (END)

        if self.model.insertRow(row):
            QMessageBox.information(self, "Success", "Test case added successfully!")
        else:
            QMessageBox.warning(self, "Warning", "Cannot add test case at this position!")

    def removeTestcase(self):
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            QMessageBox.warning(self, "Warning", "Please select a row to remove!")
            return

        row = selected_indexes[0].row()
        if self.model.deleteRow(row):
            QMessageBox.information(self, "Success", "Test case removed successfully!")
        else:
            QMessageBox.warning(self, "Warning", "Cannot remove this row!")

    def save_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            try:
                self.model.df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"File saved successfully to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Load Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.model = SpreadSheetModel(df)
                self.table.setModel(self.model)
                self.setColumnWidths()
                QMessageBox.information(self, "Success", f"File loaded successfully from {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def newSheet(self):
        reply = QMessageBox.question(
            self, "New Sheet", "Are you sure you want to create a new sheet? Unsaved changes will be lost.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.model = SpreadSheetModel()
            self.table.setModel(self.model)
            self.setColumnWidths()


def main():
    app = QApplication(sys.argv)
    editor = SpreadSheetEditor()
    editor.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
