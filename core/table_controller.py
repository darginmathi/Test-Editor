from PyQt6.QtCore import QObject, Qt
from PyQt6.QtWidgets import QApplication, QMessageBox

from .undo_commands import EditCellCommand, InsertRowsCommand, DeleteRowsCommand
from .steps_performed import get_steps_performed
from .expected_results import get_expected_results
from models.test_scenario import TestScenarioModel
from .commands import COMMANDS
from .utils import clean_object_name

class TableController(QObject):

    def __init__(self, main, undo_stack):
        super().__init__(main)
        self.main  = main
        self.undo_stack = undo_stack

    def insert_rows(self, model, position, count):
        command = InsertRowsCommand(model, position, count)
        self.undo_stack.push(command)

    def delete_rows(self, model, rows):
        if not rows:
            return
        self.undo_stack.beginMacro("Delete Rows")

        for row_index in sorted(rows, reverse=True):
            command = DeleteRowsCommand(model, row_index, 1)
            self.undo_stack.push(command)
        self.undo_stack.endMacro()

    def copy(self, model, selection):
        indexes = selection.indexes()
        if not indexes:
            return

        rows = sorted(set(index.row() for index in indexes))
        cols = sorted(set(index.column() for index in indexes))

        text_lines = []
        for row in rows:
            row_data = []
            for col in cols:
                index = model.index(row, col)
                value = model.data(index, Qt.ItemDataRole.DisplayRole)
                row_data.append(str(value) if value is not None else "")
            text_lines.append('\t'.join(row_data))

        QApplication.clipboard().setText('\n'.join(text_lines))

    def paste(self, model, start_index):
        if not start_index.isValid():
            return

        clipboard_text = QApplication.clipboard().text().strip()
        if not clipboard_text:
            return

        data = []
        for line in clipboard_text.split('\n'):
            line = line.strip()
            if line:
                if '\t' in line:
                    cells = line.split('\t')
                else:
                    cells = [line]
                data.append(cells)
        if not data:
            return

        overwrite = False

        start_row = start_index.row()
        start_col = start_index.column()

        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                target_row = start_row + i
                target_col = start_col + j
                if (target_row < model.rowCount() and
                    target_col < model.columnCount()):
                    index = model.index(target_row, target_col)
                    old_value = model.data(index, Qt.ItemDataRole.EditRole)

                    if old_value is not None and str(old_value) != "" and str(value) != str(old_value):
                        overwrite = True
                        break
            if overwrite:
                break

        if overwrite:
            reply = QMessageBox.question(
                self.main,
                "Overwrite existing cells",
                "This operation will overwrite existing data. Do you want to proceed?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply !=QMessageBox.StandardButton.Yes:
                return

        self.undo_stack.beginMacro("Paste")

        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                target_row = start_row + i
                target_col = start_col + j
                if (target_row < model.rowCount() and
                    target_col < model.columnCount()):
                    index = model.index(target_row, target_col)
                    old_value = model.data(index, Qt.ItemDataRole.EditRole)

                    if str(value) != str(old_value):
                        command = EditCellCommand(model, index, value, old_value)
                        self.undo_stack.push(command)

        self.undo_stack.endMacro()

    def cut(self, model, selection):
        self.copy(model, selection)
        self.clear(model, selection)

    def clear(self, model, selection):
        indexes = selection.indexes()
        if not indexes:
            return

        self.undo_stack.beginMacro("Clear Contents")
        for index in indexes:
            if index.isValid():
                old_value = model.data(index, Qt.ItemDataRole.EditRole)
                if str(old_value) != "":
                    command = EditCellCommand(model, index, "", old_value)
                    self.undo_stack.push(command)
        self.undo_stack.endMacro()

    def generate_test_cases(self, model):
        test_case_generated = 0
        self.undo_stack.beginMacro("Generate Test Cases")
        for row in range(model.rowCount()):
            steps_performed_index = model.index(row, TestScenarioModel.STEPS_COL)
            expected_result_index = model.index(row, TestScenarioModel.EXPECTED_COL)

            steps_performed_data = model.data(steps_performed_index, Qt.ItemDataRole.DisplayRole)
            expected_result_data = model.data(expected_result_index, Qt.ItemDataRole.DisplayRole)

            if not steps_performed_data and not expected_result_data:
                command_index = model.index(row, TestScenarioModel.COMMAND_COL)
                command = model.data(command_index, Qt.ItemDataRole.DisplayRole)

                if command in COMMANDS:

                    value1_index = model.index(row, TestScenarioModel.DATA1_COL)
                    value1 = model.data(value1_index, Qt.ItemDataRole.DisplayRole)

                    value2_index = model.index(row, TestScenarioModel.DATA2_COL)
                    value2 = model.data(value2_index, Qt.ItemDataRole.DisplayRole)

                    if value1 is None: value1 = ""
                    if value2 is None: value2 = ""

                    value1 = clean_object_name(value1)
                    value2 = value2

                    steps_performed_text = get_steps_performed(command, value1, value2)
                    expected_result_text = get_expected_results(command, value1, value2)

                    old_steps_performed = model.data(steps_performed_index, Qt.ItemDataRole.EditRole)
                    if steps_performed_text != old_steps_performed:
                        command_obj = EditCellCommand(model, steps_performed_index, steps_performed_text, old_steps_performed)
                        self.undo_stack.push(command_obj)
                        test_case_generated += 1

                    old_expected_result = model.data(expected_result_index, Qt.ItemDataRole.EditRole)
                    if expected_result_text != old_expected_result:
                        command_obj = EditCellCommand(model, expected_result_index, expected_result_text, old_expected_result)
                        self.undo_stack.push(command_obj)
        self.undo_stack.endMacro()
        return test_case_generated
