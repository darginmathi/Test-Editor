from PyQt6.QtWidgets import (QStyledItemDelegate, QLineEdit, QComboBox,
                             QCompleter, QStyleOptionViewItem, QStyle)
from PyQt6.QtCore import Qt, QModelIndex, QTimer, QStringListModel
from PyQt6.QtGui import QColor, QPalette

from core.undo_commands import EditCellCommand
from models import TestScenarioModel, ObjRepoModel
from core.commands import COMMAND_ARGS

class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, parent, undo_stack, commands, colors):
        super().__init__(parent)
        self.parent_tab = parent
        self.undo_stack = undo_stack
        self.commands = set(commands)
        self.colors = colors
        self.cached_objects = []  # Cache for objects
        self.cached_objects_set = set()
        self.objects_dirty = True
        self.duplicate_names_cache = set()
        self.duplicate_names_dirty = True

        self.editor = None
        self.current_model = None
        self.current_index = QModelIndex()

        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(750) # Auto-save delay
        self.timer.timeout.connect(self._on_timer_timeout)


    def createEditor(self, parent, option, index):
        model = index.model()
        row = index.row()
        col = index.column()

        editor = None

        if isinstance(model, TestScenarioModel):
            type_index = model.index(row, TestScenarioModel.TYPE_COL)
            type_value = model.data(type_index, Qt.ItemDataRole.EditRole)

            if type_value == "TC":
                if col == TestScenarioModel.COMMAND_COL:
                    editor = QComboBox(parent)
                    editor.setEditable(True)
                    editor.addItems(self.commands)

                    completer = QCompleter(editor)
                    string_list_model = QStringListModel(self.commands, completer)
                    completer.setModel(string_list_model)
                    completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                    completer.setFilterMode(Qt.MatchFlag.MatchContains)
                    completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                    editor.setCompleter(completer)

                elif TestScenarioModel.DATA1_COL <= col <= TestScenarioModel.DATA1_COL + 4:
                    objects = self._get_objects()
                    if objects:
                        editor = QComboBox(parent)
                        editor.setEditable(True)
                        editor.addItems(objects)

                        completer = QCompleter(editor)
                        string_list_model = QStringListModel(objects, completer)
                        completer.setModel(string_list_model)
                        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                        completer.setFilterMode(Qt.MatchFlag.MatchContains)
                        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                        editor.setCompleter(completer)
        if editor is None:
            editor = super().createEditor(parent, option, index)

        if isinstance(editor, (QLineEdit, QComboBox)):
            self.editor = editor
            self.current_model = model
            self.current_index = QModelIndex(index)

            if isinstance(editor, QLineEdit):
                editor.textChanged.connect(self.on_editor_text_changed)
            elif isinstance(editor, QComboBox):
                editor.lineEdit().textChanged.connect(self.on_editor_text_changed)

            if self.current_model:
                self.current_model.dataChanged.connect(self.on_model_data_changed)
        return editor

    def destroyEditor(self, editor, index):
        if self.current_model:
            try:
                self.current_model.dataChanged.disconnect(self.on_model_data_changed)
            except TypeError:
                pass
        if isinstance(editor, QLineEdit):
            try:
                editor.textChanged.disconnect(self.on_editor_text_changed)
            except TypeError:
                pass
        elif isinstance(editor, QComboBox):
            try:
                editor.lineEdit().textChanged.disconnect(self.on_editor_text_changed)
            except TypeError:
                pass

        if editor is self.editor:
            self.editor = None
            self.current_model = None
            self.current_index = QModelIndex()

        super().destroyEditor(editor, index)

    def on_editor_text_changed(self, text):
        self.timer.start()

    def _on_timer_timeout(self):
        if self.editor and self.current_index.isValid():
            cursor_pos = None
            if isinstance(self.editor, QLineEdit):
                cursor_pos = self.editor.cursorPosition()
            elif isinstance(self.editor, QComboBox):
                cursor_pos = self.editor.lineEdit().cursorPosition()

            self.commitData.emit(self.editor)

            if self.editor and cursor_pos is not None:
                if isinstance(self.editor, QLineEdit):
                    self.editor.setCursorPosition(cursor_pos)
                elif isinstance(self.editor, QComboBox):
                    self.editor.lineEdit().setCursorPosition(cursor_pos)

    def on_model_data_changed(self, topLeft, bottomRight, roles):
        if self.editor and self.current_index.isValid():
            try:
                self.editor.metaObject()
            except RuntimeError:
                return

            if (topLeft.row() <= self.current_index.row() <= bottomRight.row() and
                topLeft.column() <= self.current_index.column() <= bottomRight.column()):

                new_model_value = self.current_model.data(self.current_index, Qt.ItemDataRole.EditRole)
                new_model_value_str = str(new_model_value) if new_model_value is not None else ""

                editor_value = ""
                if isinstance(self.editor, QLineEdit):
                    editor_value = self.editor.text()
                elif isinstance(self.editor, QComboBox):
                    editor_value = self.editor.currentText()

                if editor_value != new_model_value_str:
                    blocked = self.editor.blockSignals(True)
                    if isinstance(self.editor, QLineEdit):
                        self.editor.setText(new_model_value_str)
                    elif isinstance(self.editor, QComboBox):
                        self.editor.setCurrentText(new_model_value_str)
                    self.editor.blockSignals(blocked)

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole)
        value_str = str(value) if value is not None else ""
        if isinstance(editor, QComboBox):
            blocked = editor.blockSignals(True)
            editor.setCurrentText(value_str)
            editor.blockSignals(blocked)
            editor.lineEdit().selectAll()
        elif isinstance(editor, QLineEdit):
            editor.setText(value_str)
        else:
            super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        new_value = ""
        if isinstance(editor, QComboBox):
            new_value = editor.currentText()
        elif isinstance(editor, QLineEdit):
            new_value = editor.text()
        else:
             super().setModelData(editor, model, index)
             return

        old_value = model.data(index, Qt.ItemDataRole.EditRole)

        if str(new_value) != str(old_value):
            command = EditCellCommand(model, index, new_value, old_value)
            self.undo_stack.push(command)

    def update_command_list(self, new_commands):
        self.commands = set(new_commands)

    def invalidate_objects_cache(self):
        self.objects_dirty = True

    def invalidate_duplicate_names_cache(self):
        self.duplicate_names_dirty = True

    def _get_objects(self):
        if self.objects_dirty:
            obj_repo_model = self.parent_tab.model2
            if obj_repo_model and not obj_repo_model.df.empty:
                self.cached_objects = obj_repo_model.df.iloc[:, ObjRepoModel.NAME_COL].dropna().astype(str).unique().tolist()
                self.cached_objects.sort()
                self.cached_objects_set = set(self.cached_objects)
            else:
                self.cached_objects = []
                self.cached_objects_set = set()
            self.objects_dirty = False
        return self.cached_objects

    def _get_duplicate_names(self, model):
        if self.duplicate_names_dirty:
            self.duplicate_names_cache = model.get_duplicate_names()
            self.duplicate_names_dirty = False
        return self.duplicate_names_cache

    def paint(self, painter, option, index):
        model = index.model()
        current_row = index.row()
        current_col = index.column()
        total_rows = model.rowCount()
        option = QStyleOptionViewItem(option)
        is_special_row = False

        # Make selected text bold
        if option.state & QStyle.StateFlag.State_Selected:
            font = option.font
            font.setBold(True)
            option.font = font

        if isinstance(model, TestScenarioModel):
            if (current_row == 0 or
                current_row == 1 or
                current_row == total_rows - 1 or
                current_row == total_rows - 2):
                option.font.setBold(True)
                is_special_row = True
            elif current_col in [TestScenarioModel.COMMAND_COL, TestScenarioModel.DESC_COL, TestScenarioModel.SKIP_COL]:
                option.font.setBold(True)
            if current_col == TestScenarioModel.STEPS_COL:
                value = model.data(index, Qt.ItemDataRole.DisplayRole)
                if value in ["Scenario Started", "Scenario Ended"]:
                    option.font.setBold(True)

        elif isinstance(model, ObjRepoModel):
            if (current_row == 0 or
                current_row == total_rows - 1):
                option.font.setBold(True)
                is_special_row = True

        final_color = option.palette.color(QPalette.ColorRole.Text)

        if not is_special_row:
            if isinstance(model, TestScenarioModel):
                value = index.data(Qt.ItemDataRole.DisplayRole)

                if value:
                    if current_col == TestScenarioModel.COMMAND_COL:
                        if value not in self.commands:
                            final_color = QColor("#E57373")

                    elif TestScenarioModel.DATA1_COL <= current_col <= TestScenarioModel.DATA5_COL:
                        if self.objects_dirty:
                            self._get_objects()
                        if value not in self.cached_objects_set:
                            final_color = QColor("#64B5F6")
            elif isinstance(model, ObjRepoModel):
                if current_col == ObjRepoModel.NAME_COL:
                    value = index.data(Qt.ItemDataRole.DisplayRole)
                    if value:
                        duplicate_names = self._get_duplicate_names(model)
                        if value in duplicate_names:
                            final_color = QColor("#E57373")

        option.palette.setColor(QPalette.ColorRole.Text, final_color)
        option.palette.setColor(QPalette.ColorRole.HighlightedText, final_color)

        # if not is_special_row and isinstance(model, TestScenarioModel):
        #     command_index = model.index(current_row, TestScenarioModel.COMMAND_COL)
        #     command_value = model.data(command_index, Qt.ItemDataRole.DisplayRole)

        #     if command_value == "StartScenario":
        #         painter.save()
        #         pen = painter.pen()
        #         pen.setColor(QColor("#B0B0B0"))
        #         pen.setWidth(1)
        #         painter.setPen(pen)
        #         painter.drawLine(option.rect.topLeft(), option.rect.topRight())
        #         painter.restore()
        #     if command_value == "EndScenario":
        #         painter.save()
        #         pen = painter.pen()
        #         pen.setColor(QColor("#B0B0B0"))
        #         pen.setWidth(1)
        #         painter.setPen(pen)
        #         painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())
        #         painter.restore()

        super().paint(painter, option, index)

        # Placeholder drawing logic
        if isinstance(model, TestScenarioModel) and not is_special_row:
            # Check if it's a data column
            if TestScenarioModel.DATA1_COL <= current_col <= TestScenarioModel.DATA5_COL:
                cell_value = model.data(index, Qt.ItemDataRole.DisplayRole)

                # Only draw placeholder if the cell is empty
                if not cell_value:
                    command_index = model.index(current_row, TestScenarioModel.COMMAND_COL)
                    command_name = model.data(command_index, Qt.ItemDataRole.DisplayRole)

                    if command_name in COMMAND_ARGS:
                        args = COMMAND_ARGS[command_name]
                        arg_index = current_col - TestScenarioModel.DATA1_COL

                        if arg_index < len(args):
                            placeholder_text = f"<{args[arg_index]}>"

                            # Set font and color for placeholder
                            font = painter.font()
                            painter.setFont(font)
                            painter.setPen(QColor(Qt.GlobalColor.gray))

                            # Draw the text
                            painter.drawText(option.rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, placeholder_text)

                            # Restore original font and pen for subsequent drawing
                            painter.setFont(option.font)
                            painter.setPen(option.palette.color(QPalette.ColorRole.Text))



