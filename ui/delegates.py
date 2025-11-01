from PyQt6.QtWidgets import (QStyledItemDelegate, QLineEdit, QComboBox,
                             QCompleter, QStyleOptionViewItem, QApplication, QStyle)
from PyQt6.QtCore import Qt, QModelIndex, QTimer, QStringListModel
from PyQt6.QtGui import QColor, QPalette

from core.undo_commands import EditCellCommand
from models import TestScenarioModel, ObjRepoModel

class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, parent, undo_stack, commands, colors):
        super().__init__(parent)
        self.parent_tab = parent
        self.undo_stack = undo_stack
        self.commands = commands
        self.colors = colors
        self.cached_objects = []  # Cache for objects
        self.objects_dirty = True

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
            self.commitData.emit(self.editor)

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
        self.commands = new_commands

    def invalidate_objects_cache(self):
        self.objects_dirty = True

    def _get_objects(self):
        if self.objects_dirty:
            obj_repo_model = self.parent_tab.model2
            if obj_repo_model and not obj_repo_model.df.empty:
                self.cached_objects = obj_repo_model.df.iloc[:, ObjRepoModel.NAME_COL].dropna().astype(str).unique().tolist()
                self.cached_objects.sort()
            else:
                self.cached_objects = []
            self.objects_dirty = False
        return self.cached_objects

    def paint(self, painter, option, index):
        if index.row() == 0:
            super().paint(painter, option, index)
            return

        option = QStyleOptionViewItem(option)

        final_color = option.palette.color(QPalette.ColorRole.Text)

        model = index.model()
        if isinstance(model, TestScenarioModel):
            col = index.column()
            value = index.data(Qt.ItemDataRole.DisplayRole)

            if value:
                if col == TestScenarioModel.COMMAND_COL:
                    if value not in self.commands:
                        final_color = QColor("#E57373")

                elif TestScenarioModel.DATA1_COL <= col <= TestScenarioModel.DATA5_COL:
                    objects = self._get_objects()
                    if value not in objects:
                        final_color = QColor("#64B5F6")

        option.palette.setColor(QPalette.ColorRole.Text, final_color)
        option.palette.setColor(QPalette.ColorRole.HighlightedText, final_color)

        super().paint(painter, option, index)


