from PyQt6.QtWidgets import QStyledItemDelegate, QCompleter, QComboBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

class AutocompleteDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.model = None

    def setModel(self, model):
        self.model = model

    def createEditor(self, parent, option, index):
        if not self.model:
            return super().createEditor(parent, option, index)

        try:
            # Check if this cell should have autocomplete
            if hasattr(self.model, 'is_autocomplete_cell') and self.model.is_autocomplete_cell(index):
                # Get autocomplete options for this cell
                current_value = self.model.data(index, Qt.ItemDataRole.DisplayRole) or ""
                options = self.model.get_autocomplete_options(index)

                if options:
                    # Create combobox with autocomplete
                    editor = QComboBox(parent)
                    editor.setEditable(True)
                    editor.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)

                    # Set font
                    font = QFont("Arial", 10)
                    editor.setFont(font)

                    # Add all options
                    editor.addItems(options)

                    # Setup completer
                    completer = QCompleter(options, editor)
                    completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                    completer.setFilterMode(Qt.MatchFlag.MatchContains)
                    completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                    editor.setCompleter(completer)

                    # Set current value
                    editor.setCurrentText(current_value)

                    return editor
        except Exception as e:
            print(f"Error creating autocomplete editor: {e}")

        return super().createEditor(parent, option, index)

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            current_value = self.model.data(index, Qt.ItemDataRole.DisplayRole) or ""
            editor.setCurrentText(current_value)
        else:
            super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)
        else:
            super().setModelData(editor, model, index)
