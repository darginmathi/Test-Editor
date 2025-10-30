import pandas as pd
import os
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtWidgets import QFileDialog
from PyQt6.QtCore import QSettings, pyqtSignal, QObject

class CommandManager(QObject):
    commandsReloaded = pyqtSignal(list)
    SETTINGS_KEY = "commandFilePath"

    def __init__(self):
        super().__init__()
        self.commands = []
        self.file_path = None

    def load_commands(self, filepath=None):
        if not filepath or not os.path.exists(filepath):
            return False
        try:
            df = pd.read_excel(filepath, sheet_name=0, usecols=[0], header=None)
            self.commands = sorted(df[0].dropna().astype(str).tolist())
            self.file_path = filepath
            self.commandsReloaded.emit(self.commands)
            return True
        except FileNotFoundError:
            QMessageBox.critical(None, "Error", f"Command file not found:\n{filepath}")
            self.commands = []
            return False
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to load commands from {filepath}:\n{str(e)}")
            self.commands = []
            return False

    def command_file_prompt(self, parent_widget=None):
        start_dir = os.path.dirname(self.file_path) if self.file_path else os.path.expanduser("~")

        filepath, _ = QFileDialog.getOpenFileName(parent_widget, "Select E2E commands file", start_dir, "Excel Files (*.xlsx *.xls)")

        if filepath:
            settings = QSettings("TestEditor", "Settings")
            settings.setValue(self.SETTINGS_KEY, filepath)
            return self.load_commands(filepath)
        else:
            return False

    def get_command_names(self):
        return self.commands

    def get_file_path(self):
        return self.file_path
