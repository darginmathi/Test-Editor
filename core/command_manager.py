from PyQt6.QtCore import pyqtSignal, QObject
from .commands import COMMANDS

class CommandManager(QObject):
    commandsReloaded = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.commands = sorted(list(COMMANDS.keys()))
        self.commandsReloaded.emit(self.commands)

    def get_command_names(self):
        return self.commands
