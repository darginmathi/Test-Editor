import os
import re
from datetime import datetime
from PyQt6.QtCore import QObject, pyqtSignal, QFileSystemWatcher, QTimer, QDateTime

class LogFinder(QObject):
    log_found = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, start_time: QDateTime, project: str, module: str, logs_dir: str):
        super().__init__()
        self._start_time = start_time
        self._project = project
        self._module = module
        self._logs_dir = logs_dir

        self._watcher = QFileSystemWatcher([self._logs_dir])
        self._watcher.directoryChanged.connect(self._on_directory_change)

        self._timeout_timer = QTimer()
        self._timeout_timer.setSingleShot(True)
        self._timeout_timer.timeout.connect(self._cleanup)
        self._timeout_timer.start(60000)

    def _on_directory_change(self, path):
        for filename in os.listdir(path):
            if self._project in filename and self._module in filename and filename.endswith("_TestLog.html"):
                filepath = os.path.join(path, filename)

                try:

                    file_timestamp_str = filename.split('_Test_')[-1].split('_TestLog.html')[0]
                    file_datetime = self._parse_timestamp(file_timestamp_str)

                    if file_datetime.isValid() and file_datetime > self._start_time:
                        self.log_found.emit(filepath)
                        self._cleanup()
                        return
                except Exception as e:
                    print(f"{e}")
                    continue

    def _parse_timestamp(self, timestamp_str):
        try:
            match = re.match(r'(\d{4})-(\d{2})-(\d{2})_(\d+)h(\d+)m(\d+)s(\d+)ms', timestamp_str)
            if not match:
                return QDateTime()

            year, month, day, hour, minute, second, millisecond = match.groups()

            dt = datetime(
                int(year), int(month), int(day),
                int(hour), int(minute), int(second),
                int(millisecond.ljust(3, '0')) * 1000  # Convert to microseconds
            )

            return QDateTime(dt)

        except Exception as e:
            print(f"Error parsing timestamp '{timestamp_str}': {e}")
            return QDateTime()

    def _cleanup(self):
        self._watcher.removePaths(self._watcher.directories())
        self._watcher.directoryChanged.disconnect()
        self._timeout_timer.stop()
        self.finished.emit()
        self.deleteLater()



