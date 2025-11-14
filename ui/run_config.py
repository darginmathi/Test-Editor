from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QComboBox, QSpinBox, QDialogButtonBox,
                             QGroupBox, QFormLayout)
from models.run_config import RunConfig

class RunConfigDialog(QDialog):
    def __init__(self, current_config, parent=None):
        super().__init__(parent)
        self.current_config = current_config
        self.setup_ui()
        self.load_current_config()

    def setup_ui(self):
        self.setWindowTitle("Run Configuration")
        self.setModal(True)
        self.resize(400, 300)

        layout = QVBoxLayout(self)

        info_group = QGroupBox("Test Information")
        info_layout = QFormLayout(info_group)

        self.project_label = QLabel(self.current_config.project_name)
        self.module_label = QLabel(self.current_config.module_name)
        self.url_label = QLabel(self.current_config.base_url)

        info_layout.addRow("Project:", self.project_label)
        info_layout.addRow("Module:", self.module_label)
        info_layout.addRow("URL:", self.url_label)

        layout.addWidget(info_group)

        settings_group = QGroupBox("Run Settings")
        settings_layout = QFormLayout(settings_group)

        self.browser_combo = QComboBox()
        self.browser_combo.addItems(["Chrome", "ChromeHL"])

        self.video_combo = QComboBox()
        self.video_combo.addItems(["NONE", "BOTH"])

        self.wait_time_spin = QSpinBox()
        self.wait_time_spin.setRange(10, 120)
        self.wait_time_spin.setSuffix(" seconds")

        settings_layout.addRow("Browser:", self.browser_combo)
        settings_layout.addRow("Video Recording:", self.video_combo)
        settings_layout.addRow("Wait Time:", self.wait_time_spin)

        layout.addWidget(settings_group)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(button_box)

    def load_current_config(self):
        self.browser_combo.setCurrentText(self.current_config.browser)
        self.video_combo.setCurrentText(self.current_config.video_option)
        self.wait_time_spin.setValue(self.current_config.wait_time)

    def get_config(self):
        return RunConfig(
            project_name=self.current_config.project_name,
            module_name=self.current_config.module_name,
            base_url=self.current_config.base_url,
            browser=self.browser_combo.currentText(),
            video_option=self.video_combo.currentText(),
            wait_time=self.wait_time_spin.value()
        )
