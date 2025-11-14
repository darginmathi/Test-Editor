from dataclasses import dataclass
from PyQt6.QtCore import QSettings

@dataclass
class RunConfig:
    project_name: str = ""
    module_name: str = ""
    base_url: str = ""
    browser: str = "ChromeHL"
    video_option: str = "NONE"
    wait_time: int = 30

    def save_to_settings(self):
        settings = QSettings("TestEditor", "RunConfig")
        settings.setValue("browser", self.browser)
        settings.setValue("video_option", self.video_option)
        settings.setValue("wait_time", self.wait_time)

    def load_from_settings(self):
        settings = QSettings("TestEditor", "RunConfig")
        self.browser = settings.value("browser", "ChromeHL")
        self.video_option = settings.value("video_option", "NONE")
        self.wait_time = settings.value("wait_time", 30, type=int)

default_config = RunConfig()
default_config.load_from_settings()


def get_base_url(project_name: str) -> str:
    return PROJECT_URLS.get(project_name, "")

PROJECT_URLS = {
    "qaoptimus": "https://qaoptimus.synoption.com/#/auth/login",
    "uatoptimus": "https://uatoptimus.synoption.com/#/auth/login",
    "qatitan": "https://qatitan.synoption.com/#/auth/login",
    "uattitan": "https://uattitan.synoption.com/#/auth/login",
    "uatmaqa": "https://uatmaqa.synoption.com/#/auth/login",
    "qaocbctitan": "https://qaocbctitan.synoption.com/#/auth/login",
    "uatocbctitan": "https://uatocbctitan.synoption.com/#/auth/login",
}
