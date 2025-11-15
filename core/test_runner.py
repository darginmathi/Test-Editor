from PyQt6.QtCore import QObject, QProcess, pyqtSignal, QProcessEnvironment, QTimer
from models.run_config import RunConfig
import os

# MAVEN_EXECUTABLE_PATH = "C:\\Users\\meetd\\Downloads\\apache-maven-3.9.10-bin\\apache-maven-3.9.10\\bin\\mvn.cmd"
MAVEN_EXECUTABLE_PATH = "mvn"

class TestRunner(QObject):
    process_started = pyqtSignal()
    process_finished = pyqtSignal(int)
    output_received = pyqtSignal(str)
    error_received = pyqtSignal(str)

    def __init__(self, e2e_dir: str):
        super().__init__()
        self.e2e_dir = e2e_dir
        self.process = None
        self.is_running = False
        self._force_kill_timer = QTimer()
        self._force_kill_timer.setSingleShot(True)
        self._force_kill_timer.timeout.connect(self._force_kill_process)

    def run_test(self, config: RunConfig):
        if self.is_running:
            self.output_received.emit("✗ Test already running! Stop it first.")
            return

        # Verify directory and pom.xml
        if not os.path.exists(self.e2e_dir):
            self.error_received.emit(f"✗ E2E directory does not exist: {self.e2e_dir}")
            return

        pom_file = os.path.join(self.e2e_dir, "pom.xml")
        if not os.path.exists(pom_file):
            self.error_received.emit(f"✗ Current Directory is not a maven project: {self.e2e_dir}")
            return

        args = self._build_command_args(config)

        self.process = QProcess()
        self.process.setWorkingDirectory(self.e2e_dir)

        # Use system environment as-is (your PATH already has everything needed)
        env = QProcessEnvironment.systemEnvironment()

        # Just set JAVA_HOME since it's missing
        if not env.contains("JAVA_HOME"):
            env.insert("JAVA_HOME", r"C:\Program Files\Java\jdk-21")

        self.process.setProcessEnvironment(env)

        self.process.readyReadStandardOutput.connect(self._on_stdout)
        self.process.readyReadStandardError.connect(self._on_stderr)
        self.process.finished.connect(self._on_finished)
        self.process.errorOccurred.connect(self._on_error)

        self.is_running = True
        self.process_started.emit()
        self.output_received.emit(f"[INFO] Starting test: {config.project_name}/{config.module_name}")

        # Use mvn.cmd explicitly
        self.process.start("mvn.cmd", args)

    def _build_command_args(self, config: RunConfig) -> list:
        exec_args = f"{config.project_name} {config.base_url} {config.browser} {config.video_option} {config.module_name} {config.wait_time}"

        return [
            "clean",
            "compile",
            "exec:java",
            f"-Dexec.mainClass=com.startAuto.StartAuto",
            f"-Dexec.args={exec_args}"
        ]

    def stop_test(self):
        if not self.is_running or not self.process:
            return

        self.output_received.emit("[INFO] Stopping test execution...")
        self.process.terminate()

        self._force_kill_timer.start(3000)

    def _force_kill_process(self):
        if self.is_running and self.process and self.process.state() == QProcess.ProcessState.Running:
            self.process.kill()
            self.output_received.emit("[INFO] Force killed process")
            self.is_running = False
            self.output_received.emit("[INFO] Test stopped")

    def _on_stdout(self):
        if self.process:
            data = self.process.readAllStandardOutput().data().decode()
            self.output_received.emit(data)

    def _on_stderr(self):
        if self.process:
            data = self.process.readAllStandardError().data().decode()
            self.error_received.emit(data)

    def _on_finished(self, exit_code):
        self.is_running = False
        self._force_kill_timer.stop()
        self.process_finished.emit(exit_code)

    def is_test_running(self):
        return self.is_running

    def cleanup(self):
        if self.process:
            self._force_kill_timer.stop()

            if self.process.state() == QProcess.ProcessState.Running:
                self.process.kill()
                self.process.waitForFinished(100)

            try:
                self.process.readyReadStandardOutput.disconnect()
                self.process.readyReadStandardError.disconnect()
                self.process.finished.disconnect()
                self.process.errorOccurred.disconnect()
            except:
                pass

            self.process = None

        self.is_running = False

    def _on_error(self, error: QProcess.ProcessError):
        error_messages = {
            QProcess.ProcessError.FailedToStart: "❌ Failed to start the process. Check if Maven and Java are properly configured.",
            QProcess.ProcessError.Crashed: "The process crashed.",
            QProcess.ProcessError.Timedout: "The process timed out.",
            QProcess.ProcessError.ReadError: "An error occurred while reading from the process.",
            QProcess.ProcessError.WriteError: "An error occurred while writing to the process.",
            QProcess.ProcessError.UnknownError: "An unknown error occurred."
            }
        self.error_received.emit(f"Process Error: {error_messages.get(error, 'An unknown error occurred.')}")

        if error == QProcess.ProcessError.FailedToStart:
            self.error_received.emit("Please ensure:")
            self.error_received.emit("1. Java JDK is installed and JAVA_HOME is set")
            self.error_received.emit("2. Maven is installed and in PATH")
            self.error_received.emit("4. The E2E directory contains a valid Maven project")
