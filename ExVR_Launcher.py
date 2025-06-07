import os
import sys
import pyuac
import json
import time
import shutil
import tempfile
import subprocess
import re
import zipfile
import requests
import argparse
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

#pyinstaller --name ExVR_Launcher --onefile --windowed --icon=./res/logo.ico --upx-dir=D:\software\upx-4.2.4-win64 launcher.py

def get_config_file_path():
    return get_resource_path("exvr_config.json")

def load_config():
    config_path = get_config_file_path()
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            log(f"Error loading config: {e}")
    return {}


def save_config(config):
    config_path = get_config_file_path()
    try:
        os.makedirs(os.path.dirname(config_path), exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        log(f"Config saved to: {config_path}")
    except Exception as e:
        log(f"Error saving config: {e}")
        raise

def delete_config():
    config_path = get_config_file_path()
    try:
        if os.path.exists(config_path):
            os.remove(config_path)
            log(f"Config deleted: {config_path}")
        else:
            log("Config file does not exist.")
    except Exception as e:
        log(f"Error deleting config: {e}")
        raise
APP_NAME = "EXVR"
APP_REG_PATH = r"SOFTWARE\EXVR"
PYTHON_VERSION = "3.11"
PYTHON_DOWNLOAD_URL = "https://mirrors.huaweicloud.com/python/3.11.9/python-3.11.9-amd64.exe"
GITHUB2_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
GITHUB_API_URL = "https://gh-proxy.com/https://api.github.com/repos/{owner}/{repo}/releases/latest"
UPDATE_CHECK_URLS = [
    "https://gh-proxy.com/raw.githubusercontent.com/ExVR-Doc/ExVR-Doc.github.io/main/docs/exvrserverdata.json",
    "https://gh-proxy.com/https://raw.githubusercontent.com/ExVR-Doc/ExVR-Doc.github.io/main/docs/exvrserverdata.json",
    "https://hub.gitmirror.com/https://raw.githubusercontent.com/ExVR-Doc/ExVR-Doc.github.io/main/docs/exvrserverdata.json",
    "https://raw.githubusercontent.com/ExVR-Doc/ExVR-Doc.github.io/main/docs/exvrserverdata.json"
]
GITHUB_REPO_OWNER = "xiaofeiyu0723"
GITHUB_REPO_NAME = "ExVR"
CUSTOM_FOLDER_NAME = "config"
REQUEST_TIMEOUT = 10
PYTHON_CHECK_TIMEOUT = 5
IGNORED_FOLDERS = []
LAU_VERSION = 1
LAU_MAPPING = {
    "modules\\palm_detection_lite.tflite": "mediapipe\\modules\\palm_detection\\palm_detection_lite.tflite",
    "modules\\hand_landmark_tracking_cpu.binarypb": "mediapipe\\modules\\hand_landmark\\hand_landmark_tracking_cpu.binarypb",
}
server_data = {}


def get_install_path():
    config = load_config()
    return config.get("InstallPath")


def get_python_path():
    config = load_config()
    return config.get("PythonPath")


def set_install_path(path):
    config = load_config()
    config["InstallPath"] = path
    save_config(config)


def set_python_path(path):
    config = load_config()
    config["PythonPath"] = path
    save_config(config)


def get_resource_path(relative_path: str) -> str:
    return os.path.join(os.getcwd(), relative_path)


# Modern QSS Stylesheet
modern_qss = """
QWidget {
    background-color: #444444; /* Darker gray background */
    color: #ffffff; /* White text for better contrast on dark background */
}

QDialog, QMessageBox, QProgressDialog {
    background-color: #555555; /* Slightly lighter gray for dialogs */
    border: 1px solid #666666;
    border-radius: 8px;
}

QLabel {
    background-color: transparent; /* Ensure labels have transparent background */
}

QLineEdit, QComboBox, QTreeView {
    border: 1px solid #777777;
    border-radius: 4px;
    padding: 2px 4px;
    background-color: #666666; /* Darker input fields */
    color: #ffffff;
}

QTreeView::item:selected {
    background-color: #777777; /* Selection color for tree view */
}

QPushButton {
    background-color: #0078d4; /* Modern blue */
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 8px; /* Rounded corners */
    font-size: 10pt;
}

QPushButton:hover {
    background-color: #005a9e; /* Darker blue on hover */
}

QPushButton:pressed {
    background-color: #003d6b; /* Even darker blue when pressed */
}

QPushButton:disabled {
    background-color: #555555; /* Gray when disabled */
    color: #888888;
}

QProgressBar {
    border: 1px solid #777777;
    border-radius: 5px;
    background-color: #ffffff; /* White background */
    text-align: center;
    color: #000000; /* Black text for contrast on white background */
}

QProgressBar::chunk {
    background-color: #0078d4; /* Blue progress chunk */
    border-radius: 4px;
    margin: 1px; /* Margin for the chunk */
}

QProgressDialog QLabel {
    background-color: transparent;
    color: #ffffff;
}

QMessageBox QLabel {
    background-color: transparent;
    color: #ffffff;
}

QMessageBox QPushButton {
    min-width: 80px; /* Ensure message box buttons have minimum width */
}

QComboBox QAbstractItemView {
    background-color: #555555;
    color: #ffffff;
    selection-background-color: #777777;
}
"""


def parse_arguments():
    parser = argparse.ArgumentParser(description='EXVR Installer')
    parser.add_argument('-log', action='store_true', help='Enable detailed logging to console')
    return parser.parse_known_args()[0]


def setup_logging(args):
    if args.log:
        import ctypes
        ctypes.windll.kernel32.AllocConsole()
        sys.stdout = open('CONOUT$', 'w')
        sys.stderr = open('CONOUT$', 'w')

    log_dir = get_resource_path("logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"installer_{time.strftime('%Y%m%d_%H%M%S')}.log")

    if args.log:
        class TeeOutput:
            def __init__(self, file_stream, console_stream):
                self.file_stream = file_stream
                self.console_stream = console_stream

            def write(self, message):
                self.file_stream.write(message)
                self.console_stream.write(message)
                self.file_stream.flush()
                self.console_stream.flush()

            def flush(self):
                self.file_stream.flush()
                self.console_stream.flush()

        log_file_stream = open(log_file, "w", encoding="utf-8")
        original_stdout = sys.stdout
        sys.stdout = TeeOutput(log_file_stream, original_stdout)
        sys.stderr = sys.stdout
    else:
        sys.stdout = open(log_file, "w", encoding="utf-8")
        sys.stderr = sys.stdout

    print(f"Logging to: {log_file}")


def log(message):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def create_tmp_folder():
    try:
        tmp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tmp")
        if not os.path.exists(tmp_dir):
            os.makedirs(tmp_dir)
        return tmp_dir
    except Exception as e:
        tmp_dir = tempfile.gettempdir()
        return tmp_dir


def clean_tmp_folder(tmp_dir):
    try:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception as e:
        log(f"Error cleaning temporary folder: {e}")


def show_error_message(title, message):
    log(f"ERROR: {title} - {message}")
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Critical)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.exec()


def show_info_message(title, message):
    log(f"INFO: {title} - {message}")
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Information)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.exec()


def ask_question(title, question):
    log(f"QUESTION: {title} - {question}")
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Question)
    msg_box.setWindowTitle(title)
    msg_box.setText(question)
    msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msg_box.setDefaultButton(QMessageBox.Yes)
    return msg_box.exec()


# 新增函数：复制文件并忽略指定文件夹
def copy_with_ignore(src, dst, ignored_folders=None):
    if ignored_folders is None:
        ignored_folders = []

    log(f"copy file : {src} to {dst}，ig: {ignored_folders}")

    if not os.path.exists(dst):
        os.makedirs(dst)

    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)

        if os.path.isdir(s):
            if item in ignored_folders:
                log(f"ig: {item}")
                if not os.path.exists(d):
                    os.makedirs(d)
                continue

            if os.path.exists(d):
                copy_with_ignore(s, d, ignored_folders)
            else:
                shutil.copytree(s, d, dirs_exist_ok=True)
        else:
            shutil.copy2(s, d)


def quote_path_if_needed(path):
    if " " in path:
        return f'"{path}"'
    return path


# 确保路径使用反斜杠
def normalize_path(path):
    return path.replace("/", "\\")


def is_admin():
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


class CustomFileDialog(QDialog):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select installation path")
        self.setMinimumSize(500, 400)
        self.selected_path = ""
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)

        path_layout = QHBoxLayout()
        path_label = QLabel("Installation path:")
        self.path_edit = QLineEdit("C:\\")
        self.path_edit.textChanged.connect(self.validate_path)

        self.drive_combo = QComboBox()
        self.populate_drives()
        self.drive_combo.currentIndexChanged.connect(self.drive_changed)

        path_layout.addWidget(path_label)
        path_layout.addWidget(self.drive_combo)
        path_layout.addWidget(self.path_edit, 1)

        self.model = QFileSystemModel()
        self.model.setFilter(QDir.AllDirs | QDir.NoDotAndDotDot)
        self.model.setRootPath("C:\\")

        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index("C:\\"))
        self.tree.setColumnWidth(0, 250)
        self.tree.clicked.connect(self.tree_item_clicked)

        for i in range(1, self.model.columnCount()):
            self.tree.hideColumn(i)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: red;")

        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Choose")
        self.ok_button.clicked.connect(self.accept)
        self.ok_button.setEnabled(False)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.status_label, 1)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(path_layout)
        layout.addWidget(self.tree)
        layout.addLayout(button_layout)

        self.validate_path()

    def populate_drives(self):
        available_drives = []
        for drive_letter in range(ord('A'), ord('Z') + 1):
            drive = chr(drive_letter) + ":\\"
            if os.path.exists(drive):
                available_drives.append(drive)

        self.drive_combo.addItems(available_drives)
        index = self.drive_combo.findText("C:\\")
        if index >= 0:
            self.drive_combo.setCurrentIndex(index)

    def drive_changed(self, index):
        drive = self.drive_combo.currentText()
        self.model.setRootPath(drive)
        self.tree.setRootIndex(self.model.index(drive))
        self.path_edit.setText(drive)

    def tree_item_clicked(self, index):
        path = self.model.filePath(index)
        self.path_edit.setText(path)

    def validate_path(self):
        path = self.path_edit.text()

        # Check if path exists
        if not os.path.exists(path):
            self.status_label.setText("Path does not exist")
            self.ok_button.setEnabled(False)
            return

        if re.search("[\u4e00-\u9fff]", path):
            self.status_label.setText("路径不能包含中文字符")
            self.ok_button.setEnabled(False)
            return

        # Check if path is writable
        if not os.access(os.path.dirname(path), os.W_OK):
            self.status_label.setText("No write permission")
            self.ok_button.setEnabled(False)
            return

        # Path is valid
        self.status_label.setText("")
        self.ok_button.setEnabled(True)
        self.selected_path = path

    def get_selected_path(self):
        return self.selected_path


# --- Worker Threads ---
class WorkerSignals(QObject):
    progress = Signal(int)
    finished = Signal()
    error = Signal(str)
    result = Signal(str)
    log = Signal(str)


class PythonCheckWorker(QThread):
    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()
        self._is_running = True

    def stop(self):
        self._is_running = False
        self.wait()

    def _check_command(self, command):
        try:
            self.signals.log.emit(f"Executing: {' '.join(command)}")
            result = subprocess.run(
                command, capture_output=True, text=True,
                timeout=PYTHON_CHECK_TIMEOUT, check=False,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode == 0 and f"Python {PYTHON_VERSION}" in result.stdout:
                self.signals.log.emit(f"Success: {result.stdout.strip()}")
                return True
            else:
                self.signals.log.emit(f"Failed or wrong version: {result.stdout.strip()} {result.stderr.strip()}")
                return False
        except Exception as e:
            self.signals.log.emit(f"Command error: {' '.join(command)} - {e}")
            return False

    def _check_registry(self):
        # 检查ExVR注册表中是否有Python路径
        try:
            python_path = get_python_path()
            if python_path and os.path.exists(python_path) and python_path.endswith("python.exe"):
                if self._check_command([python_path, "--version"]):
                    self.signals.log.emit(f"Found valid Python in JSON config: {python_path}")
                    return True
        except Exception as e:
            self.signals.log.emit(f"JSON config check error: {e}")

        return False

    def run(self):
        try:
            self.signals.log.emit("Starting Python check...")
            python_installed = False

            if self._check_registry():
                python_installed = True

            if not self._is_running:
                return

            result_status = "installed" if python_installed else "not_installed"
            self.signals.log.emit(f"Python check result: {result_status}")
            self.signals.result.emit(result_status)
            self.signals.finished.emit()
        except Exception as e:
            self.signals.log.emit(f"Python check failed: {e}")
            self.signals.error.emit(str(e))


class DownloadWorker(QThread):
    def __init__(self, url, save_path):
        super().__init__()
        self.url = url
        self.save_path = save_path
        self.signals = WorkerSignals()
        self._is_running = True

    def stop(self):
        self._is_running = False
        self.wait()

    def run(self):
        try:
            self.signals.log.emit(f"Starting download: {self.url} to {self.save_path}")
            response = requests.get(self.url, stream=True, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            total_size = int(response.headers.get("content-length", 0))
            downloaded = 0
            os.makedirs(os.path.dirname(self.save_path), exist_ok=True)
            with open(self.save_path, "wb") as file:
                for data in response.iter_content(chunk_size=8192):
                    if not self._is_running:
                        self.signals.log.emit("Download cancelled.")
                        return
                    file.write(data)
                    downloaded += len(data)
                    if total_size > 0:
                        progress = int((downloaded / total_size) * 100)
                        self.signals.progress.emit(progress)
            self.signals.log.emit("Download finished.")
            self.signals.result.emit(self.save_path)
            self.signals.finished.emit()

        except Exception as e:
            self.signals.log.emit(f"Download error: {e}")
            self.signals.error.emit(str(e))


class ExtractWorker(QThread):
    def __init__(self, zip_path, extract_path, final_path=None, ignored_folders=None):
        super().__init__()
        self.zip_path = zip_path
        self.extract_path = extract_path  # 临时解压目录
        self.final_path = final_path  # 最终目标目录
        self.ignored_folders = ignored_folders if ignored_folders else IGNORED_FOLDERS
        self.signals = WorkerSignals()
        self._is_running = True

    def stop(self):
        self._is_running = False
        self.wait()

    def run(self):
        try:
            self.signals.log.emit(f"Ready to decompress: {self.zip_path} to {self.extract_path}")
            os.makedirs(self.extract_path, exist_ok=True)

            # 解压到临时目录
            with zipfile.ZipFile(self.zip_path, "r") as zip_ref:
                total_files = len(zip_ref.namelist())
                for i, file_info in enumerate(zip_ref.infolist()):
                    if not self._is_running:
                        self.signals.log.emit("Decompression cancelled.")
                        return
                    zip_ref.extract(file_info, self.extract_path)
                    progress = int(((i + 1) / total_files) * 50)
                    self.signals.progress.emit(progress)

            if self.final_path:
                self.signals.log.emit(
                    f"Currently copying the file from the temporary directory to the final destination: {self.final_path}")

                extracted_items = os.listdir(self.extract_path)
                if len(extracted_items) == 1 and os.path.isdir(os.path.join(self.extract_path, extracted_items[0])):
                    source_dir = os.path.join(self.extract_path, extracted_items[0])
                else:
                    source_dir = self.extract_path

                copy_with_ignore(source_dir, self.final_path, self.ignored_folders)
                self.signals.progress.emit(100)
            else:
                self.signals.progress.emit(100)

            self.signals.log.emit("Decompression and copying are complete.")
            self.signals.finished.emit()
        except Exception as e:
            self.signals.log.emit(f"Decompression or copying error: {e}")
            self.signals.error.emit(str(e))


def replace_modules_with_json(install_path):
    """根据映射规则替换模块文件"""
    log("Replacing modules")
    try:
        # 获取虚拟环境路径
        venv_path = os.path.join(install_path, "venv")
        python_module_path = os.path.join(venv_path, "Lib", "site-packages")

        # 源文件路径 (EXVR安装目录中的modules文件夹)
        modules_path = install_path

        # 遍历映射关系
        for nu_file, file_path in LAU_MAPPING.items():
            nu_file_path = os.path.join(modules_path, nu_file)

            file_path = os.path.join(python_module_path, file_path)

            if os.path.exists(nu_file_path):
                log(f"Found module file: {nu_file_path}")
                log(file_path)
                try:
                    shutil.copy2(nu_file_path, file_path)
                    log(f"Successfully replaced: {nu_file_path}")
                except Exception as e:
                    log(f"Replacement failed: {e}")
            else:
                log(f"Module file not found: {nu_file_path}")
    except Exception as e:
        log(f"替换模块文件时出错: {e}")


class InstallWorker(QThread):
    def __init__(self, install_path, requirements_path):
        super().__init__()
        self.install_path = install_path
        self.requirements_path = requirements_path
        self.signals = WorkerSignals()
        self._is_running = True
        self.process = None

    def stop(self):
        self._is_running = False
        if self.process:
            try:
                self.process.terminate()
            except:
                pass
        self.wait()

    def run(self):
        try:
            self.signals.log.emit(f"Creating virtual environment at {self.install_path}...")
            venv_path = os.path.join(self.install_path, "venv")
            if not os.path.exists(venv_path):
                self.signals.log.emit("Creating new virtual environment...")

                # 使用ExVR注册表中的Python解释器
                python_path = self._get_python_from_registry()
                if not python_path:
                    delete_config()
                    raise Exception("Python interpreter not found in ExVR registry")


                self.process = subprocess.Popen(
                    [python_path, "-m", "venv", venv_path],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                self.process.wait()
                if self.process.returncode != 0:
                    raise Exception("Failed to create virtual environment")
            else:
                self.signals.log.emit("Virtual environment already exists.")

            self.signals.progress.emit(20)
            if not self._is_running: return

            pip_path = os.path.join(venv_path, "Scripts", "pip.exe")
            if not os.path.exists(pip_path): raise FileNotFoundError(f"pip not found: {pip_path}")
            if not os.path.exists(self.requirements_path): raise FileNotFoundError(
                f"requirements.txt not found: {self.requirements_path}")

            self.signals.log.emit(f"Installing requirements from {self.requirements_path} using Tsinghua mirror...")
            self.process = subprocess.Popen(
                [pip_path, "install", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "-r", self.requirements_path],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            progress = 0
            while self._is_running:
                line = self.process.stdout.readline()
                if not line and self.process.poll() is not None:
                    break
                if line:
                    self.signals.log.emit(line.strip())
                    # Increment progress based on output
                    if "Collecting" in line:
                        progress = min(progress + 2, 90)
                        self.signals.progress.emit(progress)
                    elif "Installing" in line:
                        progress = min(progress + 1, 95)
                        self.signals.progress.emit(progress)

            if not self._is_running: return

            if self.process.returncode == 0:
                self.signals.log.emit("Requirements installation completed successfully.")
                self.signals.log.emit("Extraction completed, starting module file replacement")
                replace_modules_with_json(self.install_path)
                self.signals.progress.emit(100)
                self.signals.finished.emit()
            else:
                error_msg = f"Requirements installation failed with return code {self.process.returncode}"
                self.signals.log.emit(error_msg)
                self.signals.error.emit(error_msg)


        except Exception as e:
            self.signals.log.emit(f"Installation error: {e}")
            self.signals.error.emit(str(e))

    def _get_python_from_registry(self):
        try:
            python_path = get_python_path()
            if python_path and os.path.exists(python_path):
                return python_path
        except Exception as e:
            self.signals.log.emit(f"Error getting Python path from config: {e}")
        return None


class SilentInstaller:
    def __init__(self, app, args):
        self.args = args
        self.app = app
        self.tmp_dir = create_tmp_folder()
        self.install_path = None
        self.python_installer_path = None
        self.release_zip_path = None
        self.progress_dialog = None
        self.current_worker = None
        self.user_cancelled = False
        self.show_announcement = True
        self.python_path = None

    def run(self):
        log("Starting silent installer...")
        try:
            if self._check_lau_update():
                log("laut update")
            self.install_path = get_install_path()
            self.python_path = get_python_path()
            if self.install_path and self.python_path:
                log(f"Found existing installation at: {self.install_path}")
                log(f"Found existing Python at: {self.python_path}")
                self._check_for_updates()
                return
            else:
                log("No existing installation found in config.")
        except Exception as e:
            log(f"Error reading config: {e}")

        dialog = CustomFileDialog()
        if dialog.exec() == QDialog.Accepted:
            self.install_path = dialog.get_selected_path()
            # 确保路径使用反斜杠
            self.install_path = normalize_path(self.install_path)
            log(f"Selected installation path: {self.install_path}")

            self._check_python()
        else:
            log("Installation cancelled by user.")
            self._quit_installer()

    def _stop_current_worker(self):
        """
        停掉并等待当前线程，防止"QThread: Destroyed while thread is still running"
        """
        if self.current_worker and self.current_worker.isRunning():
            log(f"正在停止线程: {type(self.current_worker).__name__}")
            try:
                # 尝试使用stop方法
                self.current_worker.stop()
            except AttributeError:
                # 如果没有stop方法，使用quit+wait
                log("线程没有stop方法，使用quit+wait")
                self.current_worker.quit()
                # 设置超时，避免无限等待
                if not self.current_worker.wait(3000):  # 等待最多3秒
                    log("线程未能在3秒内停止，强制终止")
                    self.current_worker.terminate()
                    self.current_worker.wait()
        self.current_worker = None

    def _start_worker(self, worker: QThread):
        self._stop_current_worker()
        self.current_worker = worker
        worker.start()

    def _check_python(self):
        log("Checking for Python installation...")
        self._show_progress_dialog("Check", "Check Python")
        worker = PythonCheckWorker()
        worker.signals.log.connect(log)
        worker.signals.result.connect(self._handle_python_check_result)
        worker.signals.error.connect(self._handle_error)
        self._start_worker(worker)

    def _handle_python_check_result(self, result):
        self._close_progress_dialog()
        if result == "installed":
            log("Python 3.11 is installed. Proceeding with download.")
            self._download_release()
        else:
            log("Python 3.11 not found. Automatically downloading and installing...")
            self._download_python()

    def _download_python(self):
        log(f"Downloading Python from {PYTHON_DOWNLOAD_URL}")
        self.python_installer_path = os.path.join(self.tmp_dir, "python_installer.exe")
        self._show_progress_dialog("Download Python", "Download Python 3.11...")
        worker = DownloadWorker(PYTHON_DOWNLOAD_URL, self.python_installer_path)
        worker.signals.log.connect(log)
        worker.signals.progress.connect(self._update_progress)
        worker.signals.error.connect(self._handle_error)
        worker.signals.finished.connect(self._install_python)
        self._start_worker(worker)

    def _install_python(self):
        self._close_progress_dialog()
        log("Installing Python...")

        python_install_dir = os.path.join(self.install_path, "python")
        python_install_dir = normalize_path(python_install_dir)
        os.makedirs(python_install_dir, exist_ok=True)
        log(f"Created Python installation directory: {python_install_dir}")

        try:
            log("Have admin rights, installing directly...")

            installer_path = normalize_path(self.python_installer_path)
            target_dir = quote_path_if_needed(python_install_dir)

            cmd_args = [
                installer_path,
                "/passive",
                "InstallAllUsers=0",
                "PrependPath=0",
                "Include_doc=0",
                "Include_launcher=0",
                "Include_test=0",
                "Include_dev=0",
                "AssociateFiles=0",
                "Shortcuts=0",
                f"TargetDir={target_dir}"
            ]

            log(f"Running Python installer with args: {' '.join(cmd_args)}")

            process = subprocess.Popen(
                cmd_args,
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )

            event_loop = QEventLoop()

            timer = QTimer()
            timer.setInterval(100)

            def check_process():
                if process.poll() is not None:
                    timer.stop()
                    event_loop.quit()

            timer.timeout.connect(check_process)
            timer.start()

            event_loop.exec()

            if process.returncode != 0:
                stderr = process.stderr.read().decode('utf-8', errors='ignore')
                raise Exception(f"Python installation failed with code {process.returncode}: {stderr}")

            python_exe_path = os.path.join(python_install_dir, "python.exe")
            if not os.path.exists(python_exe_path):
                raise Exception(f"Python executable not found at {python_exe_path}")

            self._store_python_path_in_registry(python_exe_path)
            self.python_path = python_exe_path

            log("Python installation completed successfully.")
            self._download_release()

        except Exception as e:
            self._handle_error(f"Failed to install Python: {e}")


    def _check_lau_update(self):
        log("check lau version...")
        try:
            remote_lau_version = server_data.get("lau_version")
            remote_lau_board = server_data.get("lau_board")

            if remote_lau_version is not None and remote_lau_version > LAU_VERSION:
                log(f"Launcher have update ({remote_lau_version})")
                self._show_lau_board(remote_lau_board)
                return True

            log("not lau version")
            return False
        except Exception as e:
            log(f"lau check update error: {e}")
            return False

    def _show_lau_board(self, content):
        log("Show lau board")
        dialog = QDialog()
        dialog.setWindowTitle("Launcher Needs Update")
        dialog.setMinimumSize(600, 400)

        dialog.setStyleSheet("""
            QDialog {
                background-color: #555555;
                color: #ffffff;
                border-radius: 8px;
            }
            QLabel {
                background-color: transparent;
                color: #ffffff;
                font-size: 12pt;
                padding: 10px;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QPushButton {
                min-width: 100px;
                font-size: 10pt;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title_label = QLabel("<h2 style='color:#ffffff;'>Update Board</h2>")
        content_label = QLabel(content or "Not Text")
        content_label.setTextFormat(Qt.RichText)
        content_label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextBrowserInteraction)
        content_label.setWordWrap(True)
        content_label.setOpenExternalLinks(True)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_label)
        scroll_area.setFrameShape(QFrame.NoFrame)

        button_box = QDialogButtonBox()
        update_button = QPushButton("Close")
        update_button.clicked.connect(lambda: sys.exit(0))
        button_box.addButton(update_button, QDialogButtonBox.ActionRole)

        layout.addWidget(title_label)
        layout.addWidget(scroll_area, 1)
        layout.addWidget(button_box)

        dialog.exec()

    def _store_python_path_in_registry(self, python_path):
        try:
            set_python_path(python_path)
            set_install_path(self.install_path)

            log(f"Stored Python path in config: {python_path}")
            log(f"Stored install path in config: {self.install_path}")
        except Exception as e:
            log(f"Error storing paths in config: {e}")
            raise

    def _download_release(self):
        try:
            self._stop_current_worker()
            release_url = None
            github_url = GITHUB_API_URL.format(owner=GITHUB_REPO_OWNER, repo=GITHUB_REPO_NAME)
            try:
                log(f"Attempting to get the latest version from GitHub: {github_url}")
                response = requests.get(github_url, timeout=5)
                if response.status_code == 200:
                    data = response.json()
                    if "zipball" in data.get('zipball_url', ''):
                        release_url = "https://gh-proxy.com/" + data.get('zipball_url')
                        log(f"Received download URL from GitHub: {release_url}")
            except Exception as e:
                log(f"Failed to get version info from GitHub: {str(e)}")

            if not release_url:
                release_url = github_url
                log(f"Using default download URL: {release_url}")

            self.release_zip_path = os.path.join(self.tmp_dir, "release.zip")
            self._show_progress_dialog("Download Application", "Downloading the latest version...")
            worker = DownloadWorker(release_url, self.release_zip_path)
            worker.signals.log.connect(log)
            worker.signals.progress.connect(self._update_progress)
            worker.signals.error.connect(self._handle_error)
            worker.signals.finished.connect(self._extract_release)
            self._start_worker(worker)
        except Exception as e:
            self._handle_error(f"Failed to get release info: {e}")

    def _extract_release(self):
        log("Extracting release...")
        extract_path = os.path.join(self.tmp_dir, "extract")
        final_path = os.path.join(self.install_path, "exvr")
        os.makedirs(extract_path, exist_ok=True)
        os.makedirs(final_path, exist_ok=True)

        self._show_progress_dialog("Extract files", "Unzipping application files...")
        worker = ExtractWorker(self.release_zip_path, extract_path, final_path)
        worker.signals.log.connect(log)
        worker.signals.progress.connect(self._update_progress)
        worker.signals.error.connect(self._handle_error)
        worker.signals.finished.connect(self._install_requirements)
        self._start_worker(worker)

    def _install_requirements(self):
        log("Installing requirements...")
        requirements_path = os.path.join(self.install_path, "exvr", "requirements.txt")
        if not os.path.exists(requirements_path):
            log("Requirements file not found. Skipping installation.")
            self._register_application()
            return

        self._show_progress_dialog("Install requirements",
                                   "Installing Python requirements... (Initial installation may be time-consuming).")
        worker = InstallWorker(os.path.join(self.install_path, "exvr"), requirements_path)
        worker.signals.log.connect(log)
        worker.signals.progress.connect(self._update_progress)
        worker.signals.error.connect(self._handle_error)
        worker.signals.finished.connect(self._register_application)
        self._start_worker(worker)

    def _register_application(self):
        self._close_progress_dialog()
        log("Registering application...")
        try:
            # 存储安装路径到JSON配置
            set_install_path(self.install_path)

            # 确保Python路径已存储
            if self.python_path and os.path.exists(self.python_path):
                set_python_path(self.python_path)

            log("Application registered successfully.")
            self._run_application()
        except Exception as e:
            self._handle_error(f"Failed to register application: {e}")

    def _check_for_updates(self):
        log("Checking for updates...")
        try:
            remote_version = server_data.get("version")

            if not remote_version:
                log("Failed to get remote version. Running current version.")
                self._run_application()
                return

            config_path = os.path.join(self.install_path, "exvr", "settings", "config.json")
            local_version = None
            if os.path.exists(config_path):
                try:
                    with open(config_path, "r") as f:
                        local_version = json.load(f).get("Version")
                    log(f"Local version: {local_version}")
                except Exception as e:
                    log(f"Error reading local config: {e}")

            if not local_version or local_version != remote_version:
                log(f"Update available: Local={local_version}, Remote={remote_version}")
                reply = ask_question("Have Update",
                                     f"Have new version ({remote_version}). Do you want to update?")
                if reply == QMessageBox.Yes:
                    log("User chose to update.")
                    self._update_application()
                else:
                    self.show_announcement = False
                    log("User declined update. Running current version.")
                    self._run_application()
            else:
                self.show_announcement = False
                log("Application is up to date.")
                self._run_application()
        except Exception as e:
            log(f"Update check process failed: {e}. Running application.")
            self._run_application()

    def _update_application(self):
        log("Starting application update...")
        self._download_release()

    def _run_application(self):
        app_log_dir = get_resource_path("logs")
        os.makedirs(app_log_dir, exist_ok=True)

        log("Preparing to run application...")
        clean_tmp_folder(self.tmp_dir)

        if self.show_announcement:
            self._show_announcement_box()

        try:
            exvr_path = os.path.join(self.install_path, "exvr")
            venv_path = os.path.join(exvr_path, "venv", "Scripts")
            main_script = os.path.join(exvr_path, "main.py")

            venv_python = os.path.join(venv_path, "python.exe")

            if not os.path.exists(venv_python):
                raise FileNotFoundError(f"Virtual environment Python not found at {venv_python}")
            if not os.path.exists(main_script):
                raise FileNotFoundError("Main application script not found.")

            log(f"Running command: {venv_python} {main_script}")
            os.chdir(exvr_path)

            flags = subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_CONSOLE

            log("Using CREATE_NO_WINDOW flag to hide initial console only")

            # 检查是否有-log参数
            cmd = [venv_python, main_script, "--log-dir", app_log_dir]
            if self.args.log:
                cmd.append("-log")

            process = subprocess.Popen(
                cmd,
                creationflags=flags,
                close_fds=True,
                shell=False,
                stdin=None,
                stdout=None,
                stderr=None
            )

            time.sleep(1)

            log("Application launched. Exiting installer.")
            self._quit_installer()

        except Exception as e:
            self._handle_error(f"Failed to run application: {e}")

    def _show_announcement_box(self):
        log("Fetching announcement board...")
        board_data = server_data.get("board")
        if not board_data:
            log("No valid announcement board found")
            return

        title = board_data.get("title", "Announcement")
        text = board_data.get("text", "")

        log(f"Showing announcement: {title}")

        dialog = QDialog()
        dialog.setWindowTitle(title)
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        dialog.setMinimumSize(500, 300)

        dialog.setStyleSheet("""
            QDialog {
                background-color: #555555;
                color: #ffffff;
                border-radius: 8px;
            }
            QLabel {
                background-color: transparent;
                color: #ffffff;
                font-size: 12pt;
                padding: 10px;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QPushButton {
                min-width: 100px;
                font-size: 10pt;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        content_label = QLabel(text)
        content_label.setTextFormat(Qt.RichText)
        content_label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextBrowserInteraction)
        content_label.setWordWrap(True)
        content_label.setOpenExternalLinks(True)

        content_label.setStyleSheet("font-size: 12pt;")

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_label)
        scroll_area.setFrameShape(QFrame.NoFrame)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)

        layout.addWidget(QLabel(f"<h2 style='color:#ffffff;'>{title}</h2>"))
        layout.addWidget(scroll_area, 1)
        layout.addWidget(button_box)

        dialog.adjustSize()

        max_width = QApplication.primaryScreen().availableGeometry().width() * 0.8
        max_height = QApplication.primaryScreen().availableGeometry().height() * 0.8
        dialog.resize(min(dialog.width(), max_width), min(dialog.height(), max_height))

        dialog.exec()

    def _show_progress_dialog(self, title, label):
        self._close_progress_dialog()
        self.progress_dialog = QProgressDialog(label, "Cancel", 0, 100, None)
        self.progress_dialog.setWindowTitle(title)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)

        self.progress_dialog.canceled.connect(self._handle_cancel_click)

        self.progress_dialog.setValue(0)
        self.progress_dialog.show()
        self.app.processEvents()

    def _handle_cancel_click(self):
        log("Cancel button clicked by user.")
        self.user_cancelled = True
        if self.current_worker:
            self.current_worker.stop()
        self._close_progress_dialog()
        show_info_message("Cancelled", "The operation has been canceled by the user.")
        self._quit_installer()

    def _update_progress(self, value):
        if self.progress_dialog and not self.user_cancelled:
            self.progress_dialog.setValue(value)
            self.app.processEvents()

    def _close_progress_dialog(self):
        if self.progress_dialog:
            try:
                self.progress_dialog.canceled.disconnect()
            except:
                pass

            self.progress_dialog.close()
            self.progress_dialog = None

    def _handle_error(self, message):
        log(f"Handling error: {message}")
        self._close_progress_dialog()
        if self.current_worker:
            self.current_worker.stop()
        show_error_message("Installation Error", message)
        self._quit_installer()

    def _quit_installer(self):
        log("Quitting installer application.")
        self._stop_current_worker()  # <== 新增
        clean_tmp_folder(self.tmp_dir)
        self.app.quit()

def get_server_data():
    global server_data
    for url in UPDATE_CHECK_URLS:
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                server_data = response.json()
        except Exception as e:
            log(f"get server data error: {e}")

def main():
    get_server_data()
    args = parse_arguments()

    setup_logging(args)

    log("Main function started.")
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except AttributeError:
        pass

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    app.setStyleSheet(modern_qss)

    installer = SilentInstaller(app, args)
    QTimer.singleShot(100, installer.run)

    sys.exit(app.exec())


if __name__ == "__main__":
    if not pyuac.isUserAdmin():
        print("This program requires administrative privileges to run.")


        def show_admin_warning():
            app = QApplication(sys.argv)

            msg_box = QMessageBox()
            msg_box.setWindowTitle("权限提示")
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setText("需要管理员权限")
            msg_box.setInformativeText(
                "请右键点击程序图标\n"
                "选择'以管理员身份运行'"
            )

            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.button(QMessageBox.Ok).setText("退出程序")
            msg_box.setDefaultButton(QMessageBox.Ok)

            msg_box.setMinimumSize(400, 200)

            result = msg_box.exec()

            sys.exit(0)


        show_admin_warning()
    else:
        main()