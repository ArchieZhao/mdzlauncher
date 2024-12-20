import sys
import os
import zipfile
import tempfile
import subprocess
import shutil
import uuid
import json
import logging
import time
from PyQt5 import QtWidgets, QtGui, QtCore
from win10toast import ToastNotifier
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# 获取当前应用程序所在目录（绝对路径）
app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# 配置文件和日志文件放在应用程序目录下
CONFIG_FILE = os.path.join(app_dir, "mdz_launcher_config.json")
LOG_FILE = os.path.join(app_dir, "mdz_launcher.log")

# 默认配置：优先使用相对路径
default_config = {
    "7zip_path": os.path.join("7-Zip", "7z.exe"),   
    "typora_path": os.path.join("typora", "Typora.exe")  
}

# 日志设置
logging.basicConfig(filename=LOG_FILE,
                    level=logging.ERROR,
                    format='%(asctime)s:%(levelname)s:%(message)s')

def load_config():
    config = default_config.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
                config.update(user_config)
        except Exception as e:
            logging.error(f"加载配置失败: {e}")
    return config

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logging.error(f"保存配置失败: {e}")

def resolve_path(path):
    # 将配置中的相对路径转换为绝对路径，如果已是绝对路径则直接返回
    if not os.path.isabs(path):
        # 检查路径是否在应用程序目录下
        potential_path = os.path.join(app_dir, path)
        if os.path.exists(potential_path):
            return potential_path
        else:
            # 如果相对路径无效，返回原路径
            return path
    return path

class DocumentSaveHandler(FileSystemEventHandler):
    """
    文件保存事件处理器，每当document.md被修改时触发打包。
    """
    def __init__(self, launcher):
        super().__init__()
        self.launcher = launcher
        self.last_modified = 0
        self.debounce_time = 1.0  # 防抖时间1秒

    def on_modified(self, event):
        if event.is_directory:
            return
        if os.path.basename(event.src_path) == "document.md":
            now = time.time()
            # 防抖处理，如果在1秒内多次修改，仅最后一次触发
            if now - self.last_modified > self.debounce_time:
                # 延迟执行打包
                QtCore.QTimer.singleShot(int(self.debounce_time*1000), self.launcher.pack_on_save)
            self.last_modified = now

class MDZLauncher(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MDZ Launcher")
        self.setGeometry(100, 100, 800, 600)  # 增大窗口尺寸以适应日志显示
        self.temp_dir = None
        self.mdz_path = None
        self.typora_process = None
        self.config = load_config()
        self.toaster = ToastNotifier()
        self.initUI()

        # 监视器和事件处理器
        self.observer = None
        self.event_handler = None

    def initUI(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('文件')

        newAction = QtWidgets.QAction('新建 .mdz 文件', self)
        newAction.triggered.connect(self.new_mdz)
        fileMenu.addAction(newAction)

        openAction = QtWidgets.QAction('打开 .mdz 文件', self)
        openAction.setShortcut('Ctrl+O')
        openAction.triggered.connect(self.open_mdz)
        fileMenu.addAction(openAction)

        settingsMenu = menubar.addMenu('设置')
        configAction = QtWidgets.QAction('配置路径', self)
        configAction.triggered.connect(self.open_settings)
        settingsMenu.addAction(configAction)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        layout = QtWidgets.QVBoxLayout()

        # 替换 QLabel 为 QTextEdit
        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("日志信息将在此显示...")
        layout.addWidget(self.log_view)

        # 添加清除日志和查看日志文件按钮
        button_layout = QtWidgets.QHBoxLayout()
        self.clear_log_button = QtWidgets.QPushButton("清除日志")
        self.clear_log_button.clicked.connect(self.clear_log)
        self.view_log_button = QtWidgets.QPushButton("查看日志文件")
        self.view_log_button.clicked.connect(self.view_log_file)
        button_layout.addWidget(self.clear_log_button)
        button_layout.addWidget(self.view_log_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        central_widget.setLayout(layout)

    def append_log(self, message, level="INFO"):
        """
        将日志消息追加到日志视图中，支持不同级别的日志。
        level: "INFO", "WARNING", "ERROR"
        """
        timestamp = QtCore.QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        if level == "INFO":
            color = "black"
        elif level == "WARNING":
            color = "orange"
        elif level == "ERROR":
            color = "red"
        else:
            color = "black"
        self.log_view.append(f'<span style="color:{color};">[{timestamp}] {message}</span>')
        # 自动滚动到文档末尾
        self.log_view.moveCursor(QtGui.QTextCursor.End)

    def open_settings(self):
        dialog = SettingsDialog(self.config, self)
        if dialog.exec_():
            new_config = dialog.get_config()
            # 更新配置并保存
            self.config.update(new_config)
            save_config(self.config)
            self.append_log("配置已更新", level="INFO")

    def new_mdz(self):
        """
        实现新建 .mdz 文件功能：
        1. 弹出保存对话框，让用户选择 .mdz 文件的保存位置。
        2. 在临时目录创建一个空的 document.md 文件和 document.assets 文件夹。
        3. 不立即打包，直接调用 launch_typora 供用户编辑。
        4. 用户编辑完成后关闭 Typora，程序自动打包为 .mdz 文件。
        """
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "新建 .mdz 文件", "", "MDZ Files (*.mdz);;All Files (*)", options=options)
        if file_path:
            if not file_path.lower().endswith('.mdz'):
                file_path += '.mdz'
            self.mdz_path = file_path
            self.temp_dir = os.path.join(tempfile.gettempdir(), f"mdz_temp_{uuid.uuid4()}")
            os.makedirs(self.temp_dir, exist_ok=True)
            doc_path = os.path.join(self.temp_dir, "document.md")
            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write("# 新文档\n\n这里是新建的 MDZ 文档，您可以开始编辑...")
            assets_dir = os.path.join(self.temp_dir, "document.assets")
            os.makedirs(assets_dir, exist_ok=True)
            self.append_log(f"新建 .mdz 文件: {self.mdz_path}", level="INFO")
            self.launch_typora()

    def open_mdz(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择 .mdz 文件", "", "MDZ Files (*.mdz);;All Files (*)", options=options)
        if file_path:
            self.mdz_path = file_path
            self.append_log(f"打开 .mdz 文件: {self.mdz_path}", level="INFO")
            self.unpack_mdz()
            self.launch_typora()

    def unpack_mdz(self):
        seven_zip = resolve_path(self.config["7zip_path"])
        if not os.path.exists(seven_zip):
            self.append_log(f"错误: 7-Zip 未找到: {seven_zip}", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"7-Zip 未找到: {seven_zip}")
            return

        self.temp_dir = os.path.join(tempfile.gettempdir(), f"mdz_temp_{uuid.uuid4()}")
        os.makedirs(self.temp_dir, exist_ok=True)

        try:
            subprocess.run([seven_zip, "x", "-y", f"-o{self.temp_dir}", self.mdz_path],
                           check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.append_log(f"已解压到临时目录: {self.temp_dir}", level="INFO")
        except subprocess.CalledProcessError as e:
            logging.error(f"解压失败: {e}")
            self.append_log(f"错误: 解压 .mdz 文件失败: {e}", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "解压失败", f"解压 .mdz 文件失败: {e}")
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = None

    def launch_typora(self):
        typora_path = resolve_path(self.config["typora_path"])
        if not os.path.exists(typora_path):
            self.append_log(f"错误: Typora 未找到: {typora_path}", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"Typora 未找到: {typora_path}")
            return

        document_path = os.path.join(self.temp_dir, "document.md")
        if not os.path.exists(document_path):
            self.append_log("错误: document.md 未找到在临时目录中", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", "document.md 未找到在临时目录中")
            return

        # 启动文件监控
        self.start_file_monitor()

        try:
            self.typora_process = subprocess.Popen([typora_path, document_path])
            self.append_log("Typora 已启动，保存时将自动打包...", level="INFO")
            self.monitor_typora()
        except Exception as e:
            logging.error(f"启动 Typora 失败: {e}")
            self.append_log(f"错误: 无法启动 Typora: {e}", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"无法启动 Typora: {e}")

    def start_file_monitor(self):
        # 启动对 temp_dir 的document.md文件监控
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
        self.event_handler = DocumentSaveHandler(self)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, self.temp_dir, recursive=False)
        self.observer.start()
        self.append_log("文件监控已启动。", level="INFO")

    def monitor_typora(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.check_typora)
        self.timer.start(1000)

    def check_typora(self):
        if self.typora_process and self.typora_process.poll() is not None:
            self.timer.stop()
            # Typora已关闭，进行最终打包并清理
            self.pack_mdz(final=True)
            # 停止监控
            if self.observer:
                self.observer.stop()
                self.observer.join()
                self.observer = None
            self.append_log("Typora 已关闭，最终打包完成。", level="INFO")

    def pack_on_save(self):
        # 在保存事件中调用的打包，这不是最终关闭打包，因此final=False
        self.pack_mdz(final=False)

    def pack_mdz(self, final=True):
        if not self.temp_dir or not self.mdz_path:
            self.append_log("错误: 临时目录或 .mdz 文件路径未设置。", level="ERROR")
            return
        seven_zip = resolve_path(self.config["7zip_path"])
        if not os.path.exists(seven_zip):
            self.append_log(f"错误: 7-Zip 未找到: {seven_zip}", level="ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"7-Zip 未找到: {seven_zip}")
            return

        try:
            # 这里先打包到一个临时文件，然后替换原mdz，避免中途损坏
            temp_mdz = self.mdz_path + ".temp"
            if os.path.exists(temp_mdz):
                os.remove(temp_mdz)

            subprocess.run([seven_zip, "a", "-tzip", "-mx9", temp_mdz, f"{self.temp_dir}\\*"],
                           check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            # 替换原文件
            shutil.move(temp_mdz, self.mdz_path)

            if final:
                self.append_log("已重新打包 .mdz 文件（最终关闭）。", level="INFO")
                self.toaster.show_toast("MDZ Launcher", "成功更新 .mdz 文件", duration=5, threaded=True)
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.temp_dir = None
            else:
                # 中间保存打包，不清理temp_dir，不给出关闭提示
                self.append_log("已自动打包 .mdz 文件（保存时）。", level="INFO")

        except subprocess.CalledProcessError as e:
            logging.error(f"打包失败: {e}")
            self.append_log(f"错误: 重新打包 .mdz 文件失败: {e}", level="ERROR")
            if final and self.temp_dir:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.temp_dir = None

    def clear_log(self):
        """
        清除日志视图中的所有内容。
        """
        self.log_view.clear()
        self.append_log("日志已清除。", level="INFO")

    def view_log_file(self):
        """
        打开日志文件供用户查看。
        """
        if os.path.exists(LOG_FILE):
            try:
                subprocess.Popen(["notepad.exe", LOG_FILE])
                self.append_log("已打开日志文件。", level="INFO")
            except Exception as e:
                self.append_log(f"错误: 无法打开日志文件: {e}", level="ERROR")
        else:
            self.append_log("日志文件不存在。", level="WARNING")

class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置路径")
        self.config = config.copy()
        self.initUI()

    def initUI(self):
        layout = QtWidgets.QVBoxLayout()

        self.zip_label = QtWidgets.QLabel("7-Zip 路径:")
        self.zip_path = QtWidgets.QLineEdit(self.config.get("7zip_path", ""))
        self.zip_browse = QtWidgets.QPushButton("浏览")
        self.zip_browse.clicked.connect(self.browse_7zip)
        zip_layout = QtWidgets.QHBoxLayout()
        zip_layout.addWidget(self.zip_path)
        zip_layout.addWidget(self.zip_browse)

        self.typora_label = QtWidgets.QLabel("Typora 路径:")
        self.typora_path = QtWidgets.QLineEdit(self.config.get("typora_path", ""))
        self.typora_browse = QtWidgets.QPushButton("浏览")
        self.typora_browse.clicked.connect(self.browse_typora)
        typora_layout = QtWidgets.QHBoxLayout()
        typora_layout.addWidget(self.typora_path)
        typora_layout.addWidget(self.typora_browse)

        self.save_button = QtWidgets.QPushButton("保存")
        self.save_button.clicked.connect(self.accept)
        self.cancel_button = QtWidgets.QPushButton("取消")
        self.cancel_button.clicked.connect(self.reject)
        buttons_layout = QtWidgets.QHBoxLayout()
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.save_button)
        buttons_layout.addWidget(self.cancel_button)

        layout.addWidget(self.zip_label)
        layout.addLayout(zip_layout)
        layout.addWidget(self.typora_label)
        layout.addLayout(typora_layout)
        layout.addStretch()
        layout.addLayout(buttons_layout)
        self.setLayout(layout)

    def browse_7zip(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择 7z.exe", "", "Executable Files (*.exe);;All Files (*)")
        if file_path:
            try:
                rel_path = os.path.relpath(file_path, app_dir)
                if not rel_path.startswith(".."):
                    # 说明在app_dir范围内，可以使用相对路径
                    self.zip_path.setText(rel_path)
                else:
                    self.zip_path.setText(file_path)
            except ValueError as ve:
                logging.error(f"浏览7-Zip路径时出错: {ve}")
                self.zip_path.setText(file_path)

    def browse_typora(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择 Typora.exe", "", "Executable Files (*.exe);;All Files (*)")
        if file_path:
            try:
                rel_path = os.path.relpath(file_path, app_dir)
                if not rel_path.startswith(".."):
                    self.typora_path.setText(rel_path)
                else:
                    self.typora_path.setText(file_path)
            except ValueError as ve:
                logging.error(f"浏览Typora路径时出错: {ve}")
                self.typora_path.setText(file_path)

    def get_config(self):
        return {
            "7zip_path": self.zip_path.text(),
            "typora_path": self.typora_path.text()
        }

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MDZLauncher()

    # 检查是否有命令行参数传入（即 .mdz 文件路径）
    if len(sys.argv) > 1:
        mdz_file = sys.argv[1]
        if os.path.isfile(mdz_file) and mdz_file.lower().endswith('.mdz'):
            window.mdz_path = mdz_file
            window.append_log(f"通过命令行打开 .mdz 文件: {window.mdz_path}", level="INFO")
            window.unpack_mdz()
            window.launch_typora()
        else:
            QtWidgets.QMessageBox.warning(window, "无效文件", "传入的文件不是有效的 .mdz 文件。")

    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
