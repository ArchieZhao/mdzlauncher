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
import traceback

from PyQt5 import QtWidgets, QtGui, QtCore
from win10toast import ToastNotifier
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileMovedEvent

#######################################
# 全局配置
#######################################
app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG_FILE = os.path.join(app_dir, "mdz_launcher_config.json")
LOG_FILE = os.path.join(app_dir, "mdz_launcher.log")

default_config = {
    "7zip_path": os.path.join("7-Zip", "7z.exe"),
    "typora_path": os.path.join("typora", "Typora.exe")
}

# 日志文件设置
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.ERROR,
    format="%(asctime)s:%(levelname)s:%(message)s"
)

#######################################
# 工具函数
#######################################
def load_config():
    config = default_config.copy()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                user_config = json.load(f)
                config.update(user_config)
        except Exception as e:
            logging.error(f"加载配置失败: {e}")
    return config

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logging.error(f"保存配置失败: {e}")

def resolve_path(path):
    if not os.path.isabs(path):
        potential_path = os.path.join(app_dir, path)
        if os.path.exists(potential_path):
            return potential_path
        else:
            return path
    return path

def safe_move(src, dst, attempts=5, wait=0.5, logger=None):
    """
    尝试多次 move src -> dst，以处理云同步软件锁定等问题。
    :param src: 源文件路径
    :param dst: 目标文件路径
    :param attempts: 最大重试次数
    :param wait: 每次失败后等待的秒数
    :param logger: 用于记录日志的函数
    :return: True 表示成功移动，False 表示多次重试后仍失败
    """
    for i in range(attempts):
        try:
            shutil.move(src, dst)
            if logger:
                logger(f"[DEBUG] safe_move succeeded on attempt #{i+1}", "INFO")
            return True
        except (PermissionError, OSError) as e:
            if logger:
                logger(f"[WARN] safe_move attempt #{i+1} failed: {e}", "WARNING")
            time.sleep(wait)
    if logger:
        logger("[ERROR] safe_move failed after multiple attempts", "ERROR")
    return False

#######################################
# 文件保存事件处理器
#######################################
class DocumentSaveHandler(FileSystemEventHandler):
    def __init__(self, launcher):
        super().__init__()
        self.launcher = launcher

    def maybe_trigger_pack(self, file_path):
        base_name = os.path.basename(file_path).lower()
        self.launcher.append_log(
            f"[DEBUG] maybe_trigger_pack => base_name={base_name}", "INFO"
        )
        if "document.md" in base_name:
            self.launcher.saveCount += 1
            self.launcher.append_log(
                f"[SAVE EVENT] Detected save #{self.launcher.saveCount} on {base_name} => docSaveDirty = True",
                "INFO"
            )
            self.launcher.docSaveDirty = True
            # 如果当前正在打包 => 只记录 docSaveDirty
            if self.launcher.packInProgress:
                self.launcher.append_log("[DEBUG] packInProgress=True => skip reset_timer", "INFO")
                return
            # 否则正常触发防抖逻辑
            self.launcher.reset_doc_save_timer()

    def on_modified(self, event):
        if event.is_directory:
            return
        self.launcher.append_log(f"[DEBUG] on_modified => {event.src_path}", "INFO")
        self.maybe_trigger_pack(event.src_path)

    def on_created(self, event):
        if event.is_directory:
            return
        self.launcher.append_log(f"[DEBUG] on_created => {event.src_path}", "INFO")
        self.maybe_trigger_pack(event.src_path)

    def on_moved(self, event):
        if isinstance(event, FileMovedEvent):
            self.launcher.append_log(
                f"[DEBUG] on_moved => {event.src_path} -> {event.dest_path}",
                "INFO"
            )
            self.maybe_trigger_pack(event.dest_path)

#######################################
# 主窗体
#######################################
class MDZLauncher(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MDZ Launcher")
        self.setGeometry(100, 100, 800, 600)

        self.config = load_config()
        self.temp_dir = None
        self.mdz_path = None
        self.typora_process = None
        self.toaster = ToastNotifier()

        # 保存事件 & 打包相关
        self.docSaveDirty = False
        self.packInProgress = False  # 是否正在打包
        self.saveCount = 0

        self.resetCount = 0
        self.max_resets = 5
        self.lastResetStart = 0.0
        self.max_delay_sec = 10

        self.initUI()

        # watchdog
        self.observer = None
        self.event_handler = None

        # 防抖定时器
        self.docSaveTimer = QtCore.QTimer()
        self.docSaveTimer.setSingleShot(True)
        self.docSaveTimer.setInterval(2000)
        self.docSaveTimer.timeout.connect(self.onDocSaveTimerTimeout)

    def initUI(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu("文件")

        newAction = QtWidgets.QAction("新建 .mdz 文件", self)
        newAction.triggered.connect(self.new_mdz)
        fileMenu.addAction(newAction)

        openAction = QtWidgets.QAction("打开 .mdz 文件", self)
        openAction.setShortcut("Ctrl+O")
        openAction.triggered.connect(self.open_mdz)
        fileMenu.addAction(openAction)

        settingsMenu = menubar.addMenu("设置")
        configAction = QtWidgets.QAction("配置路径", self)
        configAction.triggered.connect(self.open_settings)
        settingsMenu.addAction(configAction)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        layout = QtWidgets.QVBoxLayout()

        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("日志信息将在此显示...")
        layout.addWidget(self.log_view)

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
        timestamp = QtCore.QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        color_map = {
            "INFO": "black",
            "WARNING": "orange",
            "ERROR": "red"
        }
        color = color_map.get(level, "black")
        self.log_view.append(f'<span style="color:{color};">[{timestamp}] {message}</span>')
        self.log_view.moveCursor(QtGui.QTextCursor.End)

    ##################################
    # 防抖定时器：重置逻辑
    ##################################
    def reset_doc_save_timer(self):
        self.resetCount += 1
        if self.resetCount == 1:
            self.lastResetStart = time.time()

        forced = False
        if self.resetCount >= self.max_resets:
            forced = True
            self.append_log(
                f"[FORCE] Reached max_resets={self.max_resets}. Forcing immediate pack_on_save",
                "INFO"
            )
        elif time.time() - self.lastResetStart > self.max_delay_sec:
            forced = True
            self.append_log(
                f"[FORCE] Exceeded max_delay_sec={self.max_delay_sec}s => immediate pack_on_save",
                "INFO"
            )

        if forced:
            self.docSaveDirty = False
            self.resetCount = 0
            self.lastResetStart = 0.0
            self.pack_on_save()
            return

        # 使用定时器
        if self.docSaveTimer.isActive():
            self.append_log("[DEBUG] docSaveTimer stopped for reset", "INFO")
            self.docSaveTimer.stop()
        self.docSaveTimer.start()

    def onDocSaveTimerTimeout(self):
        self.append_log("[DEBUG] docSaveTimer timeout triggered", "INFO")
        if self.docSaveDirty:
            self.docSaveDirty = False
            self.resetCount = 0
            self.lastResetStart = 0.0
            self.append_log("[DEBUG] docSaveDirty is True => calling pack_on_save", "INFO")
            self.pack_on_save()
        else:
            self.append_log("[DEBUG] docSaveDirty is False => do nothing", "INFO")

    ##################################
    # 主逻辑
    ##################################
    def open_settings(self):
        dialog = SettingsDialog(self.config, self)
        if dialog.exec_():
            new_config = dialog.get_config()
            self.config.update(new_config)
            save_config(self.config)
            self.append_log("配置已更新", "INFO")

    def new_mdz(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "新建 .mdz 文件", "",
            "MDZ Files (*.mdz);;All Files (*)",
            options=options
        )
        if file_path:
            if not file_path.lower().endswith(".mdz"):
                file_path += ".mdz"
            self.mdz_path = file_path
            self.temp_dir = os.path.join(tempfile.gettempdir(), f"mdz_temp_{uuid.uuid4()}")
            os.makedirs(self.temp_dir, exist_ok=True)

            doc_path = os.path.join(self.temp_dir, "document.md")
            with open(doc_path, "w", encoding="utf-8") as f:
                f.write("# 新文档\n\n这里是新建的 MDZ 文档，您可以开始编辑...")

            assets_dir = os.path.join(self.temp_dir, "document.assets")
            os.makedirs(assets_dir, exist_ok=True)

            self.append_log(f"新建 .mdz 文件: {self.mdz_path}", "INFO")
            self.launch_typora()

    def open_mdz(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择 .mdz 文件", "",
            "MDZ Files (*.mdz);;All Files (*)",
            options=options
        )
        if file_path:
            self.mdz_path = file_path
            self.append_log(f"打开 .mdz 文件: {self.mdz_path}", "INFO")
            self.unpack_mdz()
            self.launch_typora()

    def unpack_mdz(self):
        seven_zip = resolve_path(self.config["7zip_path"])
        if not os.path.exists(seven_zip):
            self.append_log(f"错误: 7-Zip 未找到: {seven_zip}", "ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"7-Zip 未找到: {seven_zip}")
            return

        self.temp_dir = os.path.join(tempfile.gettempdir(), f"mdz_temp_{uuid.uuid4()}")
        os.makedirs(self.temp_dir, exist_ok=True)

        try:
            subprocess.run(
                [seven_zip, "x", "-y", f"-o{self.temp_dir}", self.mdz_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            self.append_log(f"已解压到临时目录: {self.temp_dir}", "INFO")
        except subprocess.CalledProcessError as e:
            logging.error(f"解压失败: {e}")
            self.append_log(f"错误: 解压 .mdz 文件失败: {e}", "ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"解压 .mdz 文件失败: {e}")
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = None

    def launch_typora(self):
        typora_path = resolve_path(self.config["typora_path"])
        if not os.path.exists(typora_path):
            self.append_log(f"错误: Typora 未找到: {typora_path}", "ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"Typora 未找到: {typora_path}")
            return

        document_path = os.path.join(self.temp_dir, "document.md")
        if not os.path.exists(document_path):
            self.append_log("错误: document.md 未找到在临时目录中", "ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", "document.md 未找到在临时目录中")
            return

        self.start_file_monitor()

        try:
            self.typora_process = subprocess.Popen([typora_path, document_path])
            self.append_log("Typora 已启动，保存时将自动打包...", "INFO")
            self.monitor_typora()
        except Exception as e:
            logging.error(f"启动 Typora 失败: {e}")
            self.append_log(f"错误: 无法启动 Typora: {e}", "ERROR")
            QtWidgets.QMessageBox.critical(self, "错误", f"无法启动 Typora: {e}")

    def start_file_monitor(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None

        self.event_handler = DocumentSaveHandler(self)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, self.temp_dir, recursive=True)
        self.observer.start()
        self.append_log("文件监控已启动。", "INFO")

    def monitor_typora(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.check_typora)
        self.timer.start(1000)

    def check_typora(self):
        if self.typora_process and self.typora_process.poll() is not None:
            self.timer.stop()

            # 若 docSaveDirty=True => 先来一次 "保存时打包"
            if self.docSaveDirty:
                self.docSaveDirty = False
                self.resetCount = 0
                self.lastResetStart = 0.0
                self.append_log("[DEBUG] docSaveDirty was True at close => pack_on_save", "INFO")
                self.pack_on_save()

            # 最终关闭打包
            self.pack_mdz(final=True)
            if self.observer:
                self.observer.stop()
                self.observer.join()
                self.observer = None
            self.append_log("Typora 已关闭，最终打包完成。", "INFO")

    ##################################
    # 打包相关
    ##################################
    def pack_on_save(self):
        self.append_log("[PACK] pack_on_save triggered => calling pack_mdz(final=False)", "INFO")

        # 若已在打包 => 不再重复进入
        if self.packInProgress:
            self.docSaveDirty = True  # 记录还有一次修改
            self.append_log("[WARN] pack_on_save called but packInProgress=True => skipping", "WARNING")
            return

        self.pack_mdz(final=False)

    def pack_mdz(self, final=True):
        self.packInProgress = True
        try:
            if not self.temp_dir or not self.mdz_path:
                self.append_log("错误: 临时目录或 .mdz 文件路径未设置。", "ERROR")
                return

            seven_zip = resolve_path(self.config["7zip_path"])
            if not os.path.exists(seven_zip):
                self.append_log(f"错误: 7-Zip 未找到: {seven_zip}", "ERROR")
                QtWidgets.QMessageBox.critical(self, "错误", f"7-Zip 未找到: {seven_zip}")
                return

            temp_mdz = self.mdz_path + ".temp"
            if os.path.exists(temp_mdz):
                os.remove(temp_mdz)

            # 调用 7-zip 打包
            subprocess.run(
                [seven_zip, "a", "-tzip", "-mx9", temp_mdz, f"{self.temp_dir}\\*"],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )

            # 多次重试移动 temp_mdz -> self.mdz_path
            success = safe_move(
                src=temp_mdz,
                dst=self.mdz_path,
                attempts=5,
                wait=0.5,
                logger=self.append_log
            )
            if not success:
                self.append_log("[ERROR] Failed to move .temp => .mdz after multiple attempts", "ERROR")
                return

            if final:
                self.append_log("已重新打包 .mdz 文件（最终关闭）。", "INFO")
                self.toaster.show_toast("MDZ Launcher", "成功更新 .mdz 文件", duration=5, threaded=True)
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.temp_dir = None
            else:
                self.append_log("已自动打包 .mdz 文件（保存时）。", "INFO")

        except subprocess.CalledProcessError as e:
            logging.error(f"打包失败: {e}")
            self.append_log(f"错误: 重新打包 .mdz 文件失败: {e}", "ERROR")
            if final and self.temp_dir:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.temp_dir = None

        finally:
            self.packInProgress = False
            # 若在打包过程中又出现了新的 docSaveDirty, 并且本次不是 final
            if self.docSaveDirty and not final:
                self.append_log("[DEBUG] Another docSaveDirty arrived while packing => reset_doc_save_timer", "INFO")
                self.reset_doc_save_timer()

    ##################################
    # UI操作
    ##################################
    def clear_log(self):
        self.log_view.clear()
        self.append_log("日志已清除。", "INFO")

    def view_log_file(self):
        if os.path.exists(LOG_FILE):
            try:
                subprocess.Popen(["notepad.exe", LOG_FILE])
                self.append_log("已打开日志文件。", "INFO")
            except Exception as e:
                self.append_log(f"错误: 无法打开日志文件: {e}", "ERROR")
        else:
            self.append_log("日志文件不存在。", "WARNING")

#######################################
# 配置对话框
#######################################
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
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择 7z.exe", "",
            "Executable Files (*.exe);;All Files (*)"
        )
        if file_path:
            try:
                rel_path = os.path.relpath(file_path, app_dir)
                if not rel_path.startswith(".."):
                    self.zip_path.setText(rel_path)
                else:
                    self.zip_path.setText(file_path)
            except ValueError as ve:
                logging.error(f"浏览7-Zip路径时出错: {ve}")
                self.zip_path.setText(file_path)

    def browse_typora(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择 Typora.exe", "",
            "Executable Files (*.exe);;All Files (*)"
        )
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

#######################################
# main
#######################################
def main():
    """
    主入口，如有异常写入 crash.log。
    """
    try:
        app = QtWidgets.QApplication(sys.argv)
        window = MDZLauncher()

        if len(sys.argv) > 1:
            mdz_file = sys.argv[1]
            if os.path.isfile(mdz_file) and mdz_file.lower().endswith(".mdz"):
                window.mdz_path = mdz_file
                window.append_log(f"通过命令行打开 .mdz 文件: {window.mdz_path}", "INFO")
                window.unpack_mdz()
                window.launch_typora()
            else:
                QtWidgets.QMessageBox.warning(window, "无效文件", "传入的文件不是有效的 .mdz 文件。")

        window.show()
        sys.exit(app.exec_())

    except Exception as e:
        err_msg = f"An unhandled exception occurred: {e}\n{traceback.format_exc()}"
        with open("crash.log", "w", encoding="utf-8") as f:
            f.write(err_msg)
        print(err_msg)
        input("Press Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main()
