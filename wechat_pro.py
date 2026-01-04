import sys
import os
import time
import random
import threading
import shutil
import tempfile
import ctypes
import math
import logging
import platform
from typing import Optional, List, Tuple, Dict
from dataclasses import dataclass, field

# === 依赖库检查与环境准备 ===
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
LOG_FILE = os.path.join(LOG_DIR, "wechat_pro.log")


def _format_missing_message(missing: List[str]) -> str:
    return f"缺少必要库: {', '.join(missing)}\n请运行: pip install {' '.join(missing)}"

def _init_logger():
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR, exist_ok=True)

    logger = logging.getLogger("wechat_pro")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger

logger = _init_logger()

def check_dependencies(exit_on_fail: bool = True) -> Tuple[bool, List[str]]:
    missing = []
    try:
        import uiautomation as auto  # noqa: F401
    except ImportError:
        missing.append("uiautomation")

    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        missing.append("pillow")

    if missing:
        msg = _format_missing_message(missing)
        logger.error(msg)
        print(msg, file=sys.stderr)

        if platform.system() == "Windows":
            try:
                ctypes.windll.user32.MessageBoxW(0, msg, "环境缺失", 0x10)
            except Exception:
                logger.debug("无法弹出 Windows 消息框，可能处于无图形环境")

        if exit_on_fail:
            return False, missing
        return False, missing

    logger.info("依赖检测通过：uiautomation、pillow 可用")
    return True, []

EXIT_ON_DEP_FAIL = os.getenv("WECHAT_PRO_EXIT_ON_DEP_FAIL", "1") != "0"
DEP_OK, DEP_MISSING = check_dependencies(exit_on_fail=EXIT_ON_DEP_FAIL)
if not DEP_OK:
    missing_msg = _format_missing_message(DEP_MISSING)
    if EXIT_ON_DEP_FAIL:
        sys.exit(1)
    raise ImportError(missing_msg)

# 核心库导入
import uiautomation as auto
import pyautogui
import pyperclip
from pynput import keyboard
from PIL import Image

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QTextEdit, 
                             QSpinBox, QDoubleSpinBox, QPushButton, QGroupBox, 
                             QMessageBox, QProgressBar, QCheckBox,
                             QFileDialog, QListWidget, QAbstractItemView, 
                             QSizePolicy, QDialog, QDateTimeEdit, QDialogButtonBox) 
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSettings, QUrl, QMimeData, QDateTime, QTime
from PyQt6.QtGui import QKeySequence, QShortcut, QIcon

# ==========================================
# 1. 基础设施层 (Infrastructure)
# ==========================================

class ClipboardScope:
    """[架构师视角] 剪贴板上下文管理器"""
    def __init__(self):
        self.original_data = ""

    def __enter__(self):
        try:
            self.original_data = pyperclip.paste()
        except Exception:
            self.original_data = ""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            time.sleep(0.05) 
            pyperclip.copy(self.original_data)
        except Exception:
            pass

class SemanticEngine:
    """[文本] 智能语义隐形引擎 2.0"""
    def __init__(self):
        self.invisible_chars = ["\u200b", "\u200c", "\u200d", "\u2060"]

    def humanize(self, base_content: str, count_threshold: int, current_idx: int, use_stealth: bool = True) -> str:
        if not use_stealth or not base_content:
            return base_content

        if count_threshold <= 1:
            return base_content

        if count_threshold < 5:
            noise = "".join(random.choices(self.invisible_chars, k=random.randint(1, 2)))
            return f"{base_content}{noise}"

        prefix = "".join(random.choices(self.invisible_chars, k=random.randint(1, 3)))
        suffix = "".join(random.choices(self.invisible_chars, k=random.randint(1, 3)))
        return f"{prefix}{base_content}{suffix}"

class ImageStealthEngine:
    """[多媒体] 图像隐形引擎"""
    def __init__(self):
        self.temp_dir = os.path.join(tempfile.gettempdir(), "wechat_pro_stealth_cache")
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        self.current_batch_files = []

    def process_batch(self, file_paths: List[str]) -> List[str]:
        self.cleanup_last_batch()
        new_paths = []
        video_exts = ['.mp4', '.mov', '.avi', '.mkv', '.wmv']
        
        for path in file_paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in video_exts:
                new_paths.append(path)
                continue
                
            try:
                processed_path = self._process_single_file(path)
                new_paths.append(processed_path)
                self.current_batch_files.append(processed_path)
            except Exception as e:
                new_paths.append(path)
        return new_paths

    def _process_single_file(self, path: str) -> str:
        ext = os.path.splitext(path)[1].lower()
        filename = f"stealth_{int(time.time()*1000)}_{random.randint(1000,9999)}{ext}"
        save_path = os.path.join(self.temp_dir, filename)

        if ext in ['.gif', '.webp']:
            self._inject_binary_noise(path, save_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
            self._perturb_pixels(path, save_path)
        else:
            shutil.copy2(path, save_path)
        return save_path

    def _perturb_pixels(self, src: str, dst: str):
        try:
            with Image.open(src) as img:
                if img.mode not in ('RGB', 'RGBA'):
                    img = img.convert('RGB')
                pixel = list(img.getpixel((0, 0)))
                pixel[0] = min(255, pixel[0] + 1) if pixel[0] < 255 else 254
                img.putpixel((0, 0), tuple(pixel))
                img.save(dst, quality=95, optimize=True)
        except Exception:
            self._inject_binary_noise(src, dst)

    def _inject_binary_noise(self, src: str, dst: str):
        shutil.copy2(src, dst)
        with open(dst, 'ab') as f:
            f.write(os.urandom(random.randint(4, 8)))

    def cleanup_last_batch(self):
        for p in self.current_batch_files:
            try:
                if os.path.exists(p): os.remove(p)
            except Exception:
                pass
        self.current_batch_files = []

# ==========================================
# 2. 驱动层 (Driver Layer - Human Mimicry)
# ==========================================

class HumanMimicry:
    """[核心升级] 真人拟态控制模块"""
    @staticmethod
    def random_jitter():
        """微小的鼠标抖动"""
        try:
            x, y = pyautogui.position()
            offset_x = random.randint(-2, 2)
            offset_y = random.randint(-2, 2)
            pyautogui.moveTo(x + offset_x, y + offset_y, duration=0.05, _pause=False)
        except Exception:
            pass
    
    @staticmethod
    def smooth_move_to(target_x, target_y):
        """[核心升级] 模拟真人手部移动鼠标"""
        try:
            duration = random.uniform(0.3, 0.6)
            target_x += random.randint(-3, 3)
            target_y += random.randint(-3, 3)
            pyautogui.moveTo(target_x, target_y, duration=duration, tween=pyautogui.easeOutQuad)
        except Exception:
            pass

class WeChatDriver:
    def __init__(self):
        self.wechat_window = None
        self.hwnd = 0 
    
    def connect(self) -> bool:
        queries = [
            {"Name": "微信", "ClassName": "WeChatMainWndForPC"}, 
            {"Name": "微信"},                                     
            {"Name": "WeChat"},                                   
            {"ClassName": "WeChatMainWndForPC"}                   
        ]
        for q in queries:
            try:
                win = auto.WindowControl(searchDepth=1, **q)
                if win.Exists(maxSearchSeconds=0.5):
                    self.wechat_window = win
                    self.hwnd = win.NativeWindowHandle 
                    return True
            except Exception:
                continue
        return False

    def activate(self, force: bool = False):
        if self.wechat_window:
            try:
                if force or not self.wechat_window.HasKeyboardFocus():
                    if self.wechat_window.GetWindowPattern().WindowVisualState == auto.WindowVisualState.Minimized:
                        self.wechat_window.GetWindowPattern().SetWindowVisualState(auto.WindowVisualState.Normal)
                    self.wechat_window.SetFocus()
                    if force: time.sleep(0.1)
            except Exception:
                pass

    def minimize_async(self):
        """[核心修复] 异步最小化窗口，不等待，不阻塞"""
        if self.hwnd:
            try:
                ctypes.windll.user32.PostMessageW(self.hwnd, 0x0112, 0xF020, 0)
            except Exception:
                pass

    def focus_input_box(self, enable_human: bool = False):
        if not self.wechat_window: return
        try:
            edit = self.wechat_window.EditControl(Name="输入")
            target_control = None
            
            if edit.Exists(maxSearchSeconds=0.1):
                target_control = edit
            else:
                edits = [c for c in self.wechat_window.GetChildren() if c.ControlTypeName == "EditControl"]
                if edits:
                    target_control = edits[-1]

            if target_control:
                if enable_human:
                    rect = target_control.BoundingRectangle
                    cx = (rect.left + rect.right) // 2
                    cy = (rect.top + rect.bottom) // 2
                    HumanMimicry.smooth_move_to(cx, cy)
                    pyautogui.click()
                else:
                    target_control.Click(simulateMove=False)
                return

            rect = self.wechat_window.BoundingRectangle
            if rect.width() > 0 and rect.height() > 0:
                tx = (rect.left + rect.right) // 2
                ty = rect.bottom - 60
                if enable_human:
                    HumanMimicry.smooth_move_to(tx, ty)
                    pyautogui.click()
                else:
                    pyautogui.click(tx, ty)
        except Exception:
            pass

    def search_contact(self, name: str) -> bool:
        if not self.wechat_window: return False
        try:
            self.activate(force=True)
            time.sleep(0.1)
            self.wechat_window.SendKeys('{Ctrl}f')
            time.sleep(0.2)
            pyperclip.copy(name)
            self.wechat_window.SendKeys('{Ctrl}v')
            time.sleep(0.5) 
            self.wechat_window.SendKeys('{Enter}')
            return True
        except Exception:
            return False

    def send_paste_and_enter(self, enable_human: bool = False):
        """[稳定发送] 彻底移除 Turbo，回归稳定"""
        if enable_human:
            HumanMimicry.random_jitter()
            time.sleep(random.uniform(0.05, 0.15))
            self.wechat_window.SendKeys('{Ctrl}v', waitTime=0.05)
            time.sleep(random.uniform(0.05, 0.1))
            self.wechat_window.SendKeys('{Enter}', waitTime=0.05)
        else:
            # 极速模式（稳定版）：最低间隔由外部 sleep 控制 (0.05s)
            self.wechat_window.SendKeys('{Ctrl}v', waitTime=0.01)
            # 这里的 waitTime 极小，但外部循环会有 0.05s 的保障
            self.wechat_window.SendKeys('{Enter}', waitTime=0.01)

# ==========================================
# 3. 业务逻辑层 (Service)
# ==========================================

@dataclass
class TaskConfig:
    target_list: List[Tuple[str, str, List[str]]] 
    global_msg: str
    global_files: List[str] 
    count_per_person: int 
    interval: float       
    start_delay: int
    target_timestamp: float
    enable_stealth_mode: bool 
    enable_human_simulation: bool 
    auto_minimize_done: bool 

class AutomationWorker(QThread):
    sig_log = pyqtSignal(str)
    sig_progress = pyqtSignal(int, int, str)
    sig_finished = pyqtSignal()
    sig_error = pyqtSignal(str)
    sig_set_clipboard_files = pyqtSignal(list)
    sig_clipboard_done = pyqtSignal()
    sig_countdown = pyqtSignal(int)

    def __init__(self, config: TaskConfig):
        super().__init__()
        self.config = config
        self._is_running = True
        self.driver = WeChatDriver()
        self.semantic = SemanticEngine()
        self.img_stealth = ImageStealthEngine()
        self._mutex = threading.Lock()
        self.clipboard_event = threading.Event()
        
        self.msgs_since_break = 0
        self.next_break_threshold = random.randint(15, 30)
        self.last_ui_update_time = 0.0

    def stop(self):
        with self._mutex:
            self._is_running = False

    def is_running(self):
        with self._mutex:
            return self._is_running
            
    def update_runtime_content(self, new_msg: str):
        with self._mutex:
            self.config.global_msg = new_msg
            if new_msg:
                self.sig_log.emit(f"📝 内容已更新: {new_msg[:10]}...")

    def update_runtime_files(self, new_files: List[str]):
        with self._mutex:
            self.config.global_files = new_files
            self.sig_log.emit(f"📂 附件列表已更新: 当前 {len(new_files)} 个文件")
            
    def update_runtime_params(self, new_count: int, new_interval: float):
        with self._mutex:
            self.config.count_per_person = new_count
            self.config.interval = new_interval

    def _smart_sleep(self, duration: float):
        with self._mutex:
            current_interval = self.config.interval
            if abs(duration - current_interval) < 0.001: 
                base_duration = current_interval
            else:
                base_duration = duration

        # [Limit Fix] 物理强制限速 0.05s
        if base_duration < 0.05:
            base_duration = 0.05
        
        actual_duration = base_duration
        if self.config.enable_human_simulation and base_duration > 0.1:
            actual_duration = base_duration * random.uniform(0.8, 1.3)
            actual_duration += random.uniform(0.01, 0.05)

        end_time = time.perf_counter() + actual_duration
        while time.perf_counter() < end_time:
            if not self.is_running(): return
            remaining = end_time - time.perf_counter()
            if remaining < 0.01:
                if remaining > 0.001: time.sleep(remaining / 2)
            else:
                time.sleep(0.005)

    def _check_human_break(self):
        if not self.config.enable_human_simulation:
            return
            
        self.msgs_since_break += 1
        if self.msgs_since_break >= self.next_break_threshold:
            break_time = random.uniform(3.0, 8.0)
            self.sig_log.emit(f"☕ 模拟真人疲劳: 暂停 {break_time:.1f} 秒...")
            HumanMimicry.random_jitter()
            self._smart_sleep(break_time)
            HumanMimicry.random_jitter()
            self.msgs_since_break = 0
            self.next_break_threshold = random.randint(15, 30)

    def on_clipboard_set_done(self):
        self.clipboard_event.set()

    def run(self):
        try:
            # [Fix] 移除 Turbo 判定
            if self.config.enable_human_simulation:
                self.sig_log.emit("🍃 真人拟态: 开启")
                pyautogui.PAUSE = 0.3
            else:
                self.sig_log.emit(f"⚡ 稳定极速: 开启 (Limit: 0.05s)")
                pyautogui.PAUSE = 0.05 
            
            # 1. 倒计时
            if self.config.target_timestamp > 0:
                self.sig_log.emit(f"⏳ 引擎已锁定！")
                while True:
                    if not self.is_running(): return
                    now = time.time()
                    remaining = self.config.target_timestamp - now
                    if remaining <= 0:
                        self.sig_countdown.emit(0)
                        break
                    self.sig_countdown.emit(int(remaining))
                    time.sleep(0.1 if remaining < 2 else 0.5)
            
            # 2. 连接微信
            self.sig_log.emit("🔗 正在连接微信...")
            if not self.driver.connect():
                raise Exception("未找到微信窗口！请确保PC微信已登录并显示在桌面上。")
            self.driver.activate()
            self._smart_sleep(0.5)

            total_targets = len(self.config.target_list)
            ops_done = 0 
            
            stealth_desc = "智能分级 (Auto-Leveling)" if self.config.enable_stealth_mode else "关闭"
            self.sig_log.emit(f"✅ 任务开始 | 目标: {total_targets} | 隐形系统: {stealth_desc}")
            
            # 3. 循环执行
            for idx, (name, custom_msg, custom_files) in enumerate(self.config.target_list):
                if not self.is_running(): break
                
                with self._mutex:
                    initial_msg = self.config.global_msg 
                    initial_files = self.config.global_files 
                    current_count_setting = self.config.count_per_person
                
                total_ops_est = total_targets * current_count_setting
                
                is_custom_mode = bool(custom_msg or custom_files)
                
                if not is_custom_mode and not initial_msg and not initial_files:
                    ops_done += current_count_setting
                    self.sig_log.emit(f"⚠️ 跳过 [{name}]: 内容为空")
                    self.sig_progress.emit(ops_done, total_ops_est, f"跳过: {name}")
                    continue

                try:
                    with ClipboardScope():
                        need_search = True
                        if total_targets == 1 and (not name or name == "当前窗口"):
                            need_search = False
                            self.sig_log.emit("📍 锁定当前窗口")
                        
                        if need_search:
                            self.sig_log.emit(f"🔍 切换: {name}")
                            if not self.driver.search_contact(name):
                                self.sig_log.emit(f"⚠️ 找不到: {name}")
                                ops_done += current_count_setting
                                self.sig_progress.emit(ops_done, total_ops_est, f"失败: {name}")
                                continue
                            self._smart_sleep(0.5)

                        self.driver.focus_input_box(self.config.enable_human_simulation)

                        sent_count_for_this_person = 0
                        
                        while True:
                            if not self.is_running(): break
                            
                            with self._mutex:
                                current_limit = self.config.count_per_person
                                current_interval_val = self.config.interval
                            
                            if sent_count_for_this_person >= current_limit:
                                break
                            
                            active_msg = ""
                            active_files = []
                            
                            if is_custom_mode:
                                active_msg = custom_msg
                                active_files = custom_files
                            else:
                                with self._mutex:
                                    active_msg = self.config.global_msg
                                    active_files = self.config.global_files
                            
                            if active_msg:
                                is_stealth = self.config.enable_stealth_mode
                                final_msg = self.semantic.humanize(active_msg, current_limit, sent_count_for_this_person, is_stealth)
                                
                                pyperclip.copy(final_msg)
                                self.driver.send_paste_and_enter(enable_human=self.config.enable_human_simulation)
                                self._check_human_break()
                                if active_files: self._smart_sleep(0.05)

                            if active_files:
                                final_files = active_files
                                if self.config.enable_stealth_mode:
                                    final_files = self.img_stealth.process_batch(active_files)
                                
                                self.clipboard_event.clear()
                                self.sig_set_clipboard_files.emit(final_files)
                                self.clipboard_event.wait(timeout=5.0)
                                
                                self.driver.send_paste_and_enter(enable_human=self.config.enable_human_simulation)
                                self._check_human_break()

                            sent_count_for_this_person += 1
                            ops_done += 1
                            est_total = total_targets * current_limit
                            
                            now_time = time.time()
                            is_last_item = (ops_done >= est_total) or (sent_count_for_this_person >= current_limit)
                            
                            if is_last_item or (now_time - self.last_ui_update_time > 0.2):
                                self.sig_progress.emit(ops_done, est_total, f"发送 -> {name} ({sent_count_for_this_person})")
                                self.last_ui_update_time = now_time

                            if sent_count_for_this_person < current_limit:
                                self._smart_sleep(current_interval_val)
                            
                    # 组间间隔
                    self._smart_sleep(0.5)

                except Exception as inner_e:
                    self.sig_log.emit(f"❌ 错误: {inner_e}")
                    self._smart_sleep(1)
            
            if self.config.auto_minimize_done and self.is_running():
                # [Fix] 冷却时间：正常等待 1.0 秒
                self.sig_log.emit("❄️ 冷却输入流 (1秒)...")
                for _ in range(10):
                    if not self.is_running(): break
                    time.sleep(0.1)
                
                if self.is_running():
                    self.sig_log.emit("📉 任务完成，发送归位信号...")
                    self.driver.minimize_async()
            
            self.img_stealth.cleanup_last_batch()

        except Exception as e:
            self.sig_error.emit(str(e))
        finally:
            try:
                self.sig_progress.disconnect()
                self.sig_log.disconnect()
                self.sig_countdown.disconnect()
            except Exception:
                pass
            self.sig_finished.emit()

# ==========================================
# 4. 表现层 (UI)
# ==========================================

class WeChatProUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.admin_suffix = " [ADMIN]" if ctypes.windll.shell32.IsUserAnAdmin() else " [USER]"
        self.setWindowTitle(f"Mr.Lu's WeChat Pro 2026 (Titan Edition){self.admin_suffix}")
        
        self.settings = QSettings("MrLu_Tools", "WeChatPro2026_Titan")
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        else:
            self.resize(980, 680) 
            
        self.worker: Optional[AutomationWorker] = None
        self.target_list = []
        self.target_datetime: Optional[QDateTime] = None 
        
        self._init_ui()
        self._init_style()
        self._restore_settings()
        self._setup_shortcuts()
        
        internal_icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
        if os.path.exists(internal_icon):
            self.setWindowIcon(QIcon(internal_icon))
        elif os.path.exists("app.ico"):
            self.setWindowIcon(QIcon("app.ico"))

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Left Panel
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(15)

        title = QLabel("WeChat Automation Pro 2026")
        title.setObjectName("Title")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(title)
        
        subtitle = QLabel("Titan Kernel | Adaptive Stealth | Human Mimicry")
        subtitle.setObjectName("Subtitle")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(subtitle)

        group_target = QGroupBox("🎯 目标管理 (拖拽 .txt 到此)")
        l_target = QVBoxLayout()
        h_file = QHBoxLayout()
        self.txt_file_path = QLineEdit()
        self.txt_file_path.setPlaceholderText("名单路径... (为空则对当前窗口发送)")
        self.txt_file_path.setReadOnly(True)
        btn_load = QPushButton("📂")
        btn_load.setFixedWidth(40)
        btn_load.clicked.connect(self._load_file_dialog)
        btn_reset = QPushButton("↺")
        btn_reset.setFixedWidth(40)
        btn_reset.clicked.connect(self._reset_mode)
        h_file.addWidget(self.txt_file_path)
        h_file.addWidget(btn_load)
        h_file.addWidget(btn_reset)
        l_target.addLayout(h_file)
        self.lbl_target_info = QLabel("模式: 单人手动 (昵称留空 = 轰炸当前窗口)")
        self.lbl_target_info.setStyleSheet("color: #aaa;")
        l_target.addWidget(self.lbl_target_info)
        self.input_single_name = QLineEdit()
        self.input_single_name.setPlaceholderText("在此输入好友昵称...")
        l_target.addWidget(self.input_single_name)
        group_target.setLayout(l_target)
        left_layout.addWidget(group_target)

        group_msg = QGroupBox("💬 发送内容 (支持多图/多视频)")
        l_msg = QVBoxLayout()
        l_msg.setSpacing(8)
        self.txt_msg = QTextEdit()
        self.txt_msg.setPlaceholderText("输入文字消息... (支持任务运行中实时修改)")
        self.txt_msg.setMinimumHeight(100)
        self.txt_msg.textChanged.connect(self._on_text_changed)
        l_msg.addWidget(self.txt_msg)
        
        self.list_images = QListWidget()
        self.list_images.setMinimumHeight(100) 
        self.list_images.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.list_images.setToolTip("支持图片和视频，可拖入多个文件")
        
        h_attach_head = QHBoxLayout()
        h_attach_head.addWidget(QLabel("📸 附件列表:"))
        btn_add_img = QPushButton("➕ 添加文件")
        btn_add_img.setFixedWidth(80)
        btn_add_img.setStyleSheet("padding: 2px; font-size: 11px;")
        btn_add_img.clicked.connect(self._open_media_dialog)
        h_attach_head.addWidget(btn_add_img)
        h_attach_head.addStretch()
        l_msg.addLayout(h_attach_head)
        
        l_msg.addWidget(self.list_images)
        self.list_images.itemDoubleClicked.connect(self._remove_list_item)

        h_checks = QHBoxLayout()
        self.chk_stealth = QCheckBox("🔰 分级智能隐形 (Auto-Leveling)")
        self.chk_stealth.setToolTip("Core 2.0:\n<5条: 轻量混淆\n>=5条: 深度包裹混淆\n自动图片 Hash 重构")
        self.chk_stealth.setStyleSheet("color: #81c784; font-weight: bold;")
        h_checks.addWidget(self.chk_stealth)
        l_msg.addLayout(h_checks)
        
        self.chk_human_sim = QCheckBox("真人拟态 (Anti-Bot)")
        self.chk_human_sim.setStyleSheet("color: #81c784; font-weight: bold;")
        self.chk_human_sim.setToolTip("模拟鼠标抖动、非线性节奏、防远程下线")
        l_msg.addWidget(self.chk_human_sim)
        
        group_msg.setLayout(l_msg)
        left_layout.addWidget(group_msg)

        # Right Panel
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(15)

        group_param = QGroupBox("⚙️ 核心参数")
        l_param = QVBoxLayout()
        h_p1 = QHBoxLayout()
        h_p1.addWidget(QLabel("发送次数:"))
        self.spin_count = QSpinBox()
        self.spin_count.setRange(1, 99999)
        self.spin_count.setValue(10)
        # [核心升级] 监听参数变化
        self.spin_count.valueChanged.connect(self._on_params_changed)
        h_p1.addWidget(self.spin_count)
        l_param.addLayout(h_p1)
        
        h_p2 = QHBoxLayout()
        h_p2.addWidget(QLabel("发送间隔(秒):"))
        self.spin_interval = QDoubleSpinBox()
        self.spin_interval.setRange(0.05, 100.00) # [Fix] 上限调整为 100秒
        self.spin_interval.setValue(0.50) # [Fix] 默认值调整为 0.5秒
        self.spin_interval.setSingleStep(0.1) # [Fix] 步长调整为 0.1 更方便调节
        self.spin_interval.setDecimals(2) # [Fix] 显示两位小数 (0.05 而不是 0.050)
        # [核心升级] 监听参数变化
        self.spin_interval.valueChanged.connect(self._on_params_changed)
        h_p2.addWidget(self.spin_interval)
        l_param.addLayout(h_p2)
        
        h_p3 = QHBoxLayout()
        h_p3.addWidget(QLabel("启动倒计时(秒):"))
        self.spin_delay = QSpinBox()
        self.spin_delay.setRange(0, 86400 * 30)
        self.spin_delay.setValue(3)
        self.spin_delay.setMinimumWidth(100)
        
        btn_calc = QPushButton("🕒 定时")
        btn_calc.clicked.connect(self._open_time_calculator)
        h_p3.addWidget(self.spin_delay)
        h_p3.addWidget(btn_calc)
        l_param.addLayout(h_p3)
        
        self.chk_auto_minimize = QCheckBox("完成自动归位 (Auto-Homing)")
        self.chk_auto_minimize.setToolTip("任务完成自动最小化微信，主程序回弹")
        self.chk_auto_minimize.setChecked(True)
        self.chk_auto_minimize.setStyleSheet("color: #e57373; font-weight: bold;")
        l_param.addWidget(self.chk_auto_minimize)
        
        self.lbl_schedule_time = QLabel("")
        self.lbl_schedule_time.setStyleSheet("color: #81c784; font-weight: bold; font-size: 12px;")
        self.lbl_schedule_time.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_schedule_time.hide() 
        l_param.addWidget(self.lbl_schedule_time)
        
        self.spin_delay.valueChanged.connect(self._on_manual_delay_change)
        
        group_param.setLayout(l_param)
        right_layout.addWidget(group_param)

        h_ctrl = QHBoxLayout()
        self.btn_run = QPushButton("🚀 启动 (Enter)")
        self.btn_run.setObjectName("BtnRun")
        self.btn_run.setMinimumHeight(60)
        self.btn_run.clicked.connect(self._start)
        # [Fix] 修复回车键启动
        self.btn_run.setDefault(True)
        self.btn_run.setAutoDefault(True)
        
        self.btn_stop = QPushButton("🛑 停止 (Esc)")
        self.btn_stop.setObjectName("BtnStop")
        self.btn_stop.setMinimumHeight(60)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)
        h_ctrl.addWidget(self.btn_run)
        h_ctrl.addWidget(self.btn_stop)
        right_layout.addLayout(h_ctrl)

        self.pbar = QProgressBar()
        self.pbar.setTextVisible(True)
        self.pbar.setMinimumHeight(25)
        right_layout.addWidget(self.pbar)
        
        right_layout.addWidget(QLabel("📜 运行日志:"))
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setObjectName("Log")
        self.txt_log.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_layout.addWidget(self.txt_log)

        main_layout.addWidget(left_widget, 55)
        main_layout.addWidget(right_widget, 45)
        self.setAcceptDrops(True)

    def _init_style(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #212121; color: #EEE; }
            QGroupBox { border: 1px solid #555; border-radius: 8px; margin-top: 12px; font-weight: bold; color: #81c784; padding-top: 20px; font-size: 13px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
            QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QListWidget { background-color: #333; color: #FFF; border: 1px solid #555; padding: 6px; border-radius: 4px; font-size: 12px;}
            QLineEdit:focus, QTextEdit:focus, QListWidget:focus { border: 1px solid #81c784; }
            QPushButton { background-color: #424242; color: white; border-radius: 6px; border: none; font-size: 14px; font-weight: bold; padding: 5px;}
            QPushButton:hover { background-color: #616161; }
            #BtnRun { background-color: #2e7d32; }
            #BtnRun:hover { background-color: #388e3c; }
            #BtnStop { background-color: #c62828; }
            #BtnStop:hover { background-color: #d32f2f; }
            #Title { font-size: 26px; color: #81c784; font-weight: bold; margin-bottom: 5px;}
            #Subtitle { font-size: 13px; color: #888; margin-bottom: 15px; }
            #Log { font-family: Consolas; font-size: 12px; background-color: #1a1a1a; color: #a5d6a7; border: none;}
            QProgressBar { border: none; background: #333; height: 18px; border-radius: 9px; text-align: center; color: black; font-weight: bold;}
            QProgressBar::chunk { background: #81c784; border-radius: 9px; }
            QListWidget::item { padding: 5px; }
            QListWidget::item:selected { background-color: #2e7d32; color: white; }
            QCheckBox { color: #bbb; font-weight: normal; }
            QCheckBox::indicator { width: 16px; height: 16px; }
        """)

    def _setup_shortcuts(self):
        self.listener = keyboard.Listener(on_press=self._on_key_press)
        self.listener.start()
        # [Fix] 绑定 Ctrl+Return 快捷键
        QShortcut(QKeySequence("Ctrl+Return"), self).activated.connect(lambda: self.btn_run.click())

    def _restore_settings(self):
        try:
            self.spin_count.setValue(int(self.settings.value("count", 10)))
            self.spin_interval.setValue(float(self.settings.value("interval", 0.05)))
            self.spin_delay.setValue(int(self.settings.value("delay", 3)))
            
            def to_bool(v, default=False):
                if isinstance(v, bool): return v
                return str(v).lower() == 'true'

            self.chk_stealth.setChecked(to_bool(self.settings.value("stealth", True)))
            self.chk_human_sim.setChecked(to_bool(self.settings.value("human_sim", False)))
            self.chk_auto_minimize.setChecked(to_bool(self.settings.value("auto_minimize", True)))
        except Exception:
            pass

    def _save_settings(self):
        self.settings.setValue("count", self.spin_count.value())
        self.settings.setValue("interval", self.spin_interval.value())
        self.settings.setValue("delay", self.spin_delay.value())
        self.settings.setValue("stealth", self.chk_stealth.isChecked())
        self.settings.setValue("human_sim", self.chk_human_sim.isChecked())
        self.settings.setValue("auto_minimize", self.chk_auto_minimize.isChecked())

    def _try_trigger_start(self):
        if self.btn_run.isEnabled():
            self._start()

    def _on_key_press(self, key):
        if key == keyboard.Key.esc:
            if self.worker and self.worker.is_running():
                self._stop()
    
    def _on_text_changed(self):
        if self.worker and self.worker.is_running():
            new_text = self.txt_msg.toPlainText().strip()
            self.worker.update_runtime_content(new_text)

    def _on_files_changed(self):
        if self.worker and self.worker.is_running():
            files = [self.list_images.item(i).text() for i in range(self.list_images.count())]
            self.worker.update_runtime_files(files)

    # [核心升级] 响应参数变更 + 智能联动
    def _on_params_changed(self):
        current_interval = self.spin_interval.value()
        
        # [Smart Logic] 如果间隔 > 1.0秒 且 未开启真人拟态 -> 自动勾选真人拟态
        # 这是一种保护机制：既然慢下来了，不如装得更像一点
        if current_interval > 1.0 and not self.chk_human_sim.isChecked():
            self.chk_human_sim.setChecked(True)
            self._log("💡 检测到间隔 > 1.0s，智能开启【真人拟态】")

        if self.worker and self.worker.is_running():
            new_count = self.spin_count.value()
            self.worker.update_runtime_params(new_count, current_interval)

    def _remove_list_item(self, item):
        self.list_images.takeItem(self.list_images.row(item))
        self._on_files_changed()

    def _open_media_dialog(self):
        img_exts = "*.png *.jpg *.jpeg *.gif *.bmp *.webp"
        vid_exts = "*.mp4 *.mov *.avi *.mkv *.wmv"
        filters = f"媒体文件 ({img_exts} {vid_exts});;图片 ({img_exts});;视频 ({vid_exts});;所有文件 (*.*)"
        
        files, _ = QFileDialog.getOpenFileNames(self, "选择发送的图片或视频", "", filters)
        if files:
            count = 0
            for f in files:
                items = [self.list_images.item(i).text() for i in range(self.list_images.count())]
                if f not in items:
                    self.list_images.addItem(f)
                    count += 1
            if count > 0:
                self._log(f"📂 已手动添加 {count} 个媒体文件")
                self._on_files_changed()

    def _open_time_calculator(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("定时启动计算器 (精确锁定)")
        dialog.setMinimumWidth(300)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("请选择预计启动的日期和时间："))
        
        dt_edit = QDateTimeEdit(QDateTime.currentDateTime())
        dt_edit.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
        dt_edit.setCalendarPopup(True)
        layout.addWidget(dt_edit)
        
        h_quick = QHBoxLayout()
        btn_1h = QPushButton("+1小时")
        btn_1h.clicked.connect(lambda: dt_edit.setDateTime(dt_edit.dateTime().addSecs(3600)))
        btn_tmr_9am = QPushButton("明天9点")
        def set_tmr_9am():
            now = QDateTime.currentDateTime()
            tmr = now.addDays(1)
            target = QDateTime(tmr.date(), QTime(9, 0))
            dt_edit.setDateTime(target)
        btn_tmr_9am.clicked.connect(set_tmr_9am)
        h_quick.addWidget(btn_1h)
        h_quick.addWidget(btn_tmr_9am)
        layout.addLayout(h_quick)
        
        lbl_preview = QLabel("预计等待: 0 秒")
        lbl_preview.setStyleSheet("color: #81c784; font-weight: bold;")
        layout.addWidget(lbl_preview)
        
        preview_timer = QTimer(dialog)
        def update_preview():
            now = QDateTime.currentDateTime()
            target = dt_edit.dateTime()
            seconds = now.secsTo(target)
            if seconds < 0:
                lbl_preview.setText("⚠️ 目标时间已过期")
                lbl_preview.setStyleSheet("color: #e57373;")
            else:
                m, s = divmod(seconds, 60)
                h, m = divmod(m, 60)
                lbl_preview.setText(f"预计等待: {seconds} 秒 ({int(h)}小时 {int(m)}分 {int(s)}秒)")
                lbl_preview.setStyleSheet("color: #81c784;")
        
        dt_edit.dateTimeChanged.connect(update_preview)
        preview_timer.timeout.connect(update_preview)
        preview_timer.start(1000) 
        update_preview()

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        def on_accept():
            has_msg = bool(self.txt_msg.toPlainText().strip())
            has_files = self.list_images.count() > 0
            has_batch = bool(self.target_list)
            
            if not has_batch and not has_msg and not has_files:
                QMessageBox.warning(dialog, "拒绝锁定", "❌ 请先输入发送内容或拖入文件！\n\n空内容无法启动定时任务。")
                return 

            dialog.accept()

        buttons.accepted.disconnect() 
        buttons.accepted.connect(on_accept)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            target = dt_edit.dateTime()
            self.target_datetime = target
            self.lbl_schedule_time.setText(f"⏰ 已锁定于 {target.toString('MM-dd HH:mm:ss')} 启动")
            self.lbl_schedule_time.show()
            
            secs = QDateTime.currentDateTime().secsTo(target)
            self.spin_delay.blockSignals(True)
            self.spin_delay.setValue(max(0, secs))
            self.spin_delay.blockSignals(False)
            self._log(f"⏰ 已设定精确时间点: {target.toString('yyyy-MM-dd HH:mm:ss')}")
            
            # [Fix] 增加检查：如果已经有任务在跑，记录一下
            if self.worker:
                self._log("⚠️ 检测到冲突：正在覆盖旧的定时任务...")

            self._start()

    def _on_manual_delay_change(self):
        if hasattr(self, 'target_datetime') and self.target_datetime:
            self.target_datetime = None
            if hasattr(self, 'lbl_schedule_time'): self.lbl_schedule_time.hide()

    def _update_countdown_display(self, seconds_left):
        if hasattr(self, 'lbl_schedule_time'):
            self.lbl_schedule_time.show()
            self.lbl_schedule_time.setText(f"🔥 引擎启动倒计时: {seconds_left} 秒")
            if seconds_left <= 5:
                self.lbl_schedule_time.setStyleSheet("color: #ff5252; font-weight: bold; font-size: 16px;")
            else:
                self.lbl_schedule_time.setStyleSheet("color: #81c784; font-weight: bold; font-size: 14px;")

    def set_clipboard_files(self, file_paths):
        try:
            clipboard = QApplication.clipboard()
            mime_data = QMimeData()
            urls = [QUrl.fromLocalFile(p) for p in file_paths]
            mime_data.setUrls(urls)
            clipboard.setMimeData(mime_data)
        except Exception as e:
            self._log(f"⚠️ 复制文件失败: {e}")
        finally:
            if self.worker:
                self.worker.on_clipboard_set_done()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
    
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if not files: return
        
        txt_files = [f for f in files if f.lower().endswith('.txt')]
        media_exts = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.mp4', '.mov', '.avi', '.mkv', '.wmv')
        media_files = [f for f in files if f.lower().endswith(media_exts)]
        
        if txt_files:
            self._load_file(txt_files[0])
            if media_files: self._log(f"➕ 同时检测到媒体文件，已添加 {len(media_files)} 个")
        
        if media_files:
            for f in media_files:
                items = [self.list_images.item(i).text() for i in range(self.list_images.count())]
                if f not in items:
                    self.list_images.addItem(f)
            self._log(f"🖼️ 已添加 {len(media_files)} 个图片/视频文件")
            self._on_files_changed()

    def _load_file_dialog(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择名单", "", "Txt (*.txt)")
        if fname: self._load_file(fname)

    def _load_file(self, path):
        try:
            targets = []
            custom_count = 0
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line: continue
                    parts = line.split('|')
                    name = parts[0].strip()
                    content = ""
                    files = []
                    if len(parts) > 1:
                        content = parts[1].strip()
                    if len(parts) > 2:
                        file_str = parts[2].strip()
                        if file_str:
                            files = [p.strip() for p in file_str.split(';') if p.strip()]
                    if content or files: custom_count += 1
                    targets.append((name, content, files))
            
            if not targets:
                 QMessageBox.warning(self, "提示", "文件为空或格式错误")
                 return

            self.target_list = targets
            self.txt_file_path.setText(os.path.basename(path))
            
            info_text = f"✅ 批量模式: {len(targets)} 人 (专属: {custom_count})"
            self.lbl_target_info.setText(info_text)
            self.lbl_target_info.setStyleSheet("color: #81c784")
            self.input_single_name.setEnabled(False)
            self._log(f"📂 名单加载成功: {len(targets)} 人")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取失败: {e}")
    
    def _reset_mode(self):
        self.target_list = []
        self.txt_file_path.clear()
        self.input_single_name.setEnabled(True)
        self.lbl_target_info.setText("模式: 单人手动")
        self.lbl_target_info.setStyleSheet("color: #aaa")
        self._on_files_changed() 
        self._log("🔄 状态已重置")

    def _log(self, msg):
        t = time.strftime("%H:%M:%S")
        self.txt_log.append(f"[{t}] {msg}")
        sb = self.txt_log.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _start(self):
        self._save_settings()
        
        # [Smart Logic] 启动前再次检查：如果间隔 > 1.0s，强制开启真人拟态
        if self.spin_interval.value() > 1.0 and not self.chk_human_sim.isChecked():
             self.chk_human_sim.setChecked(True)
             self._log("💡 启动检查：间隔 > 1.0s，已自动增强为【真人拟态】")

        self.txt_log.clear()
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)

        msg = self.txt_msg.toPlainText().strip()
        global_files = [self.list_images.item(i).text() for i in range(self.list_images.count())]

        targets = []
        is_batch = False
        
        if self.target_list:
            targets = self.target_list
            is_batch = True
        else:
            name = self.input_single_name.text().strip()
            if not name: name = "当前窗口"
            targets = [(name, "", [])]

        if is_batch:
            missing_count = 0
            for _, custom_msg, custom_files in targets:
                if (not custom_msg and not custom_files) and (not msg and not global_files):
                    missing_count += 1
            if missing_count > 0:
                if missing_count == len(targets):
                    QMessageBox.warning(self, "拒绝执行", "所有目标均无内容！")
                    self.btn_run.setEnabled(True)
                    self.btn_stop.setEnabled(False)
                    return
                else:
                    reply = QMessageBox.question(self, "预警", f"有 {missing_count} 人内容为空，是否跳过？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if reply == QMessageBox.StandardButton.No:
                        self.btn_run.setEnabled(True)
                        self.btn_stop.setEnabled(False)
                        return
        else:
            if not msg and not global_files:
                QMessageBox.warning(self, "提示", "请输入文字或拖入文件")
                self.btn_run.setEnabled(True)
                self.btn_stop.setEnabled(False)
                return

        target_ts = 0.0
        start_delay_secs = 0
        
        if self.target_datetime:
            target_ts = self.target_datetime.toMSecsSinceEpoch() / 1000.0
            now_ts = time.time()
            start_delay_secs = max(0, int(target_ts - now_ts))
        else:
            start_delay_secs = self.spin_delay.value()
            if start_delay_secs > 0:
                target_ts = time.time() + start_delay_secs

        config = TaskConfig(
            target_list=targets,
            global_msg=msg,
            global_files=global_files, 
            count_per_person=self.spin_count.value(),
            interval=self.spin_interval.value(),
            start_delay=start_delay_secs,
            target_timestamp=target_ts,
            enable_stealth_mode=self.chk_stealth.isChecked(),
            enable_human_simulation=self.chk_human_sim.isChecked(),
            auto_minimize_done=self.chk_auto_minimize.isChecked() 
        )

        self.pbar.setValue(0)
        ops_per_person = 0
        if msg: ops_per_person += 1
        if global_files: ops_per_person += 1
        if ops_per_person == 0: ops_per_person = 1
        self.pbar.setMaximum(len(targets) * self.spin_count.value() * ops_per_person)
        
        if self.worker:
            self._log("⚠️ 正在强制覆盖旧任务...")
            # 1. 核心步骤：切断旧线程的所有信号，防止其尸体触发 on_finished 导致按钮状态错误
            try:
                self.worker.sig_finished.disconnect()
                self.worker.sig_progress.disconnect()
                self.worker.sig_log.disconnect()
                self.worker.sig_countdown.disconnect()
            except Exception:
                pass
            
            # 2. 停止并强制终结
            self.worker.stop()
            if not self.worker.wait(1000): # 等待1秒让它体面退出
                self.worker.terminate() # 否则强制清理
            self.worker = None

        self.worker = AutomationWorker(config)
        self.worker.sig_log.connect(self._log)
        self.worker.sig_progress.connect(self.update_progress)
        self.worker.sig_finished.connect(self.on_finished)
        self.worker.sig_set_clipboard_files.connect(self.set_clipboard_files)
        self.worker.sig_countdown.connect(self._update_countdown_display)
        
        self.worker.start()

    def _stop(self):
        if self.worker:
            self.worker.stop()
            self._log("🛑 正在尝试紧急刹车...")

    def update_progress(self, current, total, info):
        self.pbar.setMaximum(total)
        self.pbar.setValue(current)
        self.setWindowTitle(f"WeChat Pro - {info}")

    def on_finished(self):
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.setWindowTitle("WeChat Pro 2026")
        self._log("🏁 任务完成")
        
        # [Fix] 异步归位：不在信号回调里直接操作窗口，而是延时0.5秒执行
        # 这能避开Windows窗口消息队列的死锁风险，确保主线程已经闲下来了
        if self.chk_auto_minimize.isChecked():
            QTimer.singleShot(500, self._perform_minimize_logic)
        
        self.target_datetime = None 
        if hasattr(self, 'lbl_schedule_time'): self.lbl_schedule_time.hide()

    def _perform_minimize_logic(self):
        """真正的归位逻辑，在主线程空闲时执行"""
        try:
            self._log("✨ 自动归位已触发")
            self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized | Qt.WindowState.WindowActive)
            self.showNormal()
            self.activateWindow()
            self.raise_()
        except Exception as e:
            self._log(f"⚠️ 归位微调: {e}")

    def closeEvent(self, event):
        self.settings.setValue("geometry", self.saveGeometry())
        self._save_settings() 
        if self.listener: self.listener.stop()
        if self.worker: self.worker.stop()
        event.accept()

if __name__ == "__main__":
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

    app = QApplication(sys.argv)
    win = WeChatProUI()
    win.show()
    sys.exit(app.exec())