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
import sqlite3
import re
from abc import ABC, abstractmethod
from typing import Optional, List, Tuple, Dict
from dataclasses import dataclass, field

# === 依赖库检查与环境准备 ===

def get_app_dir() -> str:
    """获取应用程序基础目录
    
    统一处理打包环境和开发环境的路径问题。
    Nuitka 打包后使用 sys.executable 所在目录，开发环境使用 __file__ 所在目录。
    
    Returns:
        str: 应用程序基础目录的绝对路径
    """
    if getattr(sys, 'frozen', False) or (hasattr(sys, 'compiled') and sys.compiled):
        # Nuitka/PyInstaller 打包环境
        return os.path.dirname(sys.executable)
    elif hasattr(sys, '_MEIPASS'):
        # PyInstaller 临时目录
        return sys._MEIPASS
    else:
        # 开发环境
        return os.path.dirname(os.path.abspath(__file__))

def get_data_dir() -> str:
    """获取应用程序数据目录（智能路径选择）
    
    优先使用 EXE 所在目录的 data 子目录。
    如果该目录不可写（如在 Program Files 中），则使用用户文档目录。
    
    Returns:
        str: 数据目录的绝对路径
    """
    # 获取基础目录
    if getattr(sys, 'frozen', False) or (hasattr(sys, 'compiled') and sys.compiled):
        # 打包后：使用 EXE 所在目录
        base_dir = os.path.dirname(sys.executable)
    else:
        # 开发环境：使用脚本所在目录
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 优先使用 EXE 同级的 data 目录
    data_dir = os.path.join(base_dir, 'data')
    
    # 检查是否可写
    try:
        if not os.path.exists(data_dir):
            os.makedirs(data_dir, exist_ok=True)
        
        # 测试是否可写
        test_file = os.path.join(data_dir, '.write_test')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        
        # 可写，返回此目录
        return data_dir
    except (OSError, PermissionError, IOError):
        # 不可写，使用用户文档目录
        import ctypes.wintypes
        
        CSIDL_PERSONAL = 5  # My Documents
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
        
        # 在文档目录创建 WeChatPro 子目录
        data_dir = os.path.join(buf.value, 'WeChatPro', 'data')
        os.makedirs(data_dir, exist_ok=True)
        return data_dir

def get_logs_dir() -> str:
    """获取应用程序日志目录
    
    优先使用 EXE 所在目录的 logs 子目录。
    如果该目录不可写，则使用临时目录。
    
    Returns:
        str: 日志目录的绝对路径
    """
    # 获取基础目录
    if getattr(sys, 'frozen', False) or (hasattr(sys, 'compiled') and sys.compiled):
        # 打包后：使用 EXE 所在目录
        base_dir = os.path.dirname(sys.executable)
    else:
        # 开发环境：使用脚本所在目录
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 优先使用 EXE 同级的 logs 目录
    logs_dir = os.path.join(base_dir, 'logs')
    
    # 检查是否可写
    try:
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir, exist_ok=True)
        
        # 测试是否可写
        test_file = os.path.join(logs_dir, '.write_test')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        
        # 可写，返回此目录
        return logs_dir
    except (OSError, PermissionError, IOError):
        # 不可写，使用临时目录
        logs_dir = os.path.join(tempfile.gettempdir(), 'WeChatPro', 'logs')
        os.makedirs(logs_dir, exist_ok=True)
        return logs_dir

def get_database_path() -> str:
    """获取数据库文件路径
    
    使用智能数据目录，确保数据持久化保存。
    
    Returns:
        str: 数据库文件的绝对路径
    """
    data_dir = get_data_dir()
    return os.path.join(data_dir, 'history.db')

def check_dependencies():
    missing = []
    try:
        import uiautomation as auto
    except ImportError:
        missing.append("uiautomation")
    
    try:
        from PIL import Image
    except ImportError:
        missing.append("pillow")
        
    if missing:
        msg = f"缺少必要库: {', '.join(missing)}\n请运行: pip install {' '.join(missing)}"
        ctypes.windll.user32.MessageBoxW(0, msg, "环境缺失", 0x10)
        sys.exit(1)

check_dependencies()

def setup_logging():
    log_dir = get_logs_dir()
    
    log_file = os.path.join(log_dir, f"wechat_pro_{time.strftime('%Y%m%d')}.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
        ]
    )
    return logging.getLogger('WeChatPro')

logger = setup_logging()

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
# 0. 抽象接口与常量定义
# ==========================================

class IMessageDriver(ABC):
    """消息驱动抽象接口
    
    定义消息发送驱动必须实现的方法，支持依赖注入和单元测试。
    """
    
    @abstractmethod
    def connect(self) -> bool:
        """连接目标应用窗口
        
        Returns:
            bool: 连接成功返回 True
        """
        pass
    
    @abstractmethod
    def activate(self, force: bool = False) -> None:
        """激活目标窗口
        
        Args:
            force: 是否强制激活
        """
        pass
    
    @abstractmethod
    def search_contact(self, name: str) -> bool:
        """搜索联系人
        
        Args:
            name: 联系人名称
            
        Returns:
            bool: 搜索成功返回 True
        """
        pass
    
    @abstractmethod
    def send_paste_and_enter(self, enable_human: bool = False) -> None:
        """粘贴并发送消息
        
        Args:
            enable_human: 是否启用真人拟态
        """
        pass


class IStealthEngine(ABC):
    """隐形引擎抽象接口
    
    定义消息隐形处理必须实现的方法。
    """
    
    @abstractmethod
    def humanize(self, content: str, **kwargs) -> str:
        """对内容进行隐形处理
        
        Args:
            content: 原始内容
            **kwargs: 额外参数
            
        Returns:
            处理后的内容
        """
        pass


class IImageStealthEngine(ABC):
    """图像隐形引擎抽象接口
    
    定义图像隐形处理必须实现的方法。
    """
    
    @abstractmethod
    def process_batch(self, file_paths: List[str]) -> List[str]:
        """批量处理图像文件
        
        Args:
            file_paths: 原始文件路径列表
            
        Returns:
            处理后的文件路径列表
        """
        pass
    
    @abstractmethod
    def cleanup_last_batch(self) -> None:
        """清理上一批临时文件"""
        pass


class TimingConfig:
    """时间配置常量
    
    集中管理所有时间相关的配置值，避免魔法数字。
    """
    CLIPBOARD_RESTORE_DELAY = 0.05
    HUMAN_MOVE_MIN_DURATION = 0.3
    HUMAN_MOVE_MAX_DURATION = 0.8
    FATIGUE_BREAK_MIN = 3.0
    FATIGUE_BREAK_MAX = 8.0
    FATIGUE_THRESHOLD_MIN = 15
    FATIGUE_THRESHOLD_MAX = 30
    CLIPBOARD_TIMEOUT = 5.0
    THREAD_WAIT_TIMEOUT = 3000
    COUNTDOWN_PRECISION = 0.1
    UI_UPDATE_INTERVAL = 0.2


class MediaExtensions:
    """媒体文件扩展名常量
    
    集中管理所有支持的文件扩展名。
    """
    IMAGE = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}
    VIDEO = {'.mp4', '.mov', '.avi', '.mkv', '.wmv'}
    ALL = IMAGE | VIDEO
    
    @classmethod
    def get_filter(cls) -> str:
        """获取文件对话框过滤器
        
        Returns:
            str: Qt 文件对话框格式的过滤器字符串
        """
        img = " ".join(f"*{ext}" for ext in cls.IMAGE)
        vid = " ".join(f"*{ext}" for ext in cls.VIDEO)
        return f"媒体文件 ({img} {vid});;图片 ({img});;视频 ({vid});;所有文件 (*)"

# ==========================================
# 1. 基础设施层 (Infrastructure)
# ==========================================

class ClipboardScope:
    """剪贴板上下文管理器
    
    在操作剪贴板前保存原始内容，操作完成后自动恢复。
    确保程序不会破坏用户的剪贴板数据。
    
    Example:
        >>> with ClipboardScope():
        ...     pyperclip.copy("临时内容")
        >>> # 剪贴板内容已恢复
    """
    def __init__(self):
        self.original_data = ""

    def __enter__(self):
        try:
            self.original_data = pyperclip.paste()
        except Exception as e:
            self.original_data = ""
            logger.debug(f"ClipboardScope.__enter__: {e}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            time.sleep(0.05) 
            pyperclip.copy(self.original_data)
        except Exception as e:
            logger.debug(f"ClipboardScope.__exit__: {e}")

class MessageTemplate:
    """消息模板引擎
    
    支持变量替换、随机选择等模板功能。
    
    支持的模板语法：
    - {name}: 变量替换
    - {date}: 当前日期
    - {time}: 当前时间
    - {random:选项1|选项2|选项3}: 随机选择
    
    Example:
        >>> template = MessageTemplate()
        >>> result = template.render("你好，{name}！", name="张三")
        >>> result
        '你好，张三！'
        >>> result = template.render("{random:早上好|下午好|晚上好}")
        >>> result in ["早上好", "下午好", "晚上好"]
        True
    """
    
    import re
    _pattern = re.compile(r'\{(\w+)(?::([^}]+))?\}')
    
    def __init__(self):
        self._variables = {}
    
    def set_variable(self, name: str, value: str):
        """设置模板变量
        
        Args:
            name: 变量名
            value: 变量值
        """
        self._variables[name] = value
    
    def render(self, template: str, **kwargs) -> str:
        """渲染模板
        
        Args:
            template: 模板字符串
            **kwargs: 变量值（覆盖预设变量）
            
        Returns:
            渲染后的字符串
        """
        variables = {**self._variables, **kwargs}
        
        def replace(match):
            name = match.group(1)
            options = match.group(2)
            
            if name == 'random' and options:
                choices = options.split('|')
                return random.choice(choices)
            
            if name == 'date':
                return time.strftime('%Y-%m-%d')
            
            if name == 'time':
                return time.strftime('%H:%M:%S')
            
            if name == 'datetime':
                return time.strftime('%Y-%m-%d %H:%M:%S')
            
            return str(variables.get(name, match.group(0)))
        
        return self._pattern.sub(replace, template)

class HistoryManager:
    """发送历史记录管理器
    
    使用 SQLite 存储发送记录，提供统计分析功能。
    
    Attributes:
        db_path: 数据库文件路径
        
    Example:
        >>> manager = HistoryManager()
        >>> manager.record_send("张三", "你好", success=True)
        >>> stats = manager.get_today_stats()
    """
    
    def __init__(self, db_path: Optional[str] = None):
        if db_path is None:
            # 使用智能路径选择
            db_path = get_database_path()
        
        self.db_path = db_path
        self._init_db()
        # 去重缓存：记录最近几秒内的发送，避免重复记录
        self._dedupe_cache = []
        self._dedupe_lock = threading.Lock()
        self._dedupe_window = 10.0  # 去重时间窗口（秒）
    
    def clear_dedupe_cache(self):
        """清空去重缓存，用于每次任务开始时"""
        with self._dedupe_lock:
            self._dedupe_cache = []
    
    def _init_db(self):
        """初始化数据库表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS send_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp REAL NOT NULL,
                target TEXT NOT NULL,
                content TEXT,
                has_attachment INTEGER DEFAULT 0,
                success INTEGER DEFAULT 1,
                error_message TEXT
            )
        ''')
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_timestamp 
            ON send_history(timestamp)
        ''')
        # 复合索引优化查询性能
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_target_timestamp
            ON send_history(target, timestamp)
        ''')
        conn.commit()
        conn.close()
    
    def record_send(self, target: str, content: str, 
                    has_attachment: bool = False, 
                    success: bool = True, 
                    error_message: str = None):
        """记录发送事件
        
        Args:
            target: 目标名称
            content: 发送内容
            has_attachment: 是否包含附件
            success: 是否成功
            error_message: 错误信息
        """
        current_time = time.time()
        
        # 写入数据库
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO send_history 
                (timestamp, target, content, has_attachment, success, error_message)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (current_time, target, content, int(has_attachment), int(success), error_message))
            conn.commit()
            conn.close()
            logger.debug(f"✅ 记录已保存: target={target}, content_len={len(content) if content else 0}")
        except Exception as e:
            logger.error(f"❌ 记录保存失败: {e}")
    
    def get_today_stats(self) -> Dict:
        """获取今日统计
        
        Returns:
            包含今日发送量、成功率等统计信息的字典
        """
        import datetime
        now = datetime.datetime.now()
        today_start = datetime.datetime(now.year, now.month, now.day)
        today_end = today_start + datetime.timedelta(days=1) - datetime.timedelta(seconds=1)
        
        ts_start = today_start.timestamp()
        ts_end = today_end.timestamp()
        
        logger.debug(f"📊 统计查询: {today_start.strftime('%Y-%m-%d %H:%M:%S')} - {today_end.strftime('%Y-%m-%d %H:%M:%S')}")
        logger.debug(f"📊 时间戳范围: {ts_start} - {ts_end}")
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT COUNT(*), SUM(success), COUNT(DISTINCT target)
            FROM send_history 
            WHERE timestamp >= ? AND timestamp <= ?
        ''', (ts_start, ts_end))
        
        row = cursor.fetchone()
        
        # 调试：查看原始数据
        cursor.execute('SELECT COUNT(*) FROM send_history')
        all_count = cursor.fetchone()[0]
        logger.debug(f"📊 数据库总记录数: {all_count}")
        
        conn.close()
        
        total, success_count, unique_targets = row
        success_count = success_count or 0
        
        logger.debug(f"📊 今日统计结果: total={total}, success={success_count}, targets={unique_targets}")
        
        return {
            'total': total or 0,
            'success': success_count,
            'failed': (total or 0) - success_count,
            'success_rate': (success_count / total * 100) if total else 0,
            'unique_targets': unique_targets or 0
        }
    
    def get_history(self, limit: int = 100, offset: int = 0) -> List[Dict]:
        """获取历史记录
        
        Args:
            limit: 返回记录数量限制
            offset: 偏移量
            
        Returns:
            历史记录列表
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, timestamp, target, content, has_attachment, success, error_message
            FROM send_history
            ORDER BY timestamp DESC
            LIMIT ? OFFSET ?
        ''', (limit, offset))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [
            {
                'id': row[0],
                'timestamp': row[1],
                'time_str': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(row[1])),
                'target': row[2],
                'content': row[3],
                'has_attachment': bool(row[4]),
                'success': bool(row[5]),
                'error_message': row[6]
            }
            for row in rows
        ]
    
    def clear_old_records(self, days: int = 30):
        """清理旧记录
        
        Args:
            days: 保留天数
        """
        cutoff = time.time() - days * 24 * 3600
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM send_history WHERE timestamp < ?', (cutoff,))
        conn.commit()
        conn.close()

class SemanticEngine(IStealthEngine):
    """智能语义隐形引擎
    
    通过注入不可见字符实现消息差异化，绕过平台重复检测。
    支持 24 种不可见字符随机注入，根据消息长度动态调整注入量。
    
    Attributes:
        INVISIBLE_CHARS: 可用的不可见字符列表（24种）
        
    Example:
        >>> engine = SemanticEngine()
        >>> result = engine.humanize("Hello", count_threshold=10)
        >>> len(result) > len("Hello")  # 包含不可见字符
        True
    """
    INVISIBLE_CHARS = [
        "\u200b", "\u200c", "\u200d", "\u2060",
        "\u2061", "\u2062", "\u2063", "\u2064",
        "\ufeff", "\u00ad", "\u034f", "\u180e",
        "\u200e", "\u200f", "\u202a", "\u202b",
        "\u202c", "\u202d", "\u202e", "\u2066",
        "\u2067", "\u2068", "\u2069", "\u00a0",
    ]

    def __init__(self):
        pass

    def humanize(self, base_content: str, count_threshold: int, current_idx: int, use_stealth: bool = True) -> str:
        """对消息内容进行隐形处理
        
        通过在消息中随机位置注入不可见字符，使每条消息在 Hash 层面不同，
        从而绕过平台的重复消息检测。
        
        Args:
            base_content: 原始消息内容
            count_threshold: 总发送次数阈值，用于判断是否需要处理
            current_idx: 当前发送索引（未使用，保留扩展）
            use_stealth: 是否启用隐形模式
            
        Returns:
            处理后的消息内容，包含随机注入的不可见字符
            
        Note:
            - 短消息（<10字符）：注入 1-2 个不可见字符
            - 长消息（>=10字符）：注入 2-4 个不可见字符
        """
        if not use_stealth or not base_content:
            return base_content

        if count_threshold <= 1:
            return base_content

        content_len = len(base_content)
        
        if content_len < 10:
            num_injections = random.randint(1, 2)
        else:
            num_injections = random.randint(2, 4)

        if num_injections >= content_len:
            noise = "".join(random.choices(self.INVISIBLE_CHARS, k=num_injections))
            return f"{base_content}{noise}"

        positions = sorted(random.sample(range(content_len), num_injections))
        
        result = list(base_content)
        offset = 0
        for pos in positions:
            char = random.choice(self.INVISIBLE_CHARS)
            result.insert(pos + offset, char)
            offset += 1
        
        return "".join(result)

class ImageStealthEngine(IImageStealthEngine):
    """图像隐形引擎
    
    通过修改图片像素和二进制数据实现图片差异化，绕过平台重复检测。
    支持多像素随机扰动和二进制噪声注入。
    
    Attributes:
        temp_dir: 临时文件存储目录
        current_batch_files: 当前批次处理的临时文件列表
        
    Example:
        >>> engine = ImageStealthEngine()
        >>> processed = engine.process_batch(["image1.png", "image2.jpg"])
        >>> engine.cleanup_last_batch()  # 清理临时文件
    """
    def __init__(self):
        self.temp_dir = os.path.join(tempfile.gettempdir(), "wechat_pro_stealth_cache")
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        self.current_batch_files = []

    def process_batch(self, file_paths: List[str]) -> List[str]:
        """批量处理图片文件
        
        对图片文件进行隐形处理，视频文件直接跳过。
        
        Args:
            file_paths: 待处理的文件路径列表
            
        Returns:
            处理后的文件路径列表（图片为临时文件路径，视频为原路径）
        """
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
                
                width, height = img.size
                num_pixels = min(random.randint(3, 5), (width * height) // 1000)
                
                pixels_to_modify = [
                    (random.randint(0, width - 1), random.randint(0, height - 1))
                    for _ in range(num_pixels)
                ]
                
                for x, y in pixels_to_modify:
                    pixel = list(img.getpixel((x, y)))
                    delta = random.randint(-3, 3)
                    if delta == 0:
                        delta = 1
                    for i in range(min(3, len(pixel))):
                        new_val = pixel[i] + delta
                        pixel[i] = max(0, min(255, new_val))
                    img.putpixel((x, y), tuple(pixel))
                
                img.save(dst, quality=95, optimize=True)
        except Exception as e:
            logger.debug(f"ImageStealthEngine._perturb_pixels: {e}")
            self._inject_binary_noise(src, dst)

    def _inject_binary_noise(self, src: str, dst: str):
        shutil.copy2(src, dst)
        with open(dst, 'ab') as f:
            f.write(os.urandom(random.randint(4, 8)))

    def cleanup_last_batch(self):
        for p in self.current_batch_files:
            try:
                if os.path.exists(p): os.remove(p)
            except Exception as e:
                logger.debug(f"ImageStealthEngine.cleanup_last_batch: 删除 {p} 失败 - {e}")
        self.current_batch_files = []

# ==========================================
# 2. 驱动层 (Driver Layer - Human Mimicry)
# ==========================================

class HumanMimicry:
    """真人拟态控制模块
    
    模拟人类操作行为，包括鼠标移动轨迹、抖动等。
    使用贝塞尔曲线实现自然的鼠标移动，避免被检测为自动化操作。
    
    Attributes:
        EASING_FUNCTIONS: 可用的缓动函数列表（5种）
        
    Example:
        >>> HumanMimicry.smooth_move_to(500, 300)  # 贝塞尔曲线移动
        >>> HumanMimicry.random_jitter()  # 随机抖动
    """
    
    EASING_FUNCTIONS = [
        lambda t: t,
        lambda t: t * t,
        lambda t: t * t * (3 - 2 * t),
        lambda t: 1 - (1 - t) * (1 - t),
        lambda t: 1 - (1 - t) ** 3,
    ]

    @staticmethod
    def random_jitter():
        """执行微小的鼠标抖动
        
        模拟人类手持鼠标时的自然抖动，随机偏移 ±2 像素。
        """
        try:
            x, y = pyautogui.position()
            offset_x = random.randint(-2, 2)
            offset_y = random.randint(-2, 2)
            pyautogui.moveTo(x + offset_x, y + offset_y, duration=0.05, _pause=False)
        except Exception as e:
            logger.debug(f"HumanMimicry.random_jitter: {e}")
    
    @staticmethod
    def smooth_move_to(target_x, target_y):
        """使用贝塞尔曲线模拟真人手部移动鼠标
        
        通过三阶贝塞尔曲线生成自然的鼠标移动轨迹，
        随机生成控制点和选择缓动函数，模拟人类操作。
        
        Args:
            target_x: 目标 X 坐标
            target_y: 目标 Y 坐标
            
        Note:
            - 移动时间：0.3-0.8 秒随机
            - 步数：20-40 步随机
            - 目标位置会有 ±3 像素的随机偏移
        """
        try:
            start_x, start_y = pyautogui.position()
            
            target_x += random.randint(-3, 3)
            target_y += random.randint(-3, 3)
            
            dist = math.sqrt((target_x - start_x) ** 2 + (target_y - start_y) ** 2)
            duration = random.uniform(0.3, 0.8)
            
            ctrl1_x = start_x + (target_x - start_x) * random.uniform(0.2, 0.4) + random.randint(-50, 50)
            ctrl1_y = start_y + (target_y - start_y) * random.uniform(0.2, 0.4) + random.randint(-50, 50)
            ctrl2_x = start_x + (target_x - start_x) * random.uniform(0.6, 0.8) + random.randint(-50, 50)
            ctrl2_y = start_y + (target_y - start_y) * random.uniform(0.6, 0.8) + random.randint(-50, 50)
            
            steps = random.randint(20, 40)
            easing_func = random.choice(HumanMimicry.EASING_FUNCTIONS)
            
            for i in range(steps + 1):
                t = easing_func(i / steps)
                
                x = (1-t)**3 * start_x + 3*(1-t)**2*t * ctrl1_x + 3*(1-t)*t**2 * ctrl2_x + t**3 * target_x
                y= (1-t)**3 * start_y+ 3*(1-t)**2*t * ctrl1_y+ 3*(1-t)*t**2 * ctrl2_y+ t**3 * target_y
                
                pyautogui.moveTo(int(x), int(y), duration=duration/steps, _pause=False)
                
        except Exception as e:
            logger.debug(f"HumanMimicry.smooth_move_to: {e}")

class WeChatDriver(IMessageDriver):
    """微信窗口驱动
    
    封装微信窗口的控制操作，包括连接、激活、搜索联系人、发送消息等。
    使用 uiautomation 库实现 Windows UI 自动化。
    
    Attributes:
        wechat_window: 微信窗口控件对象
        hwnd: 微信窗口句柄
        
    Example:
        >>> driver = WeChatDriver()
        >>> if driver.connect():
        ...     driver.search_contact("张三")
        ...     driver.send_paste_and_enter()
    """
    def __init__(self):
        self.wechat_window = None
        self.hwnd = 0 
    
    def connect(self) -> bool:
        """连接微信窗口
        
        尝试通过多种方式查找并连接微信窗口。
        
        Returns:
            bool: 连接成功返回 True，否则返回 False
            
        Note:
            支持的窗口查询方式：
            - 中文名称 + 类名
            - 中文名称
            - 英文名称
            - 类名
        """
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
            except Exception as e:
                logger.debug(f"WeChatDriver.connect: 查询 {q} 失败 - {e}")
                continue
        return False

    def activate(self, force: bool = False):
        """激活微信窗口
        
        将微信窗口置于前台并获得焦点。
        
        Args:
            force: 是否强制激活（即使已有焦点）
        """
        if self.wechat_window:
            try:
                if force or not self.wechat_window.HasKeyboardFocus():
                    if self.wechat_window.GetWindowPattern().WindowVisualState == auto.WindowVisualState.Minimized:
                        self.wechat_window.GetWindowPattern().SetWindowVisualState(auto.WindowVisualState.Normal)
                    self.wechat_window.SetFocus()
                    if force: time.sleep(0.1)
            except Exception as e:
                logger.debug(f"WeChatDriver.activate: {e}")

    def minimize_async(self):
        """异步最小化窗口
        
        使用 Windows API 异步最小化微信窗口，不等待、不阻塞。
        适用于任务完成后自动归位场景。
        """
        if self.hwnd:
            try:
                ctypes.windll.user32.PostMessageW(self.hwnd, 0x0112, 0xF020, 0)
            except Exception as e:
                logger.debug(f"WeChatDriver.minimize_async: {e}")

    def focus_input_box(self, enable_human: bool = False):
        """聚焦微信输入框
        
        尝试定位并聚焦微信的输入框控件。
        
        Args:
            enable_human: 是否启用真人拟态模式（贝塞尔曲线移动）
        """
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
                    time.sleep(1)
                    pyautogui.click()
                else:
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.click()
                return

            rect = self.wechat_window.BoundingRectangle
            if rect.width() > 0 and rect.height() > 0:
                tx = (rect.left + rect.right) // 2
                ty = rect.bottom - 60
                if enable_human:
                    HumanMimicry.smooth_move_to(tx, ty)
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.click()
                else:
                    pyautogui.click(tx, ty)
                    time.sleep(1)
                    pyautogui.click(tx, ty)
        except Exception as e:
            logger.debug(f"WeChatDriver.focus_input_box: {e}")

    def search_contact(self, name: str) -> bool:
        """搜索并定位联系人
        
        通过微信的搜索功能定位指定联系人。
        
        Args:
            name: 联系人名称或昵称
            
        Returns:
            bool: 搜索成功返回 True，否则返回 False
        """
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
        """通过剪贴板发送消息
        
        将剪贴板内容粘贴到微信输入框并发送。
        使用 Ctrl+V 粘贴 + Enter 发送的方式。
        
        Args:
            enable_human: 是否启用真人拟态模式
        """
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
    """任务配置数据类
    
    封装自动化任务的所有配置参数。
    
    Attributes:
        target_list: 目标列表，每个元素为 (名称, 专属消息, 附件列表) 元组
        global_msg: 全局消息内容
        global_files: 全局附件文件路径列表
        count_per_person: 每人发送次数
        interval: 发送间隔（秒）
        start_delay: 启动延迟（秒）
        target_timestamp: 目标启动时间戳（精确启动模式）
        enable_stealth_mode: 是否启用隐形模式
        enable_human_simulation: 是否启用真人拟态
        auto_minimize_done: 任务完成后是否自动最小化
    """
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
    """自动化任务工作线程
    
    在独立线程中执行消息发送任务，支持：
    - 倒计时启动和精确时间点启动
    - 批量发送消息和附件
    - 隐形模式和真人拟态
    - 运行时参数修改
    - 疲劳模拟（随机休息）
    
    Signals:
        sig_log: 日志信号，参数为日志消息
        sig_progress: 进度信号，参数为 (当前, 总数, 描述)
        sig_finished: 完成信号
        sig_error: 错误信号，参数为错误消息
        sig_set_clipboard_files: 设置剪贴板文件信号
        sig_clipboard_done: 剪贴板操作完成信号
        sig_countdown: 倒计时信号，参数为剩余秒数
        
    Example:
        >>> config = TaskConfig(...)
        >>> worker = AutomationWorker(config)
        >>> worker.sig_log.connect(print)
        >>> worker.start()
    """
    sig_log = pyqtSignal(str)
    sig_progress = pyqtSignal(int, int, str)
    sig_finished = pyqtSignal()
    sig_error = pyqtSignal(str)
    sig_set_clipboard_files = pyqtSignal(list)
    sig_clipboard_done = pyqtSignal()
    sig_countdown = pyqtSignal(int)

    def __init__(self, config: TaskConfig, 
                 driver: Optional[IMessageDriver] = None,
                 stealth_engine: Optional[IStealthEngine] = None,
                 image_stealth_engine: Optional[IImageStealthEngine] = None):
        """初始化自动化工作线程
        
        Args:
            config: 任务配置对象
            driver: 消息驱动实例（可选，默认使用 WeChatDriver）
            stealth_engine: 文本隐形引擎实例（可选，默认使用 SemanticEngine）
            image_stealth_engine: 图像隐形引擎实例（可选，默认使用 ImageStealthEngine）
        """
        super().__init__()
        self.config = config
        self._is_running = True
        self.driver = driver or WeChatDriver()
        self.semantic = stealth_engine or SemanticEngine()
        self.img_stealth = image_stealth_engine or ImageStealthEngine()
        self._mutex = threading.Lock()
        self.clipboard_event = threading.Event()
        self.history_manager: Optional[HistoryManager] = None
        
        self.msgs_since_break = 0
        self.next_break_threshold = random.randint(TimingConfig.FATIGUE_THRESHOLD_MIN, TimingConfig.FATIGUE_THRESHOLD_MAX)
        self.last_ui_update_time = 0.0
    
    def set_history_manager(self, manager: 'HistoryManager'):
        """设置历史记录管理器
        
        Args:
            manager: 历史记录管理器实例
        """
        self.history_manager = manager

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
                                if not self.clipboard_event.wait(timeout=5.0):
                                    self.sig_log.emit("⚠️ 剪贴板设置超时，跳过本次附件发送")
                                    logger.warning(f"Clipboard timeout for files: {final_files}")
                                    continue
                                
                                self.driver.send_paste_and_enter(enable_human=self.config.enable_human_simulation)
                                self._check_human_break()

                            sent_count_for_this_person += 1
                            ops_done += 1
                            est_total = total_targets * current_limit
                            
                            if self.history_manager:
                                self.history_manager.record_send(
                                    target=name,
                                    content=active_msg[:100] if active_msg else "",
                                    has_attachment=bool(active_files),
                                    success=True
                                )
                            
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
            except Exception as e:
                logger.debug(f"AutomationWorker.run finally: 断开信号失败 - {e}")
            self.sig_finished.emit()

# ==========================================
# 4. 表现层 (UI)
# ==========================================

class SettingsManager:
    """设置管理器
    
    负责应用程序设置的持久化存储和读取。
    
    Attributes:
        settings: QSettings 实例
        
    Example:
        >>> manager = SettingsManager()
        >>> manager.save("count", 10)
        >>> manager.load("count", 10)
        10
    """
    
    def __init__(self):
        self.settings = QSettings("MrLu_Tools", "WeChatPro2026_Titan")
    
    def save(self, key: str, value):
        """保存设置值
        
        Args:
            key: 设置键名
            value: 设置值
        """
        self.settings.setValue(key, value)
    
    def load(self, key: str, default=None):
        """加载设置值
        
        Args:
            key: 设置键名
            default: 默认值
            
        Returns:
            设置值，如果不存在则返回默认值
        """
        return self.settings.value(key, default)
    
    def save_geometry(self, geometry: bytes):
        """保存窗口几何信息
        
        Args:
            geometry: 窗口几何数据
        """
        self.settings.setValue("geometry", geometry)
    
    def load_geometry(self):
        """加载窗口几何信息
        
        Returns:
            窗口几何数据，如果不存在则返回 None
        """
        return self.settings.value("geometry")
    
    def save_all(self, config: dict):
        """批量保存设置
        
        Args:
            config: 设置字典
        """
        for key, value in config.items():
            self.settings.setValue(key, value)
    
    @staticmethod
    def to_bool(value, default: bool = False) -> bool:
        """将值转换为布尔值
        
        Args:
            value: 待转换的值
            default: 默认值
            
        Returns:
            布尔值
        """
        if isinstance(value, bool):
            return value
        return str(value).lower() == 'true'


class FileHandler:
    """文件处理器
    
    负责名单文件和媒体文件的处理。
    
    Attributes:
        target_list: 目标列表
        
    Example:
        >>> handler = FileHandler()
        >>> targets = handler.load_target_list("names.txt")
        >>> len(targets)
        10
    """
    
    MEDIA_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', 
                        '.mp4', '.mov', '.avi', '.mkv', '.wmv')
    
    def __init__(self):
        self.target_list: List[Tuple[str, str, List[str]]] = []
    
    def load_target_list(self, path: str) -> Tuple[bool, str, int]:
        """加载目标名单文件
        
        Args:
            path: 名单文件路径
            
        Returns:
            Tuple[成功标志, 消息, 目标数量]
        """
        try:
            targets = []
            custom_count = 0
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
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
                    if content or files:
                        custom_count += 1
                    targets.append((name, content, files))
            
            if not targets:
                return False, "文件为空或格式错误", 0
            
            self.target_list = targets
            return True, f"名单加载成功: {len(targets)} 人 (专属: {custom_count})", len(targets)
            
        except Exception as e:
            return False, f"读取失败: {e}", 0
    
    def reset(self):
        """重置目标列表"""
        self.target_list = []
    
    @classmethod
    def is_media_file(cls, path: str) -> bool:
        """检查是否为媒体文件
        
        Args:
            path: 文件路径
            
        Returns:
            是否为媒体文件
        """
        return path.lower().endswith(cls.MEDIA_EXTENSIONS)
    
    @classmethod
    def is_text_file(cls, path: str) -> bool:
        """检查是否为文本文件
        
        Args:
            path: 文件路径
            
        Returns:
            是否为文本文件
        """
        return path.lower().endswith('.txt')
    
    @classmethod
    def filter_files(cls, files: List[str]) -> Tuple[List[str], List[str]]:
        """过滤文件列表
        
        Args:
            files: 文件路径列表
            
        Returns:
            Tuple[文本文件列表, 媒体文件列表]
        """
        txt_files = [f for f in files if cls.is_text_file(f)]
        media_files = [f for f in files if cls.is_media_file(f)]
        return txt_files, media_files


class TaskController:
    """任务控制器
    
    负责任务的启动、停止和状态管理。
    
    Attributes:
        worker: 当前工作线程
        is_running: 任务是否正在运行
        
    Example:
        >>> controller = TaskController()
        >>> controller.start(config)
        >>> controller.stop()
    """
    
    def __init__(self):
        self.worker: Optional[AutomationWorker] = None
        self._is_running = False
    
    def is_running(self) -> bool:
        """检查任务是否正在运行
        
        Returns:
            是否正在运行
        """
        return self.worker is not None and self.worker.is_running()
    
    def start(self, config: TaskConfig, 
              on_log=None, on_progress=None, on_finished=None,
              on_clipboard=None, on_countdown=None,
              history_manager=None) -> AutomationWorker:
        """启动任务
        
        Args:
            config: 任务配置
            on_log: 日志回调
            on_progress: 进度回调
            on_finished: 完成回调
            on_clipboard: 剪贴板回调
            on_countdown: 倒计时回调
            history_manager: 历史记录管理器
            
        Returns:
            工作线程实例
        """
        if self.worker:
            self._disconnect_signals()
            self.worker.stop()
            # 增加等待时间到10秒，确保旧 worker 完全停止
            wait_success = self.worker.wait(10000)
            if not wait_success:
                logger.warning("旧 worker 未能在10秒内完全停止，可能存在并发问题")
            # 清除旧 worker 引用
            self.worker = None
        
        self.worker = AutomationWorker(config)
        
        # 在启动前设置 history_manager
        if history_manager:
            self.worker.set_history_manager(history_manager)
        
        if on_log:
            self.worker.sig_log.connect(on_log)
        if on_progress:
            self.worker.sig_progress.connect(on_progress)
        if on_finished:
            self.worker.sig_finished.connect(on_finished)
        if on_clipboard:
            self.worker.sig_set_clipboard_files.connect(on_clipboard)
        if on_countdown:
            self.worker.sig_countdown.connect(on_countdown)
        
        self.worker.start()
        return self.worker
    
    def stop(self, wait: bool = True, timeout: int = 10000) -> bool:
        """停止任务
        
        Args:
            wait: 是否等待 worker 完全停止
            timeout: 等待超时时间（毫秒）
            
        Returns:
            bool: worker 是否成功停止
        """
        if self.worker:
            self.worker.stop()
            if wait:
                success = self.worker.wait(timeout)
                if not success:
                    logger.warning(f"worker 未能在 {timeout} 毫秒内停止")
                return success
        return True
    
    def _disconnect_signals(self):
        """断开所有信号连接"""
        if self.worker:
            try:
                self.worker.sig_finished.disconnect()
                self.worker.sig_progress.disconnect()
                self.worker.sig_log.disconnect()
                self.worker.sig_set_clipboard_files.disconnect()
                self.worker.sig_countdown.disconnect()
            except Exception:
                pass
    
    def update_runtime_content(self, content: str):
        """更新运行时内容
        
        Args:
            content: 新内容
        """
        if self.worker and self.worker.is_running():
            self.worker.update_runtime_content(content)
    
    def update_runtime_files(self, files: List[str]):
        """更新运行时文件列表
        
        Args:
            files: 文件列表
        """
        if self.worker and self.worker.is_running():
            self.worker.update_runtime_files(files)
    
    def update_runtime_params(self, count: int, interval: float):
        """更新运行时参数
        
        Args:
            count: 发送次数
            interval: 发送间隔
        """
        if self.worker and self.worker.is_running():
            self.worker.update_runtime_params(count, interval)
    
    def on_clipboard_set_done(self):
        """剪贴板设置完成回调"""
        if self.worker:
            self.worker.on_clipboard_set_done()


class LogBuffer:
    """日志缓冲器
    
    提供日志缓冲、行数限制和导出功能。
    
    Attributes:
        max_lines: 最大日志行数
        buffer: 日志缓冲区
        
    Example:
        >>> buffer = LogBuffer(max_lines=1000)
        >>> buffer.append("Hello")
        >>> buffer.export_to_file("log.txt")
    """
    
    def __init__(self, max_lines: int = 1000):
        self.max_lines = max_lines
        self._buffer: List[str] = []
        self._lock = threading.Lock()
    
    def append(self, line: str):
        """添加日志行
        
        Args:
            line: 日志行内容
        """
        with self._lock:
            self._buffer.append(line)
            if len(self._buffer) > self.max_lines:
                excess = len(self._buffer) - self.max_lines
                self._buffer = self._buffer[excess:]
    
    def get_all(self) -> List[str]:
        """获取所有日志行
        
        Returns:
            日志行列表
        """
        with self._lock:
            return self._buffer.copy()
    
    def clear(self):
        """清空日志缓冲区"""
        with self._lock:
            self._buffer.clear()
    
    def export_to_file(self, path: str) -> bool:
        """导出日志到文件
        
        Args:
            path: 目标文件路径
            
        Returns:
            是否导出成功
        """
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.get_all()))
            return True
        except Exception:
            return False
    
    def get_line_count(self) -> int:
        """获取当前日志行数
        
        Returns:
            日志行数
        """
        with self._lock:
            return len(self._buffer)


class WeChatProUI(QMainWindow):
    """微信 Pro 主窗口
    
    提供图形用户界面，包括：
    - 目标管理：单个目标、批量名单导入
    - 发送内容：消息编辑、附件管理
    - 核心参数：发送次数、间隔、定时启动
    - 高级选项：隐形模式、真人拟态、自动归位
    - 实时日志：进度显示、操作记录
    
    Attributes:
        worker: 当前运行的工作线程
        target_list: 目标列表
        target_datetime: 定时启动的目标时间
        
    Example:
        >>> app = QApplication(sys.argv)
        >>> window = WeChatProUI()
        >>> window.show()
        >>> sys.exit(app.exec())
    """
    def __init__(self):
        super().__init__()
        self.admin_suffix = " [ADMIN]" if ctypes.windll.shell32.IsUserAnAdmin() else " [USER]"
        self.setWindowTitle(f"Mr.Lu's WeChat Pro 2026 (Titan Edition){self.admin_suffix}")
        
        self.settings_manager = SettingsManager()
        self.file_handler = FileHandler()
        self.task_controller = TaskController()
        self.log_buffer = LogBuffer(max_lines=500)
        self.history_manager = HistoryManager()
        
        geometry = self.settings_manager.load_geometry()
        if geometry:
            self.restoreGeometry(geometry)
        else:
            self.resize(980, 680) 
            
        self.target_datetime: Optional[QDateTime] = None 
        
        self._init_ui()
        self._init_style()
        self._restore_settings()
        self._setup_shortcuts()
        
        internal_icon = os.path.join(get_app_dir(), "app.ico")
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
        
        group_stats = QGroupBox("📊 今日统计")
        l_stats = QVBoxLayout()
        
        h_stats1 = QHBoxLayout()
        h_stats1.addWidget(QLabel("发送量:"))
        self.lbl_stat_total = QLabel("0")
        self.lbl_stat_total.setStyleSheet("color: #81c784; font-weight: bold; font-size: 14px;")
        h_stats1.addWidget(self.lbl_stat_total)
        h_stats1.addStretch()
        l_stats.addLayout(h_stats1)
        
        h_stats2 = QHBoxLayout()
        h_stats2.addWidget(QLabel("成功率:"))
        self.lbl_stat_rate = QLabel("0%")
        self.lbl_stat_rate.setStyleSheet("color: #81c784; font-weight: bold; font-size: 14px;")
        h_stats2.addWidget(self.lbl_stat_rate)
        h_stats2.addStretch()
        l_stats.addLayout(h_stats2)
        
        h_stats3 = QHBoxLayout()
        h_stats3.addWidget(QLabel("活跃目标:"))
        self.lbl_stat_targets = QLabel("0")
        self.lbl_stat_targets.setStyleSheet("color: #81c784; font-weight: bold; font-size: 14px;")
        h_stats3.addWidget(self.lbl_stat_targets)
        h_stats3.addStretch()
        l_stats.addLayout(h_stats3)
        
        group_stats.setLayout(l_stats)
        right_layout.addWidget(group_stats)

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
        
        h_log_header = QHBoxLayout()
        h_log_header.addWidget(QLabel("📜 运行日志:"))
        btn_export_log = QPushButton("📤 导出")
        btn_export_log.setFixedWidth(60)
        btn_export_log.setStyleSheet("padding: 2px; font-size: 11px;")
        btn_export_log.clicked.connect(self._export_log)
        h_log_header.addWidget(btn_export_log)
        btn_clear_log = QPushButton("🗑️ 清空")
        btn_clear_log.setFixedWidth(60)
        btn_clear_log.setStyleSheet("padding: 2px; font-size: 11px;")
        btn_clear_log.clicked.connect(self._clear_log)
        h_log_header.addWidget(btn_clear_log)
        right_layout.addLayout(h_log_header)
        
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
            QLabel { color: #E0E0E0; }
            QGroupBox { border: 1px solid #555; border-radius: 8px; margin-top: 12px; font-weight: bold; color: #81c784; padding-top: 20px; font-size: 13px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
            QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QListWidget { background-color: #333; color: #FFF; border: 1px solid #555; padding: 6px; border-radius: 4px; font-size: 12px;}
            QLineEdit:focus, QTextEdit:focus, QListWidget:focus { border: 1px solid #81c784; }
            QPushButton { background-color: #424242; color: #FFFFFF; border-radius: 6px; border: none; font-size: 14px; font-weight: bold; padding: 5px;}
            QPushButton:hover { background-color: #616161; }
            #BtnRun { background-color: #2e7d32; }
            #BtnRun:hover { background-color: #388e3c; }
            #BtnStop { background-color: #c62828; }
            #BtnStop:hover { background-color: #d32f2f; }
            #Title { font-size: 26px; color: #81c784; font-weight: bold; margin-bottom: 5px;}
            #Subtitle { font-size: 13px; color: #888; margin-bottom: 15px; }
            #Log { font-family: Consolas; font-size: 12px; background-color: #1a1a1a; color: #a5d6a7; border: none;}
            QProgressBar { border: none; background: #333; height: 18px; border-radius: 9px; text-align: center; color: white; font-weight: bold;}
            QProgressBar::chunk { background: #81c784; border-radius: 9px; }
            QListWidget::item { padding: 5px; }
            QListWidget::item:selected { background-color: #2e7d32; color: white; }
            QCheckBox { color: #E0E0E0; font-weight: normal; }
            QCheckBox::indicator { width: 16px; height: 16px; }
            QDialog { background-color: #2b2b2b; }
        """)
    
    def _show_message_box(self, icon, title, text, buttons=QMessageBox.StandardButton.Ok, default_button=QMessageBox.StandardButton.Ok):
        """显示带样式的消息框
        
        Args:
            icon: 消息图标
            title: 标题
            text: 内容
            buttons: 按钮
            default_button: 默认按钮
            
        Returns:
            用户选择的按钮
        """
        msg_box = QMessageBox(self)
        msg_box.setIcon(icon)
        msg_box.setWindowTitle(title)
        msg_box.setText(text)
        msg_box.setStandardButtons(buttons)
        msg_box.setDefaultButton(default_button)
        
        msg_box.setStyleSheet("""
            QMessageBox { background-color: #2b2b2b; }
            QLabel { color: #E0E0E0; font-size: 13px; }
            QPushButton { background-color: #424242; color: #FFFFFF; border-radius: 6px; border: none; font-size: 13px; font-weight: bold; padding: 8px 20px; min-width: 80px; }
            QPushButton:hover { background-color: #616161; }
        """)
        
        return msg_box.exec()

    def _setup_shortcuts(self):
        self.listener = keyboard.Listener(on_press=self._on_key_press)
        self.listener.start()
        QShortcut(QKeySequence("Ctrl+Return"), self).activated.connect(lambda: self.btn_run.click())
        self._update_stats()
    
    def _update_stats(self):
        """更新统计面板"""
        try:
            stats = self.history_manager.get_today_stats()
            self.lbl_stat_total.setText(str(stats['total']))
            self.lbl_stat_rate.setText(f"{stats['success_rate']:.1f}%")
            self.lbl_stat_targets.setText(str(stats['unique_targets']))
        except Exception:
            self.lbl_stat_total.setText("0")
            self.lbl_stat_rate.setText("0%")
            self.lbl_stat_targets.setText("0")

    def _restore_settings(self):
        try:
            self.spin_count.setValue(int(self.settings_manager.load("count", 10)))
            self.spin_interval.setValue(float(self.settings_manager.load("interval", 0.05)))
            self.spin_delay.setValue(int(self.settings_manager.load("delay", 3)))
            
            self.chk_stealth.setChecked(SettingsManager.to_bool(self.settings_manager.load("stealth", True)))
            self.chk_human_sim.setChecked(SettingsManager.to_bool(self.settings_manager.load("human_sim", False)))
            self.chk_auto_minimize.setChecked(SettingsManager.to_bool(self.settings_manager.load("auto_minimize", True)))
        except Exception:
            pass

    def _save_settings(self):
        self.settings_manager.save_all({
            "count": self.spin_count.value(),
            "interval": self.spin_interval.value(),
            "delay": self.spin_delay.value(),
            "stealth": self.chk_stealth.isChecked(),
            "human_sim": self.chk_human_sim.isChecked(),
            "auto_minimize": self.chk_auto_minimize.isChecked()
        })

    def _try_trigger_start(self):
        if self.btn_run.isEnabled():
            self._start()

    def _on_key_press(self, key):
        if key == keyboard.Key.esc:
            if self.task_controller.is_running():
                self._stop()
    
    def _on_text_changed(self):
        if self.task_controller.is_running():
            new_text = self.txt_msg.toPlainText().strip()
            self.task_controller.update_runtime_content(new_text)

    def _on_files_changed(self):
        if self.task_controller.is_running():
            files = [self.list_images.item(i).text() for i in range(self.list_images.count())]
            self.task_controller.update_runtime_files(files)

    def _on_params_changed(self):
        current_interval = self.spin_interval.value()
        
        if current_interval > 1.0 and not self.chk_human_sim.isChecked():
            self.chk_human_sim.setChecked(True)
            self._log("💡 检测到间隔 > 1.0s，智能开启【真人拟态】")

        if self.task_controller.is_running():
            new_count = self.spin_count.value()
            self.task_controller.update_runtime_params(new_count, current_interval)

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
            has_batch = bool(self.file_handler.target_list)
            
            if not has_batch and not has_msg and not has_files:
                msg_box = QMessageBox(dialog)
                msg_box.setIcon(QMessageBox.Icon.Warning)
                msg_box.setWindowTitle("拒绝锁定")
                msg_box.setText("❌ 请先输入发送内容或拖入文件！\n\n空内容无法启动定时任务。")
                msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg_box.setStyleSheet("""
                    QMessageBox { background-color: #2b2b2b; }
                    QLabel { color: #E0E0E0; font-size: 13px; }
                    QPushButton { background-color: #424242; color: #FFFFFF; border-radius: 6px; border: none; font-size: 13px; font-weight: bold; padding: 8px 20px; min-width: 80px; }
                    QPushButton:hover { background-color: #616161; }
                """)
                msg_box.exec()
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
            
            if self.task_controller.is_running():
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
            self.task_controller.on_clipboard_set_done()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
    
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if not files: return
        
        txt_files, media_files = FileHandler.filter_files(files)
        
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
        success, message, count = self.file_handler.load_target_list(path)
        
        if not success:
            self._show_message_box(QMessageBox.Icon.Warning, "提示", message)
            return
        
        self.txt_file_path.setText(os.path.basename(path))
        
        info_text = f"✅ 批量模式: {count} 人 (专属: {message.split('专属: ')[1] if '专属: ' in message else 0})"
        self.lbl_target_info.setText(info_text)
        self.lbl_target_info.setStyleSheet("color: #81c784")
        self.input_single_name.setEnabled(False)
        self._log(f"📂 {message}")
    
    def _reset_mode(self):
        self.file_handler.reset()
        self.txt_file_path.clear()
        self.input_single_name.setEnabled(True)
        self.lbl_target_info.setText("模式: 单人手动")
        self.lbl_target_info.setStyleSheet("color: #aaa")
        self._on_files_changed() 
        self._log("🔄 状态已重置")

    def _log(self, msg):
        t = time.strftime("%H:%M:%S")
        log_line = f"[{t}] {msg}"
        self.log_buffer.append(log_line)
        self.txt_log.append(log_line)
        sb = self.txt_log.verticalScrollBar()
        sb.setValue(sb.maximum())
    
    def _export_log(self):
        """导出日志到文件"""
        from PyQt6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getSaveFileName(
            self, "导出日志", 
            f"wechat_pro_log_{time.strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt)"
        )
        if path:
            if self.log_buffer.export_to_file(path):
                self._log(f"📄 日志已导出: {path}")
            else:
                self._log("⚠️ 日志导出失败")
    
    def _clear_log(self):
        """清空日志"""
        self.log_buffer.clear()
        self.txt_log.clear()
        self._log("🗑️ 日志已清空")

    def _start(self):
        self._save_settings()
        self.history_manager.clear_dedupe_cache()
        
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
        
        if self.file_handler.target_list:
            targets = self.file_handler.target_list
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
                    self._show_message_box(QMessageBox.Icon.Warning, "拒绝执行", "所有目标均无内容！")
                    self.btn_run.setEnabled(True)
                    self.btn_stop.setEnabled(False)
                    return
                else:
                    reply = self._show_message_box(
                        QMessageBox.Icon.Question, 
                        "预警", 
                        f"有 {missing_count} 人内容为空，是否跳过？", 
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.No:
                        self.btn_run.setEnabled(True)
                        self.btn_stop.setEnabled(False)
                        return
        else:
            if not msg and not global_files:
                self._show_message_box(QMessageBox.Icon.Warning, "提示", "请输入文字或拖入文件")
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
        
        if self.task_controller.is_running():
            self._log("⚠️ 正在强制覆盖旧任务...")
        
        self.task_controller.start(
            config,
            on_log=self._log,
            on_progress=self.update_progress,
            on_finished=self.on_finished,
            on_clipboard=self.set_clipboard_files,
            on_countdown=self._update_countdown_display,
            history_manager=self.history_manager
        )

    def _stop(self):
        if self.task_controller.is_running():
            self.task_controller.stop()
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
        self._update_stats()
        
        if self.chk_auto_minimize.isChecked():
            QTimer.singleShot(500, self._perform_minimize_logic)
        
        self.target_datetime = None 
        if hasattr(self, 'lbl_schedule_time'): self.lbl_schedule_time.hide()

    def _perform_minimize_logic(self):
        try:
            self._log("✨ 自动归位已触发")
            self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized | Qt.WindowState.WindowActive)
            self.showNormal()
            self.activateWindow()
            self.raise_()
        except Exception as e:
            self._log(f"⚠️ 归位微调: {e}")

    def closeEvent(self, event):
        self.settings_manager.save_geometry(self.saveGeometry())
        self._save_settings() 
        if self.listener: self.listener.stop()
        self.task_controller.stop()
        event.accept()

if __name__ == "__main__":
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    win = WeChatProUI()
    win.show()
    sys.exit(app.exec())
