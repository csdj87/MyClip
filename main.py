import sys
import os
import sqlite3
import time
import re
import ctypes
import winreg
from ctypes import wintypes
import datetime

# ==========================================
# 依赖库检查
# ==========================================
try:
    import win32gui
    import win32con
    import win32api
    import win32process
    import win32clipboard
except ImportError:
    print("错误：缺少 pywin32 库。请运行 'pip install pywin32'")
    pass

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QListWidgetItem, QLineEdit, QStackedWidget,
                             QSystemTrayIcon, QMenu, QFrame, QGraphicsDropShadowEffect,
                             QFileIconProvider, QAbstractItemView, QSpinBox, 
                             QKeySequenceEdit, QSizePolicy, QGroupBox, QCheckBox)
from PyQt6.QtCore import (Qt, QPoint, QTimer, QSettings, QBuffer, QIODevice, 
                          QFileInfo, pyqtSignal, QSize, QThread, QRect, 
                          QPropertyAnimation, QEasingCurve, QEvent, QMimeData, QUrl, QObject)
from PyQt6.QtGui import (QIcon, QColor, QPixmap, QImage, QKeySequence, 
                         QCursor, QAction, QFont, QFontMetrics, QPainter)

# ==========================================
# 0. 路径与工具函数
# ==========================================
def get_asset_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

def get_data_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

def create_tinted_icon(svg_path, color_hex):
    if not os.path.exists(svg_path):
        return QIcon()
    src_pixmap = QPixmap(svg_path)
    if src_pixmap.isNull():
        return QIcon()
    tgt_pixmap = QPixmap(src_pixmap.size())
    tgt_pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(tgt_pixmap)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.drawPixmap(0, 0, src_pixmap)
    painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceIn)
    painter.fillRect(tgt_pixmap.rect(), QColor(color_hex))
    painter.end()
    return QIcon(tgt_pixmap)

# ==========================================
# 1. 样式表
# ==========================================
DARK_STYLE = """
QWidget { font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif; color: #e0e0e0; background-color: #1e1e1e; font-size: 13px; }
QScrollBar:vertical { border: none; background: #2b2b2b; width: 8px; margin: 0px; }
QScrollBar::handle:vertical { background: #555; min-height: 20px; border-radius: 4px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { background: none; }
#CentralWidget { background-color: #1e1e1e; border: 1px solid #444; border-radius: 8px; }
#TopBtn, #CloseBtn { background: transparent; border: none; border-radius: 4px; }
#TopBtn:hover { background-color: rgba(255, 255, 255, 30); }
#TopBtn:checked { background-color: rgba(255, 255, 255, 50); }
#CloseBtn:hover { background-color: #e81123; }
QLineEdit { background-color: #2b2b2b; color: white; border: 1px solid #555; border-radius: 4px; padding: 4px; }
QLineEdit:focus { border: 1px solid #0078d7; }
QListWidget { background: transparent; border: none; outline: none; }
#ItemContainer { background-color: #2b2b2b; border: 1px solid #444; }
#ItemContainer:hover { border: 1px solid transparent; background-color: #333; }
QGroupBox { border: 1px solid #555; border-radius: 5px; margin-top: 10px; font-weight: bold; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 3px; color: #aaa; }
QSpinBox, QKeySequenceEdit { background-color: #333; color: white; border: 1px solid #555; border-radius: 3px; padding: 3px; }
QCheckBox { spacing: 5px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #555; border-radius: 3px; background: #333; }
QCheckBox::indicator:checked { background-color: #0078d7; border-color: #0078d7; }
QPushButton#ActionBtn { background-color: #0078d7; color: white; border: none; padding: 6px 12px; border-radius: 4px; }
QPushButton#ActionBtn:hover { background-color: #0089f7; }
QMenu { background-color: #2b2b2b; color: white; border: 1px solid #555; padding: 5px; }
QMenu::item { padding: 5px 20px; border-radius: 4px; }
QMenu::item:selected { background-color: #444; }
"""

LIGHT_STYLE = """
QWidget { font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif; color: #333333; background-color: #f5f5f5; font-size: 13px; }
QScrollBar:vertical { border: none; background: #e0e0e0; width: 8px; margin: 0px; }
QScrollBar::handle:vertical { background: #aaa; min-height: 20px; border-radius: 4px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { background: none; }
#CentralWidget { background-color: #f5f5f5; border: 1px solid #aaaaaa; border-radius: 8px; }
#TopBtn, #CloseBtn { background: transparent; border: none; border-radius: 4px; }
#TopBtn:hover { background-color: rgba(0, 0, 0, 20); }
#TopBtn:checked { background-color: rgba(0, 0, 0, 40); }
#CloseBtn:hover { background-color: #e81123; }
QLineEdit { background-color: #ffffff; color: #333; border: 1px solid #ccc; border-radius: 4px; padding: 4px; }
QLineEdit:focus { border: 1px solid #0078d7; }
QListWidget { background: transparent; border: none; outline: none; }
#ItemContainer { background-color: #ffffff; border: 1px solid #ccc; }
#ItemContainer:hover { border: 1px solid transparent; background-color: #fff; }
QGroupBox { border: 1px solid #ccc; border-radius: 5px; margin-top: 10px; font-weight: bold; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 3px; color: #333; }
QSpinBox, QKeySequenceEdit { background-color: white; color: #333; border: 1px solid #ccc; border-radius: 3px; padding: 3px; }
QCheckBox { spacing: 5px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #ccc; border-radius: 3px; background: white; }
QCheckBox::indicator:checked { background-color: #0078d7; border-color: #0078d7; }
QPushButton#ActionBtn { background-color: #0078d7; color: white; border: none; padding: 6px 12px; border-radius: 4px; }
QPushButton#ActionBtn:hover { background-color: #006cc1; }
QMenu { background-color: #ffffff; color: #333; border: 1px solid #ccc; padding: 5px; }
QMenu::item { padding: 5px 20px; border-radius: 4px; }
QMenu::item:selected { background-color: #e5e5e5; }
"""

# ==========================================
# 2. 全局信号与数据库
# ==========================================
class GlobalSignals(QObject):
    database_changed = pyqtSignal(str) 

global_signals = GlobalSignals()

class DBManager:
    def __init__(self, db_name="clipboard.db"):
        self.db_path = get_data_path(db_name)
        self.conn = None
        self.init_db()

    def get_conn(self):
        if not self.conn:
            self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
            self.conn.row_factory = sqlite3.Row
        return self.conn

    def init_db(self):
        conn = self.get_conn()
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type TEXT, content_text TEXT, content_html TEXT, 
                content_blob BLOB, search_text TEXT, hash_val TEXT,
                is_pinned INTEGER DEFAULT 0, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_search ON history(search_text)')
        conn.commit()

    # emit_signal: 是否发送信号刷新UI。批量操作或后台静默更新时建议设为 False
    def add_item(self, type_, text=None, html=None, blob=None, filepath=None, hash_val=None, emit_signal=True):
        conn = self.get_conn()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM history WHERE hash_val = ?", (hash_val,))
        row = cursor.fetchone()
        
        updated = False
        if row:
            cursor.execute("UPDATE history SET created_at = CURRENT_TIMESTAMP WHERE id = ?", (row['id'],))
            updated = True
            row_id = row['id']
        else:
            search_text = text if text else (filepath if filepath else "")
            cursor.execute('''
                INSERT INTO history (type, content_text, content_html, content_blob, search_text, hash_val)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (type_, text or filepath, html, blob, search_text, hash_val))
            row_id = cursor.lastrowid
            updated = False
            
        conn.commit()
        if emit_signal:
            global_signals.database_changed.emit('history')
        return row_id, updated

    def get_items(self, limit=50, search_query=None, only_pinned=False):
        conn = self.get_conn()
        sql = "SELECT * FROM history WHERE 1=1"
        params = []
        if only_pinned: sql += " AND is_pinned = 1"
        if search_query:
            sql += " AND search_text LIKE ?"
            params.append(f"%{search_query}%")
        sql += " ORDER BY created_at DESC LIMIT ?"
        params.append(limit)
        return conn.execute(sql, params).fetchall()

    def set_pinned(self, item_id, is_pinned):
        self.get_conn().execute("UPDATE history SET is_pinned = ? WHERE id = ?", (1 if is_pinned else 0, item_id)).connection.commit()
        global_signals.database_changed.emit('all')

    def delete_item(self, item_id):
        self.get_conn().execute("DELETE FROM history WHERE id = ?", (item_id,)).connection.commit()
        global_signals.database_changed.emit('all')
        
    def clear_all(self, include_pinned=False):
        sql = "DELETE FROM history" if include_pinned else "DELETE FROM history WHERE is_pinned = 0"
        self.get_conn().execute(sql).connection.commit()
        global_signals.database_changed.emit('all')

class NativeHotkeyThread(QThread):
    sig_trigger = pyqtSignal()
    def __init__(self, hotkey_str):
        super().__init__()
        self.hotkey_str = hotkey_str
        self.running = True
        self.thread_id = None
        
    def run(self):
        if not sys.platform.startswith('win'): return
        try:
            user32 = ctypes.windll.user32
            self.thread_id = ctypes.windll.kernel32.GetCurrentThreadId()
            MODS = {'CTRL': 2, 'SHIFT': 4, 'ALT': 1, 'WIN': 8}
            parts = self.hotkey_str.upper().split('+')
            mod, vk = 0, 0
            for p in parts:
                p = p.strip()
                if p in MODS: mod |= MODS[p]
                elif len(p)==1: vk = ord(p)
                elif p.startswith('F') and p[1:].isdigit(): vk = 0x70 + int(p[1:]) - 1
            
            if vk and user32.RegisterHotKey(None, 1, mod, vk):
                msg = wintypes.MSG()
                while self.running:
                    res = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                    if res == 0 or res == -1: break
                    if msg.message == win32con.WM_HOTKEY: self.sig_trigger.emit()
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
                user32.UnregisterHotKey(None, 1)
        except Exception as e:
            print(f"Hotkey Error: {e}")

    def stop(self):
        self.running = False
        if self.thread_id and sys.platform.startswith('win'):
            try:
                ctypes.windll.user32.PostThreadMessageW(self.thread_id, win32con.WM_QUIT, 0, 0)
            except: pass
        self.wait()

# ==========================================
# 3. 自定义列表项
# ==========================================
class ClipboardItemWidget(QWidget):
    MAX_HEIGHT = 120

    def __init__(self, data_row, search_keyword="", parent=None):
        super().__init__(parent)
        self.data = data_row
        self.search_keyword = search_keyword
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4) 
        layout.setSpacing(0)

        self.container = QFrame()
        self.container.setObjectName("ItemContainer")
        self.container.setMaximumHeight(self.MAX_HEIGHT)
        
        type_color = self.get_color_by_type()
        self.container.setStyleSheet(f"""
            #ItemContainer:hover {{
                border: 1px solid {type_color};
            }}
        """)

        self.con_layout = QHBoxLayout(self.container)
        self.con_layout.setContentsMargins(0, 0, 0, 0)
        self.con_layout.setSpacing(0)

        self.color_strip = QLabel()
        self.color_strip.setFixedWidth(4)
        self.color_strip.setStyleSheet(f"background-color: {self.get_color_by_type()};")
        self.color_strip.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        self.con_layout.addWidget(self.color_strip)

        self.content_lbl = QLabel()
        self.content_lbl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.content_lbl.setStyleSheet("border: none; background: transparent; padding: 4px 4px;") 
        self.con_layout.addWidget(self.content_lbl, 1)
        
        self.render_content()
        layout.addWidget(self.container)

    def get_color_by_type(self):
        t = self.data['type']
        if t == 'text': return "#34A853"
        if t == 'html': return "#4285F4"
        if t == 'image': return "#EA4335"
        if t == 'file': return "#FBBC05"
        return "#9AA0A6"

    def render_content(self):
        t = self.data['type']
        blob = self.data['content_blob']
        if blob:
            img = QImage.fromData(blob)
            if not img.isNull():
                pix = QPixmap.fromImage(img)
                target_h = self.MAX_HEIGHT - 10 
                if pix.height() > target_h:
                    pix = pix.scaledToHeight(target_h, Qt.TransformationMode.SmoothTransformation)
                if pix.width() > 320: 
                    pix = pix.scaledToWidth(320, Qt.TransformationMode.SmoothTransformation)
                
                self.content_lbl.setPixmap(pix)
                self.content_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft)
                return

        if t == 'file':
            path = self.data['content_text']
            file_info = QFileInfo(path)
            icon_provider = QFileIconProvider()
            icon = icon_provider.icon(file_info)
            pix = icon.pixmap(28, 28)
            filename = os.path.basename(path)
            if self.search_keyword: filename = self.highlight_text(filename)
            
            icon_lbl = QLabel()
            icon_lbl.setPixmap(pix)
            icon_lbl.setFixedSize(32, 32)
            icon_lbl.setStyleSheet("background: transparent; border: none; padding-left: 4px;")
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignVCenter)
            self.con_layout.insertWidget(1, icon_lbl)
            
            self.content_lbl.setText(filename)
            self.content_lbl.setTextFormat(Qt.TextFormat.RichText)
            self.content_lbl.setWordWrap(True)
            self.content_lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            return

        text = self.data['content_text'] or ""
        display_text = text[:800] 
        if self.search_keyword:
            display_text = self.highlight_text(display_text).replace("\n", "<br>")
            self.content_lbl.setTextFormat(Qt.TextFormat.RichText)
        else:
            self.content_lbl.setTextFormat(Qt.TextFormat.PlainText)
        self.content_lbl.setText(display_text)
        self.content_lbl.setWordWrap(True)
        self.content_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

    def highlight_text(self, text):
        return re.sub(f"({re.escape(self.search_keyword)})", 
                      r"<span style='background-color: #FFEB3B; color: black;'>\1</span>", 
                      text, flags=re.IGNORECASE)

    def get_required_height(self, width):
        self.container.setFixedWidth(width - 10) 
        self.container.layout().activate()
        sh = self.sizeHint()
        self.container.setMinimumWidth(0)
        self.container.setMaximumWidth(16777215)
        return min(sh.height(), self.MAX_HEIGHT)

# ==========================================
# 4. 主窗口
# ==========================================
class ClipboardManager(QMainWindow):
    REG_APP_NAME = "MyClipboardTool"

    def __init__(self):
        super().__init__()
        self.db = DBManager()
        self.settings = QSettings("MyTools", "ClipboardManager")
        
        self.setMinimumSize(350, 450)
        self.resize(350, 450)
        self.SHADOW_MARGIN = 10
        self._margin = self.SHADOW_MARGIN + 5
        
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | 
                           Qt.WindowType.WindowSystemMenuHint | 
                           Qt.WindowType.Tool | 
                           Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setMouseTracking(True)
        
        self.is_light_theme = True
        self.last_switch_key = ""
        
        self.tray = QSystemTrayIcon(self)
        self.init_ui_elements()
        self.init_tray()
        self.init_clipboard_monitor()
        self.update_hotkey()
        
        self.init_startup_theme()
        
        self.theme_timer = QTimer(self)
        self.theme_timer.timeout.connect(self.check_scheduled_theme_switch)
        self.theme_timer.start(60000) 
        
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self.refresh_list)

        self.refresh_list()
        global_signals.database_changed.connect(self.on_database_changed)

    def init_startup_theme(self):
        hour = datetime.datetime.now().hour
        should_be_light = 6 <= hour < 18
        self.set_theme(should_be_light)

    def check_scheduled_theme_switch(self):
        now = datetime.datetime.now()
        if now.minute != 0: return
        current_key = f"{now.year}-{now.month}-{now.day}-{now.hour}"
        if current_key == self.last_switch_key: return

        if now.hour == 6:
            self.set_theme(True)
            self.last_switch_key = current_key
        elif now.hour == 18:
            self.set_theme(False)
            self.last_switch_key = current_key

    def set_theme(self, is_light):
        self.is_light_theme = is_light
        self.setStyleSheet(LIGHT_STYLE if is_light else DARK_STYLE)
        icon_color = "#333333" if is_light else "#e0e0e0"
        self.update_icons(icon_color)

    def update_icons(self, color_hex):
        self.btn_search.setIcon(create_tinted_icon(get_asset_path("search.svg"), color_hex))
        self.btn_pin.setIcon(create_tinted_icon(get_asset_path("pin.svg"), color_hex))
        self.btn_settings.setIcon(create_tinted_icon(get_asset_path("set.svg"), color_hex))
        close_color = "#555555" if self.is_light_theme else "#cccccc" 
        self.btn_close.setIcon(create_tinted_icon(get_asset_path("x.svg"), close_color))
        
        tray_icon = create_tinted_icon(get_asset_path("myclip.svg"), color_hex)
        if not tray_icon.isNull():
            self.tray.setIcon(tray_icon)
        else:
            self.tray.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_DriveFDIcon))

    def init_ui_elements(self):
        self.main_wrapper = QWidget()
        self.main_wrapper.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCentralWidget(self.main_wrapper)
        
        self.wrapper_layout = QVBoxLayout(self.main_wrapper)
        self.wrapper_layout.setContentsMargins(self.SHADOW_MARGIN, self.SHADOW_MARGIN, self.SHADOW_MARGIN, self.SHADOW_MARGIN)
        
        self.central_widget = QWidget()
        self.central_widget.setObjectName("CentralWidget")
        self.wrapper_layout.addWidget(self.central_widget)
        
        self.shadow_effect = QGraphicsDropShadowEffect(self)
        self.shadow_effect.setBlurRadius(15)
        self.shadow_effect.setXOffset(0)
        self.shadow_effect.setYOffset(0)
        self.shadow_effect.setColor(QColor(0, 0, 0, 150))
        self.central_widget.setGraphicsEffect(self.shadow_effect)

        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(5, 5, 5, 5)
        self.main_layout.setSpacing(5)

        self.title_container = QWidget()
        self.title_container.setFixedHeight(35)
        self.title_layout = QHBoxLayout(self.title_container)
        self.title_layout.setContentsMargins(0, 0, 0, 0)
        self.title_layout.setSpacing(5)

        self.btn_search = QPushButton("")
        self.btn_search.setIconSize(QSize(18, 18))
        self.btn_pin = QPushButton("")
        self.btn_pin.setIconSize(QSize(18, 18))
        self.btn_settings = QPushButton("")
        self.btn_settings.setIconSize(QSize(18, 18))
        
        for btn in [self.btn_search, self.btn_pin, self.btn_settings]:
            btn.setObjectName("TopBtn")
            btn.setCheckable(True)
            btn.setAutoExclusive(True)
            btn.setFixedSize(32, 30)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            self.title_layout.addWidget(btn)

        self.btn_search.clicked.connect(self.on_search_btn_clicked)
        self.btn_pin.clicked.connect(lambda: self.switch_tab(1))
        self.btn_settings.clicked.connect(lambda: self.switch_tab(2))

        self.title_layout.addStretch()

        self.btn_close = QPushButton("")
        self.btn_close.setIconSize(QSize(24, 24))
        self.btn_close.setObjectName("CloseBtn")
        self.btn_close.setFixedSize(32, 30)
        self.btn_close.clicked.connect(self.hide)
        self.title_layout.addWidget(self.btn_close)

        self.main_layout.addWidget(self.title_container)

        self.stack = QStackedWidget()
        
        # List Page
        self.page_list_container = QWidget()
        layout_list = QVBoxLayout(self.page_list_container)
        layout_list.setContentsMargins(0, 0, 0, 0)
        layout_list.setSpacing(5)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索剪贴板...")
        self.search_input.textChanged.connect(self.on_search_text_changed)
        self.search_input.setVisible(False)
        layout_list.addWidget(self.search_input)
        
        self.list_widget = QListWidget()
        self.list_widget.setSpacing(0) 
        self.list_widget.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.list_widget.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.list_widget.setUniformItemSizes(False)
        self.list_widget.itemDoubleClicked.connect(self.on_item_double_click)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # 安装事件过滤器以支持回车键粘贴
        self.list_widget.installEventFilter(self)
        
        layout_list.addWidget(self.list_widget)
        self.stack.addWidget(self.page_list_container)

        # Pin Page
        self.page_pin = QWidget()
        layout_pin = QVBoxLayout(self.page_pin)
        layout_pin.setContentsMargins(0, 0, 0, 0)
        self.list_pin = QListWidget()
        self.list_pin.setSpacing(0)
        self.list_pin.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.list_pin.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.list_pin.itemDoubleClicked.connect(self.on_item_double_click)
        self.list_pin.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_pin.customContextMenuRequested.connect(lambda p: self.show_context_menu(p, is_pinned_page=True))
        self.list_pin.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # 安装事件过滤器以支持回车键粘贴
        self.list_pin.installEventFilter(self)
        
        layout_pin.addWidget(self.list_pin)
        self.stack.addWidget(self.page_pin)

        # Settings Page
        self.page_settings = QWidget()
        layout_set = QVBoxLayout(self.page_settings)
        layout_set.setContentsMargins(10, 5, 10, 5)
        layout_set.setSpacing(15)

        grp_general = QGroupBox("常规设置")
        layout_gen = QVBoxLayout(grp_general)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("最大历史条数(10-500):"))
        self.spin_max_items = QSpinBox()
        self.spin_max_items.setRange(10, 500)
        self.spin_max_items.setValue(self.settings.value("max_items", 100, int))
        self.spin_max_items.valueChanged.connect(lambda v: self.settings.setValue("max_items", v))
        row1.addWidget(self.spin_max_items)
        layout_gen.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("显示热键:"))
        self.key_edit = QKeySequenceEdit()
        self.key_edit.setKeySequence(QKeySequence(self.settings.value("hotkey", "Ctrl+Shift+V")))
        self.key_edit.editingFinished.connect(self.on_hotkey_changed)
        row2.addWidget(self.key_edit)
        layout_gen.addLayout(row2)

        self.check_autostart = QCheckBox("开机自动启动")
        self.check_autostart.setChecked(self.is_autostart_enabled())
        self.check_autostart.toggled.connect(self.set_autostart)
        layout_gen.addWidget(self.check_autostart)
        layout_set.addWidget(grp_general)

        grp_action = QGroupBox("操作")
        layout_act = QVBoxLayout(grp_action)
        btn_theme = QPushButton("手动切换主题")
        btn_theme.setObjectName("ActionBtn")
        btn_theme.clicked.connect(lambda: self.set_theme(not self.is_light_theme))
        layout_act.addWidget(btn_theme)
        
        btn_clear_db = QPushButton("清空所有未固定历史")
        btn_clear_db.setObjectName("ActionBtn")
        btn_clear_db.setStyleSheet("background-color: #d9534f;")
        btn_clear_db.clicked.connect(self.clear_database)
        layout_act.addWidget(btn_clear_db)
        layout_set.addWidget(grp_action)
        layout_set.addStretch()
        self.stack.addWidget(self.page_settings)

        self.main_layout.addWidget(self.stack)
        self.btn_search.setChecked(True)
        self.stack.setCurrentIndex(0)

    # ==========================================
    # [NEW] 事件过滤器：处理回车键粘贴
    # ==========================================
    def eventFilter(self, source, event):
        if event.type() == QEvent.Type.KeyPress and source in [self.list_widget, self.list_pin]:
            if event.key() in [Qt.Key.Key_Return, Qt.Key.Key_Enter]:
                item = source.currentItem()
                if item:
                    self.on_item_double_click(item)
                    return True
        return super().eventFilter(source, event)

    def is_autostart_enabled(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, self.REG_APP_NAME)
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False

    def set_autostart(self, checked):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE)
            if checked:
                if getattr(sys, 'frozen', False):
                    app_path = f'"{sys.executable}"'
                else:
                    app_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                winreg.SetValueEx(key, self.REG_APP_NAME, 0, winreg.REG_SZ, app_path)
            else:
                try: winreg.DeleteValue(key, self.REG_APP_NAME)
                except WindowsError: pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"AutoStart Registry Error: {e}")

    def _get_edge(self, pos: QPoint):
        m = self._margin
        w, h = self.width(), self.height()
        x, y = pos.x(), pos.y()
        edge = 0
        if x <= m: edge |= Qt.Edge.LeftEdge.value
        if x >= w - m: edge |= Qt.Edge.RightEdge.value
        if y <= m: edge |= Qt.Edge.TopEdge.value
        if y >= h - m: edge |= Qt.Edge.BottomEdge.value
        return edge

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            edges = self._get_edge(event.position().toPoint())
            if edges:
                try: self.windowHandle().startSystemResize(Qt.Edge(edges))
                except: pass
            else:
                if event.position().y() < (45 + self.SHADOW_MARGIN): 
                    self.windowHandle().startSystemMove()

    def mouseMoveEvent(self, event):
        edges = self._get_edge(event.position().toPoint())
        if edges: self.setCursor(Qt.CursorShape.SizeAllCursor)
        else: self.setCursor(Qt.CursorShape.ArrowCursor)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(10, self.update_all_list_item_sizes)

    def changeEvent(self, event):
        if event.type() == QEvent.Type.ActivationChange:
            if self.isActiveWindow(): 
                self.shadow_effect.setColor(QColor(0, 0, 0, 150))
            else: 
                self.shadow_effect.setColor(QColor(0, 0, 0, 60))
                self.hide()
        super().changeEvent(event)

    def on_search_btn_clicked(self):
        current_idx = self.stack.currentIndex()
        if current_idx == 0:
            is_vis = self.search_input.isVisible()
            self.search_input.setVisible(not is_vis)
            if not is_vis: self.search_input.setFocus()
            else: self.search_input.clear()
        else:
            self.switch_tab(0)
            self.search_input.setVisible(True)
            self.search_input.setFocus()
        self.btn_search.setChecked(True)

    def switch_tab(self, index):
        self.stack.setCurrentIndex(index)
        self.btn_search.setChecked(index == 0)
        self.btn_pin.setChecked(index == 1)
        self.btn_settings.setChecked(index == 2)
        if index == 0: 
            QTimer.singleShot(0, self.list_widget.scrollToTop)
        elif index == 1: 
            self.refresh_pinned_list()
        QTimer.singleShot(50, self.update_all_list_item_sizes)

    def update_all_list_item_sizes(self):
        current_list = None
        if self.stack.currentIndex() == 0:
            current_list = self.list_widget
        elif self.stack.currentIndex() == 1:
            current_list = self.list_pin
        
        if not current_list or not current_list.isVisible(): return
        viewport_width = current_list.viewport().width()
        if viewport_width <= 0: return

        for i in range(current_list.count()):
            item = current_list.item(i)
            widget = current_list.itemWidget(item)
            if widget:
                new_height = widget.get_required_height(viewport_width)
                current_size = item.sizeHint()
                if current_size.height() != new_height or current_size.width() != viewport_width:
                    item.setSizeHint(QSize(viewport_width, new_height))
        current_list.doItemsLayout()

    def on_database_changed(self, change_type):
        if change_type in ['history', 'all']:
            self.refresh_list()
        if change_type in ['pinned', 'all']:
            self.refresh_pinned_list()

    def refresh_list(self):
        query = self.search_input.text()
        items = self.db.get_items(limit=self.spin_max_items.value(), search_query=query, only_pinned=False)
        self.populate_list(self.list_widget, items, query)

    def refresh_pinned_list(self):
        items = self.db.get_items(limit=1000, only_pinned=True)
        self.populate_list(self.list_pin, items, "")

    def populate_list(self, list_widget, items, query):
        list_widget.setUpdatesEnabled(False)
        list_widget.clear()
        viewport_width = list_widget.viewport().width()
        if viewport_width <= 0: viewport_width = 300

        for row in items:
            item = QListWidgetItem(list_widget)
            widget = ClipboardItemWidget(dict(row), query)
            list_widget.setItemWidget(item, widget)
            height = widget.get_required_height(viewport_width)
            item.setSizeHint(QSize(viewport_width, height))
            
        list_widget.setUpdatesEnabled(True)

    def on_search_text_changed(self, text):
        self.search_timer.start()

    def init_clipboard_monitor(self):
        self.clipboard = QApplication.clipboard()
        self.clipboard.dataChanged.connect(self.on_clipboard_change)
        self.is_pasting = False

    def on_clipboard_change(self):
        if self.is_pasting: return
        mime = self.clipboard.mimeData()
        
        db_type, text, html, blob, path = 'text', None, None, None, None
        h_val = ""
        
        try:
            if mime.hasUrls() and mime.urls() and mime.urls()[0].isLocalFile():
                db_type, path = 'file', mime.urls()[0].toLocalFile()
                h_val = str(hash(path))
            elif mime.hasText():
                db_type = 'text'
                text = mime.text()
                if not text: return
                if mime.hasHtml(): 
                    db_type = 'html'
                    html = mime.html()
                h_val = str(hash(text))
                if mime.hasImage():
                    img = mime.imageData()
                    if img:
                        ba = QBuffer()
                        ba.open(QIODevice.OpenModeFlag.WriteOnly)
                        img.save(ba, "PNG")
                        blob = ba.data().data()
            elif mime.hasImage():
                db_type = 'image'
                img = mime.imageData()
                if img:
                    ba = QBuffer()
                    ba.open(QIODevice.OpenModeFlag.WriteOnly)
                    img.save(ba, "PNG")
                    blob = ba.data().data()
                    h_val = str(hash(blob))
            else: return

            if h_val:
                self.db.add_item(db_type, text, html, blob, path, h_val, emit_signal=True)
        except Exception: pass

    def on_item_double_click(self, item):
        widget = self.list_widget.itemWidget(item) or self.list_pin.itemWidget(item)
        if widget:
            self.do_paste(widget.data)

    # ==========================================
    # [MODIFIED] 增加 as_plain_text 参数
    # ==========================================
    def do_paste(self, row, as_plain_text=False):
        # ========================================================
        # 极速粘贴逻辑 (Zero-Latency Paste)
        # ========================================================
        
        # 1. 立即隐藏窗口
        self.hide()
        QApplication.processEvents()

        # 2. 设置剪贴板
        self.is_pasting = True
        new_mime = QMimeData()
        
        # 如果强制纯文本，或者原本就是文本类型且没有其他特殊格式
        if as_plain_text:
             new_mime.setText(row['content_text'])
        else:
            if row['type'] == 'text': 
                new_mime.setText(row['content_text'])
            elif row['type'] == 'html': 
                new_mime.setHtml(row['content_html'])
                new_mime.setText(row['content_text'])
            elif row['type'] == 'image': 
                if row['content_blob']:
                    img = QImage.fromData(row['content_blob'])
                    new_mime.setImageData(img)
            elif row['type'] == 'file':
                url = QUrl.fromLocalFile(row['content_text'])
                new_mime.setUrls([url])

        QApplication.clipboard().setMimeData(new_mime)
        
        # 3. 模拟粘贴按键 (Ctrl + V)
        time.sleep(0.05) 
        if sys.platform.startswith('win'):
            try:
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                win32api.keybd_event(ord('V'), 0, 0, 0)
                win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            except Exception as e: 
                print(f"Paste Error: {e}")

        # 4. 延迟更新数据库与界面，确保粘贴动作流畅
        QTimer.singleShot(500, lambda: self.deferred_update_after_paste(row))

    def deferred_update_after_paste(self, row):
        """延迟执行的后台任务：更新数据库顺序并重绘界面"""
        self.db.add_item(
            type_=row['type'],
            text=row['content_text'] if row['type'] != 'file' else None,
            html=row['content_html'],
            blob=row['content_blob'],
            filepath=row['content_text'] if row['type'] == 'file' else None,
            hash_val=row['hash_val'],
            emit_signal=False 
        )
        self.refresh_list()
        self.is_pasting = False

    def show_context_menu(self, position, is_pinned_page=False):
        s_list = self.list_pin if is_pinned_page else self.list_widget
        item = s_list.itemAt(position)
        if not item: return
        widget = s_list.itemWidget(item)
        if not widget: return
        
        row_id = widget.data['id']
        
        menu = QMenu()
        
        # ==========================================
        # [NEW] 增加“粘贴为纯文本”选项
        # ==========================================
        if widget.data['type'] == 'html':
            act_plain = menu.addAction("粘贴为纯文本")
        else:
            act_plain = None

        act_pin = menu.addAction("取消固定" if is_pinned_page else "固定")
        act_del = menu.addAction("删除")
        menu.addSeparator()
        act_clear = menu.addAction("清空所有")
        action = menu.exec(s_list.mapToGlobal(position))
        
        if act_plain and action == act_plain:
            self.do_paste(widget.data, as_plain_text=True)
        elif action == act_pin:
            self.db.set_pinned(row_id, not is_pinned_page)
        elif action == act_del:
            self.db.delete_item(row_id)
        elif action == act_clear:
            self.db.clear_all(False)

    def clear_database(self):
        self.db.clear_all(False)
        
    def on_hotkey_changed(self):
        seq = self.key_edit.keySequence().toString()
        self.settings.setValue("hotkey", seq)
        self.update_hotkey()

    def update_hotkey(self):
        hk = self.settings.value("hotkey", "Ctrl+Shift+V")
        if hasattr(self, 'hk_thread'): 
            self.hk_thread.stop()
        self.hk_thread = NativeHotkeyThread(hk)
        self.hk_thread.sig_trigger.connect(self.toggle_visible)
        self.hk_thread.start()

    def init_tray(self):
        menu = QMenu()
        act_toggle = menu.addAction("显示/隐藏")
        act_toggle.triggered.connect(self.toggle_visible)
        act_settings = menu.addAction("设置")
        act_settings.triggered.connect(self.open_settings_from_tray)
        menu.addSeparator()
        menu.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.toggle_visible()
            
    def open_settings_from_tray(self):
        self.show()
        self.switch_tab(2) 
        self.activateWindow()
        self.raise_()

    def toggle_visible(self):
        if self.isVisible(): 
            self.hide()
        else:
            cp = QCursor.pos()
            screen = QApplication.screenAt(cp).availableGeometry()
            x = min(cp.x(), screen.right() - self.width())
            y = min(cp.y(), screen.bottom() - self.height())
            x = max(screen.left(), x)
            y = max(screen.top(), y)
            self.move(x, y)
            
            if self.stack.currentIndex() != 0:
                self.switch_tab(0)
            
            if self.search_input.isVisible():
                self.search_input.clear()
                self.search_input.setVisible(False)
                self.list_widget.setFocus()
            
            self.list_widget.scrollToTop()
            if self.list_widget.count() > 0:
                self.list_widget.setCurrentRow(0)
                
            self.show()
            self.activateWindow()
            self.raise_()
            QTimer.singleShot(10, self.update_all_list_item_sizes) 
            self.list_widget.setFocus()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_SCALE_FACTOR"] = "1"
    win = ClipboardManager()
    sys.exit(app.exec())