import os
import sys
import time
import shutil
import sqlite3
import subprocess
import threading
import queue
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# 导入托盘与图标组件
try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# 导入 Windows 注册表组件（用于开机自启）
try:
    import winreg
    HAS_WINREG = True
except ImportError:
    HAS_WINREG = False

# 导入 Excel 导出组件
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ================= 核心配置区域 =================
APP_NAME = "FileSmartTracker"

# 1. 系统标准的软件专属数据存储目录: C:\Users\<用户名>\AppData\Roaming\FileSmartTracker
APPDATA_ROOT = os.environ.get("APPDATA") or os.path.expanduser("~")
DATA_DIR = os.path.join(APPDATA_ROOT, APP_NAME)
os.makedirs(DATA_DIR, exist_ok=True)

# 2. 获取当前 .exe 或脚本所在旧目录（仅用于自动迁移历史数据）
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_FILE = os.path.join(DATA_DIR, "file_index.db")
DEFAULT_EXCEL_FILE = os.path.join(DATA_DIR, "下载文件索引台账.xlsx")
DEFAULT_WATCH_DIR = os.path.expanduser("~/Desktop/File Collection")

# 自动迁移：如果新目录还没有数据库，而旧 .exe 旁边有历史数据库，自动复制过来保留全部记录
LEGACY_DB = os.path.join(EXE_DIR, "file_index.db")
if not os.path.exists(DB_FILE) and os.path.exists(LEGACY_DB):
    try:
        shutil.copy2(LEGACY_DB, DB_FILE)
    except Exception:
        pass
# ===============================================


def get_file_fingerprint(file_path):
    try:
        st = os.stat(file_path)
        return str(st.st_ino), st.st_size
    except Exception:
        return "", 0


class Database:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_and_migrate_db()

    def _init_and_migrate_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS file_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    download_time TEXT,
                    file_name TEXT,
                    source_app TEXT,
                    sender TEXT,
                    note TEXT,
                    file_path TEXT UNIQUE,
                    file_id TEXT,
                    file_size INTEGER
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS frequent_senders (
                    name TEXT PRIMARY KEY
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS frequent_apps (
                    name TEXT PRIMARY KEY
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)

            conn.execute("""
                DELETE FROM frequent_senders 
                WHERE name IN ('未知', '蒋工', 'Vincent', 'Kelvin', 'Dennis', 'Kay')
                AND name NOT IN (SELECT DISTINCT sender FROM file_records WHERE sender IS NOT NULL AND sender != '')
            """)

            conn.execute("""
                INSERT OR IGNORE INTO frequent_apps (name)
                SELECT DISTINCT source_app FROM file_records WHERE source_app IS NOT NULL AND source_app != ''
            """)
            conn.execute("""
                INSERT OR IGNORE INTO frequent_senders (name)
                SELECT DISTINCT sender FROM file_records WHERE sender IS NOT NULL AND sender != '' AND sender != '未知'
            """)

            for col, col_type in [("file_id", "TEXT"), ("file_size", "INTEGER")]:
                try:
                    conn.execute(f"ALTER TABLE file_records ADD COLUMN {col} {col_type}")
                except sqlite3.OperationalError:
                    pass

    def get_config(self, key, default_val=""):
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT value FROM settings WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else default_val

    def set_config(self, key, value):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))

    def get_apps(self):
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT name FROM frequent_apps ORDER BY rowid ASC")
            return [row[0] for row in cur.fetchall()]

    def add_app(self, name):
        if name and name.strip():
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("INSERT OR IGNORE INTO frequent_apps (name) VALUES (?)", (name.strip(),))

    def get_senders(self):
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT name FROM frequent_senders ORDER BY rowid ASC")
            return [row[0] for row in cur.fetchall()]

    def add_sender(self, name):
        if name and name.strip():
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("INSERT OR IGNORE INTO frequent_senders (name) VALUES (?)", (name.strip(),))

    def exists(self, file_path):
        norm_p = os.path.normpath(file_path)
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute(
                "SELECT 1 FROM file_records WHERE LOWER(file_path) = LOWER(?)",
                (norm_p,)
            )
            return cur.fetchone() is not None

    def is_inside_tracked_folder(self, file_path):
        """检查当前文件路径是否属于已按文件夹整包归档的子文件"""
        norm_p = os.path.normpath(file_path).lower()
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT file_path FROM file_records")
            for (rec_path,) in cur.fetchall():
                rec_norm = os.path.normpath(rec_path).lower()
                if norm_p != rec_norm and norm_p.startswith(rec_norm + os.sep):
                    return True
        return False

    def find_record_by_name(self, file_name):
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute(
                "SELECT id, file_path FROM file_records WHERE LOWER(file_name) = LOWER(?)",
                (file_name,)
            )
            return cur.fetchone()

    def update_path_and_time_by_id(self, record_id, new_path):
        norm_p = os.path.normpath(new_path)
        new_name = os.path.basename(norm_p)
        if os.path.isdir(norm_p):
            new_name = f"📁 {new_name}"
        file_id, file_size = get_file_fingerprint(norm_p)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                UPDATE file_records 
                SET file_path = ?, file_name = ?, file_id = ?, file_size = ? 
                WHERE id = ?
            """, (norm_p, new_name, file_id, file_size, record_id))

    def add_record(self, download_time, file_name, source_app, sender, note, file_path):
        norm_p = os.path.normpath(file_path)
        file_id, file_size = get_file_fingerprint(norm_p)
        self.add_app(source_app)
        self.add_sender(sender)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                INSERT OR REPLACE INTO file_records 
                (download_time, file_name, source_app, sender, note, file_path, file_id, file_size)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (download_time, file_name, source_app, sender, note, norm_p, file_id, file_size))

    def update_path_on_move(self, old_path, new_path):
        norm_old = os.path.normpath(old_path)
        norm_new = os.path.normpath(new_path)
        new_name = os.path.basename(norm_new)
        if os.path.isdir(norm_new):
            new_name = f"📁 {new_name}"
        file_id, file_size = get_file_fingerprint(norm_new)
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute(
                "UPDATE file_records SET file_path = ?, file_name = ?, file_id = ?, file_size = ? WHERE LOWER(file_path) = LOWER(?)",
                (norm_new, new_name, file_id, file_size, norm_old)
            )
            return cur.rowcount > 0

    def update_path_by_id(self, record_id, new_path):
        self.update_path_and_time_by_id(record_id, new_path)

    def get_all(self, keyword=""):
        with sqlite3.connect(self.db_path) as conn:
            if keyword:
                kw = f"%{keyword}%"
                cur = conn.execute("""
                    SELECT id, download_time, file_name, source_app, sender, note, file_path, file_id, file_size
                    FROM file_records 
                    WHERE file_name LIKE ? OR source_app LIKE ? OR sender LIKE ? OR note LIKE ?
                    ORDER BY id DESC
                """, (kw, kw, kw, kw))
            else:
                cur = conn.execute("""
                    SELECT id, download_time, file_name, source_app, sender, note, file_path, file_id, file_size
                    FROM file_records ORDER BY id DESC
                """)
            return cur.fetchall()

    def delete_record(self, record_id):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM file_records WHERE id = ?", (record_id,))


class SandboxHandler(FileSystemEventHandler):
    def __init__(self, event_queue, db, root_dir, refresh_callback):
        self.event_queue = event_queue
        self.db = db
        self.root_dir = os.path.normpath(root_dir)
        self.refresh_callback = refresh_callback
        self.recent_items = {}
        self.lock = threading.Lock()

    def _is_valid_item(self, item_path):
        name = os.path.basename(item_path)
        if name.startswith(('~$', '._', '~')):
            return False
        if os.path.isdir(item_path):
            return True
        _, ext = os.path.splitext(name)
        if not ext or ext.lower() in {'.tmp', '.crdownload', '.part', '.opdownload', '.dwl', '.dwl2', '.log', '.bak'}:
            return False
        if name.endswith(".db") or name == os.path.basename(DEFAULT_EXCEL_FILE):
            return False
        return True

    def _resolve_target_entity(self, raw_path):
        norm_p = os.path.normpath(raw_path)
        try:
            rel = os.path.relpath(norm_p, self.root_dir)
        except ValueError:
            return norm_p

        if rel == "." or rel.startswith(".."):
            return None

        parts = rel.split(os.sep)

        if len(parts) == 1 and os.path.isdir(norm_p):
            return None

        if len(parts) >= 3:
            package_folder = os.path.join(self.root_dir, parts[0], parts[1])
            if os.path.isdir(package_folder):
                return package_folder

        return norm_p

    def _handle_event(self, raw_path):
        target_path = self._resolve_target_entity(raw_path)
        if not target_path or not self._is_valid_item(target_path):
            return

        if self.db.is_inside_tracked_folder(target_path):
            return

        is_dir = os.path.isdir(target_path)
        base_name = os.path.basename(target_path)
        display_name = f"📁 {base_name}" if is_dir else base_name

        if self.db.exists(target_path):
            existing = self.db.find_record_by_name(display_name)
            if existing:
                self.db.update_path_and_time_by_id(existing[0], target_path)
                self.refresh_callback()
            return

        existing_by_name = self.db.find_record_by_name(display_name)
        if existing_by_name:
            self.db.update_path_and_time_by_id(existing_by_name[0], target_path)
            self.refresh_callback()
            return

        now = time.time()
        with self.lock:
            if target_path in self.recent_items and now - self.recent_items[target_path] < 4.0:
                return
            self.recent_items[target_path] = now

        def delayed_check():
            time.sleep(1.5)
            if not os.path.exists(target_path) or not self._is_valid_item(target_path):
                return

            curr_base = os.path.basename(target_path)
            curr_is_dir = os.path.isdir(target_path)
            curr_display = f"📁 {curr_base}" if curr_is_dir else curr_base

            existing = self.db.find_record_by_name(curr_display)
            if existing:
                self.db.update_path_and_time_by_id(existing[0], target_path)
                self.refresh_callback()
                return

            if not self.db.exists(target_path):
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.event_queue.put((timestamp, curr_display, "", "", target_path))

        threading.Thread(target=delayed_check, daemon=True).start()

    def on_created(self, event):
        self._handle_event(event.src_path)

    def on_moved(self, event):
        if not self._is_valid_item(event.dest_path):
            return

        dest_path = os.path.normpath(event.dest_path)
        dest_base = os.path.basename(dest_path)
        dest_display = f"📁 {dest_base}" if os.path.isdir(dest_path) else dest_base

        existing = self.db.find_record_by_name(dest_display)
        if existing:
            self.db.update_path_and_time_by_id(existing[0], dest_path)
            self.refresh_callback()
            return

        if self._is_valid_item(event.src_path):
            if self.db.update_path_on_move(event.src_path, dest_path):
                self.refresh_callback()
                return

        self._handle_event(dest_path)


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件智能台账系统")
        self.root.geometry("1120x560")

        self.db = Database(DB_FILE)
        self.event_queue = queue.Queue()
        
        self.active_popup = None
        self.active_batch_items = []
        self.merge_mode_var = tk.BooleanVar(value=False)
        self.active_title_var = None
        self.active_list_var = None
        self.active_merge_chk = None

        saved_dir = self.db.get_config("watch_dir", DEFAULT_WATCH_DIR)
        self.current_watch_dir = os.path.normpath(saved_dir)

        self._build_ui()
        self._check_and_create_sandbox(self.current_watch_dir)
        
        self.auto_heal_on_start()
        self._load_table()
        self._start_watcher(self.current_watch_dir)

        self.tray_icon = None
        self._init_tray()

        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)
        self.root.withdraw()

        self.root.after(400, self._check_queue)

    def _create_tray_image(self):
        image = Image.new('RGB', (64, 64), color=(26, 82, 118))
        dc = ImageDraw.Draw(image)
        dc.rectangle([14, 10, 46, 54], fill=(245, 247, 250))
        dc.polygon([(36, 10), (46, 20), (36, 20)], fill=(180, 190, 205))
        dc.rectangle([18, 16, 32, 21], fill=(41, 128, 185))
        dc.rectangle([18, 26, 42, 29], fill=(127, 140, 141))
        dc.rectangle([18, 34, 42, 37], fill=(127, 140, 141))
        dc.rectangle([18, 42, 36, 45], fill=(127, 140, 141))
        return image

    def _init_tray(self):
        if not HAS_TRAY:
            return
        menu = pystray.Menu(
            pystray.MenuItem("打开主界面", self._show_window_from_tray, default=True),
            pystray.MenuItem("退出程序", self._quit_from_tray)
        )
        self.tray_icon = pystray.Icon("FileTracker", self._create_tray_image(), "文件智能台账系统", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _hide_to_tray(self):
        self.root.withdraw()

    def _show_window_from_tray(self, icon=None, item=None):
        def show():
            self.auto_heal_on_start()
            self._load_table()
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        self.root.after(0, show)

    def _quit_from_tray(self, icon=None, item=None):
        try:
            if hasattr(self, 'observer') and self.observer:
                self.observer.stop()
                self.observer.join(timeout=1.0)
        except Exception:
            pass

        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass

        try:
            self.root.after(0, self.root.destroy)
        except Exception:
            pass

        os._exit(0)

    def _check_and_create_sandbox(self, path):
        if not os.path.exists(path):
            try:
                os.makedirs(path)
            except Exception:
                pass

    def _is_auto_start_enabled(self):
        if not HAS_WINREG:
            return False
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            val, _ = winreg.QueryValueEx(key, APP_NAME)
            winreg.CloseKey(key)
            return os.path.exists(val.strip('"'))
        except Exception:
            return False

    def _toggle_auto_start(self):
        if not HAS_WINREG:
            messagebox.showerror("错误", "当前环境不支持修改注册表。")
            return
        
        exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        target_val = f'"{exe_path}"'

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            if self.autostart_var.get():
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, target_val)
                messagebox.showinfo("成功", "已成功开启开机自启动！")
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
                messagebox.showinfo("成功", "已关闭开机自启动。")
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("操作失败", f"设置开机自启失败: {e}")

    def _build_ui(self):
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="🔍 搜索:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._load_table())
        ttk.Entry(top_frame, textvariable=self.search_var, width=16).pack(side=tk.LEFT, padx=5)

        ttk.Button(top_frame, text="🔧 修复断链", command=self.heal_all_broken_links).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="📊 同步 Excel", command=lambda: self.export_to_excel(DEFAULT_EXCEL_FILE)).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="🔄 刷新", command=self._load_table).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="📂 数据目录", command=lambda: os.startfile(DATA_DIR)).pack(side=tk.LEFT, padx=4)

        self.autostart_var = tk.BooleanVar(value=self._is_auto_start_enabled())
        autostart_chk = ttk.Checkbutton(top_frame, text="🚀 开机自启", variable=self.autostart_var, command=self._toggle_auto_start)
        autostart_chk.pack(side=tk.LEFT, padx=8)

        path_frame = ttk.Frame(self.root, padding=(10, 0, 10, 5))
        path_frame.pack(fill=tk.X)

        ttk.Label(path_frame, text="监控中转站路径:").pack(side=tk.LEFT)
        self.path_var = tk.StringVar(value=self.current_watch_dir)
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=65)
        self.path_entry.pack(side=tk.LEFT, padx=6)

        ttk.Button(path_frame, text="📁 浏览...", command=self._browse_dir).pack(side=tk.LEFT, padx=4)
        ttk.Button(path_frame, text="💾 保存路径", command=self._apply_new_path).pack(side=tk.LEFT, padx=4)

        columns = ("seq", "status", "time", "name", "app", "sender", "note", "path")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="extended")

        headers_map = [
            ("seq", "序号"),
            ("status", "状态"),
            ("time", "日期时间"),
            ("name", "文件 / 文件夹名称"),
            ("app", "来源"),
            ("sender", "联系人"),
            ("note", "备注"),
            ("path", "当前路径")
        ]
        for col, text in headers_map:
            self.tree.heading(col, text=text, command=lambda c=col: self._sort_column(c, False))

        self.tree.column("seq", width=45, anchor=tk.CENTER)
        self.tree.column("status", width=80, anchor=tk.CENTER)
        self.tree.column("time", width=135, anchor=tk.CENTER)
        self.tree.column("name", width=190)
        self.tree.column("app", width=120, anchor=tk.CENTER)
        self.tree.column("sender", width=120, anchor=tk.CENTER)
        self.tree.column("note", width=110)
        self.tree.column("path", width=215)

        self.tree.tag_configure("broken", foreground="#CC3333")
        self.tree.tag_configure("normal", foreground="#000000")

        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(5, 10))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=(5, 10))

        self.tree.bind("<Double-1>", lambda e: self._open_file())
        self.tree.bind("<Button-3>", self._on_right_click)

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📂 打开所在目录并高亮定位", command=self._open_containing_folder)
        self.context_menu.add_command(label="📄 直接打开文件 / 文件夹", command=self._open_file)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🔍 重新定位此条目", command=self._relocate_single_file)
        self.context_menu.add_command(label="🗑️ 从表中移除选中条目", command=self._delete_selected)

    def _renumber_rows(self):
        """动态连续重新顺排序号（永远从 1 开始连续递增）"""
        for idx, item_id in enumerate(self.tree.get_children(''), start=1):
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = idx
            self.tree.item(item_id, values=vals)

    def _sort_column(self, col, reverse):
        items = self.tree.get_children('')
        if col == "seq":
            # 点击序号列时，直接反转当前表格顺序并重新顺排 1, 2, 3...
            l = list(reversed(items))
            for index, k in enumerate(l):
                self.tree.move(k, '', index)
        else:
            l = [(self.tree.set(k, col), k) for k in items]
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
            for index, (val, k) in enumerate(l):
                self.tree.move(k, '', index)

        # 排序后保持动态连续序号 1, 2, 3...
        self._renumber_rows()

        headers_base = {
            "seq": "序号", "status": "状态", "time": "日期时间", "name": "文件 / 文件夹名称",
            "app": "来源", "sender": "联系人", "note": "备注", "path": "当前路径"
        }
        for c, base_text in headers_base.items():
            arrow = ""
            if c == col:
                arrow = " ▼" if reverse else " ▲"
            self.tree.heading(c, text=base_text + arrow, command=lambda _c=c: self._sort_column(_c, not reverse))

    def _browse_dir(self):
        d = filedialog.askdirectory(title="选择自定义监控文件夹")
        if d:
            self.path_var.set(os.path.normpath(d))

    def _apply_new_path(self):
        new_path = os.path.normpath(self.path_var.get().strip())
        if not new_path:
            messagebox.showwarning("提示", "路径不能为空！")
            return
        
        self._check_and_create_sandbox(new_path)
        self.current_watch_dir = new_path
        self.db.set_config("watch_dir", new_path)

        if hasattr(self, 'observer'):
            self.observer.stop()
        self._start_watcher(new_path)
        messagebox.showinfo("成功", f"监控路径已更新并保存：\n{new_path}")

    def _load_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        records = self.db.get_all(self.search_var.get().strip())
        for idx, row in enumerate(records, start=1):
            rec_id, d_time, f_name, app, sender, note, f_path, f_id, f_size = row
            exists = os.path.exists(f_path)
            status_text = "✅ 正常" if exists else "⚠️ 断链"
            tag = "normal" if exists else "broken"
            # iid 直接绑定数据库的真实主键 ID，第 1 列显示连续递增的动态序号 idx (1, 2, 3...)
            self.tree.insert(
                "", tk.END, iid=str(rec_id),
                values=(idx, status_text, d_time, f_name, app, sender, note, f_path),
                tags=(tag,)
            )

    def _get_selected_items_with_db_id(self):
        """返回 [(真实数据库ID, 行数据元组), ...]"""
        selected = self.tree.selection()
        if not selected:
            return []
        return [(int(iid), self.tree.item(iid, "values")) for iid in selected]

    def _search_file_in_workspaces(self, target_name, target_id, target_size):
        clean_name = target_name.replace("📁 ", "")
        search_roots = [self.current_watch_dir, os.path.expanduser("~/Desktop"), os.path.expanduser("~/Documents")]
        fallback_match = None
        for root_dir in search_roots:
            if not os.path.exists(root_dir):
                continue
            for dirpath, dirnames, filenames in os.walk(root_dir):
                for name in filenames + dirnames:
                    full_path = os.path.join(dirpath, name)
                    try:
                        st = os.stat(full_path)
                        if target_id and target_id != "0" and str(st.st_ino) == str(target_id):
                            return full_path
                        if name == clean_name:
                            fallback_match = full_path
                    except Exception:
                        continue
        return fallback_match

    def auto_heal_on_start(self):
        records = self.db.get_all()
        for rec in records:
            rec_id, _, f_name, _, _, _, old_path, _, _ = rec
            clean_name = f_name.replace("📁 ", "")
            if not os.path.exists(old_path):
                if os.path.exists(self.current_watch_dir):
                    for dirpath, dirnames, filenames in os.walk(self.current_watch_dir):
                        if clean_name in filenames or clean_name in dirnames:
                            new_path = os.path.join(dirpath, clean_name)
                            self.db.update_path_and_time_by_id(rec_id, new_path)
                            break
        
        with sqlite3.connect(self.db.db_path) as conn:
            conn.execute("DELETE FROM file_records WHERE file_name LIKE '%.tmp'")

    def heal_all_broken_links(self):
        self.auto_heal_on_start()
        records = self.db.get_all()
        broken_list = [r for r in records if not os.path.exists(r[6])]
        
        for rec in broken_list:
            rec_id, _, f_name, _, _, _, old_path, f_id, f_size = rec
            found_path = self._search_file_in_workspaces(f_name, f_id, f_size)
            if found_path:
                self.db.update_path_and_time_by_id(rec_id, found_path)

        self._load_table()
        if not broken_list:
            messagebox.showinfo("状态良好", "所有索引文件均在原位！")
        else:
            messagebox.showinfo("断链自愈完成", "成功自动寻回并修复断链记录！")

    def _relocate_single_file(self):
        items = self._get_selected_items_with_db_id()
        if not items:
            return
        rec_id, vals = items[0]
        f_name = vals[3]
        clean_name = f_name.replace("📁 ", "")
        if f_name.startswith("📁 "):
            new_path = filedialog.askdirectory(title=f"请重新定位文件夹: {clean_name}")
        else:
            new_path = filedialog.askopenfilename(title=f"请重新定位文件: {clean_name}", initialfile=clean_name)
        if new_path:
            self.db.update_path_and_time_by_id(rec_id, os.path.normpath(new_path))
            self._load_table()

    def _open_file(self):
        items = self._get_selected_items_with_db_id()
        if not items:
            return
        _, vals = items[0]
        path = vals[7]
        if os.path.exists(path):
            os.startfile(path)
        else:
            if messagebox.askyesno("路径已失效", "原路径找不到目标，是否立即搜索修复？"):
                self._relocate_single_file()

    def _open_containing_folder(self):
        items = self._get_selected_items_with_db_id()
        if not items:
            return
        _, vals = items[0]
        path = vals[7]
        if os.path.exists(path):
            subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')
        else:
            if messagebox.askyesno("路径已失效", "原路径找不到目标，是否立即搜索修复？"):
                self._relocate_single_file()

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            if row_id not in self.tree.selection():
                self.tree.selection_set(row_id)
            self.context_menu.post(event.x_root, event.y_root)

    def _delete_selected(self):
        items = self._get_selected_items_with_db_id()
        if not items:
            return
        
        count = len(items)
        if messagebox.askyesno("确认删除", f"确定要从表中移除选中的 {count} 条索引记录吗？"):
            for rec_id, _ in items:
                self.db.delete_record(rec_id)
            # 重新加载后，序号会自动从 1 开始重新连续顺排
            self._load_table()

    def export_to_excel(self, excel_path):
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少组件", "请先安装 openpyxl。")
            return

        records = self.db.get_all()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "下载文件索引台账"

        headers = ["序号", "日期时间", "文件/文件夹名称", "来源", "联系人", "备注", "状态", "快捷打开", "快捷打开目录", "完整路径"]
        ws.append(headers)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
        link_font = Font(name="Microsoft YaHei", color="0563C1", underline="single")
        normal_font = Font(name="Microsoft YaHei")

        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=col_num)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        for seq_num, rec in enumerate(records, start=1):
            row_idx = seq_num + 1
            rec_id, d_time, f_name, app, sender, note, f_path, _, _ = rec
            exists = os.path.exists(f_path)
            status_str = "正常" if exists else "断链"
            folder_path = os.path.dirname(f_path)

            file_link_formula = f'=HYPERLINK("{f_path.replace('"', '""')}", "📄 点击打开")'
            folder_link_formula = f'=HYPERLINK("{folder_path.replace('"', '""')}", "📂 打开所在目录")'

            ws.append([seq_num, d_time, f_name, app, sender, note, status_str, file_link_formula, folder_link_formula, f_path])

            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_num)
                cell.font = link_font if col_num in (8, 9) else normal_font

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for i, width in enumerate([6, 20, 32, 16, 18, 22, 10, 18, 18, 55], start=1):
            ws.column_dimensions[get_column_letter(i)].width = width

        try:
            wb.save(excel_path)
            if messagebox.askyesno("导出成功", f"台账已同步至专属数据目录：\n{os.path.abspath(excel_path)}\n\n是否立即打开？"):
                os.startfile(os.path.abspath(excel_path))
        except PermissionError:
            messagebox.showwarning("文件被占用", "请先关闭正在打开的 Excel 文件！")

    def _start_watcher(self, watch_dir):
        self.observer = Observer()
        handler = SandboxHandler(
            self.event_queue, self.db, watch_dir,
            refresh_callback=lambda: self.root.after(0, self._load_table)
        )
        if os.path.exists(watch_dir):
            self.observer.schedule(handler, watch_dir, recursive=True)
        self.observer.start()

    def _check_queue(self):
        try:
            while not self.event_queue.empty():
                if self.active_popup is None or not self.active_popup.winfo_exists():
                    item = self.event_queue.get_nowait()
                    self._open_batch_popup(item)
                else:
                    if self.merge_mode_var.get():
                        item = self.event_queue.get_nowait()
                        self._append_to_active_popup(item)
                    else:
                        break
        except queue.Empty:
            pass
        self.root.after(400, self._check_queue)

    def _append_to_active_popup(self, item):
        if any(existing[4] == item[4] for existing in self.active_batch_items):
            return
        self.active_batch_items.append(item)
        count = len(self.active_batch_items)
        names_preview = "、".join([x[1] for x in self.active_batch_items[:3]])
        if count > 3:
            names_preview += f" 等共 {count} 项"

        if self.active_title_var:
            self.active_title_var.set(f"批量归档（已合并 {count} 个项目）")
        if self.active_list_var:
            self.active_list_var.set(f"包含: {names_preview}")
        if self.active_merge_chk:
            self.active_merge_chk.config(text=f"合并归档后续连续保存的文件（已追加 {count - 1} 个）")

    def _open_batch_popup(self, first_item):
        self.active_batch_items = [first_item]
        self.merge_mode_var.set(False)

        popup = tk.Toplevel(self.root)
        self.active_popup = popup
        popup.title("检测到新文件 - 智能归档")
        popup.geometry("460x315")
        popup.attributes("-topmost", True)

        frame = ttk.Frame(popup, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        timestamp, file_name, _, _, file_path = first_item

        self.active_title_var = tk.StringVar(value=f"新项目: {file_name}")
        ttk.Label(frame, textvariable=self.active_title_var, font=("Microsoft YaHei", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4)
        )

        self.active_list_var = tk.StringVar(value=f"路径: {file_path[:54]}...")
        ttk.Label(frame, textvariable=self.active_list_var, foreground="gray").grid(
            row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10)
        )

        ttk.Label(frame, text="来源:").grid(row=2, column=0, sticky=tk.W, pady=4)
        app_var = tk.StringVar(value="")
        app_combo = ttk.Combobox(frame, textvariable=app_var, values=self.db.get_apps(), width=34)
        app_combo.grid(row=2, column=1, sticky=tk.W, pady=4)
        app_combo.focus_set()

        ttk.Label(frame, text="联系人:").grid(row=3, column=0, sticky=tk.W, pady=4)
        sender_var = tk.StringVar(value="")
        sender_combo = ttk.Combobox(frame, textvariable=sender_var, values=self.db.get_senders(), width=34)
        sender_combo.grid(row=3, column=1, sticky=tk.W, pady=4)

        ttk.Label(frame, text="备注信息:").grid(row=4, column=0, sticky=tk.W, pady=4)
        note_var = tk.StringVar()
        note_entry = ttk.Entry(frame, textvariable=note_var, width=37)
        note_entry.grid(row=4, column=1, sticky=tk.W, pady=4)

        def on_toggle_merge():
            if self.merge_mode_var.get():
                while not self.event_queue.empty():
                    try:
                        nxt = self.event_queue.get_nowait()
                        self._append_to_active_popup(nxt)
                    except queue.Empty:
                        break

        self.active_merge_chk = ttk.Checkbutton(
            frame,
            text="合并归档后续连续保存的文件（共享来源、联系人与备注）",
            variable=self.merge_mode_var,
            command=on_toggle_merge
        )
        self.active_merge_chk.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(10, 12))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=2)

        def close_popup():
            self.active_popup = None
            self.active_batch_items = []
            popup.destroy()

        def on_confirm(event=None):
            entered_app = app_var.get().strip()
            entered_sender = sender_var.get().strip()
            entered_note = note_var.get().strip()

            for item in self.active_batch_items:
                t_stamp, f_name, _, _, f_path = item
                self.db.add_record(
                    t_stamp, f_name,
                    entered_app,
                    entered_sender,
                    entered_note, f_path
                )
            self._load_table()
            close_popup()

        popup.protocol("WM_DELETE_WINDOW", close_popup)
        popup.bind("<Return>", on_confirm)
        popup.bind("<Escape>", lambda e: close_popup())

        ttk.Button(btn_frame, text="✅ 记录到表格 (Enter)", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="❌ 忽略本次 (Esc)", command=close_popup).pack(side=tk.LEFT, padx=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()