import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import requests
import webbrowser
import os
import sys
import time
import re
import shutil
import subprocess
import json
from threading import Thread, Event
from datetime import datetime

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

try:
    from PIL import Image, ImageDraw, ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False


# =====================================================================
#  常量与路径
# =====================================================================
CURRENT_VERSION = "v1.1.1"
GITHUB_USERNAME = "RyanTanC"
GITHUB_REPO_NAME = "HNUST-Exam-System"

_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".hnust_exam")
_CONFIG_FILE = os.path.join(_CONFIG_DIR, "config.json")
_PROGRESS_FILE = os.path.join(_CONFIG_DIR, "progress.json")
_SKIP_VERSION_FILE = os.path.join(_CONFIG_DIR, "skip_ver")


# =====================================================================
#  配置与进度管理
# =====================================================================
def _ensure_config_dir():
    os.makedirs(_CONFIG_DIR, exist_ok=True)


def _load_config():
    _ensure_config_dir()
    defaults = {
        "font_scale": 1.0,
        "dark_mode": False,
        "show_answer_immediately": False,
        "user_python_path": "",
    }
    try:
        if os.path.exists(_CONFIG_FILE):
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            defaults.update(data)
    except Exception:
        pass
    return defaults


def _save_config(config):
    _ensure_config_dir()
    try:
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _load_progress():
    _ensure_config_dir()
    try:
        if os.path.exists(_PROGRESS_FILE):
            with open(_PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_progress(progress):
    _ensure_config_dir()
    try:
        with open(_PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _get_log_dir():
    try:
        docs = os.path.expanduser("~/Documents")
        if not os.path.isdir(docs):
            docs = os.path.expanduser("~")
        log_dir = os.path.join(docs, "HNUST_Exam_Logs")
        os.makedirs(log_dir, exist_ok=True)
        return log_dir
    except Exception:
        return os.path.dirname(os.path.abspath(sys.argv[0]))


# =====================================================================
#  版本更新检查
# =====================================================================
def _version_tuple(v):
    v = v.lstrip("vV")
    parts = v.split("-")
    main = tuple(int(x) for x in parts[0].split(".") if x.isdigit())
    return main


def _load_skip_version():
    try:
        if os.path.exists(_SKIP_VERSION_FILE):
            with open(_SKIP_VERSION_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
    except Exception:
        pass
    return ""


def _save_skip_version(ver):
    try:
        _ensure_config_dir()
        with open(_SKIP_VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(ver)
    except Exception:
        pass


def _fetch_update_info():
    try:
        repo_api_url = (
            f"https://api.github.com/repos/{GITHUB_USERNAME}"
            f"/{GITHUB_REPO_NAME}/releases/latest"
        )
        response = requests.get(repo_api_url, timeout=5)
        response.raise_for_status()
        data = response.json()
        latest_version = data.get("tag_name", "")
        if not latest_version:
            return None
        if _version_tuple(latest_version) <= _version_tuple(CURRENT_VERSION):
            return None
        if _load_skip_version() == latest_version:
            return None

        release_notes = (data.get("body", "") or "").strip() or "暂无更新日志"
        download_url = ""
        assets = data.get("assets", [])
        if assets:
            download_url = assets[0].get("browser_download_url", "")
        if not download_url:
            download_url = data.get("html_url", "")

        published = data.get("published_at", "")
        if published:
            try:
                dt = datetime.strptime(published, "%Y-%m-%dT%H:%M:%SZ")
                published = dt.strftime("%Y年%m月%d日 %H:%M")
            except Exception:
                pass

        return {
            "latest_ver": latest_version,
            "release_notes": release_notes,
            "download_url": download_url,
            "published_at": published,
        }
    except Exception:
        pass
    return None


# =====================================================================
#  更新对话框（深色模式适配）
# =====================================================================
def _show_update_dialog(parent, info):
    win = tk.Toplevel(parent)
    win.title("发现新版本")
    win.resizable(True, True)
    win.configure(bg=Theme.WHITE)
    win.attributes("-topmost", True)

    icon_path = get_resource_path("icon.ico")
    if os.path.exists(icon_path):
        try:
            win.iconbitmap(icon_path)
        except Exception:
            pass

    BG = Theme.WHITE
    PRIMARY = Theme.PRIMARY
    PRIMARY_HOVER = Theme.PRIMARY_HOVER
    SURFACE = Theme.SURFACE
    BORDER = Theme.BORDER
    TEXT = Theme.TEXT
    MUTED = Theme.MUTED
    GREEN = Theme.SUCCESS

    current_ver = CURRENT_VERSION
    latest_ver = info["latest_ver"]
    release_notes = info["release_notes"]
    download_url = info["download_url"]
    published_at = info.get("published_at", "")

    # 顶部标题栏
    header = tk.Frame(win, bg=PRIMARY, height=100)
    header.pack(fill=tk.X)
    header.pack_propagate(False)

    icon_circle = tk.Canvas(header, width=56, height=56, bg=PRIMARY,
                            highlightthickness=0)
    icon_circle.pack(side=tk.LEFT, padx=(30, 15), pady=22)
    icon_circle.create_oval(2, 2, 54, 54, fill="#ffffff", outline="#ffffff")
    icon_circle.create_text(28, 28, text="↑",
                            font=("微软雅黑", 22, "bold"), fill=PRIMARY)

    header_text = tk.Frame(header, bg=PRIMARY)
    header_text.pack(side=tk.LEFT, pady=22)
    tk.Label(header_text, text="发现新版本",
             font=("微软雅黑", 18, "bold"), bg=PRIMARY,
             fg="#ffffff").pack(anchor="w")
    tk.Label(header_text,
             text=f"v{current_ver.lstrip('v')} → {latest_ver}",
             font=("微软雅黑", 11), bg=PRIMARY,
             fg=Theme.HEADER_SUB_TEXT).pack(anchor="w")

    # 主体
    body = tk.Frame(win, bg=BG)
    body.pack(fill=tk.BOTH, expand=True)

    # 版本信息卡片
    info_card = tk.Frame(body, bg=SURFACE, bd=0)
    info_card.pack(fill=tk.X, padx=24, pady=(20, 0))
    info_grid = tk.Frame(info_card, bg=SURFACE)
    info_grid.pack(padx=16, pady=14, anchor="w")

    info_items = [("当前版本", current_ver, MUTED),
                  ("最新版本", latest_ver, GREEN)]
    if published_at:
        info_items.append(("发布时间", published_at, MUTED))
    for i, (label, value, color) in enumerate(info_items):
        tk.Label(info_grid, text=label + "：",
                 font=("微软雅黑", 10), bg=SURFACE, fg=MUTED,
                 anchor="e", width=10).grid(row=i, column=0,
                                            sticky="e", pady=3)
        tk.Label(info_grid, text=value,
                 font=("微软雅黑", 10, "bold"), bg=SURFACE, fg=color,
                 anchor="w").grid(row=i, column=1, sticky="w",
                                 pady=3, padx=(8, 0))

    tk.Label(body, text="更新日志",
             font=("微软雅黑", 11, "bold"), bg=BG, fg=TEXT,
             anchor="w").pack(fill=tk.X, padx=24, pady=(16, 6))

    # 日志滚动区
    notes_frame = tk.Frame(body, bg=BG, bd=1, relief=tk.SOLID,
                           highlightbackground=BORDER,
                           highlightthickness=1)
    notes_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 12))

    notes_canvas = tk.Canvas(notes_frame, bg=Theme.NOTES_BG,
                             highlightthickness=0)
    notes_scrollbar = ttk.Scrollbar(notes_frame, orient="vertical",
                                    command=notes_canvas.yview)
    notes_inner = tk.Frame(notes_canvas, bg=Theme.NOTES_BG)
    notes_canvas.create_window((0, 0), window=notes_inner, anchor="nw")
    notes_canvas.configure(yscrollcommand=notes_scrollbar.set)
    notes_inner.bind(
        "<Configure>",
        lambda e: notes_canvas.configure(scrollregion=notes_canvas.bbox("all")))
    notes_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    notes_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _render_notes(wrap_val):
        for w in notes_inner.winfo_children():
            w.destroy()
        for line in release_notes.splitlines():
            line = line.strip()
            if not line:
                tk.Label(notes_inner, text="", bg=Theme.NOTES_BG,
                         height=1).pack()
                continue
            if line.startswith("### "):
                tk.Label(notes_inner, text=line[4:],
                         font=("微软雅黑", 10, "bold"), bg=Theme.NOTES_BG,
                         fg=TEXT, anchor="w", wraplength=wrap_val,
                         justify="left").pack(fill=tk.X, padx=12,
                                              pady=(8, 2))
            elif line.startswith("## "):
                tk.Label(notes_inner, text=line[3:],
                         font=("微软雅黑", 11, "bold"), bg=Theme.NOTES_BG,
                         fg=TEXT, anchor="w", wraplength=wrap_val,
                         justify="left").pack(fill=tk.X, padx=12,
                                              pady=(8, 2))
            elif line.startswith("- ") or line.startswith("* "):
                tk.Label(notes_inner, text=f"  •  {line[2:]}",
                         font=("微软雅黑", 10), bg=Theme.NOTES_BG, fg=TEXT,
                         anchor="w", wraplength=max(100, wrap_val - 10),
                         justify="left").pack(fill=tk.X, padx=16, pady=1)
            else:
                tk.Label(notes_inner, text=line,
                         font=("微软雅黑", 10), bg=Theme.NOTES_BG, fg=TEXT,
                         anchor="w", wraplength=wrap_val,
                         justify="left").pack(fill=tk.X, padx=12, pady=1)

    def _on_notes_resize(event):
        _render_notes(max(200, event.width - 40))

    notes_frame.bind("<Configure>", _on_notes_resize)

    def _mw(e):
        if sys.platform == "darwin":
            notes_canvas.yview_scroll(-e.delta, "units")
        else:
            notes_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    notes_canvas.bind("<MouseWheel>", _mw)
    notes_canvas.bind("<Button-4>",
                      lambda e: notes_canvas.yview_scroll(-3, "units"))
    notes_canvas.bind("<Button-5>",
                      lambda e: notes_canvas.yview_scroll(3, "units"))

    # 按钮栏
    btn_bar = tk.Frame(win, bg=BG, height=80)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Frame(btn_bar, bg=BORDER, height=1).pack(fill=tk.X)
    btn_inner = tk.Frame(btn_bar, bg=BG)
    btn_inner.pack(pady=12)

    def _on_update():
        if download_url:
            webbrowser.open(download_url)
        win.destroy()

    def _on_skip():
        _save_skip_version(latest_ver)
        win.destroy()

    def _on_continue():
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", _on_continue)

    skip_btn = tk.Label(btn_inner, text="  跳过此版本  ",
                        font=("微软雅黑", 9), bg=BG, fg=MUTED,
                        padx=10, pady=6, cursor="hand2", relief=tk.FLAT)
    skip_btn.pack(side=tk.LEFT, padx=(0, 8))
    skip_btn.bind("<Enter>", lambda e: skip_btn.config(fg=TEXT))
    skip_btn.bind("<Leave>", lambda e: skip_btn.config(fg=MUTED))
    skip_btn.bind("<Button-1>", lambda e: _on_skip())

    later_btn = tk.Label(btn_inner, text="  暂不更新  ",
                         font=("微软雅黑", 10), bg=SURFACE, fg=MUTED,
                         padx=14, pady=6, cursor="hand2", relief=tk.FLAT)
    later_btn.pack(side=tk.LEFT, padx=(0, 12))
    later_btn.bind("<Enter>", lambda e: later_btn.config(fg=Theme.DANGER))
    later_btn.bind("<Leave>", lambda e: later_btn.config(fg=MUTED))
    later_btn.bind("<Button-1>", lambda e: _on_continue())

    update_btn = tk.Label(btn_inner, text="   立即更新   ",
                          font=("微软雅黑", 11, "bold"), bg=PRIMARY,
                          fg="#ffffff", padx=24, pady=8,
                          cursor="hand2", relief=tk.FLAT)
    update_btn.pack(side=tk.LEFT)
    update_btn.bind("<Enter>", lambda e: update_btn.config(bg=PRIMARY_HOVER))
    update_btn.bind("<Leave>", lambda e: update_btn.config(bg=PRIMARY))
    update_btn.bind("<Button-1>", lambda e: _on_update())

    tk.Label(body,
             text="建议更新到最新版本以获得最佳体验（也可以稍后更新）",
             font=("微软雅黑", 8), bg=BG, fg=MUTED).pack(pady=(0, 4))

    # 自适应尺寸
    win.withdraw()
    win.update_idletasks()
    req_w = win.winfo_reqwidth()
    req_h = win.winfo_reqheight()
    scr_w = win.winfo_screenwidth()
    scr_h = win.winfo_screenheight()
    W = max(500, min(req_w + 40, int(scr_w * 0.85)))
    H = max(400, min(req_h + 20, int(scr_h * 0.8)))
    x = (scr_w - W) // 2
    y = (scr_h - H) // 2
    win.geometry(f"{W}x{H}+{x}+{y}")
    win.minsize(500, 400)
    win.deiconify()

    win.grab_set()
    win.focus_force()
    parent.wait_window(win)


# =====================================================================
#  主题（深色模式全面优化）
# =====================================================================
class Theme:
    PRIMARY = "#0078d7"
    PRIMARY_HOVER = "#005fa3"
    ACCENT = "#ff9900"
    BG = "#f0f0f0"
    WHITE = "#ffffff"
    TEXT = "#333333"
    HINT_BG = "#e6f2ff"
    HINT_TEXT = "#0066cc"
    DANGER = "#ff6600"
    SUCCESS = "#28a745"
    MUTED = "#999999"
    BORDER = "#e0e0e0"
    NAV_ACTIVE = "#cce0ff"
    NAV_CURRENT = "#0050a0"
    NAV_ANSWERED_BG = "#d4edda"
    NAV_ANSWERED_FG = "#155724"
    NAV_MARKED_BG = "#f8d7da"
    NAV_MARKED_FG = "#a71d2a"
    NAV_HEADER_BG = "#e8f0fe"
    SURFACE = "#f8f9fa"
    CARD_BG = "#ffffff"
    INPUT_BG = "#ffffff"
    PROGRESS_BG = "#e0e0e0"
    KB_HINT_BG = "#f8f8f8"
    TOOLTIP_BG = "#ffffe0"
    TOOLTIP_FG = "#333333"
    NOTES_BG = "#fafbfc"
    ANSWER_BG = "#e6ffe6"
    WARN_BG = "#fff3cd"
    WARN_TEXT = "#856404"
    WARN_BORDER = "#ffc107"
    HEADER_SUB_TEXT = "#bfdbfe"
    CRASH_BG = "#ffffff"
    CRASH_HEADER = "#dc2626"

    _is_dark = False
    _font_scale = 1.0

    FONT = ("微软雅黑", 11)
    FONT_BOLD = ("微软雅黑", 11, "bold")
    FONT_TITLE = ("微软雅黑", 12, "bold")
    FONT_HUGE = ("微软雅黑", 16, "bold")
    FONT_SMALL = ("微软雅黑", 9)
    FONT_TINY = ("微软雅黑", 7)

    _LIGHT = {
        "PRIMARY": "#0078d7", "PRIMARY_HOVER": "#005fa3",
        "ACCENT": "#ff9900", "BG": "#f0f0f0",
        "WHITE": "#ffffff", "TEXT": "#333333", "HINT_BG": "#e6f2ff",
        "HINT_TEXT": "#0066cc",
        "DANGER": "#ff6600", "SUCCESS": "#28a745", "MUTED": "#999999",
        "BORDER": "#e0e0e0", "NAV_ACTIVE": "#cce0ff",
        "NAV_CURRENT": "#0050a0", "NAV_ANSWERED_BG": "#d4edda",
        "NAV_ANSWERED_FG": "#155724", "NAV_MARKED_BG": "#f8d7da",
        "NAV_MARKED_FG": "#a71d2a", "NAV_HEADER_BG": "#e8f0fe",
        "SURFACE": "#f8f9fa", "CARD_BG": "#ffffff",
        "INPUT_BG": "#ffffff", "PROGRESS_BG": "#e0e0e0",
        "KB_HINT_BG": "#f8f8f8",
        "TOOLTIP_BG": "#ffffe0", "TOOLTIP_FG": "#333333",
        "NOTES_BG": "#fafbfc", "ANSWER_BG": "#e6ffe6",
        "WARN_BG": "#fff3cd", "WARN_TEXT": "#856404",
        "WARN_BORDER": "#ffc107", "HEADER_SUB_TEXT": "#bfdbfe",
        "CRASH_BG": "#ffffff", "CRASH_HEADER": "#dc2626",
    }

    _DARK = {
        "PRIMARY": "#4da6ff", "PRIMARY_HOVER": "#3d8adf",
        "ACCENT": "#ffaa33", "BG": "#1e1e1e",
        "WHITE": "#2d2d2d", "TEXT": "#e0e0e0", "HINT_BG": "#1a2a3a",
        "HINT_TEXT": "#66aaff",
        "DANGER": "#ff6644", "SUCCESS": "#44cc66", "MUTED": "#888888",
        "BORDER": "#444444", "NAV_ACTIVE": "#2a3a5a",
        "NAV_CURRENT": "#cc7700", "NAV_ANSWERED_BG": "#1a3a2a",
        "NAV_ANSWERED_FG": "#44cc66", "NAV_MARKED_BG": "#3a1a1a",
        "NAV_MARKED_FG": "#ff6644", "NAV_HEADER_BG": "#1a2a3a",
        "SURFACE": "#2a2a2a", "CARD_BG": "#2d2d2d",
        "INPUT_BG": "#333333", "PROGRESS_BG": "#333333",
        "KB_HINT_BG": "#252525",
        "TOOLTIP_BG": "#3a3a2a", "TOOLTIP_FG": "#e0e0e0",
        "NOTES_BG": "#252525", "ANSWER_BG": "#1a3a2a",
        "WARN_BG": "#3a3020", "WARN_TEXT": "#ffcc66",
        "WARN_BORDER": "#886600", "HEADER_SUB_TEXT": "#d0e8ff",
        "CRASH_BG": "#1e1e1e", "CRASH_HEADER": "#b72626",
    }

    @classmethod
    def set_dark_mode(cls, enabled):
        cls._is_dark = enabled
        colors = cls._DARK if enabled else cls._LIGHT
        for k, v in colors.items():
            setattr(cls, k, v)

    @classmethod
    def _update_fonts(cls):
        s = cls._font_scale
        cls.FONT = ("微软雅黑", max(8, int(11 * s)))
        cls.FONT_BOLD = ("微软雅黑", max(8, int(11 * s)), "bold")
        cls.FONT_TITLE = ("微软雅黑", max(9, int(12 * s)), "bold")
        cls.FONT_HUGE = ("微软雅黑", max(12, int(16 * s)), "bold")
        cls.FONT_SMALL = ("微软雅黑", max(7, int(9 * s)))
        cls.FONT_TINY = ("微软雅黑", max(6, int(7 * s)))


# =====================================================================
#  工具函数
# =====================================================================
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def bind_mousewheel(widget, canvas):
    def _on_wheel(event):
        if sys.platform == "darwin":
            canvas.yview_scroll(-event.delta, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_wheel_up(event):
        canvas.yview_scroll(-3, "units")

    def _on_wheel_down(event):
        canvas.yview_scroll(3, "units")

    def _bind_recursive(w):
        w.bind("<MouseWheel>", _on_wheel)
        w.bind("<Button-4>", _on_wheel_up)
        w.bind("<Button-5>", _on_wheel_down)
        for child in w.winfo_children():
            _bind_recursive(child)

    _bind_recursive(widget)


def find_system_python():
    if sys.platform != "win32":
        return None

    NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    candidates = []

    try:
        py_path = shutil.which("py")
        if py_path:
            result = subprocess.run(
                [py_path, "-3", "-c", "import sys; print(sys.executable)"],
                capture_output=True, text=True, timeout=5,
                creationflags=NO_WINDOW)
            if result.returncode == 0:
                exe = result.stdout.strip()
                if exe and os.path.isfile(exe) and exe not in candidates:
                    candidates.append(exe)
    except Exception:
        pass

    for name in ["python", "python3"]:
        try:
            p = shutil.which(name)
            if p and os.path.isfile(p) and p not in candidates:
                candidates.append(p)
        except Exception:
            pass

    try:
        result = subprocess.run(
            ["where", "python"],
            capture_output=True, text=True, timeout=5,
            creationflags=NO_WINDOW)
        if result.returncode == 0:
            for line in result.stdout.strip().splitlines():
                line = line.strip()
                if line and os.path.isfile(line) and line not in candidates:
                    candidates.append(line)
    except Exception:
        pass

    try:
        import winreg
        for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            for key_path in [
                r"SOFTWARE\Python\PythonCore",
                r"SOFTWARE\WOW6432Node\Python\PythonCore"
            ]:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        for i in range(winreg.QueryInfoKey(key)[0]):
                            version = winreg.EnumKey(key, i)
                            try:
                                with winreg.OpenKey(
                                        key, f"{version}\\InstallPath") as ik:
                                    path = winreg.QueryValue(ik, "")
                                    exe = os.path.join(path, "python.exe")
                                    if os.path.isfile(exe) \
                                            and exe not in candidates:
                                        candidates.append(exe)
                            except Exception:
                                continue
                except Exception:
                    continue
    except Exception:
        pass

    local_app = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
    program_files_x86 = os.environ.get(
        "ProgramFiles(x86)", "C:\\Program Files (x86)")
    home = os.path.expanduser("~")

    common_paths = []
    for v in range(8, 20):
        common_paths.extend([
            os.path.join(local_app,
                         f"Programs\\Python\\Python3{v}\\python.exe"),
            f"C:\\Python3{v}\\python.exe",
            os.path.join(program_files,
                         f"Python3{v}\\python.exe"),
            os.path.join(program_files_x86,
                         f"Python3{v}\\python.exe"),
        ])
    common_paths.extend([
        os.path.join(local_app, "anaconda3", "python.exe"),
        os.path.join(local_app, "miniconda3", "python.exe"),
        os.path.join(home, "anaconda3", "python.exe"),
        os.path.join(home, "miniconda3", "python.exe"),
        os.path.join(program_files, "anaconda3", "python.exe"),
        os.path.join(program_files, "miniconda3", "python.exe"),
        r"C:\Anaconda3\python.exe",
        r"C:\Miniconda3\python.exe",
    ])
    for p in common_paths:
        if os.path.isfile(p) and p not in candidates:
            candidates.append(p)

    for dir_path in os.environ.get("PATH", "").split(os.pathsep):
        dir_path = dir_path.strip()
        if not dir_path:
            continue
        exe = os.path.join(dir_path, "python.exe")
        if os.path.isfile(exe) and exe not in candidates:
            candidates.append(exe)

    return candidates[0] if candidates else None


# =====================================================================
#  主应用类
# =====================================================================
class HNUSTExamSystem:

    # ─────────────────────────────────────────────────────────────────
    #  初始化
    # ─────────────────────────────────────────────────────────────────
    def __init__(self, root):
        self.root = root
        self.root.title("HNUST仿真平台")
        self.root.geometry("1200x800")
        self.root.configure(bg=Theme.BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if sys.platform == "win32":
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "HNUST.ExamSystem.V1.0")

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        # 全局 ttk 深色样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self._apply_ttk_styles()

        cfg = _load_config()
        Theme._font_scale = cfg.get("font_scale", 1.0)
        Theme._update_fonts()
        Theme.set_dark_mode(cfg.get("dark_mode", False))
        self.root.configure(bg=Theme.BG)
        self.show_answer_immediately = cfg.get("show_answer_immediately", False)
        self._user_python_path = cfg.get("user_python_path", "")

        self._add_anti_sale_watermark()

        self._current_view = "welcome"
        self.current_exam_file = None
        self.questions = []
        self.question_groups = {}
        self.current_index = 0
        self.user_answers = {}
        self.score = 0
        self.marked_questions = set()
        self.exam_time = 60 * 60
        self.remaining_time = self.exam_time
        self.timer_running = False
        self._stop_timer = Event()
        self._timer_thread = None
        self.answer_var = None
        self.answer_text = None
        self.answer_entry = None
        self._choice_buttons = {}
        self.exam_submitted = False
        self._backup_dir = None
        self._progress_data = _load_progress()

        self.question_type_order = ["单选", "填空", "判断", "程序填空", "程序改错"]
        self.is_pure_program_exam = False
        self.active_type_order = []

        self.create_welcome_window()
        self._check_updates_async()

    def _apply_ttk_styles(self):
        self.style.configure("Vertical.TScrollbar",
                             background=Theme.BORDER,
                             troughcolor=Theme.SURFACE,
                             bordercolor=Theme.BORDER,
                             darkcolor=Theme.BORDER,
                             lightcolor=Theme.BORDER,
                             arrowcolor=Theme.MUTED)
        self.style.map("Vertical.TScrollbar",
                       background=[("active", Theme.PRIMARY)])
        self.style.configure("TEntry",
                             fieldbackground=Theme.INPUT_BG,
                             foreground=Theme.TEXT,
                             bordercolor=Theme.BORDER,
                             darkcolor=Theme.BORDER,
                             lightcolor=Theme.BORDER)
        self.style.configure("TSeparator", background=Theme.BORDER)

    # ─────────────────────────────────────────────────────────────────
    #  更新检查
    # ─────────────────────────────────────────────────────────────────
    def _check_updates_async(self):
        def _worker():
            try:
                info = _fetch_update_info()
                if info:
                    try:
                        self.root.after(1000, lambda: self._prompt_update(info))
                    except tk.TclError:
                        pass
            except Exception:
                pass
        Thread(target=_worker, daemon=True).start()

    def _prompt_update(self, info):
        try:
            _show_update_dialog(self.root, info)
        except tk.TclError:
            pass

    # ─────────────────────────────────────────────────────────────────
    #  关闭
    # ─────────────────────────────────────────────────────────────────
    def _on_close(self):
        if self.timer_running and not self.exam_submitted:
            if not messagebox.askyesno(
                    "确认退出",
                    "考试正在进行中，确定要退出吗？\n未交卷的答案将不会保存。"):
                return
        self.timer_running = False
        self._stop_timer.set()
        if self._timer_thread and self._timer_thread.is_alive():
            self._timer_thread.join(timeout=2)
        self._cleanup_backup()
        self.root.destroy()

    # ─────────────────────────────────────────────────────────────────
    #  防售卖水印
    # ─────────────────────────────────────────────────────────────────
    def _add_anti_sale_watermark(self):
        self._wm = tk.Label(
            self.root, text="该程序免费 禁止商用售卖",
            font=Theme.FONT_TINY, bg=Theme.BG, fg=Theme.MUTED, bd=0)
        self._wm.place(relx=0.99, rely=0.99, anchor="se")
        self._wm._wm_protected = True

        self._wm2 = tk.Label(
            self.root, text="HNUST Exam · Free · 禁止售卖",
            font=("微软雅黑", 7), bg=Theme.BG, fg=Theme.MUTED, bd=0)
        self._wm2.place(relx=0.01, rely=0.99, anchor="sw")
        self._wm2._wm_protected = True

        self._protected_title = "HNUST仿真平台 | 免费使用 禁止售卖"
        self.root.title(self._protected_title)
        self.root.after(5000, self._wm_check_once)

    def _wm_check_once(self):
        try:
            if not self.root.winfo_exists():
                return
            if not (hasattr(self, '_wm') and self._wm.winfo_exists()):
                self._wm = tk.Label(
                    self.root, text="该程序免费 禁止商用售卖",
                    font=Theme.FONT_TINY, bg=Theme.BG, fg=Theme.MUTED, bd=0)
                self._wm.place(relx=0.99, rely=0.99, anchor="se")
                self._wm._wm_protected = True
            if not (hasattr(self, '_wm2') and self._wm2.winfo_exists()):
                self._wm2 = tk.Label(
                    self.root, text="HNUST Exam · Free · 禁止售卖",
                    font=("微软雅黑", 7), bg=Theme.BG, fg=Theme.MUTED, bd=0)
                self._wm2.place(relx=0.01, rely=0.99, anchor="sw")
                self._wm2._wm_protected = True
            cur = self.root.title()
            if "免费" not in cur and "禁止" not in cur:
                self.root.title(self._protected_title)
        except tk.TclError:
            pass

    # ─────────────────────────────────────────────────────────────────
    #  辅助方法
    # ─────────────────────────────────────────────────────────────────
    def show_analysis(self):
        win = tk.Toplevel(self.root)
        win.title("试题解析")
        win.configure(bg=Theme.WHITE)
        win.resizable(False, False)
        W, H = 380, 220
        win.update_idletasks()
        x = (win.winfo_screenwidth() - W) // 2
        y = (win.winfo_screenheight() - H) // 2
        win.geometry(f"{W}x{H}+{x}+{y}")
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                win.iconbitmap(icon_path)
            except Exception:
                pass
        tk.Label(win, text="试题解析",
                 font=("微软雅黑", 14, "bold"),
                 bg=Theme.WHITE, fg=Theme.PRIMARY).pack(pady=(25, 10))
        tk.Label(win, text="暂时没有解析，问问豆包吧",
                 font=("微软雅黑", 12),
                 bg=Theme.WHITE, fg=Theme.TEXT).pack(pady=5)
        btn_frame = tk.Frame(win, bg=Theme.WHITE)
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="打开豆包",
                  font=("微软雅黑", 10, "bold"),
                  bg=Theme.PRIMARY, fg="white",
                  padx=15, pady=6, bd=0, cursor="hand2",
                  command=lambda: [webbrowser.open("https://www.doubao.com"),
                                   win.destroy()]).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="关闭",
                  font=("微软雅黑", 10),
                  bg=Theme.BG, fg=Theme.TEXT, padx=15, pady=6,
                  bd=1, relief=tk.SOLID, cursor="hand2",
                  command=win.destroy).pack(side=tk.LEFT, padx=10)
        win.grab_set()

    def _clear_window(self):
        for widget in self.root.winfo_children():
            if not getattr(widget, '_wm_protected', False):
                widget.destroy()

    def _clear_children(self, parent):
        for widget in parent.winfo_children():
            widget.destroy()

    # ─────────────────────────────────────────────────────────────────
    #  欢迎页
    # ─────────────────────────────────────────────────────────────────
    def create_welcome_window(self):
        self._current_view = "welcome"
        self._clear_window()

        title_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=80)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 24, "bold")).pack(side=tk.LEFT, padx=30,
                                                    pady=15)
        tk.Label(title_bar, text=CURRENT_VERSION,
                 bg=Theme.PRIMARY, fg=Theme.HEADER_SUB_TEXT,
                 font=("微软雅黑", 12)).pack(side=tk.LEFT, padx=10, pady=15)

        main_frame = tk.Frame(self.root, bg=Theme.WHITE, bd=1,
                              relief=tk.SOLID)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=100, pady=40)

        tk.Label(main_frame, text="欢迎使用 HNUST 考试仿真平台",
                 font=("微软雅黑", 18, "bold"), bg=Theme.WHITE,
                 fg=Theme.PRIMARY).pack(pady=(30, 20))

        canvas = tk.Canvas(main_frame, bg=Theme.WHITE, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical",
                                  command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=Theme.WHITE)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(40, 0))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 40))
        bind_mousewheel(scroll_frame, canvas)
        bind_mousewheel(canvas, canvas)

        sections = [
            ("软件介绍", [
                "本软件是模仿学校机房万维考试系统开发的免费练习工具",
                "专为HNUST同学设计，让你在宿舍也能随时随地进行考试模拟练习",
                "完美还原考试界面和操作流程，提前熟悉考试环境",
                "如若想使用程序设计、程序改错、填空三种功能，"
                "你的电脑需预先配置好 Python 运行环境。"
            ]),
            ("功能特点", [
                "支持单选、判断、填空、程序填空、程序改错、程序设计等多种题型",
                "自动判分，即时显示答题结果和正确答案",
                "一键打开程序文件，自动用IDLE编辑",
                "程序文件自动备份，支持一键重置",
                "题目导航，快速跳转到未答题目",
                "考试计时，时间到自动交卷",
                "试卷使用记录，查看练习历史",
                "支持深色模式和字体缩放",
            ]),
            ("开源声明", [
                "本软件完全免费开源，代码将托管在GitHub上",
                "任何人都可以自由下载、使用、修改和分发",
                "严禁任何形式的商用售卖，违者必究",
            ]),
            ("开发说明", [
                "本项目采用 AI 辅助开发模式完成",
                "特别感谢：Claude code 、豆包、DeepSeek、小米MIMO 提供的AI编程支持",
                "如果觉得好用，欢迎给个Star支持一下作者"
            ]),
            ("问题反馈", [
                "该应用为学生开发，可能存在一些bug和不完善的地方",
                "如果遇到任何问题或有改进建议",
                "欢迎通过GitHub提交Issue或在频道私信作者",
            ]),
            ("免责声明", [
                "本软件仅供学习交流使用，与学校官方考试系统无关",
                "题库内容由用户自行提供，作者不承担任何版权责任",
                "使用本软件产生的任何后果由用户自行承担"
            ])
        ]

        for title, content in sections:
            tk.Label(scroll_frame, text=title,
                     font=("微软雅黑", 14, "bold"), bg=Theme.WHITE,
                     fg=Theme.TEXT, anchor="w").pack(fill=tk.X, pady=(15, 8))
            for line in content:
                tk.Label(scroll_frame, text=line,
                         font=("微软雅黑", 11), bg=Theme.WHITE,
                         fg=Theme.TEXT, wraplength=800,
                         justify="left", anchor="w").pack(fill=tk.X, pady=2)

        bottom_frame = tk.Frame(self.root, bg=Theme.BG)
        bottom_frame.pack(fill=tk.X, padx=100, pady=(0, 40))

        self.agree_var = tk.BooleanVar(value=False)
        agree_frame = tk.Frame(bottom_frame, bg=Theme.BG)
        agree_frame.pack(side=tk.LEFT, padx=100, pady=10)

        BOX_SIZE = 26
        self.agree_canvas = tk.Canvas(agree_frame, width=BOX_SIZE,
                                      height=BOX_SIZE, bg=Theme.BG,
                                      highlightthickness=0)
        self.agree_canvas.pack(side=tk.LEFT, padx=(0, 10))

        def draw_box(checked=False):
            self.agree_canvas.delete("all")
            self.agree_canvas.create_rectangle(
                2, 2, BOX_SIZE - 2, BOX_SIZE - 2,
                outline=Theme.MUTED, width=2, fill=Theme.WHITE)
            if checked:
                self.agree_canvas.create_line(
                    7, 13, 11, 18, 19, 6,
                    width=3, fill=Theme.PRIMARY,
                    capstyle=tk.ROUND, joinstyle=tk.ROUND)

        draw_box(False)

        def toggle_agree(event=None):
            new_val = not self.agree_var.get()
            self.agree_var.set(new_val)
            draw_box(new_val)

        self.agree_canvas.bind("<Button-1>", toggle_agree)
        agree_label = tk.Label(agree_frame, text="我已阅读并同意以上所有条款",
                               font=("微软雅黑", 13, "bold"),
                               bg=Theme.BG, fg=Theme.TEXT, cursor="hand2")
        agree_label.pack(side=tk.LEFT)
        agree_label.bind("<Button-1>", toggle_agree)

        def enter_system():
            if not self.agree_var.get():
                messagebox.showwarning("提示", "请先阅读并同意以上条款")
                return
            self.create_select_window()

        self.enter_btn = tk.Button(
            bottom_frame, text="进入系统（请等待 10 秒）",
            font=("微软雅黑", 14, "bold"), command=enter_system,
            bg="#cccccc", fg="white", padx=40, pady=10,
            bd=0, cursor="arrow", state=tk.DISABLED)
        self.enter_btn.pack(side=tk.RIGHT, padx=20)
        self._welcome_countdown(10)

    def _welcome_countdown(self, remaining):
        if remaining > 0:
            self.enter_btn.config(text=f"进入系统（请等待 {remaining} 秒）")
            self.root.after(
                1000, lambda: self._welcome_countdown(remaining - 1))
        else:
            self.enter_btn.config(
                text="进入系统", bg=Theme.PRIMARY, fg="white",
                cursor="hand2", state=tk.NORMAL)

    # ─────────────────────────────────────────────────────────────────
    #  选卷页
    # ─────────────────────────────────────────────────────────────────
    def create_select_window(self):
        self._current_view = "select"
        self._clear_window()
        self._progress_data = _load_progress()

        title_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=60)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(side=tk.LEFT, padx=20,
                                                    pady=10)

        settings_btn = tk.Label(title_bar, text="⚙ 设置",
                                bg=Theme.PRIMARY, fg="white",
                                font=("微软雅黑", 11), cursor="hand2")
        settings_btn.pack(side=tk.RIGHT, padx=20, pady=10)
        settings_btn.bind("<Button-1>",
                          lambda e: self._show_settings())
        settings_btn.bind("<Enter>",
                          lambda e: settings_btn.config(fg=Theme.HEADER_SUB_TEXT))
        settings_btn.bind("<Leave>",
                          lambda e: settings_btn.config(fg="white"))

        main_frame = tk.Frame(self.root, bg=Theme.BG)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="请选择试卷",
                 font=("微软雅黑", 24, "bold"), bg=Theme.BG,
                 fg=Theme.TEXT).pack(pady=80)

        exam_dir = get_resource_path("题库")
        if not os.path.exists(exam_dir):
            external_exam_dir = "题库"
            if not os.path.exists(external_exam_dir):
                os.makedirs(external_exam_dir)
                messagebox.showinfo(
                    "提示", "题库文件夹已创建，请将Excel试卷文件放入其中")
            exam_dir = external_exam_dir

        self.exam_files = [f for f in os.listdir(exam_dir)
                           if f.endswith(".xlsx")]
        if not self.exam_files:
            tk.Label(main_frame, text="题库文件夹中没有找到试卷文件",
                     font=("微软雅黑", 14), fg="red",
                     bg=Theme.BG).pack(pady=20)
            return

        list_frame = tk.Frame(main_frame, bg=Theme.WHITE, bd=1,
                              relief=tk.SOLID)
        list_frame.pack(pady=20)

        self.exam_listbox = tk.Listbox(
            list_frame, font=("微软雅黑", 14), width=50, height=12,
            bd=0, highlightthickness=0, selectbackground=Theme.PRIMARY,
            selectforeground="white", bg=Theme.WHITE, fg=Theme.TEXT)

        for file in self.exam_files:
            display = os.path.splitext(file)[0]
            entry = self._progress_data.get(file, {})
            status = entry.get("status", "")
            if status == "completed":
                score = entry.get("best_score", "")
                display += f"  ✓ {score}%" if score else "  ✓"
            elif status == "started":
                display += "  ○"
            self.exam_listbox.insert(tk.END, display)

        self.exam_listbox.selection_set(0)
        self.exam_listbox.pack(padx=2, pady=2)
        self.exam_listbox.bind("<Double-Button-1>",
                               lambda e: self.start_exam())
        self.exam_listbox.bind("<Return>",
                               lambda e: self.start_exam())

        # 悬停提示
        self._exam_list_tooltip = None

        def _on_list_motion(event):
            idx = self.exam_listbox.nearest(event.y)
            if 0 <= idx < len(self.exam_files):
                file = self.exam_files[idx]
                entry = self._progress_data.get(file, {})
                lines = []
                if entry.get("last_completed"):
                    lines.append(f"上次完成：{entry['last_completed']}")
                if entry.get("best_score") is not None:
                    lines.append(f"最高得分：{entry['best_score']}%")
                if entry.get("last_started"):
                    lines.append(f"上次开始：{entry['last_started']}")
                if lines:
                    self._show_tooltip(event, "\n".join(lines))
                else:
                    self._hide_tooltip()
            else:
                self._hide_tooltip()

        self.exam_listbox.bind("<Motion>", _on_list_motion)
        self.exam_listbox.bind("<Leave>",
                               lambda e: self._hide_tooltip())

        tk.Button(main_frame, text="开始考试",
                  font=("微软雅黑", 16, "bold"), command=self.start_exam,
                  bg=Theme.PRIMARY, fg="white",
                  padx=40, pady=12, bd=0,
                  cursor="hand2").pack(pady=(40, 10))

        tk.Label(main_frame,
                 text="该程序免费提供给HNUST学生使用，禁止任何形式的商用售卖",
                 font=("微软雅黑", 8), bg=Theme.BG,
                 fg=Theme.MUTED).pack(side=tk.BOTTOM, pady=10)

    def _show_mark_legend(self):
        win = tk.Toplevel(self.root)
        win.title("标记说明")
        win.configure(bg=Theme.WHITE)
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                win.iconbitmap(icon_path)
            except Exception:
                pass

        hdr = tk.Frame(win, bg=Theme.PRIMARY, height=50)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="标记说明",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 14, "bold")).pack(
            side=tk.LEFT, padx=20, pady=10)

        body = tk.Frame(win, bg=Theme.WHITE)
        body.pack(fill=tk.BOTH, padx=24, pady=20)

        items = [
            ("✓", "已完成", "已交卷的试卷，后面显示最高得分", Theme.SUCCESS),
            ("○", "进行中", "已开始但尚未交卷的试卷", Theme.PRIMARY),
            ("（无标记）", "未开始", "尚未打开过的试卷", Theme.MUTED),
        ]

        for icon, label, desc, color in items:
            row = tk.Frame(body, bg=Theme.WHITE)
            row.pack(fill=tk.X, pady=8)
            tk.Label(row, text=icon, font=("微软雅黑", 16),
                     bg=Theme.WHITE, fg=color, width=4,
                     anchor="center").pack(side=tk.LEFT)
            col = tk.Frame(row, bg=Theme.WHITE)
            col.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(col, text=label, font=("微软雅黑", 12, "bold"),
                     bg=Theme.WHITE, fg=Theme.TEXT,
                     anchor="w").pack(fill=tk.X)
            tk.Label(col, text=desc, font=("微软雅黑", 10),
                     bg=Theme.WHITE, fg=Theme.MUTED,
                     anchor="w").pack(fill=tk.X)

        tk.Button(win, text="知道了",
                  font=("微软雅黑", 11), bg=Theme.PRIMARY, fg="white",
                  padx=30, pady=6, bd=0, cursor="hand2",
                  command=win.destroy).pack(pady=(0, 20))

        win.update_idletasks()
        W = max(win.winfo_reqwidth() + 40, 360)
        H = win.winfo_reqheight() + 20
        x = (win.winfo_screenwidth() - W) // 2
        y = (win.winfo_screenheight() - H) // 2
        win.geometry(f"{W}x{H}+{x}+{y}")

    def _clear_all_marks(self):
        if not self._progress_data:
            messagebox.showinfo("提示", "当前没有任何标记记录")
            return False
        if messagebox.askyesno(
                "确认清除",
                "确定要清除所有试卷的完成标记吗？\n\n"
                "清除后，所有试卷的完成状态和得分记录将被重置。\n"
                "此操作不可撤销。"):
            _save_progress({})
            self._progress_data = {}
            messagebox.showinfo("完成", "所有标记已清除")
            return True
        return False

    # ─────────────────────────────────────────────────────────────────
    #  设置页
    # ─────────────────────────────────────────────────────────────────
    def _show_settings(self):
        win = tk.Toplevel(self.root)
        win.title("个性化设置")
        win.configure(bg=Theme.WHITE)
        win.resizable(False, False)

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                win.iconbitmap(icon_path)
            except Exception:
                pass

        orig_dark = Theme._is_dark
        orig_scale = Theme._font_scale
        marks_cleared = {"v": False}

        header = tk.Frame(win, bg=Theme.PRIMARY, height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="⚙ 个性化设置",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(side=tk.LEFT, padx=20,
                                                    pady=15)

        body = tk.Frame(win, bg=Theme.WHITE)
        body.pack(fill=tk.BOTH, expand=True, padx=24, pady=24)

        # 卡片1：即时反馈
        card1 = self._make_settings_card(
            body, "答题后立即显示对错",
            "开启后，单选题和判断题选择答案后立即显示对错及正确答案\n"
            "关闭后，仅记录答案，交卷后统一评判")
        state1 = {"on": self.show_answer_immediately}
        self._make_toggle(card1, state1,
                          lambda v: setattr(self, 'show_answer_immediately',
                                            v))

        # 卡片2：深色模式
        card2 = self._make_settings_card(
            body, "深色模式",
            "切换深色/浅色主题，减少视觉疲劳\n"
            "更改将在关闭设置后生效")
        state2 = {"on": Theme._is_dark}
        self._make_toggle(card2, state2,
                          lambda v: Theme.set_dark_mode(v))

        # 卡片3：字体缩放
        card3 = tk.Frame(body, bg=Theme.SURFACE,
                         highlightbackground=Theme.BORDER,
                         highlightthickness=1)
        card3.pack(fill=tk.X, pady=(0, 12))

        text3 = tk.Frame(card3, bg=Theme.SURFACE)
        text3.pack(side=tk.LEFT, padx=20, pady=18, fill=tk.X, expand=True)
        tk.Label(text3, text="字体大小",
                 font=("微软雅黑", 12, "bold"), bg=Theme.SURFACE,
                 fg=Theme.TEXT, anchor="w").pack(fill=tk.X)
        tk.Label(text3, text="调整界面文字大小（80% ~ 150%）",
                 font=("微软雅黑", 9), bg=Theme.SURFACE, fg=Theme.MUTED,
                 anchor="w").pack(fill=tk.X, pady=(8, 0))

        # 字体预览区域
        preview_frame = tk.Frame(body, bg=Theme.SURFACE,
                                 highlightbackground=Theme.BORDER,
                                 highlightthickness=1)
        preview_frame.pack(fill=tk.X, pady=(0, 12))
        tk.Label(preview_frame, text="预览效果",
                 font=("微软雅黑", 9, "bold"), bg=Theme.SURFACE,
                 fg=Theme.MUTED, anchor="w").pack(
            fill=tk.X, padx=16, pady=(10, 4))
        preview_title = tk.Label(
            preview_frame, text="这是标题文字的预览效果",
            font=Theme.FONT_TITLE, bg=Theme.SURFACE,
            fg=Theme.TEXT, anchor="w")
        preview_title.pack(fill=tk.X, padx=16, pady=2)
        preview_body = tk.Label(
            preview_frame, text="这是正文内容的预览效果，用于确认字体大小是否合适",
            font=Theme.FONT, bg=Theme.SURFACE,
            fg=Theme.TEXT, anchor="w")
        preview_body.pack(fill=tk.X, padx=16, pady=2)
        preview_small = tk.Label(
            preview_frame, text="这是小字标注的预览效果",
            font=Theme.FONT_SMALL, bg=Theme.SURFACE,
            fg=Theme.MUTED, anchor="w")
        preview_small.pack(fill=tk.X, padx=16, pady=(2, 10))

        scale_frame = tk.Frame(card3, bg=Theme.SURFACE)
        scale_frame.pack(side=tk.RIGHT, padx=20, pady=18)
        scale_var = tk.DoubleVar(value=Theme._font_scale * 100)
        scale_label = tk.Label(scale_frame,
                               text=f"{int(scale_var.get())}%",
                               font=("微软雅黑", 11, "bold"),
                               bg=Theme.SURFACE, fg=Theme.TEXT, width=5)
        scale_label.pack()

        def _on_scale(val):
            pct = float(val)
            Theme._font_scale = pct / 100.0
            Theme._update_fonts()
            scale_label.config(text=f"{int(pct)}%")
            preview_title.config(font=Theme.FONT_TITLE)
            preview_body.config(font=Theme.FONT)
            preview_small.config(font=Theme.FONT_SMALL)

        tk.Scale(scale_frame, from_=80, to=150, orient=tk.HORIZONTAL,
                 variable=scale_var, command=_on_scale,
                 length=150, bg=Theme.SURFACE, fg=Theme.TEXT,
                 highlightthickness=0,
                 troughcolor=Theme.BORDER).pack()

        # 卡片4：试卷标记管理
        card4 = tk.Frame(body, bg=Theme.SURFACE,
                         highlightbackground=Theme.BORDER,
                         highlightthickness=1)
        card4.pack(fill=tk.X, pady=(0, 12))
        text4 = tk.Frame(card4, bg=Theme.SURFACE)
        text4.pack(side=tk.LEFT, padx=20, pady=18, fill=tk.X, expand=True)
        tk.Label(text4, text="试卷标记管理",
                 font=("微软雅黑", 12, "bold"), bg=Theme.SURFACE,
                 fg=Theme.TEXT, anchor="w").pack(fill=tk.X)
        tk.Label(text4, text="查看标记含义说明，或清除所有试卷的完成标记",
                 font=("微软雅黑", 9), bg=Theme.SURFACE, fg=Theme.MUTED,
                 anchor="w").pack(fill=tk.X, pady=(8, 0))
        mark_btn_frame = tk.Frame(card4, bg=Theme.SURFACE)
        mark_btn_frame.pack(side=tk.RIGHT, padx=20, pady=12)
        tk.Button(mark_btn_frame, text="标记说明",
                  font=("微软雅黑", 9), bg=Theme.SURFACE, fg=Theme.TEXT,
                  bd=1, relief=tk.SOLID, padx=12, pady=3,
                  cursor="hand2",
                  command=self._show_mark_legend).pack(side=tk.LEFT, padx=4)
        tk.Button(mark_btn_frame, text="清除所有标记",
                  font=("微软雅黑", 9), bg=Theme.SURFACE, fg=Theme.DANGER,
                  bd=1, relief=tk.SOLID, padx=12, pady=3,
                  cursor="hand2",
                  command=lambda: marks_cleared.update(
                      {"v": True}) if self._clear_all_marks() else None
                  ).pack(side=tk.LEFT, padx=4)

        # 状态提示
        status_hint = tk.Label(body, text="", font=("微软雅黑", 9),
                               bg=Theme.WHITE, fg=Theme.MUTED)
        status_hint.pack(fill=tk.X, pady=(5, 0))

        def _update_hint():
            try:
                if not win.winfo_exists():
                    return
                mode = "即时反馈" if self.show_answer_immediately \
                    else "考试模式"
                theme = "深色" if Theme._is_dark else "浅色"
                status_hint.config(
                    text=f"答题：{mode}  |  主题：{theme}  |  "
                         f"字体：{int(Theme._font_scale * 100)}%",
                    fg=Theme.SUCCESS if self.show_answer_immediately
                    else Theme.PRIMARY)
                win.after(300, _update_hint)
            except tk.TclError:
                pass

        _update_hint()

        # 底部按钮
        btn_bar = tk.Frame(win, bg=Theme.WHITE)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=24, pady=(0, 20))

        def _on_done():
            _save_config({
                "font_scale": Theme._font_scale,
                "dark_mode": Theme._is_dark,
                "show_answer_immediately": self.show_answer_immediately,
                "user_python_path": self._user_python_path,
            })
            win.destroy()
            if Theme._is_dark != orig_dark or abs(
                    Theme._font_scale - orig_scale) > 0.01 or marks_cleared["v"]:
                self._recreate_current_view()

        def _on_cancel():
            if abs(Theme._font_scale - orig_scale) > 0.01:
                Theme._font_scale = orig_scale
                Theme._update_fonts()
            if Theme._is_dark != orig_dark:
                Theme.set_dark_mode(orig_dark)
            win.destroy()
            if marks_cleared["v"]:
                self._recreate_current_view()

        done_btn = tk.Button(btn_bar, text="完 成",
                             font=("微软雅黑", 11),
                             bg=Theme.PRIMARY, fg="white",
                             activeforeground="white",
                             activebackground=Theme.ACCENT,
                             relief=tk.FLAT, padx=36, pady=7,
                             cursor="hand2", command=_on_done)
        done_btn.pack(side=tk.RIGHT)
        done_btn.bind("<Enter>",
                      lambda e: done_btn.config(bg=Theme.ACCENT))
        done_btn.bind("<Leave>",
                      lambda e: done_btn.config(bg=Theme.PRIMARY))

        win.protocol("WM_DELETE_WINDOW", _on_cancel)

        win.withdraw()
        win.update_idletasks()
        W = max(win.winfo_reqwidth() + 40, 640)
        H = win.winfo_reqheight() + 20
        x = (win.winfo_screenwidth() - W) // 2
        y = (win.winfo_screenheight() - H) // 2
        win.geometry(f"{W}x{H}+{x}+{y}")
        win.deiconify()

        win.grab_set()
        win.focus_force()

    def _make_settings_card(self, parent, title, desc):
        card = tk.Frame(parent, bg=Theme.SURFACE,
                        highlightbackground=Theme.BORDER,
                        highlightthickness=1)
        card.pack(fill=tk.X, pady=(0, 12))
        tf = tk.Frame(card, bg=Theme.SURFACE)
        tf.pack(side=tk.LEFT, padx=20, pady=18, fill=tk.X, expand=True)
        tk.Label(tf, text=title,
                 font=("微软雅黑", 12, "bold"), bg=Theme.SURFACE,
                 fg=Theme.TEXT, anchor="w").pack(fill=tk.X)
        tk.Label(tf, text=desc,
                 font=("微软雅黑", 9), bg=Theme.SURFACE, fg=Theme.MUTED,
                 anchor="w", justify="left").pack(fill=tk.X, pady=(8, 0))
        return card

    def _make_toggle(self, parent, state, on_change):
        frame = tk.Frame(parent, bg=Theme.SURFACE, width=76, height=44)
        frame.pack(side=tk.RIGHT, padx=20, pady=18)
        frame.pack_propagate(False)

        if _HAS_PIL:
            SS, TW, TH = 3, 52, 26

            def _render(is_on):
                w, h = TW * SS, TH * SS
                img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
                draw = ImageDraw.Draw(img)
                tc = (76, 175, 80) if is_on else (189, 189, 189)
                draw.rounded_rectangle([0, 0, w - 1, h - 1],
                                       radius=h // 2, fill=tc)
                pad = 3 * SS
                diameter = h - 2 * pad
                r = diameter // 2
                cx = (w - pad - r) if is_on else (pad + r)
                cy = h // 2
                draw.ellipse([cx - r + SS, cy - r + SS,
                              cx + r + SS, cy + r + SS],
                             fill=(0, 0, 0, 22))
                draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                             fill="white")
                img = img.resize((TW, TH), Image.LANCZOS)
                return ImageTk.PhotoImage(img)

            photo = _render(state["on"])
            lbl = tk.Label(frame, image=photo, bg=Theme.SURFACE,
                           cursor="hand2")
            lbl.image = photo
            lbl.place(relx=0.5, rely=0.5, anchor="center")

            def _flip(event=None):
                state["on"] = not state["on"]
                on_change(state["on"])
                p = _render(state["on"])
                lbl.configure(image=p)
                lbl.image = p

            lbl.bind("<Button-1>", _flip)
        else:
            var = tk.BooleanVar(value=state["on"])

            def _toggle():
                state["on"] = var.get()
                on_change(state["on"])

            tk.Checkbutton(frame, variable=var, command=_toggle,
                           bg=Theme.SURFACE,
                           activebackground=Theme.SURFACE).place(
                relx=0.5, rely=0.5, anchor="center")

    def _recreate_current_view(self):
        self.root.configure(bg=Theme.BG)
        self._apply_ttk_styles()
        if self._current_view == "exam":
            messagebox.showinfo("提示", "主题更改将在返回主菜单后生效")
        elif self._current_view == "select":
            self.create_select_window()
        else:
            self.create_welcome_window()

    # ─────────────────────────────────────────────────────────────────
    #  试卷进度记录
    # ─────────────────────────────────────────────────────────────────
    def _save_exam_progress(self, status, score_pct=None):
        progress = _load_progress()
        exam_key = os.path.basename(self.current_exam_file)
        entry = progress.get(exam_key, {})
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        if status == "completed":
            entry["status"] = "completed"
            entry["last_completed"] = now_str
            if score_pct is not None:
                prev_best = entry.get("best_score", 0)
                entry["best_score"] = max(prev_best, round(score_pct, 1))
        elif status == "started":
            if entry.get("status") != "completed":
                entry["status"] = "started"
            entry["last_started"] = now_str

        progress[exam_key] = entry
        _save_progress(progress)
        self._progress_data = progress

    # ─────────────────────────────────────────────────────────────────
    #  开始考试
    # ─────────────────────────────────────────────────────────────────
    def start_exam(self):
        selected = self.exam_listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一份试卷")
            return

        exam_file = self.exam_files[selected[0]]
        exam_name = os.path.splitext(exam_file)[0]

        internal_exam_file = get_resource_path(
            os.path.join("题库", exam_name + ".xlsx"))
        external_exam_file = os.path.join("题库", exam_name + ".xlsx")

        if os.path.exists(internal_exam_file):
            self.current_exam_file = internal_exam_file
        elif os.path.exists(external_exam_file):
            self.current_exam_file = external_exam_file
        else:
            messagebox.showerror("错误", "找不到试卷文件")
            return

        try:
            df = pd.read_excel(self.current_exam_file)
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            for col in df.columns:
                df[col] = df[col].astype(str).str.strip()
            df = df[df["题号"] != ""]
            df = df[df["题目"] != ""]

            required_cols = {"题号", "题型", "题目", "正确答案", "分值"}
            missing = required_cols - set(df.columns)
            if missing:
                messagebox.showerror(
                    "错误",
                    f"Excel 缺少必要列：{', '.join(missing)}\n"
                    f"当前列：{', '.join(df.columns)}")
                return

            if "程序文件" in df.columns:
                df["程序文件"] = df["程序文件"].astype(str).str.strip()
            else:
                df["程序文件"] = ""

            dup_nums = df[df.duplicated(subset=["题号"], keep=False)][
                "题号"].unique()
            if len(dup_nums) > 0:
                preview = ", ".join(dup_nums[:5].tolist())
                messagebox.showwarning(
                    "警告",
                    f"发现重复题号：{preview}"
                    f"{'...' if len(dup_nums) > 5 else ''}\n"
                    "可能导致答案被覆盖，建议检查Excel文件。")

            self.questions = df.to_dict("records")

            all_question_types = set(df["题型"].unique())
            if all_question_types == {"程序设计"}:
                self.is_pure_program_exam = True
                self.question_groups = {"程序设计": []}
                for idx, q in enumerate(self.questions):
                    q["_global_idx"] = idx
                    self.question_groups["程序设计"].append(q)
            else:
                self.is_pure_program_exam = False
                self.question_groups = {}
                for idx, q in enumerate(self.questions):
                    q["_global_idx"] = idx
                    q_type = q["题型"]
                    self.question_groups.setdefault(q_type, []).append(q)

            self.active_type_order = []
            for q in self.questions:
                t = q["题型"]
                if t not in self.active_type_order:
                    self.active_type_order.append(t)

            self.user_answers = {}
            self.marked_questions = set()
            self.current_index = 0
            self.score = 0
            self.remaining_time = self.exam_time
            self.exam_submitted = False
            self._stop_timer = Event()

            program_types = {"程序设计", "程序填空", "程序改错"}
            has_program = any(
                q["题型"] in program_types for q in self.questions)
            if has_program:
                py = self._user_python_path or find_system_python()
                if not py or not os.path.isfile(py):
                    result = self._ask_user_for_python()
                    if result:
                        self._user_python_path = result

            self._init_backup()
            self._save_exam_progress("started")
            self.create_exam_window()
            self.start_timer()

        except Exception as e:
            messagebox.showerror("错误", f"读取试卷失败：{str(e)}")

    def _cleanup_backup(self):
        if self._backup_dir and os.path.exists(self._backup_dir):
            try:
                shutil.rmtree(self._backup_dir)
            except Exception:
                pass
        self._backup_dir = None

    def _init_backup(self):
        exam_dir = os.path.dirname(self.current_exam_file)
        source_dir = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(source_dir):
            source_dir = exam_dir
        self._backup_dir = os.path.join(exam_dir, "_backup_programs")
        if os.path.exists(self._backup_dir):
            shutil.rmtree(self._backup_dir)
        os.makedirs(self._backup_dir, exist_ok=True)
        referenced_files = set()
        for q in self.questions:
            pf = q.get("程序文件", "").strip()
            if pf:
                referenced_files.add(pf)
        for pf in referenced_files:
            src = os.path.join(source_dir, pf)
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(self._backup_dir, pf))

    # ─────────────────────────────────────────────────────────────────
    #  考试界面（左右键导航 + 深色适配）
    # ─────────────────────────────────────────────────────────────────
    def create_exam_window(self):
        self._current_view = "exam"
        self._clear_window()

        # 顶部栏
        top_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=50)
        top_bar.pack(fill=tk.X)
        tk.Label(top_bar, text="HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 12, "bold")).pack(side=tk.LEFT, padx=15,
                                                    pady=10)
        info_frame = tk.Frame(top_bar, bg=Theme.PRIMARY)
        info_frame.pack(side=tk.RIGHT, padx=15)
        exam_name = os.path.splitext(
            os.path.basename(self.current_exam_file))[0]
        tk.Label(info_frame, text="姓名：xxx  学号：xxxxxxxxxxx",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 10)).pack(anchor="e")
        tk.Label(info_frame, text=f"{exam_name} · 练习",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 10)).pack(anchor="e")

        # 进度条
        progress_area = tk.Frame(self.root, bg=Theme.BG)
        progress_area.pack(fill=tk.X, padx=10, pady=(5, 0))
        self.progress_label = tk.Label(
            progress_area,
            text=f"已完成 0 / {len(self.questions)} 题",
            font=Theme.FONT_SMALL, bg=Theme.BG, fg=Theme.MUTED)
        self.progress_label.pack(side=tk.RIGHT, padx=(10, 0))
        self.progress_frame = tk.Frame(progress_area, bg=Theme.PROGRESS_BG,
                                       height=8)
        self.progress_frame.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.progress_bar = tk.Frame(self.progress_frame, bg=Theme.SUCCESS,
                                     height=8)
        self.progress_bar.place(x=0, y=0, relheight=1.0, relwidth=0)

        main_frame = tk.Frame(self.root, bg=Theme.BG)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧题目区
        self.left_frame = tk.Frame(main_frame, bg=Theme.WHITE, bd=1,
                                   relief=tk.SOLID)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                             padx=10, pady=10)

        kb_hint = tk.Frame(self.left_frame, bg=Theme.KB_HINT_BG, height=20)
        kb_hint.pack(fill=tk.X, side=tk.BOTTOM)
        kb_hint.pack_propagate(False)
        tk.Label(kb_hint,
                 text="快捷键  ← → 切换题目  |  "
                      "Ctrl+N 下一未答  |  Ctrl+A 查看答案  |  "
                      "A B C D 选择选项",
                 font=("微软雅黑", 8), bg=Theme.KB_HINT_BG,
                 fg=Theme.MUTED,
                 anchor="center").pack(fill=tk.BOTH, expand=True)

        left_scroll_area = tk.Frame(self.left_frame, bg=Theme.WHITE)
        left_scroll_area.pack(fill=tk.BOTH, expand=True)

        self._lp_canvas = tk.Canvas(left_scroll_area, bg=Theme.WHITE,
                                    highlightthickness=0)
        lp_scrollbar = ttk.Scrollbar(left_scroll_area, orient="vertical",
                                     command=self._lp_canvas.yview)
        self._lp_canvas.configure(yscrollcommand=lp_scrollbar.set)
        self._lp_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lp_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._lp_inner = tk.Frame(self._lp_canvas, bg=Theme.WHITE)
        lp_canvas_win = self._lp_canvas.create_window(
            (0, 0), window=self._lp_inner, anchor="nw")

        self._lp_cached_content_h = 1
        self._lp_cached_visible_h = 1

        def _on_lp_canvas_resize(e):
            self._lp_canvas.itemconfig(lp_canvas_win, width=e.width)
            self._lp_cached_visible_h = e.height

        self._lp_canvas.bind("<Configure>", _on_lp_canvas_resize)

        def _on_lp_inner_resize(e):
            self._lp_canvas.configure(
                scrollregion=self._lp_canvas.bbox("all"))
            self._lp_cached_content_h = e.height

        self._lp_inner.bind("<Configure>", _on_lp_inner_resize)

        # 右侧导航
        self.right_frame = tk.Frame(main_frame, bg=Theme.WHITE, bd=1,
                                    relief=tk.SOLID, width=250)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y,
                              padx=(0, 10), pady=10)
        self.right_frame.pack_propagate(False)
        self._build_nav_panels()

        # 题目内容区
        self.q_title_bar = tk.Frame(self._lp_inner, bg=Theme.PRIMARY)
        self.q_title_bar.pack(fill=tk.X, padx=20, pady=(20, 10))
        self.q_instruction = tk.Frame(self._lp_inner, bg=Theme.WHITE)
        self.q_content = tk.Frame(self._lp_inner, bg=Theme.WHITE)
        self.q_content.pack(fill=tk.X, padx=20, pady=(10, 5))
        self.q_answer_area = tk.Frame(self._lp_inner, bg=Theme.WHITE)
        self.q_answer_area.pack(fill=tk.X, padx=20, pady=10)
        self.q_feedback = tk.Frame(self._lp_inner, bg=Theme.WHITE)
        self.q_feedback.pack(fill=tk.X, padx=20, pady=5)

        self._setup_left_scroll()

        # 底部按钮栏
        bottom_bar = tk.Frame(self.root, bg=Theme.BG, height=60)
        bottom_bar.pack(fill=tk.X, padx=10, pady=(0, 10))
        bottom_bar.pack_propagate(False)

        btn_style = {"font": ("微软雅黑", 10), "bd": 1,
                     "relief": tk.SOLID, "cursor": "hand2",
                     "padx": 10, "pady": 5}
        left_btns = tk.Frame(bottom_bar, bg=Theme.BG)
        left_btns.pack(side=tk.LEFT, padx=5)

        button_configs = [
            (0, "上题", self.prev_question, Theme.BG, Theme.TEXT),
            (1, "下题", self.next_question, Theme.BG, Theme.TEXT),
            (2, "下一未答", self.jump_next_unanswered,
             Theme.ACCENT, "white"),
            (3, "答题", self.open_program_file, Theme.ACCENT, "white"),
            (4, "试题文件夹", self.open_exam_folder, Theme.BG, Theme.TEXT),
            (5, "重做", self.redo_question, Theme.BG, Theme.TEXT),
            (6, "标记试题", self.toggle_mark, Theme.BG, Theme.TEXT),
            (7, "答案", self.show_answer, Theme.BG, Theme.TEXT),
            (8, "试题解析", self.show_analysis, Theme.BG, Theme.TEXT),
        ]

        for col, text, cmd, bg, fg in button_configs:
            btn = tk.Button(left_btns, text=text, **btn_style,
                            bg=bg, fg=fg, command=cmd)
            btn.grid(row=0, column=col, padx=5)
            if text == "答题":
                self.answer_btn = btn
            elif text == "标记试题":
                self.mark_btn = btn

        right_bottom = tk.Frame(bottom_bar, bg=Theme.BG)
        right_bottom.pack(side=tk.RIGHT)
        self.time_label = tk.Label(right_bottom, text="01:00:00",
                                   font=("微软雅黑", 14, "bold"),
                                   bg=Theme.BG, fg=Theme.TEXT)
        self.time_label.pack(side=tk.LEFT, padx=20)
        tk.Button(right_bottom, text="交卷",
                  font=("微软雅黑", 12, "bold"),
                  bg=Theme.DANGER, fg="white", bd=0,
                  padx=25, pady=8, cursor="hand2",
                  command=self.submit_exam).pack(side=tk.RIGHT)

        tk.Label(bottom_bar, text="免费使用 禁止售卖",
                 font=Theme.FONT_TINY, bg=Theme.BG,
                 fg=Theme.MUTED).pack(side=tk.LEFT, padx=10)

        self.answer_var = tk.StringVar()
        self.answer_text = None
        self.answer_entry = None
        self._choice_buttons = {}

        # 快捷键绑定
        self.root.bind("<Left>", self._on_left_arrow)
        self.root.bind("<Right>", self._on_right_arrow)
        self.root.bind("<Control-Left>",
                       lambda e: self.prev_question())
        self.root.bind("<Control-Right>",
                       lambda e: self.next_question())
        self.root.bind("<Control-Return>",
                       lambda e: self.submit_exam())
        self.root.bind("<Control-n>",
                       lambda e: self.jump_next_unanswered())
        self.root.bind("<Control-a>",
                       lambda e: self.show_answer())
        self.root.bind("<KeyPress>", self._on_key_press)

        self.show_question()

    # ─────────────────────────────────────────────────────────────────
    #  键盘事件
    # ─────────────────────────────────────────────────────────────────
    def _on_key_press(self, event):
        if self.exam_submitted:
            return
        focused = self.root.focus_get()
        if focused is not None:
            w_class = focused.winfo_class()
            if w_class in ('Entry', 'TEntry', 'Text', 'Spinbox',
                           'TSpinbox', 'Combobox', 'TCombobox'):
                return
        if not self.questions:
            return
        current_q = self.questions[self.current_index]
        q_type = current_q["题型"]
        if q_type == "单选":
            key = event.char.upper()
            opts = [o[0] for o in
                    self._get_options_for_question(current_q)]
            if key in opts:
                self._choose(current_q["题号"], key)
        elif q_type == "判断":
            if event.char.lower() in ("t", "y", "1"):
                self._choose(current_q["题号"], "A")
            elif event.char.lower() in ("f", "n", "0"):
                self._choose(current_q["题号"], "B")

    def _on_left_arrow(self, event):
        if self._current_view != "exam" or self.exam_submitted:
            return
        focused = self.root.focus_get()
        if focused is not None:
            w_class = focused.winfo_class()
            if w_class in ('Entry', 'TEntry', 'Text', 'Spinbox',
                           'TSpinbox', 'Combobox', 'TCombobox'):
                return
        self.prev_question()
        return "break"

    def _on_right_arrow(self, event):
        if self._current_view != "exam" or self.exam_submitted:
            return
        focused = self.root.focus_get()
        if focused is not None:
            w_class = focused.winfo_class()
            if w_class in ('Entry', 'TEntry', 'Text', 'Spinbox',
                           'TSpinbox', 'Combobox', 'TCombobox'):
                return
        self.next_question()
        return "break"

    # ─────────────────────────────────────────────────────────────────
    #  程序文件操作
    # ─────────────────────────────────────────────────────────────────
    def open_program_file(self):
        current_q = self.questions[self.current_index]
        program_file = current_q.get("程序文件", "").strip()
        if not program_file:
            messagebox.showinfo("提示", "该题目没有对应的程序文件")
            return

        if ".." in program_file:
            messagebox.showerror("错误", "文件路径中不允许包含 '..'")
            return
        if program_file.startswith(("/", "\\")):
            messagebox.showerror("错误", "不允许使用绝对路径")
            return
        if ":" in program_file:
            messagebox.showerror(
                "错误",
                f"不允许使用绝对路径（包含盘符）：{program_file}")
            return
        if program_file.startswith("\\\\"):
            messagebox.showerror("错误", "不允许使用网络路径")
            return

        exam_dir = os.path.dirname(self.current_exam_file)
        base_dir = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(base_dir):
            base_dir = exam_dir

        program_path = os.path.normpath(
            os.path.join(base_dir, program_file))

        real_base = os.path.realpath(base_dir)
        real_path = os.path.realpath(program_path)
        if not (real_path.startswith(real_base + os.sep)
                or real_path == real_base):
            messagebox.showerror("错误",
                                 f"文件路径越界：{program_file}")
            return

        if not os.path.exists(program_path):
            alt_path = os.path.normpath(
                os.path.join(exam_dir, program_file))
            real_alt = os.path.realpath(alt_path)
            if (real_alt.startswith(
                    os.path.realpath(exam_dir) + os.sep)
                    and os.path.exists(alt_path)):
                program_path = alt_path
            else:
                messagebox.showerror(
                    "错误",
                    f"找不到程序文件：{program_file}\n"
                    f"请确保文件放在Excel同目录的\"试题文件夹\"中"
                    f"或Excel同目录下")
                return

        self.answer_btn.config(state=tk.DISABLED, text="启动中...")

        progress_win = tk.Toplevel(self.root)
        progress_win.title("启动中")
        progress_win.configure(bg=Theme.WHITE)
        progress_win.resizable(False, False)
        progress_win.attributes("-topmost", True)
        pw_w, pw_h = 300, 100
        progress_win.update_idletasks()
        px = (progress_win.winfo_screenwidth() - pw_w) // 2
        py = (progress_win.winfo_screenheight() - pw_h) // 2
        progress_win.geometry(f"{pw_w}x{pw_h}+{px}+{py}")
        tk.Label(progress_win, text="正在启动编辑器，请稍候...",
                 font=Theme.FONT, bg=Theme.WHITE,
                 fg=Theme.TEXT).pack(pady=(25, 10))

        def _open_thread():
            opened_with_idle = False
            success = False
            error_msg = ""
            try:
                if program_file.lower().endswith(".py"):
                    opened_with_idle = self._try_open_with_idle(
                        program_path)
                    if opened_with_idle:
                        success = True
                if not success:
                    if sys.platform == "win32":
                        os.startfile(program_path)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", program_path],
                                       check=True)
                    else:
                        subprocess.run(["xdg-open", program_path],
                                       check=True)
                    success = True
            except Exception as e:
                error_msg = str(e)

            def _on_done():
                try:
                    progress_win.destroy()
                except tk.TclError:
                    pass
                self.answer_btn.config(state=tk.NORMAL, text="答题")
                if success and opened_with_idle:
                    messagebox.showinfo(
                        "提示",
                        f"已用IDLE打开：{program_file}\n"
                        "修改完成后按Ctrl+S保存，"
                        "然后回到本系统输入答案")
                elif success:
                    messagebox.showinfo(
                        "提示",
                        f"已用默认程序打开：{program_file}\n"
                        "（未检测到Python IDLE）\n"
                        "修改完成后保存文件，然后回到本系统输入答案")
                else:
                    messagebox.showerror(
                        "错误",
                        f"打开文件失败：{error_msg or '未知错误'}\n\n"
                        f"文件路径：{program_path}")

            self.root.after(0, _on_done)

        Thread(target=_open_thread, daemon=True).start()

    def _try_open_with_idle(self, file_path):
        abs_path = os.path.abspath(file_path)
        NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        user_py = self._user_python_path
        if user_py and os.path.isfile(user_py):
            try:
                subprocess.Popen([user_py, "-m", "idlelib", abs_path],
                                 creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        if not hasattr(sys, '_MEIPASS'):
            try:
                subprocess.Popen(
                    [sys.executable, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        try:
            if shutil.which("py"):
                check = subprocess.run(
                    ["py", "-3", "--version"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=NO_WINDOW)
                if check.returncode == 0 and "Python" in check.stdout:
                    subprocess.Popen(
                        ["py", "-3", "-m", "idlelib", abs_path],
                        creationflags=NO_WINDOW)
                    return True
        except Exception:
            pass

        for cmd_name in ["python", "python3"]:
            python_path = shutil.which(cmd_name)
            if python_path:
                try:
                    check = subprocess.run(
                        [python_path, "--version"],
                        capture_output=True, text=True, timeout=5,
                        creationflags=NO_WINDOW)
                    if check.returncode == 0 \
                            and "Python" in check.stdout:
                        subprocess.Popen(
                            [python_path, "-m", "idlelib", abs_path],
                            creationflags=NO_WINDOW)
                        return True
                except Exception:
                    continue

        python_exe = find_system_python()
        if python_exe:
            try:
                subprocess.Popen(
                    [python_exe, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass
            python_dir = os.path.dirname(python_exe)
            idle_pyw = os.path.join(python_dir, "Lib", "idlelib",
                                    "idle.pyw")
            if os.path.isfile(idle_pyw):
                try:
                    subprocess.Popen(
                        [python_exe, idle_pyw, abs_path],
                        creationflags=NO_WINDOW)
                    return True
                except Exception:
                    pass

        return False

    def _ask_user_for_python(self):
        win = tk.Toplevel(self.root)
        win.title("Python 环境配置")
        win.configure(bg=Theme.WHITE)
        win.resizable(False, False)

        W, H = 480, 380
        win.update_idletasks()
        x = (win.winfo_screenwidth() - W) // 2
        y = (win.winfo_screenheight() - H) // 2
        win.geometry(f"{W}x{H}+{x}+{y}")

        result = {"path": None}

        hdr = tk.Frame(win, bg=Theme.DANGER, height=60)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="未找到 Python 环境",
                 bg=Theme.DANGER, fg="white",
                 font=("微软雅黑", 14, "bold")).pack(
            side=tk.LEFT, padx=20, pady=12)

        body = tk.Frame(win, bg=Theme.WHITE)
        body.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

        tk.Label(body,
                 text="本试卷包含程序题，需要 Python 环境。\n"
                      "请选择以下操作：",
                 font=Theme.FONT, bg=Theme.WHITE, fg=Theme.TEXT,
                 wraplength=420,
                 justify="center").pack(pady=(0, 15))

        def _auto_search():
            python_path = find_system_python()
            if python_path:
                result["path"] = python_path
                self._user_python_path = python_path
                messagebox.showinfo("找到 Python",
                                    f"已找到：\n{python_path}",
                                    parent=win)
                win.destroy()
            else:
                messagebox.showwarning(
                    "未找到",
                    "自动搜索未找到 Python。\n请手动选择或安装。",
                    parent=win)

        def _manual_select():
            initial_dir = os.environ.get("ProgramFiles", "C:\\")
            for cand in [
                os.path.join(os.environ.get("LOCALAPPDATA", ""),
                             "Programs\\Python"),
                os.environ.get("ProgramFiles", "C:\\Program Files"),
            ]:
                if os.path.isdir(cand):
                    initial_dir = cand
                    break
            path = filedialog.askopenfilename(
                title="选择 python.exe",
                filetypes=[("python.exe", "python.exe"),
                           ("所有文件", "*.*")],
                initialdir=initial_dir, parent=win)
            if path and os.path.isfile(path):
                result["path"] = path
                self._user_python_path = path
                win.destroy()

        def _download():
            webbrowser.open("https://www.python.org/downloads/")
            messagebox.showinfo(
                "提示",
                "下载安装后请重新打开本程序。\n"
                "安装时务必勾选 \"Add Python to PATH\"！",
                parent=win)

        def _skip():
            win.destroy()

        btn_configs = [
            ("自动搜索 Python", Theme.PRIMARY, "white", True,
             _auto_search),
            ("手动选择 python.exe", Theme.BG, Theme.TEXT, False,
             _manual_select),
            ("去官网下载 Python", Theme.BG, Theme.TEXT, False,
             _download),
        ]
        for text, bg, fg, is_primary, cmd in btn_configs:
            tk.Button(body, text=text, font=Theme.FONT_BOLD,
                      bg=bg, fg=fg, padx=20, pady=8,
                      bd=0 if is_primary else 1,
                      relief=tk.SOLID if not is_primary else tk.FLAT,
                      cursor="hand2",
                      command=cmd).pack(fill=tk.X, pady=4)

        tk.Button(body, text="跳过（仅做非程序题）",
                  font=Theme.FONT_SMALL, bg=Theme.WHITE, fg=Theme.MUTED,
                  bd=0, cursor="hand2",
                  command=_skip).pack(pady=(10, 0))

        win.protocol("WM_DELETE_WINDOW", _skip)
        win.grab_set()
        win.wait_window()
        return result["path"]

    def open_exam_folder(self):
        exam_dir = os.path.dirname(self.current_exam_file)
        exam_folder = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(exam_folder):
            exam_folder = exam_dir
        try:
            if sys.platform == "win32":
                os.startfile(exam_folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", exam_folder], check=True)
            else:
                subprocess.run(["xdg-open", exam_folder], check=True)
        except Exception as e:
            messagebox.showerror("错误",
                                 f"打开文件夹失败：{str(e)}")

    # ─────────────────────────────────────────────────────────────────
    #  重做
    # ─────────────────────────────────────────────────────────────────
    def _restore_program_file(self, program_file):
        if not self._backup_dir or not os.path.exists(self._backup_dir):
            return
        backup_file = os.path.join(self._backup_dir, program_file)
        if not os.path.exists(backup_file):
            return
        exam_dir = os.path.dirname(self.current_exam_file)
        target_path = os.path.join(exam_dir, "试题文件夹", program_file)
        if not os.path.exists(target_path):
            target_path = os.path.join(exam_dir, program_file)
        try:
            shutil.copy2(backup_file, target_path)
        except Exception:
            pass

    def redo_question(self):
        current_q = self.questions[self.current_index]
        global_num = current_q["题号"]
        q_type = current_q["题型"]
        program_file = current_q.get("程序文件", "").strip()

        if program_file and q_type in ("程序设计", "程序填空", "程序改错"):
            if not messagebox.askyesno(
                    "确认重做",
                    "确定要重做此题吗？\n"
                    "答案和程序文件都将恢复为初始状态。"):
                return
            self._restore_program_file(program_file)
        else:
            if not messagebox.askyesno(
                    "确认重做",
                    "确定要重做此题吗？当前答案将被清空。"):
                return

        self.user_answers.pop(global_num, None)
        self._clear_children(self.q_feedback)
        self.answer_label = tk.Label(
            self.q_feedback, text="",
            font=("微软雅黑", 11, "bold"), bg=Theme.WHITE,
            fg=Theme.PRIMARY, justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)

        if q_type in ("单选", "判断"):
            self.answer_var.set("")
            for opt_letter, btn in self._choice_buttons.items():
                try:
                    btn.config(bg=Theme.BG, fg=Theme.TEXT)
                except tk.TclError:
                    pass
        else:
            self.answer_var.set("")
            if self.answer_text is not None:
                try:
                    self.answer_text.delete("1.0", tk.END)
                except tk.TclError:
                    pass
            if self.answer_entry is not None:
                try:
                    self.answer_entry.delete(0, tk.END)
                except tk.TclError:
                    pass

        self._update_nav_status()

    # ─────────────────────────────────────────────────────────────────
    #  导航面板
    # ─────────────────────────────────────────────────────────────────
    def _build_nav_panels(self):
        nav_header = tk.Frame(self.right_frame, bg=Theme.WHITE)
        nav_header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(nav_header, text="题目导航",
                 font=Theme.FONT_TITLE, bg=Theme.WHITE, fg=Theme.TEXT,
                 anchor="center").pack(fill=tk.X, pady=(12, 6), padx=5)
        tk.Frame(nav_header, bg=Theme.BORDER,
                 height=1).pack(fill=tk.X, padx=8)

        nav_footer = tk.Frame(self.right_frame, bg=Theme.WHITE)
        nav_footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(nav_footer, bg=Theme.BORDER,
                 height=1).pack(fill=tk.X, padx=8)
        self.status_label = tk.Label(
            nav_footer,
            text=f"未答 {len(self.questions)}，已答 0，标记 0",
            font=("微软雅黑", 9), bg=Theme.WHITE, fg=Theme.MUTED)
        self.status_label.pack(pady=(6, 10))

        scroll_area = tk.Frame(self.right_frame, bg=Theme.WHITE)
        scroll_area.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        canvas = tk.Canvas(scroll_area, bg=Theme.WHITE,
                           highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_area, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        nav_inner = tk.Frame(canvas, bg=Theme.WHITE)
        canvas_win = canvas.create_window(
            (0, 0), window=nav_inner, anchor="nw")

        self._ns_cached_content_h = 1
        self._ns_cached_visible_h = 1

        def _on_canvas_resize(e):
            canvas.itemconfig(canvas_win, width=e.width)
            self._ns_cached_visible_h = e.height

        canvas.bind("<Configure>", _on_canvas_resize)

        def _on_inner_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            self._ns_cached_content_h = e.height

        nav_inner.bind("<Configure>", _on_inner_resize)

        self.nav_panels = {}
        self.nav_q_buttons = {}

        first_type = True
        for idx, q_type in enumerate(self.active_type_order):
            count = len(self.question_groups.get(q_type, []))
            if count == 0:
                continue
            if not first_type:
                tk.Frame(nav_inner, bg=Theme.BORDER,
                         height=1).pack(fill=tk.X, pady=4)
            first_type = False

            header = tk.Frame(nav_inner, bg=Theme.NAV_HEADER_BG)
            header.pack(fill=tk.X, padx=4, pady=1)

            is_first_panel = (len(self.nav_panels) == 0)
            arrow_var = tk.StringVar(
                value="▼" if is_first_panel else "▶")
            arrow_lbl = tk.Label(header, textvariable=arrow_var,
                                 font=("微软雅黑", 9),
                                 bg=Theme.NAV_HEADER_BG, fg=Theme.TEXT,
                                 width=2, anchor="center")
            arrow_lbl.pack(side=tk.LEFT, padx=(4, 0))

            title_lbl = tk.Label(header,
                                 text=f"{q_type}（{count}题）",
                                 font=("微软雅黑", 10, "bold"),
                                 bg=Theme.NAV_HEADER_BG, fg=Theme.TEXT,
                                 anchor="w")
            title_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

            body = tk.Frame(nav_inner, bg=Theme.WHITE)
            self.nav_panels[q_type] = {
                "header": header, "arrow": arrow_var,
                "body": body, "items": []
            }

            if q_type in self.question_groups:
                for type_idx, q in enumerate(
                        self.question_groups[q_type]):
                    global_idx = q["_global_idx"]
                    global_num = q["题号"]
                    preview = q["题目"][:80] + (
                        "..." if len(q["题目"]) > 80 else "")
                    btn = tk.Button(
                        body,
                        text=f"  第{type_idx + 1}题（{global_num}）",
                        font=("微软雅黑", 9), bg=Theme.WHITE,
                        fg=Theme.TEXT, bd=0, anchor="w", padx=20,
                        cursor="hand2",
                        activebackground=Theme.NAV_ACTIVE,
                        command=lambda gi=global_idx:
                        self._nav_jump(gi))
                    btn.bind("<Enter>",
                             lambda e, p=preview:
                             self._show_tooltip(e, p))
                    btn.bind("<Leave>",
                             lambda e: self._hide_tooltip())
                    btn.pack(fill=tk.X, pady=0)
                    self.nav_panels[q_type]["items"].append(btn)
                    self.nav_q_buttons[global_idx] = btn

            def _toggle(al=arrow_var, bd=body, hd=header):
                if bd.winfo_ismapped():
                    bd.pack_forget()
                    al.set("▶")
                else:
                    bd.pack(fill=tk.X, after=hd)
                    al.set("▼")
                canvas.after(30, lambda: self._ns_clamp(canvas))

            for widget in (header, arrow_lbl, title_lbl):
                widget.bind("<Button-1>",
                            lambda e, fn=_toggle: fn())

            if is_first_panel:
                body.pack(fill=tk.X, after=header)

        self._ns_canvas = canvas
        self._ns_inner = nav_inner
        self._ns_state = {'vel': 0.0, 'aid': None,
                          'bounce': False, 'lt': 0.0}

        _FRICTION = 0.80
        _STOP_THR = 0.00002
        _MAX_VEL = 0.18
        _PX_PER_NOTCH = 38.0

        def _on_wheel(event):
            if self._ns_state['bounce']:
                return "break"
            if event.num == 4:
                notch = -1
            elif event.num == 5:
                notch = 1
            elif sys.platform == "darwin":
                notch = -event.delta
            else:
                notch = -event.delta / 120.0
            content_h = self._ns_cached_content_h
            if content_h <= 0:
                return "break"
            frac_per_px = 1.0 / content_h
            vel_delta = notch * _PX_PER_NOTCH * frac_per_px
            self._ns_state['vel'] += vel_delta
            self._ns_state['vel'] = max(
                -_MAX_VEL, min(_MAX_VEL, self._ns_state['vel']))
            if self._ns_state['aid'] is None:
                self._ns_state['lt'] = time.perf_counter()
                self._ns_state['aid'] = canvas.after(14, _step)
            return "break"

        def _step():
            try:
                now = time.perf_counter()
                dt = now - self._ns_state['lt']
                self._ns_state['lt'] = now
                n_frames = max(0.3, dt * 60.0)
                friction = _FRICTION ** n_frames
                vel = self._ns_state['vel'] * friction
                if abs(vel) < _STOP_THR:
                    self._ns_state['vel'] = 0.0
                    self._ns_state['aid'] = None
                    return
                self._ns_state['vel'] = vel
                content_h = self._ns_cached_content_h
                visible_h = self._ns_cached_visible_h
                if content_h <= visible_h:
                    self._ns_state['vel'] = 0.0
                    self._ns_state['aid'] = None
                    canvas.yview_moveto(0)
                    return
                max_frac = 1.0 - visible_h / content_h
                current = canvas.yview()[0]
                new_pos = current + vel
                if new_pos < -0.008:
                    self._ns_state['vel'] = 0.0
                    self._ns_state['aid'] = None
                    self._ns_bounce_top()
                    return
                if new_pos > max_frac + 0.008:
                    self._ns_state['vel'] = 0.0
                    self._ns_state['aid'] = None
                    self._ns_bounce_bottom()
                    return
                new_pos = max(0, min(max_frac, new_pos))
                canvas.yview_moveto(new_pos)
                self._ns_state['aid'] = canvas.after(14, _step)
            except tk.TclError:
                self._ns_state['aid'] = None

        canvas.bind("<MouseWheel>", _on_wheel)
        canvas.bind("<Button-4>", _on_wheel)
        canvas.bind("<Button-5>", _on_wheel)

        def _bind_mw(w):
            w.bind("<MouseWheel>", _on_wheel)
            w.bind("<Button-4>", _on_wheel)
            w.bind("<Button-5>", _on_wheel)
            for child in w.winfo_children():
                _bind_mw(child)

        _bind_mw(nav_inner)

    def _ns_clamp(self, canvas):
        try:
            canvas.update_idletasks()
            sr = canvas.bbox("all")
            if sr:
                canvas.configure(scrollregion=sr)
                self._ns_cached_content_h = sr[3] - sr[1]
            content_h = self._ns_cached_content_h
            visible_h = canvas.winfo_height()
            self._ns_cached_visible_h = visible_h
            if content_h <= visible_h:
                canvas.yview_moveto(0)
                self._ns_state['vel'] = 0
            else:
                max_frac = 1.0 - visible_h / content_h
                current = canvas.yview()[0]
                if current > max_frac:
                    canvas.yview_moveto(max_frac)
                elif current < 0:
                    canvas.yview_moveto(0)
        except tk.TclError:
            pass

    def _ns_bounce_top(self):
        if self._ns_state['bounce']:
            return
        self._ns_state['bounce'] = True
        canvas = self._ns_canvas
        inner = self._ns_inner
        BOUNCE_PX = 40
        children = inner.winfo_children()
        spacer = tk.Frame(inner, bg=Theme.WHITE, height=0)
        if children:
            spacer.pack(fill=tk.X, before=children[0])
        else:
            spacer.pack(fill=tk.X)

        def _grow(h):
            try:
                if h >= BOUNCE_PX:
                    canvas.after(15, lambda: _shrink(BOUNCE_PX))
                    return
                spacer.configure(height=h)
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
                canvas.after(7, lambda: _grow(h + 7))
            except tk.TclError:
                pass

        def _shrink(h):
            try:
                if h <= 1:
                    spacer.destroy()
                    canvas.after(15, _cleanup)
                    return
                spacer.configure(height=max(0, h))
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
                canvas.after(8, lambda: _shrink(int(h * 0.45)))
            except tk.TclError:
                pass

        def _cleanup():
            try:
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
            except tk.TclError:
                pass
            self._ns_state['bounce'] = False

        _grow(4)

    def _ns_bounce_bottom(self):
        if self._ns_state['bounce']:
            return
        self._ns_state['bounce'] = True
        canvas = self._ns_canvas
        inner = self._ns_inner
        BOUNCE_PX = 40
        spacer = tk.Frame(inner, bg=Theme.WHITE, height=0)
        spacer.pack(fill=tk.X)

        def _grow(h):
            try:
                if h >= BOUNCE_PX:
                    canvas.after(15, lambda: _shrink(BOUNCE_PX))
                    return
                spacer.configure(height=h)
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(1)
                canvas.after(7, lambda: _grow(h + 7))
            except tk.TclError:
                pass

        def _shrink(h):
            try:
                if h <= 1:
                    spacer.destroy()
                    canvas.after(15, _cleanup)
                    return
                spacer.configure(height=max(0, h))
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(1)
                canvas.after(8, lambda: _shrink(int(h * 0.45)))
            except tk.TclError:
                pass

        def _cleanup():
            try:
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._ns_cached_content_h = sr[3] - sr[1]
            except tk.TclError:
                pass
            self._ns_state['bounce'] = False

        _grow(4)

    # ─────────────────────────────────────────────────────────────────
    #  左侧滚动
    # ─────────────────────────────────────────────────────────────────
    def _setup_left_scroll(self):
        canvas = self._lp_canvas
        inner = self._lp_inner
        self._lp_state = {'vel': 0.0, 'aid': None,
                          'bounce': False, 'lt': 0.0}
        _FRICTION = 0.80
        _STOP_THR = 0.00002
        _MAX_VEL = 0.18
        _PX_PER_NOTCH = 38.0

        def _step():
            try:
                now = time.perf_counter()
                dt = now - self._lp_state['lt']
                self._lp_state['lt'] = now
                n_frames = max(0.3, dt * 60.0)
                friction = _FRICTION ** n_frames
                vel = self._lp_state['vel'] * friction
                if abs(vel) < _STOP_THR:
                    self._lp_state['vel'] = 0.0
                    self._lp_state['aid'] = None
                    return
                self._lp_state['vel'] = vel
                content_h = self._lp_cached_content_h
                visible_h = self._lp_cached_visible_h
                if content_h <= visible_h:
                    self._lp_state['vel'] = 0.0
                    self._lp_state['aid'] = None
                    canvas.yview_moveto(0)
                    return
                max_frac = 1.0 - visible_h / content_h
                current = canvas.yview()[0]
                new_pos = current + vel
                if new_pos < -0.008:
                    self._lp_state['vel'] = 0.0
                    self._lp_state['aid'] = None
                    self._lp_bounce_top()
                    return
                if new_pos > max_frac + 0.008:
                    self._lp_state['vel'] = 0.0
                    self._lp_state['aid'] = None
                    self._lp_bounce_bottom()
                    return
                new_pos = max(0, min(max_frac, new_pos))
                canvas.yview_moveto(new_pos)
                self._lp_state['aid'] = canvas.after(14, _step)
            except tk.TclError:
                self._lp_state['aid'] = None

        self._lp_step = _step

        def _on_wheel(event):
            if self._lp_state['bounce']:
                return "break"
            if event.num == 4:
                notch = -1
            elif event.num == 5:
                notch = 1
            elif sys.platform == "darwin":
                notch = -event.delta
            else:
                notch = -event.delta / 120.0
            content_h = self._lp_cached_content_h
            if content_h <= 0:
                return "break"
            frac_per_px = 1.0 / content_h
            vel_delta = notch * _PX_PER_NOTCH * frac_per_px
            self._lp_state['vel'] += vel_delta
            self._lp_state['vel'] = max(
                -_MAX_VEL, min(_MAX_VEL, self._lp_state['vel']))
            if self._lp_state['aid'] is None:
                self._lp_state['lt'] = time.perf_counter()
                self._lp_state['aid'] = canvas.after(14, _step)
            return "break"

        self._lp_on_wheel = _on_wheel
        canvas.bind("<MouseWheel>", _on_wheel)
        canvas.bind("<Button-4>", _on_wheel)
        canvas.bind("<Button-5>", _on_wheel)
        self._lp_bind_mw(inner)

    def _lp_bind_mw(self, widget):
        w_handler = self._lp_on_wheel
        widget.bind("<MouseWheel>", w_handler)
        widget.bind("<Button-4>", w_handler)
        widget.bind("<Button-5>", w_handler)
        for child in widget.winfo_children():
            self._lp_bind_mw(child)

    def _lp_bounce_top(self):
        if self._lp_state['bounce']:
            return
        self._lp_state['bounce'] = True
        canvas = self._lp_canvas
        inner = self._lp_inner
        BOUNCE_PX = 40
        children = inner.winfo_children()
        spacer = tk.Frame(inner, bg=Theme.WHITE, height=0)
        if children:
            spacer.pack(fill=tk.X, before=children[0])
        else:
            spacer.pack(fill=tk.X)

        def _grow(h):
            try:
                if h >= BOUNCE_PX:
                    canvas.after(15, lambda: _shrink(BOUNCE_PX))
                    return
                spacer.configure(height=h)
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
                canvas.after(7, lambda: _grow(h + 7))
            except tk.TclError:
                pass

        def _shrink(h):
            try:
                if h <= 1:
                    spacer.destroy()
                    canvas.after(15, _cleanup)
                    return
                spacer.configure(height=max(0, h))
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
                canvas.after(8, lambda: _shrink(int(h * 0.45)))
            except tk.TclError:
                pass

        def _cleanup():
            try:
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(0)
            except tk.TclError:
                pass
            self._lp_state['bounce'] = False

        _grow(4)

    def _lp_bounce_bottom(self):
        if self._lp_state['bounce']:
            return
        self._lp_state['bounce'] = True
        canvas = self._lp_canvas
        inner = self._lp_inner
        BOUNCE_PX = 40
        spacer = tk.Frame(inner, bg=Theme.WHITE, height=0)
        spacer.pack(fill=tk.X)

        def _grow(h):
            try:
                if h >= BOUNCE_PX:
                    canvas.after(15, lambda: _shrink(BOUNCE_PX))
                    return
                spacer.configure(height=h)
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(1)
                canvas.after(7, lambda: _grow(h + 7))
            except tk.TclError:
                pass

        def _shrink(h):
            try:
                if h <= 1:
                    spacer.destroy()
                    canvas.after(15, _cleanup)
                    return
                spacer.configure(height=max(0, h))
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
                canvas.yview_moveto(1)
                canvas.after(8, lambda: _shrink(int(h * 0.45)))
            except tk.TclError:
                pass

        def _cleanup():
            try:
                sr = canvas.bbox("all")
                if sr:
                    canvas.configure(scrollregion=sr)
                    self._lp_cached_content_h = sr[3] - sr[1]
            except tk.TclError:
                pass
            self._lp_state['bounce'] = False

        _grow(4)

    # ─────────────────────────────────────────────────────────────────
    #  Tooltip
    # ─────────────────────────────────────────────────────────────────
    def _show_tooltip(self, event, text):
        self._hide_tooltip()
        self._tooltip = tk.Toplevel(self.root)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.wm_attributes("-topmost", True)
        self._tooltip.wm_geometry(
            f"+{event.x_root + 15}+{event.y_root + 10}")
        tk.Label(self._tooltip, text=text,
                 font=Theme.FONT_SMALL, bg=Theme.TOOLTIP_BG,
                 fg=Theme.TOOLTIP_FG,
                 relief=tk.SOLID, bd=1, wraplength=300,
                 justify="left", padx=8, pady=4).pack()

    def _hide_tooltip(self):
        if hasattr(self, '_tooltip') and self._tooltip:
            self._tooltip.destroy()
            self._tooltip = None

    # ─────────────────────────────────────────────────────────────────
    #  导航操作
    # ─────────────────────────────────────────────────────────────────
    def _nav_jump(self, global_idx):
        self.current_index = global_idx
        self.show_question()

    def jump_next_unanswered(self):
        self._auto_save_current()
        for i in range(self.current_index + 1, len(self.questions)):
            if self.questions[i]["题号"] not in self.user_answers:
                self.current_index = i
                self.show_question()
                return
        for i in range(0, self.current_index + 1):
            if self.questions[i]["题号"] not in self.user_answers:
                self.current_index = i
                self.show_question()
                return
        messagebox.showinfo("提示", "所有题目都已作答！")

    # ─────────────────────────────────────────────────────────────────
    #  显示题目
    # ─────────────────────────────────────────────────────────────────
    def show_question(self):
        current_q = self.questions[self.current_index]
        global_num = current_q["题号"]
        q_type = current_q["题型"]
        type_questions = self.question_groups[q_type]
        type_question_num = type_questions.index(current_q) + 1

        self._clear_children(self.q_title_bar)
        title_text = (
            f"[{self.current_index + 1}/{len(self.questions)}] "
            f"{q_type} - 第{type_question_num}题"
            f"（题号：{global_num}）- {current_q['分值']}分"
            f"（共{len(self.questions)}题，共100.0分）")
        tk.Label(self.q_title_bar, text=title_text,
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=10, pady=8)

        self._clear_children(self.q_instruction)
        if q_type == "程序设计":
            if not self.q_instruction.winfo_ismapped():
                self.q_instruction.pack(fill=tk.X,
                                        before=self.q_content)
            tk.Label(self.q_instruction, text="<<答题说明>>",
                     font=("微软雅黑", 10, "bold"), bg=Theme.WHITE,
                     fg=Theme.TEXT, anchor="w").pack(fill=tk.X, padx=20,
                                                     pady=(0, 5))
            tk.Label(self.q_instruction,
                     text='1. 点击下方"答题"按钮，系统会用IDLE打开'
                          '对应的程序文件\n'
                          '2. 在程序中完成需求（如补全代码、修复bug）\n'
                          '3. 修改完成后按Ctrl+S保存文件\n'
                          '4. 回到本系统，在答案框中输入核心代码',
                     font=("微软雅黑", 9), bg=Theme.WHITE,
                     fg=Theme.TEXT, wraplength=900,
                     justify="left",
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 10))
            tk.Label(self.q_instruction,
                     text="注意：需按题目要求编写代码，"
                          "保存后再提交答案。",
                     font=("微软雅黑", 9), bg=Theme.WHITE,
                     fg=Theme.TEXT,
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 20))
            ttk.Separator(self.q_instruction,
                          orient="horizontal").pack(fill=tk.X, padx=20)
        elif q_type in ("程序改错", "程序填空"):
            if not self.q_instruction.winfo_ismapped():
                self.q_instruction.pack(fill=tk.X,
                                        before=self.q_content)
            tk.Label(self.q_instruction, text="<<答题说明>>",
                     font=("微软雅黑", 10, "bold"), bg=Theme.WHITE,
                     fg=Theme.TEXT, anchor="w").pack(fill=tk.X, padx=20,
                                                     pady=(0, 5))
            tk.Label(self.q_instruction,
                     text='1. 点击下方"答题"按钮，系统会用IDLE打开'
                          '对应的程序文件\n'
                          '   （无需安装PyCharm，Python自带IDLE即可）\n'
                          '2. 在**********FOUND**********语句的'
                          '下一行修改程序\n'
                          '3. 修改完成后按Ctrl+S保存文件\n'
                          '4. 回到本系统，在答案框中输入修改后的内容',
                     font=("微软雅黑", 9), bg=Theme.WHITE,
                     fg=Theme.TEXT, wraplength=900,
                     justify="left",
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 10))
            tk.Label(self.q_instruction,
                     text="注意：不可以增加或删除程序行，"
                          "也不可以更改程序的结构。",
                     font=("微软雅黑", 9), bg=Theme.WHITE,
                     fg=Theme.TEXT,
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 20))
            ttk.Separator(self.q_instruction,
                          orient="horizontal").pack(fill=tk.X, padx=20)
        else:
            self.q_instruction.pack_forget()

        self._clear_children(self.q_content)
        tk.Label(self.q_content, text=current_q["题目"],
                 font=("微软雅黑", 11), bg=Theme.WHITE, fg=Theme.TEXT,
                 wraplength=900, justify="left",
                 anchor="w").pack(fill=tk.X, pady=(10, 10))

        self._clear_children(self.q_answer_area)
        self.answer_text = None
        self.answer_entry = None
        self._choice_buttons = {}
        self.answer_var.set(self.user_answers.get(global_num, ""))

        if q_type == "单选":
            self._build_single_choice(current_q, global_num)
        elif q_type == "判断":
            self._build_judge(current_q, global_num)
        elif q_type in ("填空", "程序填空", "程序改错", "程序设计"):
            self._build_text_input(current_q, global_num, q_type)

        self._clear_children(self.q_feedback)
        self.answer_label = tk.Label(
            self.q_feedback, text="",
            font=("微软雅黑", 11, "bold"), bg=Theme.WHITE,
            fg=Theme.PRIMARY, justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)

        program_types = {"程序设计", "程序填空", "程序改错"}
        if q_type in program_types:
            self.answer_btn.grid()
        else:
            self.answer_btn.grid_remove()

        self._ensure_panel_open(q_type)
        self._update_nav_status()

        try:
            self._lp_state['vel'] = 0.0
            if self._lp_state.get('aid') is not None:
                self._lp_canvas.after_cancel(self._lp_state['aid'])
                self._lp_state['aid'] = None
            self._lp_canvas.yview_moveto(0)
            self._lp_bind_mw(self._lp_inner)
        except (tk.TclError, AttributeError):
            pass

    def _ensure_panel_open(self, q_type):
        if q_type in self.nav_panels:
            panel = self.nav_panels[q_type]
            if not panel["body"].winfo_ismapped():
                panel["body"].pack(fill=tk.X, after=panel["header"])
                panel["arrow"].set("▼")

    def _get_options_for_question(self, q):
        options = []
        for letter in ["A", "B", "C", "D", "E", "F"]:
            text = q.get(f"选项{letter}",
                         q.get(f"选项 {letter}", "")).strip()
            if text:
                options.append((letter, text))
        return options

    def _build_single_choice(self, q, global_num):
        options_frame = tk.Frame(self.q_answer_area, bg=Theme.WHITE)
        options_frame.pack(fill=tk.X, pady=(0, 10))
        options = self._get_options_for_question(q)
        for opt_letter, opt_text in options:
            tk.Label(options_frame, text=f"({opt_letter}){opt_text}",
                     font=("微软雅黑", 11), bg=Theme.WHITE,
                     fg=Theme.TEXT, anchor="w").pack(fill=tk.X, pady=5)
        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG,
                            bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar,
                 text="在下面选择答案（点击选项字母进行选择）",
                 bg=Theme.HINT_BG, fg=Theme.HINT_TEXT,
                 font=("微软雅黑", 9),
                 anchor="w").pack(fill=tk.X, padx=10, pady=5)
        buttons_frame = tk.Frame(self.q_answer_area, bg=Theme.WHITE)
        buttons_frame.pack(fill=tk.X, pady=10)
        current_answer = self.answer_var.get()
        for opt_letter, _ in options:
            selected = (current_answer == opt_letter)
            btn = tk.Button(
                buttons_frame, text=opt_letter,
                font=("微软雅黑", 16, "bold"),
                width=3, height=1, bd=1, relief=tk.SOLID,
                bg=Theme.PRIMARY if selected else Theme.BG,
                fg="white" if selected else Theme.TEXT,
                cursor="hand2",
                command=lambda o=opt_letter:
                self._choose(global_num, o))
            btn.pack(side=tk.LEFT, padx=20)
            self._choice_buttons[opt_letter] = btn

    def _build_judge(self, q, global_num):
        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG,
                            bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar,
                 text="在下面选择答案（点击对或错进行选择）",
                 bg=Theme.HINT_BG, fg=Theme.HINT_TEXT,
                 font=("微软雅黑", 9),
                 anchor="w").pack(fill=tk.X, padx=10, pady=5)
        buttons_frame = tk.Frame(self.q_answer_area, bg=Theme.WHITE)
        buttons_frame.pack(fill=tk.X, pady=15)
        current_answer = self.answer_var.get()
        for label, value in [("对", "A"), ("错", "B")]:
            selected = (current_answer == value)
            btn = tk.Button(
                buttons_frame, text=label,
                font=("微软雅黑", 16, "bold"),
                width=5, height=1, bd=1, relief=tk.SOLID,
                bg=Theme.PRIMARY if selected else Theme.BG,
                fg="white" if selected else Theme.TEXT,
                cursor="hand2",
                command=lambda v=value: self._choose(global_num, v))
            btn.pack(side=tk.LEFT, padx=20)
            self._choice_buttons[value] = btn

    def _build_text_input(self, q, global_num, q_type):
        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG,
                            bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar, text="在下面输入答案",
                 bg=Theme.HINT_BG, fg=Theme.HINT_TEXT,
                 font=("微软雅黑", 9),
                 anchor="w").pack(fill=tk.X, padx=10, pady=5)
        input_frame = tk.Frame(self.q_answer_area, bg=Theme.WHITE)
        input_frame.pack(fill=tk.X, pady=10)
        tk.Label(input_frame, text="答案：",
                 font=("微软雅黑", 11),
                 bg=Theme.WHITE, fg=Theme.TEXT).pack(side=tk.LEFT)
        if q_type in ("程序填空", "程序改错", "程序设计"):
            self.answer_text = tk.Text(input_frame,
                                       font=("Consolas", 11),
                                       width=60, height=5,
                                       bg=Theme.INPUT_BG,
                                       fg=Theme.TEXT,
                                       insertbackground=Theme.TEXT,
                                       selectbackground=Theme.PRIMARY,
                                       selectforeground="white",
                                       relief=tk.SOLID, bd=1,
                                       highlightbackground=Theme.BORDER,
                                       highlightthickness=1)
            self.answer_text.pack(side=tk.LEFT, padx=10)
            if global_num in self.user_answers:
                self.answer_text.insert(tk.END,
                                        self.user_answers[global_num])
            self.answer_text.bind(
                "<KeyRelease>",
                lambda e: self._save_text(global_num))
            self.answer_text.bind(
                "<FocusOut>",
                lambda e: self._save_text(global_num))
        else:
            self.answer_entry = ttk.Entry(
                input_frame, font=("微软雅黑", 11),
                width=40, textvariable=self.answer_var)
            self.answer_entry.pack(side=tk.LEFT, padx=10)
            self.answer_entry.bind(
                "<KeyRelease>",
                lambda e: self._save_var(global_num))
            self.answer_entry.bind(
                "<FocusOut>",
                lambda e: self._save_var(global_num))

    # ─────────────────────────────────────────────────────────────────
    #  答案处理
    # ─────────────────────────────────────────────────────────────────
    def _choose(self, global_num, option):
        self.user_answers[global_num] = option
        self.answer_var.set(option)
        current_q = self.questions[self.current_index]

        if self.show_answer_immediately:
            correct = self._normalize_answer(
                current_q["正确答案"], current_q["题型"])
            chosen = self._normalize_answer(option, current_q["题型"])
            for opt_letter, btn in self._choice_buttons.items():
                try:
                    normalized = self._normalize_answer(
                        opt_letter, current_q["题型"])
                    if normalized == correct:
                        btn.config(bg=Theme.SUCCESS, fg="white")
                    elif opt_letter == option and chosen != correct:
                        btn.config(bg=Theme.DANGER, fg="white")
                    else:
                        btn.config(bg=Theme.BG, fg=Theme.TEXT)
                except tk.TclError:
                    pass
            try:
                if chosen == correct:
                    self.answer_label.config(
                        text="回答正确！", fg=Theme.SUCCESS)
                else:
                    self.answer_label.config(
                        text=f"回答错误，正确答案是："
                             f"{current_q['正确答案']}",
                        fg=Theme.DANGER)
            except tk.TclError:
                pass
        else:
            for opt_letter, btn in self._choice_buttons.items():
                try:
                    if opt_letter == option:
                        btn.config(bg=Theme.PRIMARY, fg="white")
                    else:
                        btn.config(bg=Theme.BG, fg=Theme.TEXT)
                except tk.TclError:
                    pass
            try:
                self.answer_label.config(text="")
            except tk.TclError:
                pass

        self._update_nav_status()

    def _save_var(self, global_num):
        answer = self.answer_var.get().strip()
        if answer:
            self.user_answers[global_num] = answer
        else:
            self.user_answers.pop(global_num, None)
        self._update_nav_status()

    def _save_text(self, global_num):
        if self.answer_text is None:
            return
        answer = self.answer_text.get("1.0", tk.END).strip()
        if answer:
            self.user_answers[global_num] = answer
        else:
            self.user_answers.pop(global_num, None)
        self._update_nav_status()

    # ─────────────────────────────────────────────────────────────────
    #  导航状态
    # ─────────────────────────────────────────────────────────────────
    def _update_nav_status(self):
        answered_count = len(self.user_answers)
        marked_count = len(self.marked_questions)

        for q_type in self.active_type_order:
            if q_type not in self.question_groups:
                continue
            for type_idx, q in enumerate(
                    self.question_groups[q_type]):
                global_idx = q["_global_idx"]
                global_num = q["题号"]
                if global_idx not in self.nav_q_buttons:
                    continue
                btn = self.nav_q_buttons[global_idx]

                if global_idx == self.current_index:
                    bg_color = Theme.NAV_CURRENT
                    fg_color = "white"
                    font = ("微软雅黑", 9, "bold")
                elif global_num in self.user_answers:
                    bg_color = Theme.NAV_ANSWERED_BG
                    fg_color = Theme.NAV_ANSWERED_FG
                    font = ("微软雅黑", 9)
                else:
                    bg_color = Theme.WHITE
                    fg_color = Theme.TEXT
                    font = ("微软雅黑", 9)

                btn_text = f"  第{type_idx + 1}题（{global_num}）"
                if global_idx in self.marked_questions:
                    btn_text = "🚩" + btn_text.strip()
                    if (global_idx != self.current_index
                            and global_num not in self.user_answers):
                        bg_color = Theme.NAV_MARKED_BG
                        fg_color = Theme.NAV_MARKED_FG

                btn.config(text=btn_text, bg=bg_color,
                           fg=fg_color, font=font)

        if hasattr(self, 'mark_btn'):
            if self.current_index in self.marked_questions:
                self.mark_btn.config(text="取消标记",
                                     fg=Theme.NAV_MARKED_FG)
            else:
                self.mark_btn.config(text="标记试题",
                                     fg=Theme.TEXT)

        self.status_label.config(
            text=f"未答 {len(self.questions) - answered_count}，"
                 f"已答 {answered_count}，标记 {marked_count}")
        self._update_progress()

    def _update_progress(self):
        total = len(self.questions)
        if total == 0:
            return
        answered = len(self.user_answers)
        ratio = answered / total
        self.progress_bar.place(x=0, y=0, relheight=1.0, relwidth=ratio)
        if hasattr(self, 'progress_label'):
            if answered == total:
                self.progress_label.config(
                    text="已完成，可以交卷！", fg=Theme.SUCCESS)
            else:
                self.progress_label.config(
                    text=f"已完成 {answered} / {total} 题",
                    fg=Theme.MUTED)

    def _auto_save_current(self):
        if self.answer_text is not None:
            try:
                answer = self.answer_text.get("1.0", tk.END).strip()
                gn = self.questions[self.current_index]["题号"]
                if answer:
                    self.user_answers[gn] = answer
                else:
                    self.user_answers.pop(gn, None)
            except tk.TclError:
                pass
        elif self.answer_entry is not None:
            try:
                answer = self.answer_var.get().strip()
                gn = self.questions[self.current_index]["题号"]
                if answer:
                    self.user_answers[gn] = answer
                else:
                    self.user_answers.pop(gn, None)
            except tk.TclError:
                pass

    def prev_question(self):
        self._auto_save_current()
        if self.current_index > 0:
            self.current_index -= 1
            self.show_question()

    def next_question(self):
        self._auto_save_current()
        if self.current_index < len(self.questions) - 1:
            self.current_index += 1
            self.show_question()

    def toggle_mark(self):
        if self.current_index in self.marked_questions:
            self.marked_questions.remove(self.current_index)
        else:
            self.marked_questions.add(self.current_index)
        self._update_nav_status()

    def show_answer(self):
        q = self.questions[self.current_index]
        self._clear_children(self.q_feedback)
        self.answer_label = tk.Label(
            self.q_feedback, text="",
            font=("微软雅黑", 11, "bold"), bg=Theme.WHITE,
            fg=Theme.PRIMARY, justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)
        answer_frame = tk.Frame(self.q_feedback, bg=Theme.ANSWER_BG,
                                bd=1, relief=tk.SOLID)
        answer_frame.pack(fill=tk.X, pady=5)
        tk.Label(answer_frame, text="标准答案",
                 font=("微软雅黑", 10, "bold"), bg=Theme.ANSWER_BG,
                 fg=Theme.SUCCESS,
                 anchor="w").pack(fill=tk.X, padx=10, pady=(8, 2))
        tk.Label(answer_frame, text=q["正确答案"],
                 font=("Consolas", 12, "bold"), bg=Theme.ANSWER_BG,
                 fg=Theme.SUCCESS, anchor="w", wraplength=800,
                 justify="left").pack(fill=tk.X, padx=10, pady=(0, 10))

    # ─────────────────────────────────────────────────────────────────
    #  计时器
    # ─────────────────────────────────────────────────────────────────
    def start_timer(self):
        self.timer_running = True
        self._stop_timer = Event()
        self._timer_thread = Thread(target=self._tick, daemon=True)
        self._timer_thread.start()

    def _tick(self):
        while self.timer_running and self.remaining_time > 0:
            if self._stop_timer.wait(1.0):
                return
            self.remaining_time -= 1
            try:
                if self.root.winfo_exists():
                    self.root.after(0, self._update_time_display)
            except (tk.TclError, RuntimeError):
                return
        if self.remaining_time <= 0 and self.timer_running:
            try:
                if self.root.winfo_exists():
                    self.root.after(0, self._force_submit)
            except (tk.TclError, RuntimeError):
                pass

    def _update_time_display(self):
        try:
            if not (hasattr(self, 'time_label')
                    and self.time_label.winfo_exists()):
                return
            h = self.remaining_time // 3600
            m = (self.remaining_time % 3600) // 60
            s = self.remaining_time % 60
            time_str = f"{h:02d}:{m:02d}:{s:02d}"
            if self.remaining_time <= 300:
                fg_color = "#ff4444" if Theme._is_dark else "#ff0000"
            elif self.remaining_time <= 600:
                fg_color = "#ff8844" if Theme._is_dark else "#ff6600"
            else:
                fg_color = Theme.TEXT
            self.time_label.config(text=time_str, fg=fg_color)
        except tk.TclError:
            pass

    def _force_submit(self):
        if self.exam_submitted:
            return
        self.timer_running = False
        self._stop_timer.set()
        self.exam_submitted = True
        try:
            self.time_label.config(text="00:00:00")
        except tk.TclError:
            pass
        messagebox.showinfo("提示", "考试时间到！系统将自动交卷。")
        self._save_exam_progress("completed", self._compute_score_pct())
        self._do_score_and_show_result()

    # ─────────────────────────────────────────────────────────────────
    #  评分工具
    # ─────────────────────────────────────────────────────────────────
    def _normalize_answer(self, ans, q_type):
        ans = str(ans).strip().lower()
        if q_type == "判断":
            mapping = {
                "对": "a", "错": "b", "t": "a", "f": "b",
                "true": "a", "false": "b", "1": "a", "0": "b",
                "y": "a", "n": "b", "yes": "a", "no": "b",
                "正确": "a", "错误": "b", "√": "a", "×": "b",
            }
            ans = mapping.get(ans, ans)
        return ans

    def _check_fill_in(self, user_ans, correct_ans):
        user_parts = re.split(r'[,;，；\s]+', user_ans.strip())
        correct_parts = re.split(r'[,;，；\s]+', correct_ans.strip())
        user_parts = [p for p in user_parts if p]
        correct_parts = [p for p in correct_parts if p]
        if len(user_parts) != len(correct_parts):
            return False
        return all(u.strip().lower() == c.strip().lower()
                   for u, c in zip(user_parts, correct_parts))

    def _normalize_code(self, code):
        lines = [line.strip()
                 for line in str(code).strip().splitlines()]
        lines = [line for line in lines if line]
        return "\n".join(lines)

    def _normalize_code_flexible(self, code):
        code = str(code).strip()
        code = code.replace('\r\n', '\n').replace('\r', '\n')
        lines = [line.rstrip() for line in code.split('\n')]
        lines = [line for line in lines if line]
        normalized = [line.replace("'", '"') for line in lines]
        return '\n'.join(normalized)

    def _check_program_answer(self, user_ans, correct_ans):
        if self._normalize_code(user_ans) == \
                self._normalize_code(correct_ans):
            return True
        if self._normalize_code_flexible(user_ans) == \
                self._normalize_code_flexible(correct_ans):
            return True
        return False

    def _compute_score_pct(self):
        score = 0
        total_score = 0
        for q in self.questions:
            gn = q["题号"]
            q_type = q["题型"]
            try:
                qs = int(float(q["分值"]))
            except (ValueError, TypeError):
                qs = 0
            total_score += qs
            if gn in self.user_answers:
                user_ans = self._normalize_answer(
                    self.user_answers[gn], q_type)
                correct_ans = self._normalize_answer(
                    q["正确答案"], q_type)
                if q_type in ("填空", "程序填空"):
                    ok = self._check_fill_in(user_ans, correct_ans)
                elif q_type in ("程序设计", "程序改错"):
                    ok = self._check_program_answer(
                        self.user_answers[gn], q["正确答案"])
                else:
                    ok = (user_ans == correct_ans)
                if ok:
                    score += qs
        if total_score == 0:
            return 0
        return score / total_score * 100

    # ─────────────────────────────────────────────────────────────────
    #  交卷
    # ─────────────────────────────────────────────────────────────────
    def submit_exam(self):
        if self.exam_submitted:
            return
        self._auto_save_current()
        answered = len(self.user_answers)
        total = len(self.questions)
        unanswered = total - answered
        unanswered_list = [q for q in self.questions
                           if q["题号"] not in self.user_answers]

        preview_win = tk.Toplevel(self.root)
        preview_win.title("交卷确认")
        preview_win.configure(bg=Theme.WHITE)
        preview_win.transient(self.root)
        preview_win.grab_set()

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                preview_win.iconbitmap(icon_path)
            except Exception:
                pass

        preview_win.geometry("460x520")
        preview_win.minsize(460, 300)
        preview_win.resizable(True, True)

        preview_win.update_idletasks()
        screen_w = preview_win.winfo_screenwidth()
        screen_h = preview_win.winfo_screenheight()
        x = (screen_w - 460) // 2
        y = (screen_h - 520) // 2
        preview_win.geometry(f"+{x}+{y}")

        container = tk.Frame(preview_win, bg=Theme.WHITE)
        container.pack(fill=tk.BOTH, expand=True)

        # ── 顶部标题栏 ──
        hdr = tk.Frame(container, bg=Theme.PRIMARY, height=60)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="交卷前检查",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(
            side=tk.LEFT, padx=20, pady=12)

        # ── 统计卡片 ──
        stats_card = tk.Frame(container, bg=Theme.SURFACE,
                              highlightbackground=Theme.BORDER,
                              highlightthickness=1)
        stats_card.pack(fill=tk.X, padx=15, pady=(10, 5))

        stats_data = [
            ("总题数", str(total), Theme.TEXT),
            ("已作答", str(answered), Theme.SUCCESS),
            ("未作答", str(unanswered),
             Theme.DANGER if unanswered > 0 else Theme.TEXT),
            ("标记数", str(len(self.marked_questions)),
             Theme.NAV_MARKED_FG if len(self.marked_questions) > 0
             else Theme.TEXT),
        ]
        for label, value, color in stats_data:
            row = tk.Frame(stats_card, bg=Theme.SURFACE)
            row.pack(fill=tk.X, padx=16, pady=4)
            tk.Label(row, text=label, font=("微软雅黑", 11),
                     bg=Theme.SURFACE, fg=Theme.TEXT,
                     anchor="w", width=10).pack(side=tk.LEFT)
            tk.Label(row, text=value,
                     font=("微软雅黑", 12, "bold"),
                     bg=Theme.SURFACE, fg=color,
                     anchor="e").pack(side=tk.RIGHT)

        # ── 标记题提醒 ──
        if self.marked_questions:
            marked_unanswered = [
                idx for idx in self.marked_questions
                if idx < len(self.questions)
                and self.questions[idx]["题号"] not in self.user_answers
            ]
            marked_answered = [
                idx for idx in self.marked_questions
                if idx < len(self.questions)
                and self.questions[idx]["题号"] in self.user_answers
            ]
            warn_parts = []
            if marked_unanswered:
                warn_parts.append(
                    f"{len(marked_unanswered)} 道标记题尚未作答")
            if marked_answered:
                warn_parts.append(
                    f"{len(marked_answered)} 道标记题已作答待确认")
            if warn_parts:
                mark_frame = tk.Frame(container, bg=Theme.WARN_BG,
                                      highlightbackground=Theme.WARN_BORDER,
                                      highlightthickness=1)
                mark_frame.pack(fill=tk.X, padx=15, pady=(5, 4))
                tk.Label(
                    mark_frame,
                    text=f"你有 {', '.join(warn_parts)}，是否仔细检查？",
                    font=("微软雅黑", 11), bg=Theme.WARN_BG,
                    fg=Theme.WARN_TEXT,
                    anchor="w", wraplength=400,
                    justify="left").pack(fill=tk.X, padx=12, pady=8)
                if marked_unanswered:
                    tk.Button(
                        mark_frame, text="检查标记题",
                        font=("微软雅黑", 10), bg=Theme.ACCENT,
                        fg="white", bd=0, padx=12, pady=4,
                        cursor="hand2",
                        command=lambda: self._check_marked_and_close(
                            preview_win, marked_unanswered)
                    ).pack(padx=12, pady=(0, 8), anchor="w")

        # ── 未答题列表 ──
        if unanswered > 0:
            tk.Label(container,
                     text=f"以下 {unanswered} 题尚未作答：",
                     font=("微软雅黑", 12, "bold"), bg=Theme.WHITE,
                     fg=Theme.DANGER).pack(
                anchor="w", padx=15, pady=(10, 2))

            list_outer = tk.Frame(container, bg=Theme.WHITE, bd=1,
                                  relief=tk.SOLID)
            list_outer.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

            list_canvas = tk.Canvas(list_outer, bg=Theme.WHITE,
                                    highlightthickness=0)
            list_scrollbar = ttk.Scrollbar(list_outer, orient="vertical",
                                           command=list_canvas.yview)
            list_inner = tk.Frame(list_canvas, bg=Theme.WHITE)

            list_canvas.create_window((0, 0), window=list_inner, anchor="nw")
            list_canvas.configure(yscrollcommand=list_scrollbar.set)
            list_inner.bind(
                "<Configure>",
                lambda e: list_canvas.configure(
                    scrollregion=list_canvas.bbox("all")))
            list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            bind_mousewheel(list_inner, list_canvas)
            bind_mousewheel(list_canvas, list_canvas)

            for q in unanswered_list:
                tk.Label(list_inner,
                         text=f"  ✗ 第{q['题号']}题 · {q['题型']} · "
                              f"{q['分值']}分",
                         font=("微软雅黑", 10), bg=Theme.WHITE,
                         fg=Theme.DANGER,
                         anchor="w").pack(fill=tk.X, padx=5, pady=1)

            tk.Button(
                container, text="复制未答题号到剪贴板",
                font=("微软雅黑", 9), bg=Theme.BG, fg=Theme.TEXT,
                bd=1, relief=tk.SOLID, padx=10, pady=2,
                cursor="hand2",
                command=lambda: self._copy_unanswered_to_clipboard(
                    unanswered_list, preview_win)
            ).pack(padx=15, pady=(0, 4), anchor="w")

        # ── 底部提示 ──
        tk.Label(container,
                 text="注意：请勿修改程序中的其它任何内容。",
                 font=("微软雅黑", 8), bg=Theme.WHITE,
                 fg=Theme.MUTED).pack(pady=(5, 5))

        # ── 按钮栏 ──
        btn_frame = tk.Frame(container, bg=Theme.WHITE)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        def _confirm_submit():
            self.timer_running = False
            self._stop_timer.set()
            self.exam_submitted = True
            self._cleanup_backup()
            score_pct = self._compute_score_pct()
            self._save_exam_progress("completed", score_pct)
            preview_win.destroy()
            self._do_score_and_show_result()

        tk.Button(btn_frame, text="返回继续答题",
                  font=("微软雅黑", 12, "bold"), bg=Theme.BG,
                  fg=Theme.TEXT, bd=1, relief=tk.SOLID,
                  padx=20, pady=8, cursor="hand2",
                  command=preview_win.destroy).pack(side=tk.LEFT)

        tk.Button(btn_frame, text="确认交卷",
                  font=("微软雅黑", 12, "bold"), bg=Theme.DANGER,
                  fg="white", bd=0, padx=20, pady=8,
                  cursor="hand2",
                  command=_confirm_submit).pack(side=tk.RIGHT)

        # ── 居中调整 ──
        preview_win.update_idletasks()
        req_h = container.winfo_reqheight() + 40
        max_h = int(screen_h * 0.8)
        final_h = min(req_h, max_h)
        final_w = max(460, container.winfo_reqwidth() + 20)
        preview_win.geometry(f"{final_w}x{final_h}+{x}+{y}")

    def _check_marked_and_close(self, win, marked_indices):
        win.destroy()
        if marked_indices:
            self._nav_jump(marked_indices[0])

    def _copy_unanswered_to_clipboard(self, unanswered_list, win):
        nums = ", ".join(q["题号"] for q in unanswered_list)
        self.root.clipboard_clear()
        self.root.clipboard_append(nums)
        messagebox.showinfo("已复制",
                            f"未答题号已复制到剪贴板：\n{nums}",
                            parent=win)

    # ─────────────────────────────────────────────────────────────────
    #  成绩展示
    # ─────────────────────────────────────────────────────────────────
    def _do_score_and_show_result(self):
        self.score = 0
        total_score = 0
        results = []

        for q in self.questions:
            gn = q["题号"]
            q_type = q["题型"]
            try:
                qs = int(float(q["分值"]))
            except (ValueError, TypeError):
                qs = 0
            total_score += qs
            is_correct = False
            if gn in self.user_answers:
                user_ans = self._normalize_answer(
                    self.user_answers[gn], q_type)
                correct_ans = self._normalize_answer(
                    q["正确答案"], q_type)
                if q_type in ("填空", "程序填空"):
                    is_correct = self._check_fill_in(
                        user_ans, correct_ans)
                elif q_type in ("程序设计", "程序改错"):
                    is_correct = self._check_program_answer(
                        self.user_answers[gn], q["正确答案"])
                else:
                    is_correct = (user_ans == correct_ans)
                if is_correct:
                    self.score += qs
            results.append({
                "题号": gn, "题型": q_type, "分值": qs,
                "用户答案": self.user_answers.get(gn, "未作答"),
                "正确答案": q["正确答案"],
                "正确": is_correct, "题目": q["题目"],
            })

        if total_score == 0:
            total_score = 1

        self._clear_window()

        top_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=50)
        top_bar.pack(fill=tk.X)
        tk.Label(top_bar, text="考试结果",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(padx=20, pady=10)

        score_frame = tk.Frame(self.root, bg=Theme.WHITE, bd=1,
                               relief=tk.SOLID)
        score_frame.pack(fill=tk.X, padx=40, pady=20)
        pct = self.score / total_score * 100
        if pct >= 90:
            grade_color, grade_text = Theme.SUCCESS, "优秀"
        elif pct >= 60:
            grade_color, grade_text = Theme.PRIMARY, "及格"
        else:
            grade_color, grade_text = Theme.DANGER, "不及格"

        tk.Label(score_frame, text=f"{self.score} / {total_score}",
                 font=("微软雅黑", 36, "bold"), fg=grade_color,
                 bg=Theme.WHITE).pack(pady=(15, 0))
        tk.Label(score_frame,
                 text=f"正确率 {pct:.1f}%  ·  {grade_text}",
                 font=Theme.FONT_BOLD, fg=grade_color,
                 bg=Theme.WHITE).pack(pady=(0, 15))

        detail_frame = tk.Frame(self.root, bg=Theme.WHITE, bd=1,
                                relief=tk.SOLID)
        detail_frame.pack(fill=tk.BOTH, expand=True,
                          padx=40, pady=(0, 10))
        canvas = tk.Canvas(detail_frame, bg=Theme.WHITE,
                           highlightthickness=0)
        scrollbar = ttk.Scrollbar(detail_frame, orient="vertical",
                                  command=canvas.yview)
        inner = tk.Frame(canvas, bg=Theme.WHITE)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        inner.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")))
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        bind_mousewheel(inner, canvas)
        bind_mousewheel(canvas, canvas)

        for r in results:
            row = tk.Frame(inner, bg=Theme.WHITE)
            row.pack(fill=tk.X, padx=10, pady=2)
            icon = "✓" if r["正确"] else "✗"
            color = Theme.SUCCESS if r["正确"] else Theme.DANGER
            tk.Label(row, text=icon, font=Theme.FONT,
                     bg=Theme.WHITE).pack(side=tk.LEFT, padx=(5, 10))
            tk.Label(row,
                     text=f"{r['题号']} · {r['题型']} · {r['分值']}分",
                     font=Theme.FONT_SMALL, bg=Theme.WHITE,
                     fg=Theme.TEXT, width=20,
                     anchor="w").pack(side=tk.LEFT)
            tk.Label(row,
                     text=f"你的答案: {r['用户答案']}",
                     font=Theme.FONT_SMALL, bg=Theme.WHITE,
                     fg=Theme.MUTED, width=20,
                     anchor="w").pack(side=tk.LEFT)
            tk.Label(row,
                     text=f"正确答案: {r['正确答案']}",
                     font=Theme.FONT_SMALL, bg=Theme.WHITE,
                     fg=color,
                     anchor="w").pack(side=tk.LEFT, padx=(10, 0))
            tk.Button(row, text="查看",
                      font=("微软雅黑", 8), bg=Theme.BG, fg=Theme.TEXT,
                      bd=1, relief=tk.SOLID, cursor="hand2", padx=8,
                      pady=2,
                      command=lambda ri=r:
                      self._review_question(ri)).pack(
                side=tk.RIGHT, padx=5)

        btn_frame = tk.Frame(self.root, bg=Theme.BG)
        btn_frame.pack(fill=tk.X, padx=40, pady=(0, 20))
        tk.Button(btn_frame, text="返回选卷",
                  font=Theme.FONT_BOLD, bg=Theme.PRIMARY, fg="white",
                  bd=0, padx=30, pady=10, cursor="hand2",
                  command=self.create_select_window).pack(side=tk.RIGHT)

    def _review_question(self, result):
        win = tk.Toplevel(self.root)
        win.title(f"题目回顾 - {result['题号']}")
        win.configure(bg=Theme.WHITE)
        win.resizable(True, True)

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                win.iconbitmap(icon_path)
            except Exception:
                pass

        hdr = tk.Frame(win, bg=Theme.PRIMARY, height=50)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        icon_txt = "✓" if result["正确"] else "✗"
        tk.Label(hdr,
                 text=f"{icon_txt}  {result['题型']} - "
                      f"题号 {result['题号']} - {result['分值']}分",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 12, "bold")).pack(
            side=tk.LEFT, padx=20, pady=10)

        body_canvas = tk.Canvas(win, bg=Theme.WHITE,
                                highlightthickness=0)
        body_sb = ttk.Scrollbar(win, orient="vertical",
                                command=body_canvas.yview)
        body = tk.Frame(body_canvas, bg=Theme.WHITE)
        body_win_id = body_canvas.create_window(
            (0, 0), window=body, anchor="nw")
        body_canvas.configure(yscrollcommand=body_sb.set)
        body.bind(
            "<Configure>",
            lambda e: body_canvas.configure(
                scrollregion=body_canvas.bbox("all")))
        body_sb.pack(side=tk.RIGHT, fill=tk.Y)
        body_canvas.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        bind_mousewheel(body, body_canvas)
        bind_mousewheel(body_canvas, body_canvas)

        _review_labels = []

        def _update_wrap(event):
            new_wrap = max(200, event.width - 60)
            body_canvas.itemconfig(body_win_id, width=event.width)
            for lbl in _review_labels:
                try:
                    lbl.config(wraplength=new_wrap)
                except tk.TclError:
                    pass

        body_canvas.bind("<Configure>", _update_wrap)

        tk.Label(body, text="题目",
                 font=("微软雅黑", 11, "bold"), bg=Theme.WHITE,
                 fg=Theme.TEXT, anchor="w").pack(fill=tk.X, pady=(0, 5))
        q_lbl = tk.Label(body, text=result.get("题目", "暂无"),
                         font=("微软雅黑", 11), bg=Theme.WHITE,
                         fg=Theme.TEXT, wraplength=650, justify="left",
                         anchor="w")
        q_lbl.pack(fill=tk.X, pady=(0, 15))
        _review_labels.append(q_lbl)

        if result["题型"] == "单选":
            for q in self.questions:
                if q["题号"] == result["题号"]:
                    opts = self._get_options_for_question(q)
                    if opts:
                        tk.Label(body, text="选项：",
                                 font=("微软雅黑", 11, "bold"),
                                 bg=Theme.WHITE, fg=Theme.TEXT,
                                 anchor="w").pack(
                            fill=tk.X, pady=(0, 5))
                        for ol, ot in opts:
                            opt_lbl = tk.Label(
                                body, text=f"  ({ol}) {ot}",
                                font=("微软雅黑", 11),
                                bg=Theme.WHITE, fg=Theme.TEXT,
                                anchor="w", wraplength=650,
                                justify="left")
                            opt_lbl.pack(fill=tk.X, padx=20, pady=2)
                            _review_labels.append(opt_lbl)
                    break

        ttk.Separator(body, orient="horizontal").pack(fill=tk.X,
                                                      pady=10)
        color = Theme.SUCCESS if result["正确"] else Theme.DANGER
        ans_frame = tk.Frame(body, bg=Theme.SURFACE, bd=1,
                             relief=tk.SOLID)
        ans_frame.pack(fill=tk.X, pady=5)
        tk.Label(ans_frame,
                 text=f"你的答案：{result['用户答案']}",
                 font=("微软雅黑", 11, "bold"), bg=Theme.SURFACE,
                 fg=color, anchor="w", wraplength=650,
                 justify="left").pack(fill=tk.X, padx=15, pady=8)
        tk.Label(ans_frame,
                 text=f"正确答案：{result['正确答案']}",
                 font=("Consolas", 11, "bold"), bg=Theme.SURFACE,
                 fg=Theme.SUCCESS, anchor="w", wraplength=650,
                 justify="left").pack(fill=tk.X, padx=15, pady=(0, 8))

        tk.Button(body, text="关闭",
                  font=("微软雅黑", 11), bg=Theme.BG, fg=Theme.TEXT,
                  padx=20, pady=5, cursor="hand2",
                  command=win.destroy).pack(pady=20)

        win.withdraw()
        win.update_idletasks()
        req_w = win.winfo_reqwidth()
        req_h = win.winfo_reqheight()
        scr_w = win.winfo_screenwidth()
        scr_h = win.winfo_screenheight()
        W = max(500, min(req_w + 40, 800, int(scr_w * 0.8)))
        H = max(350, min(req_h + 20, int(scr_h * 0.7)))
        x = (scr_w - W) // 2
        y = (scr_h - H) // 2
        win.geometry(f"{W}x{H}+{x}+{y}")
        win.minsize(400, 300)
        win.deiconify()
        win.grab_set()


# =====================================================================
#  启动入口
# =====================================================================
if __name__ == "__main__":
    try:
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "HNUST.ExamSystem.V1.0")

        root = tk.Tk()
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(default=icon_path)
                root.iconbitmap(icon_path)
            except Exception:
                pass
        app = HNUSTExamSystem(root)
        root.mainloop()

    except Exception as e:
        import traceback

        log_dir = _get_log_dir()
        log_name = f"exam_crash_{time.strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(log_dir, log_name)

        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"{'=' * 60}\n")
                f.write(f"启动错误: "
                        f"{time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"版本: {CURRENT_VERSION}\n")
                f.write(f"系统: {sys.platform} / {os.name}\n")
                f.write(f"Python: {sys.version}\n")
                f.write(f"{'-' * 60}\n")
                f.write(traceback.format_exc())
                f.write("\n")
        except Exception:
            log_path = "(日志写入失败)"

        try:
            root_err = tk.Tk()
            root_err.withdraw()

            err_win = tk.Toplevel(root_err)
            err_win.title("启动失败")
            err_win.configure(bg=Theme.CRASH_BG)
            err_win.resizable(False, False)
            ew_w, ew_h = 520, 320
            err_win.update_idletasks()
            ex = (err_win.winfo_screenwidth() - ew_w) // 2
            ey = (err_win.winfo_screenheight() - ew_h) // 2
            err_win.geometry(f"{ew_w}x{ew_h}+{ex}+{ey}")

            hdr = tk.Frame(err_win, bg=Theme.CRASH_HEADER, height=60)
            hdr.pack(fill=tk.X)
            hdr.pack_propagate(False)
            tk.Label(hdr, text="程序启动失败",
                     bg=Theme.CRASH_HEADER, fg="white",
                     font=("微软雅黑", 14, "bold")).pack(
                side=tk.LEFT, padx=20, pady=12)

            body = tk.Frame(err_win, bg=Theme.CRASH_BG)
            body.pack(fill=tk.BOTH, expand=True, padx=24, pady=16)

            tk.Label(body, text=f"错误信息：{str(e)[:200]}",
                     font=("微软雅黑", 10), bg=Theme.CRASH_BG,
                     fg=Theme.TEXT, wraplength=460, justify="left",
                     anchor="w").pack(fill=tk.X, pady=(0, 10))
            tk.Label(body, text=f"日志文件：{log_path}",
                     font=("微软雅黑", 9), bg=Theme.CRASH_BG,
                     fg=Theme.MUTED, wraplength=460, justify="left",
                     anchor="w").pack(fill=tk.X, pady=(0, 15))

            btn_frame = tk.Frame(body, bg=Theme.CRASH_BG)
            btn_frame.pack(fill=tk.X)

            def _open_log_dir():
                try:
                    if sys.platform == "win32":
                        os.startfile(log_dir)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", log_dir])
                    else:
                        subprocess.run(["xdg-open", log_dir])
                except Exception:
                    pass

            tk.Button(btn_frame, text="打开日志文件夹",
                      font=("微软雅黑", 10), bg=Theme.BG,
                      fg=Theme.TEXT, padx=12, pady=5, cursor="hand2",
                      command=_open_log_dir).pack(
                side=tk.LEFT, padx=(0, 10))
            tk.Button(btn_frame, text="关闭",
                      font=("微软雅黑", 10), bg=Theme.CRASH_HEADER,
                      fg="white", padx=20, pady=5, cursor="hand2",
                      command=root_err.destroy).pack(side=tk.RIGHT)

            err_win.protocol("WM_DELETE_WINDOW", root_err.destroy)
            err_win.grab_set()
            root_err.mainloop()
        except Exception:
            pass
