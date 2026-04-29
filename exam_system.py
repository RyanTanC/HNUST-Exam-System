import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import requests
import webbrowser
import os
import sys
import time
import re
import shutil
import subprocess
from threading import Thread

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass


# =====================================================================
#  版本更新检查 — 精美弹窗 + 建议更新（非强制，不阻塞启动）
# =====================================================================
CURRENT_VERSION = "v1.0.4"
GITHUB_USERNAME = "RyanTanC"
GITHUB_REPO_NAME = "HNUST-Exam-System"

_SKIP_VERSION_FILE = os.path.join(os.path.expanduser("~"), ".hnust_exam_skip_ver")


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
        with open(_SKIP_VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(ver)
    except Exception:
        pass


def _fetch_update_info():
    try:
        repo_api_url = (
            f"https://api.github.com/repos/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}/releases/latest"
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
            print(f"[更新检查] 用户已跳过版本 {latest_version}，跳过提示")
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
                from datetime import datetime
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

    except requests.exceptions.Timeout:
        print("[更新检查] 请求超时，跳过")
    except requests.exceptions.ConnectionError:
        print("[更新检查] 网络连接失败，跳过")
    except requests.exceptions.HTTPError as e:
        print(f"[更新检查] HTTP 错误: {e}")
    except Exception as e:
        print(f"[更新检查] 未知错误: {type(e).__name__}: {e}")

    return None


def _show_update_dialog(parent, info):
    win = tk.Toplevel(parent)
    win.title("发现新版本")
    win.resizable(False, False)
    win.configure(bg="#ffffff")
    win.attributes("-topmost", True)

    WIN_W, WIN_H = 520, 580
    win.update_idletasks()
    sx = (win.winfo_screenwidth() - WIN_W) // 2
    sy = (win.winfo_screenheight() - WIN_H) // 2
    win.geometry(f"{WIN_W}x{WIN_H}+{sx}+{sy}")

    icon_path = get_resource_path("icon.ico")
    if os.path.exists(icon_path):
        try:
            win.iconbitmap(icon_path)
        except Exception:
            pass

    BG = "#ffffff"
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    DANGER = "#dc2626"
    SURFACE = "#f1f5f9"
    BORDER = "#e2e8f0"
    TEXT = "#1e293b"
    MUTED = "#64748b"
    GREEN = "#16a34a"

    current_ver = CURRENT_VERSION
    latest_ver = info["latest_ver"]
    release_notes = info["release_notes"]
    download_url = info["download_url"]
    published_at = info.get("published_at", "")

    header = tk.Frame(win, bg=PRIMARY, height=100)
    header.pack(fill=tk.X)
    header.pack_propagate(False)

    icon_circle = tk.Canvas(header, width=56, height=56, bg=PRIMARY, highlightthickness=0)
    icon_circle.pack(side=tk.LEFT, padx=(30, 15), pady=22)
    icon_circle.create_oval(2, 2, 54, 54, fill="#ffffff", outline="#ffffff", width=2)
    icon_circle.create_text(28, 28, text="↑", font=("微软雅黑", 22, "bold"), fill=PRIMARY)

    header_text = tk.Frame(header, bg=PRIMARY)
    header_text.pack(side=tk.LEFT, pady=22)

    tk.Label(header_text, text="发现新版本",
             font=("微软雅黑", 18, "bold"), bg=PRIMARY, fg="#ffffff").pack(anchor="w")
    tk.Label(header_text, text=f"v{current_ver.lstrip('v')} → {latest_ver}",
             font=("微软雅黑", 11), bg=PRIMARY, fg="#bfdbfe").pack(anchor="w")

    body = tk.Frame(win, bg=BG)
    body.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

    info_card = tk.Frame(body, bg=SURFACE, bd=0)
    info_card.pack(fill=tk.X, padx=24, pady=(20, 0))

    info_grid = tk.Frame(info_card, bg=SURFACE)
    info_grid.pack(padx=16, pady=14, anchor="w")

    info_items = [
        ("当前版本", current_ver, MUTED),
        ("最新版本", latest_ver, GREEN),
    ]
    if published_at:
        info_items.append(("发布时间", published_at, MUTED))

    for i, (label, value, color) in enumerate(info_items):
        tk.Label(info_grid, text=label + "：",
                 font=("微软雅黑", 10), bg=SURFACE, fg=MUTED,
                 anchor="e", width=10).grid(row=i, column=0, sticky="e", pady=3)
        tk.Label(info_grid, text=value,
                 font=("微软雅黑", 10, "bold"), bg=SURFACE, fg=color,
                 anchor="w").grid(row=i, column=1, sticky="w", pady=3, padx=(8, 0))

    tk.Label(body, text="更新日志",
             font=("微软雅黑", 11, "bold"), bg=BG, fg=TEXT,
             anchor="w").pack(fill=tk.X, padx=24, pady=(16, 6))

    notes_frame = tk.Frame(body, bg=BG, bd=1, relief=tk.SOLID, highlightbackground=BORDER,
                           highlightthickness=1)
    notes_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 12))

    notes_canvas = tk.Canvas(notes_frame, bg="#fafbfc", highlightthickness=0)
    notes_scrollbar = ttk.Scrollbar(notes_frame, orient="vertical", command=notes_canvas.yview)
    notes_inner = tk.Frame(notes_canvas, bg="#fafbfc")

    notes_canvas.create_window((0, 0), window=notes_inner, anchor="nw")
    notes_canvas.configure(yscrollcommand=notes_scrollbar.set)
    notes_inner.bind("<Configure>",
                     lambda e: notes_canvas.configure(scrollregion=notes_canvas.bbox("all")))

    notes_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    notes_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    for line in release_notes.splitlines():
        line = line.strip()
        if not line:
            tk.Label(notes_inner, text="", bg="#fafbfc", height=1).pack()
            continue
        if line.startswith("### "):
            tk.Label(notes_inner, text=line[4:],
                     font=("微软雅黑", 10, "bold"), bg="#fafbfc", fg=TEXT,
                     anchor="w", wraplength=440, justify="left").pack(fill=tk.X, padx=12, pady=(8, 2))
        elif line.startswith("## "):
            tk.Label(notes_inner, text=line[3:],
                     font=("微软雅黑", 11, "bold"), bg="#fafbfc", fg=TEXT,
                     anchor="w", wraplength=440, justify="left").pack(fill=tk.X, padx=12, pady=(8, 2))
        elif line.startswith("- ") or line.startswith("* "):
            tk.Label(notes_inner, text=f"  •  {line[2:]}",
                     font=("微软雅黑", 10), bg="#fafbfc", fg=TEXT,
                     anchor="w", wraplength=430, justify="left").pack(fill=tk.X, padx=16, pady=1)
        else:
            tk.Label(notes_inner, text=line,
                     font=("微软雅黑", 10), bg="#fafbfc", fg=TEXT,
                     anchor="w", wraplength=440, justify="left").pack(fill=tk.X, padx=12, pady=1)

    def _mw(e):
        if sys.platform == "darwin":
            notes_canvas.yview_scroll(-e.delta, "units")
        else:
            notes_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    notes_canvas.bind("<MouseWheel>", _mw)
    notes_canvas.bind("<Button-4>", lambda e: notes_canvas.yview_scroll(-3, "units"))
    notes_canvas.bind("<Button-5>", lambda e: notes_canvas.yview_scroll(3, "units"))

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
    later_btn.bind("<Enter>", lambda e: later_btn.config(fg=DANGER))
    later_btn.bind("<Leave>", lambda e: later_btn.config(fg=MUTED))
    later_btn.bind("<Button-1>", lambda e: _on_continue())

    update_btn = tk.Label(btn_inner, text="   立即更新   ",
                          font=("微软雅黑", 11, "bold"), bg=PRIMARY, fg="#ffffff",
                          padx=24, pady=8, cursor="hand2", relief=tk.FLAT)
    update_btn.pack(side=tk.LEFT)
    update_btn.bind("<Enter>", lambda e: update_btn.config(bg=PRIMARY_HOVER))
    update_btn.bind("<Leave>", lambda e: update_btn.config(bg=PRIMARY))
    update_btn.bind("<Button-1>", lambda e: _on_update())

    tk.Label(body, text="建议更新到最新版本以获得最佳体验（也可以稍后更新）",
             font=("微软雅黑", 8), bg=BG, fg=MUTED).pack(pady=(0, 4))

    win.grab_set()
    win.focus_force()
    parent.wait_window(win)


# =====================================================================


class Theme:
    PRIMARY = "#0078d7"
    ACCENT = "#ff9900"
    BG = "#f0f0f0"
    WHITE = "#ffffff"
    TEXT = "#333333"
    HINT_BG = "#e6f2ff"
    DANGER = "#ff6600"
    SUCCESS = "#28a745"
    MUTED = "#999999"
    BORDER = "#e0e0e0"
    NAV_ACTIVE = "#cce0ff"
    NAV_CURRENT = "#ff9900"
    NAV_ANSWERED = "#0078d7"
    NAV_MARKED = "#ff6600"
    FONT = ("微软雅黑", 11)
    FONT_BOLD = ("微软雅黑", 11, "bold")
    FONT_TITLE = ("微软雅黑", 12, "bold")
    FONT_HUGE = ("微软雅黑", 16, "bold")
    FONT_SMALL = ("微软雅黑", 9)
    FONT_TINY = ("微软雅黑", 7)


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
                creationflags=NO_WINDOW
            )
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
            creationflags=NO_WINDOW
        )
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
                                with winreg.OpenKey(key, f"{version}\\InstallPath") as ik:
                                    path = winreg.QueryValue(ik, "")
                                    exe = os.path.join(path, "python.exe")
                                    if os.path.isfile(exe) and exe not in candidates:
                                        candidates.append(exe)
                            except Exception:
                                continue
                except Exception:
                    continue
    except Exception:
        pass

    local_app = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
    home = os.path.expanduser("~")

    common_paths = []
    for v in range(8, 20):
        common_paths.extend([
            os.path.join(local_app, f"Programs\\Python\\Python3{v}\\python.exe"),
            f"C:\\Python3{v}\\python.exe",
            os.path.join(program_files, f"Python3{v}\\python.exe"),
            os.path.join(program_files_x86, f"Python3{v}\\python.exe"),
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

    if candidates:
        return candidates[0]
    return None


class HNUSTExamSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("HNUST仿真平台")
        self.root.geometry("1200x800")
        self.root.configure(bg=Theme.BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if sys.platform == "win32":
            import ctypes
            app_id = "HNUST.ExamSystem.V1.0"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except:
                pass
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("HNUST.ExamSystem")
        except:
            pass

        self._add_anti_sale_watermark()

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
        self._timer_thread = None
        self.answer_var = None
        self.answer_text = None
        self.answer_entry = None
        self._choice_buttons = {}
        self.exam_submitted = False

        self.question_type_order = ["单选", "填空", "判断", "程序填空", "程序改错"]
        self.is_pure_program_exam = False
        self.pure_program_type_order = ["程序设计"]
        self.active_type_order = []

        self._backup_dir = None

        self.create_welcome_window()

        self._check_updates_async()

    def _check_updates_async(self):
        def _worker():
            try:
                info = _fetch_update_info()
                if info:
                    try:
                        self.root.after(1000, lambda: self._prompt_update(info))
                    except tk.TclError:
                        pass
            except Exception as e:
                print(f"[更新检查] 后台检查失败: {e}")

        Thread(target=_worker, daemon=True).start()

    def _prompt_update(self, info):
        try:
            _show_update_dialog(self.root, info)
        except tk.TclError:
            pass

    def _on_close(self):
        self.timer_running = False
        if self._timer_thread and self._timer_thread.is_alive():
            self._timer_thread.join(timeout=2)
        self._cleanup_backup()
        self.root.destroy()

    def _add_anti_sale_watermark(self):
        watermark = tk.Label(
            self.root, text="该程序免费 禁止商用售卖",
            font=Theme.FONT_TINY, bg=Theme.BG, fg="#cccccc", bd=0)
        watermark.place(relx=0.99, rely=0.99, anchor="se")
        self.root.title("HNUST仿真平台 | 免费使用 禁止售卖")

    def not_implemented(self):
        messagebox.showinfo("提示", "该功能暂未实现，敬请期待")

    def _clear_window(self):
        for widget in self.root.winfo_children():
            if widget.winfo_class() != "Label" or "禁止商用售卖" not in widget.cget("text"):
                widget.destroy()

    def _clear_children(self, parent):
        for widget in parent.winfo_children():
            widget.destroy()

    def create_welcome_window(self):
        self._clear_window()

        title_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=80)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="🌐 HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 24, "bold")).pack(side=tk.LEFT, padx=30, pady=15)
        tk.Label(title_bar, text="v1.0.4",
                 bg=Theme.PRIMARY, fg="#cce0ff",
                 font=("微软雅黑", 12)).pack(side=tk.LEFT, padx=10, pady=15)

        main_frame = tk.Frame(self.root, bg="white", bd=1, relief=tk.SOLID)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=100, pady=40)

        tk.Label(main_frame, text="欢迎使用 HNUST 考试仿真平台",
                 font=("微软雅黑", 18, "bold"), bg="white", fg=Theme.PRIMARY).pack(pady=(30, 20))

        canvas = tk.Canvas(main_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="white")

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(40, 0))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 40))

        bind_mousewheel(scroll_frame, canvas)
        bind_mousewheel(canvas, canvas)

        sections = [
            ("📌 软件介绍", [
                "本软件是模仿学校机房万维考试系统开发的**免费练习工具**",
                "专为HNUST同学设计，让你在宿舍也能随时随地进行考试模拟练习",
                "完美还原考试界面和操作流程，提前熟悉考试环境",
                "如若想使用**程序设计、程序改错、填空**三种功能，你的电脑需预先配置好 Python 运行环境。❗❗❗"
            ]),
            ("💻 功能特点", [
                "✅ 支持单选、判断、填空、程序填空、程序改错、程序设计等多种题型",
                "✅ 自动判分，即时显示答题结果和正确答案",
                "✅ 一键打开程序文件，自动用IDLE编辑",
                "✅ 程序文件自动备份，支持一键重置",
                "✅ 题目导航，快速跳转到未答题目",
                "✅ 考试计时，时间到自动交卷"
            ]),
            ("🔓 开源声明", [
                "本软件**完全免费开源**，代码将托管在GitHub上",
                "任何人都可以自由下载、使用、修改和分发",
                "**严禁任何形式的商用售卖**，违者必究",
                "欢迎举报盗售问题，侵权必删"
            ]),
            ("🤖 开发说明", [
                "本项目采用 AI 辅助开发模式完成",
                "特别感谢：豆包、DeepSeek、小米MIMO 提供的AI编程支持",
                "如果觉得好用，欢迎给个Star支持一下作者"
            ]),
            ("📝 问题反馈", [
                "该应用为学生开发，可能存在一些bug和不完善的地方",
                "如果遇到任何问题或有改进建议",
                "欢迎通过以下方式联系作者反馈：",
                "  • 在频道私信作者",
                "  • 在GitHub提交Issue"
            ]),
            ("⚠️ 免责声明", [
                "本软件仅供学习交流使用，与学校官方考试系统无关",
                "题库内容由用户自行提供，作者不承担任何版权责任",
                "使用本软件产生的任何后果由用户自行承担"
            ])
        ]

        for title, content in sections:
            tk.Label(scroll_frame, text=title,
                     font=("微软雅黑", 14, "bold"), bg="white", fg=Theme.TEXT, anchor="w").pack(fill=tk.X, pady=(15, 8))
            for line in content:
                tk.Label(scroll_frame, text=line,
                         font=("微软雅黑", 11), bg="white", fg=Theme.TEXT,
                         wraplength=800, justify="left", anchor="w").pack(fill=tk.X, pady=2)

        bottom_frame = tk.Frame(self.root, bg=Theme.BG)
        bottom_frame.pack(fill=tk.X, padx=100, pady=(0, 40))

        self.agree_var = tk.BooleanVar(value=False)

        agree_frame = tk.Frame(bottom_frame, bg=Theme.BG)
        agree_frame.pack(side=tk.LEFT, padx=100, pady=10)

        BOX_SIZE = 26
        self.agree_canvas = tk.Canvas(agree_frame, width=BOX_SIZE, height=BOX_SIZE,
                                      bg=Theme.BG, highlightthickness=0)
        self.agree_canvas.pack(side=tk.LEFT, padx=(0, 10))

        def draw_box(checked=False):
            self.agree_canvas.delete("all")
            self.agree_canvas.create_rectangle(2, 2, BOX_SIZE - 2, BOX_SIZE - 2,
                                               outline="#999999", width=2, fill="white")
            if checked:
                self.agree_canvas.create_line(
                    7, 13, 11, 18, 19, 6,
                    width=3, fill=Theme.PRIMARY, capstyle=tk.ROUND, joinstyle=tk.ROUND)

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
            bg="#cccccc", fg="white",
            padx=40, pady=10, bd=0, cursor="arrow",
            state=tk.DISABLED)
        self.enter_btn.pack(side=tk.RIGHT, padx=20)

        self._welcome_countdown(10)

    def _welcome_countdown(self, remaining):
        if remaining > 0:
            self.enter_btn.config(text=f"进入系统（请等待 {remaining} 秒）")
            self.root.after(1000, lambda: self._welcome_countdown(remaining - 1))
        else:
            self.enter_btn.config(
                text="进入系统",
                bg=Theme.PRIMARY,
                fg="white",
                cursor="hand2",
                state=tk.NORMAL)

    def create_select_window(self):
        self._clear_window()

        title_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=60)
        title_bar.pack(fill=tk.X)
        tk.Label(title_bar, text="🌐 HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(side=tk.LEFT, padx=20, pady=10)

        main_frame = tk.Frame(self.root, bg=Theme.BG)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="请选择试卷",
                 font=("微软雅黑", 24, "bold"), bg=Theme.BG).pack(pady=80)

        exam_dir = get_resource_path("题库")
        if not os.path.exists(exam_dir):
            external_exam_dir = "题库"
            if not os.path.exists(external_exam_dir):
                os.makedirs(external_exam_dir)
                messagebox.showinfo("提示", "题库文件夹已创建，请将Excel试卷文件放入其中")
            exam_dir = external_exam_dir

        self.exam_files = [f for f in os.listdir(exam_dir) if f.endswith(".xlsx")]

        if not self.exam_files:
            tk.Label(main_frame, text="题库文件夹中没有找到试卷文件",
                     font=("微软雅黑", 14), fg="red", bg=Theme.BG).pack(pady=20)
            return

        list_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID)
        list_frame.pack(pady=20)

        self.exam_listbox = tk.Listbox(
            list_frame, font=("微软雅黑", 14), width=50, height=12,
            bd=0, highlightthickness=0, selectbackground=Theme.PRIMARY)
        for file in self.exam_files:
            self.exam_listbox.insert(tk.END, os.path.splitext(file)[0])
        self.exam_listbox.selection_set(0)
        self.exam_listbox.pack(padx=2, pady=2)

        self.exam_listbox.bind("<Double-Button-1>", lambda e: self.start_exam())
        self.exam_listbox.bind("<Return>", lambda e: self.start_exam())

        tk.Button(main_frame, text="开始考试",
                  font=("微软雅黑", 16, "bold"), command=self.start_exam,
                  bg=Theme.PRIMARY, fg="white",
                  padx=40, pady=12, bd=0, cursor="hand2").pack(pady=40)

        tk.Label(main_frame, text="该程序免费提供给HNUST学生使用，禁止任何形式的商用售卖",
                 font=("微软雅黑", 8), bg=Theme.BG, fg=Theme.MUTED).pack(side=tk.BOTTOM, pady=10)

    def start_exam(self):
        selected = self.exam_listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一份试卷")
            return

        exam_name = self.exam_listbox.get(selected[0])

        internal_exam_file = get_resource_path(os.path.join("题库", exam_name + ".xlsx"))
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

            dup_nums = df[df.duplicated(subset=["题号"], keep=False)]["题号"].unique()
            if len(dup_nums) > 0:
                preview = ", ".join(dup_nums[:5].tolist())
                messagebox.showwarning(
                    "警告",
                    f"发现重复题号：{preview}{'...' if len(dup_nums) > 5 else ''}\n"
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

            self._init_backup()

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

    # =====================================================================
    #  create_exam_window — 左侧加惯性滚动 + 底部栏固定高度
    # =====================================================================
    def create_exam_window(self):
        self._clear_window()

        top_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=50)
        top_bar.pack(fill=tk.X)
        tk.Label(top_bar, text="🌐 HNUST仿真平台",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 12, "bold")).pack(side=tk.LEFT, padx=15, pady=10)
        info_frame = tk.Frame(top_bar, bg=Theme.PRIMARY)
        info_frame.pack(side=tk.RIGHT, padx=15)
        exam_name = os.path.splitext(os.path.basename(self.current_exam_file))[0]
        tk.Label(info_frame, text="姓名：xxx  学号：xxxxxxxxxxx",
                 bg=Theme.PRIMARY, fg="white", font=("微软雅黑", 10)).pack(anchor="e")
        tk.Label(info_frame, text=f"{exam_name} · 练习",
                 bg=Theme.PRIMARY, fg="white", font=("微软雅黑", 10)).pack(anchor="e")

        main_frame = tk.Frame(self.root, bg=Theme.BG)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ========== 左侧：用 Canvas + Scrollbar 包裹，实现惯性滚动 ==========
        self.left_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_scroll_area = tk.Frame(self.left_frame, bg="white")
        left_scroll_area.pack(fill=tk.BOTH, expand=True)

        self._lp_canvas = tk.Canvas(left_scroll_area, bg="white", highlightthickness=0)
        lp_scrollbar = ttk.Scrollbar(left_scroll_area, orient="vertical",
                                     command=self._lp_canvas.yview)
        self._lp_canvas.configure(yscrollcommand=lp_scrollbar.set)
        self._lp_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lp_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._lp_inner = tk.Frame(self._lp_canvas, bg="white")
        lp_canvas_win = self._lp_canvas.create_window(
            (0, 0), window=self._lp_inner, anchor="nw")

        self._lp_cached_content_h = 1
        self._lp_cached_visible_h = 1

        def _on_lp_canvas_resize(e):
            self._lp_canvas.itemconfig(lp_canvas_win, width=e.width)
            self._lp_cached_visible_h = e.height

        self._lp_canvas.bind("<Configure>", _on_lp_canvas_resize)

        def _on_lp_inner_resize(e):
            self._lp_canvas.configure(scrollregion=self._lp_canvas.bbox("all"))
            self._lp_cached_content_h = e.height

        self._lp_inner.bind("<Configure>", _on_lp_inner_resize)
        # ==================================================================

        self.right_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID, width=250)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)
        self.right_frame.pack_propagate(False)

        self._build_nav_panels()

        # ★ 内容框架放入 self._lp_inner（可滚动区域）
        self.q_title_bar = tk.Frame(self._lp_inner, bg=Theme.PRIMARY)
        self.q_title_bar.pack(fill=tk.X, padx=20, pady=(20, 10))

        self.q_instruction = tk.Frame(self._lp_inner, bg="white")
        self.q_instruction.pack(fill=tk.X)

        self.q_content = tk.Frame(self._lp_inner, bg="white")
        self.q_content.pack(fill=tk.X, padx=20, pady=(10, 5))

        self.q_answer_area = tk.Frame(self._lp_inner, bg="white")
        self.q_answer_area.pack(fill=tk.X, padx=20, pady=10)

        self.q_feedback = tk.Frame(self._lp_inner, bg="white")
        self.q_feedback.pack(fill=tk.X, padx=20, pady=5)

        # ★ 初始化左侧惯性滚动引擎
        self._setup_left_scroll()

        # ★ 底部按钮栏 — 固定高度，不随窗口缩放
        bottom_bar = tk.Frame(self.root, bg=Theme.BG, height=60)
        bottom_bar.pack(fill=tk.X, padx=10, pady=(0, 10))
        bottom_bar.pack_propagate(False)

        btn_style = {"font": ("微软雅黑", 10), "bd": 1, "relief": tk.SOLID,
                     "cursor": "hand2", "padx": 10, "pady": 5}

        tk.Button(bottom_bar, text="上题", **btn_style, bg=Theme.BG,
                  command=self.prev_question).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="下题", **btn_style, bg=Theme.BG,
                  command=self.next_question).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="下一未答", **btn_style, bg=Theme.ACCENT, fg="white",
                  command=self.jump_next_unanswered).pack(side=tk.LEFT, padx=5)

        tk.Button(bottom_bar, text="答题", **btn_style, bg=Theme.ACCENT, fg="white",
                  command=self.open_program_file).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="试题文件夹", **btn_style, bg=Theme.BG,
                  command=self.open_exam_folder).pack(side=tk.LEFT, padx=5)

        tk.Button(bottom_bar, text="重做", **btn_style, bg=Theme.BG,
                  command=self.redo_question).pack(side=tk.LEFT, padx=5)

        self.mark_btn = tk.Button(bottom_bar, text="标记试题", **btn_style, bg=Theme.BG,
                                  command=self.toggle_mark)
        self.mark_btn.pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="答案", **btn_style, bg=Theme.BG,
                  command=self.show_answer).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="试题解析", **btn_style, bg=Theme.BG,
                  command=self.not_implemented).pack(side=tk.LEFT, padx=5)

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
                 font=Theme.FONT_TINY, bg=Theme.BG, fg="#cccccc").pack(side=tk.LEFT, padx=10)

        self.progress_frame = tk.Frame(self.root, bg="#e0e0e0", height=6)
        self.progress_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.progress_bar = tk.Frame(self.progress_frame, bg=Theme.SUCCESS, height=6)
        self.progress_bar.place(x=0, y=0, relheight=1.0, relwidth=0)

        self.answer_var = tk.StringVar()
        self.answer_text = None
        self.answer_entry = None
        self._choice_buttons = {}

        self.root.bind("<Left>", lambda e: self.prev_question())
        self.root.bind("<Right>", lambda e: self.next_question())
        self.root.bind("<Control-Return>", lambda e: self.submit_exam())
        self.root.bind("<Alt-n>", lambda e: self.jump_next_unanswered())
        self.root.bind("<Alt-a>", lambda e: self.show_answer())
        self.root.bind("<KeyPress>", self._on_key_press)

        self.show_question()

    def _on_key_press(self, event):
        if self.exam_submitted:
            return
        focused = self.root.focus_get()
        if focused is not None and isinstance(focused, (tk.Text, tk.Entry, ttk.Entry)):
            return
        if not self.questions:
            return

        current_q = self.questions[self.current_index]
        q_type = current_q["题型"]

        if q_type == "单选":
            key = event.char.upper()
            if key in [opt[0] for opt in self._get_options_for_question(current_q)]:
                self._choose(current_q["题号"], key)
        elif q_type == "判断":
            if event.char.lower() in ("t", "y", "1"):
                self._choose(current_q["题号"], "A")
            elif event.char.lower() in ("f", "n", "0"):
                self._choose(current_q["题号"], "B")

    def open_program_file(self):
        current_q = self.questions[self.current_index]
        program_file = current_q.get("程序文件", "").strip()

        if not program_file:
            messagebox.showinfo("提示", "该题目没有对应的程序文件")
            return

        if ".." in program_file or program_file.startswith(("/", "\\")):
            messagebox.showerror("错误", f"非法的文件路径：{program_file}")
            return

        exam_dir = os.path.dirname(self.current_exam_file)
        program_path = os.path.join(exam_dir, "试题文件夹", program_file)
        if not os.path.exists(program_path):
            program_path = os.path.join(exam_dir, program_file)

        if not os.path.exists(program_path):
            messagebox.showerror(
                "错误",
                f"找不到程序文件：{program_file}\n"
                f"系统尝试查找的路径：\n{program_path}\n"
                "请确保文件放在Excel同目录的\"试题文件夹\"中或Excel同目录下")
            return

        try:
            if program_file.lower().endswith(".py"):
                opened = self._open_with_idle(program_path)
                if not opened:
                    if sys.platform == "win32":
                        os.startfile(program_path)
                    else:
                        subprocess.run(["open", program_path], check=True)
                    messagebox.showinfo(
                        "提示",
                        f"已用默认程序打开：{program_file}\n"
                        "（未检测到系统安装的Python IDLE）\n"
                        "请确保已安装Python并添加到系统路径")
                else:
                    messagebox.showinfo(
                        "提示",
                        f"已用IDLE打开程序文件：{program_file}\n"
                        "在IDLE中修改代码，修改完成后按Ctrl+S保存\n"
                        "然后回到本系统输入答案")
            else:
                if sys.platform == "win32":
                    os.startfile(program_path)
                else:
                    subprocess.run(["open", program_path], check=True)
                messagebox.showinfo(
                    "提示",
                    f"已打开文件：{program_file}\n"
                    "修改完成后保存文件，然后回到本系统输入答案")
        except Exception as e:
            messagebox.showerror("错误", f"打开文件失败：{str(e)}")

    def _open_with_idle(self, file_path):
        abs_path = os.path.abspath(file_path)
        NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

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
                    creationflags=NO_WINDOW
                )
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
                        creationflags=NO_WINDOW
                    )
                    if check.returncode == 0 and "Python" in check.stdout:
                        subprocess.Popen(
                            [python_path, "-m", "idlelib", abs_path],
                            creationflags=NO_WINDOW)
                        return True
                except Exception:
                    continue

        python_exe = find_system_python()

        if python_exe is not None:
            python_dir = os.path.dirname(python_exe)

            try:
                subprocess.Popen(
                    [python_exe, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

            idle_pyw = os.path.join(python_dir, "Lib", "idlelib", "idle.pyw")
            if os.path.isfile(idle_pyw):
                try:
                    subprocess.Popen(
                        [python_exe, idle_pyw, abs_path],
                        creationflags=NO_WINDOW)
                    return True
                except Exception:
                    pass

            idle_exe = os.path.join(python_dir, "Scripts", "idle.exe")
            if os.path.isfile(idle_exe):
                try:
                    subprocess.Popen(
                        [idle_exe, abs_path],
                        creationflags=NO_WINDOW)
                    return True
                except Exception:
                    pass

            pythonw_exe = os.path.join(python_dir, "pythonw.exe")
            if os.path.isfile(pythonw_exe):
                try:
                    subprocess.Popen(
                        [pythonw_exe, "-m", "idlelib", abs_path],
                        creationflags=NO_WINDOW)
                    return True
                except Exception:
                    pass

        python_path = self._ask_user_for_python()
        if python_path:
            try:
                subprocess.Popen(
                    [python_path, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        return False

    def _ask_user_for_python(self):
        result = messagebox.askyesno(
            "未找到 Python",
            "系统未能自动找到 Python 环境。\n\n"
            "你的电脑上是否已安装 Python？\n\n"
            '  • 已安装 → 点击"是"，手动选择 python.exe\n'
            '  • 未安装 → 点击"否"，前往官网下载')

        if not result:
            messagebox.showinfo(
                "安装 Python",
                "请前往 python.org 下载安装 Python。\n\n"
                '安装时务必勾选 "Add Python to PATH" 选项！\n\n'
                "安装完成后重新打开本程序。")
            webbrowser.open("https://www.python.org/downloads/")
            return None

        from tkinter import filedialog

        initial_dir = "C:\\"
        for candidate in [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs\\Python"),
            os.environ.get("ProgramFiles", "C:\\Program Files"),
        ]:
            if os.path.isdir(candidate):
                initial_dir = candidate
                break

        python_path = filedialog.askopenfilename(
            title="找到 python.exe 并选择它",
            filetypes=[("python.exe", "python.exe"), ("所有文件", "*.*")],
            initialdir=initial_dir)

        if python_path and os.path.isfile(python_path):
            return python_path
        return None

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
            messagebox.showerror("错误", f"打开文件夹失败：{str(e)}")

    def reset_program_file(self):
        current_q = self.questions[self.current_index]
        program_file = current_q.get("程序文件", "").strip()

        if not program_file:
            messagebox.showinfo("提示", "该题目没有对应的程序文件")
            return

        if not messagebox.askyesno("确认", f"确定要重置 {program_file} 吗？\n所有修改将丢失，恢复到原始版本！"):
            return

        exam_dir = os.path.dirname(self.current_exam_file)
        target_path = os.path.join(exam_dir, "试题文件夹", program_file)
        if not os.path.exists(target_path):
            target_path = os.path.join(exam_dir, program_file)

        if self._backup_dir and os.path.exists(self._backup_dir):
            backup_file = os.path.join(self._backup_dir, program_file)
            if os.path.exists(backup_file):
                try:
                    shutil.copy2(backup_file, target_path)
                    messagebox.showinfo("成功", f"{program_file} 已恢复到原始版本！")
                    return
                except Exception as e:
                    messagebox.showerror("错误", f"恢复失败：{str(e)}")
                    return

        messagebox.showinfo("提示", "未找到备份文件，请手动从原始来源恢复")

    def redo_question(self):
        current_q = self.questions[self.current_index]
        global_num = current_q["题号"]

        self.user_answers.pop(global_num, None)

        self._clear_children(self.q_feedback)
        self.answer_label = tk.Label(
            self.q_feedback, text="",
            font=("微软雅黑", 11, "bold"), bg="white", fg="#0000ff",
            justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)

        q_type = current_q["题型"]

        if q_type in ("单选", "判断"):
            self.answer_var.set("")
            for opt_letter, btn in self._choice_buttons.items():
                try:
                    btn.config(bg=Theme.BG, fg="black")
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

    # =====================================================================
    #  题目导航栏 — 速度模型惯性滚动 + 弹性回弹 + 缓存边界
    # =====================================================================
    def _build_nav_panels(self):
        nav_header = tk.Frame(self.right_frame, bg="white")
        nav_header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(nav_header, text="题目导航",
                 font=Theme.FONT_TITLE, bg="white", anchor="center"
                 ).pack(fill=tk.X, pady=(12, 6), padx=5)
        tk.Frame(nav_header, bg=Theme.BORDER, height=1).pack(fill=tk.X, padx=8)

        nav_footer = tk.Frame(self.right_frame, bg="white")
        nav_footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(nav_footer, bg=Theme.BORDER, height=1).pack(fill=tk.X, padx=8)
        self.status_label = tk.Label(
            nav_footer,
            text=f"未答 {len(self.questions)}，已答 0，标记 0",
            font=("微软雅黑", 9), bg="white", fg="#666")
        self.status_label.pack(pady=(6, 10))

        scroll_area = tk.Frame(self.right_frame, bg="white")
        scroll_area.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        canvas = tk.Canvas(scroll_area, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_area, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        nav_inner = tk.Frame(canvas, bg="white")
        canvas_win = canvas.create_window((0, 0), window=nav_inner, anchor="nw")

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
        used_type_order = self.active_type_order

        first_type = True
        for idx, q_type in enumerate(used_type_order):
            count = len(self.question_groups.get(q_type, []))
            if count == 0:
                continue

            if not first_type:
                tk.Frame(nav_inner, bg=Theme.BORDER, height=1).pack(fill=tk.X, pady=4)
            first_type = False

            header = tk.Frame(nav_inner, bg="#e8f0fe")
            header.pack(fill=tk.X, padx=4, pady=1)

            is_first_panel = (len(self.nav_panels) == 0)
            arrow_var = tk.StringVar(value="▼" if is_first_panel else "▶")
            arrow_lbl = tk.Label(header, textvariable=arrow_var,
                                 font=("微软雅黑", 9), bg="#e8f0fe",
                                 width=2, anchor="center")
            arrow_lbl.pack(side=tk.LEFT, padx=(4, 0))

            title_lbl = tk.Label(header, text=f"{q_type}（{count}题）",
                                 font=("微软雅黑", 10, "bold"),
                                 bg="#e8f0fe", anchor="w")
            title_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

            body = tk.Frame(nav_inner, bg="white")

            self.nav_panels[q_type] = {
                "header": header, "arrow": arrow_var,
                "body": body, "items": []
            }

            if q_type in self.question_groups:
                for type_idx, q in enumerate(self.question_groups[q_type]):
                    global_idx = q["_global_idx"]
                    global_num = q["题号"]
                    preview = q["题目"][:30] + ("..." if len(q["题目"]) > 30 else "")

                    btn = tk.Button(
                        body,
                        text=f"  第{type_idx + 1}题（{global_num}）",
                        font=("微软雅黑", 9), bg="white", fg=Theme.TEXT,
                        bd=0, anchor="w", padx=20, cursor="hand2",
                        activebackground=Theme.NAV_ACTIVE,
                        command=lambda gi=global_idx: self._nav_jump(gi))
                    btn.bind("<Enter>", lambda e, p=preview: self._show_tooltip(e, p))
                    btn.bind("<Leave>", lambda e: self._hide_tooltip())
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
                widget.bind("<Button-1>", lambda e, fn=_toggle: fn())

            if is_first_panel:
                body.pack(fill=tk.X, after=header)

        self._ns_canvas = canvas
        self._ns_inner = nav_inner
        self._ns_state = {
            'vel': 0.0,
            'aid': None,
            'bounce': False,
            'lt': 0.0,
        }

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
            if self._ns_state['vel'] > _MAX_VEL:
                self._ns_state['vel'] = _MAX_VEL
            elif self._ns_state['vel'] < -_MAX_VEL:
                self._ns_state['vel'] = -_MAX_VEL

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

                if new_pos < 0:
                    new_pos = 0
                elif new_pos > max_frac:
                    new_pos = max_frac

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
        spacer = tk.Frame(inner, bg="white", height=0)
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

        spacer = tk.Frame(inner, bg="white", height=0)
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

    # =====================================================================
    #  左侧题目区域 — 惯性滚动 + 弹性回弹（与导航栏同款物理模型）
    # =====================================================================
    def _setup_left_scroll(self):
        canvas = self._lp_canvas
        inner = self._lp_inner

        self._lp_state = {
            'vel': 0.0,
            'aid': None,
            'bounce': False,
            'lt': 0.0,
        }

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

                if new_pos < 0:
                    new_pos = 0
                elif new_pos > max_frac:
                    new_pos = max_frac

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
            if self._lp_state['vel'] > _MAX_VEL:
                self._lp_state['vel'] = _MAX_VEL
            elif self._lp_state['vel'] < -_MAX_VEL:
                self._lp_state['vel'] = -_MAX_VEL

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
        spacer = tk.Frame(inner, bg="white", height=0)
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

        spacer = tk.Frame(inner, bg="white", height=0)
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

    # =====================================================================

    def _show_tooltip(self, event, text):
        if hasattr(self, '_tooltip') and self._tooltip:
            self._tooltip.destroy()

        self._tooltip = tk.Toplevel(self.root)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.wm_attributes("-topmost", True)
        self._tooltip.wm_geometry(f"+{event.x_root + 15}+{event.y_root + 10}")

        label = tk.Label(self._tooltip, text=text,
                         font=Theme.FONT_SMALL, bg="#ffffe0",
                         relief=tk.SOLID, bd=1, wraplength=300,
                         justify="left", padx=8, pady=4)
        label.pack()

    def _hide_tooltip(self):
        if hasattr(self, '_tooltip') and self._tooltip:
            self._tooltip.destroy()
            self._tooltip = None

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
                 font=("微软雅黑", 10, "bold"), anchor="w").pack(fill=tk.X, padx=10, pady=8)

        self._clear_children(self.q_instruction)
        if q_type == "程序设计":
            tk.Label(self.q_instruction, text="<<答题说明>>",
                     font=("微软雅黑", 10, "bold"), bg="white",
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 5))
            tk.Label(self.q_instruction,
                     text=('1. 点击下方"答题"按钮，系统会用IDLE打开对应的程序文件\n'
                           '2. 在程序中完成需求（如补全代码、修复bug）\n'
                           '3. 修改完成后按Ctrl+S保存文件\n'
                           '4. 回到本系统，在答案框中输入核心代码（如关键函数/循环）'),
                     font=("微软雅黑", 9), bg="white", fg=Theme.TEXT,
                     wraplength=900, justify="left", anchor="w").pack(fill=tk.X, padx=20, pady=(0, 10))
            tk.Label(self.q_instruction,
                     text="注意：需按题目要求编写代码，保存后再提交答案。",
                     font=("微软雅黑", 9), bg="white", fg=Theme.TEXT,
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 20))
            ttk.Separator(self.q_instruction, orient="horizontal").pack(fill=tk.X, padx=20)
        elif q_type in ("程序改错", "程序填空"):
            tk.Label(self.q_instruction, text="<<答题说明>>",
                     font=("微软雅黑", 10, "bold"), bg="white",
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 5))
            tk.Label(self.q_instruction,
                     text=('1. 点击下方"答题"按钮，系统会用IDLE打开对应的程序文件\n'
                           '   （无需安装PyCharm，Python自带IDLE即可）\n'
                           '2. 在**********FOUND**********语句的下一行修改程序\n'
                           '3. 修改完成后按Ctrl+S保存文件\n'
                           '4. 回到本系统，在答案框中输入修改后的内容'),
                     font=("微软雅黑", 9), bg="white", fg=Theme.TEXT,
                     wraplength=900, justify="left", anchor="w").pack(fill=tk.X, padx=20, pady=(0, 10))
            tk.Label(self.q_instruction,
                     text="注意：不可以增加或删除程序行，也不可以更改程序的结构。",
                     font=("微软雅黑", 9), bg="white", fg=Theme.TEXT,
                     anchor="w").pack(fill=tk.X, padx=20, pady=(0, 20))
            ttk.Separator(self.q_instruction, orient="horizontal").pack(fill=tk.X, padx=20)

        self._clear_children(self.q_content)
        tk.Label(self.q_content, text=current_q["题目"],
                 font=("微软雅黑", 11), bg="white",
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
            font=("微软雅黑", 11, "bold"), bg="white", fg="#0000ff",
            justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)

        self._ensure_panel_open(q_type)
        self._update_nav_status()

        # ★ 切换题目时：回到顶部 + 停止残留惯性 + 重新绑定滚轮
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
            text = q.get(f"选项{letter}", q.get(f"选项 {letter}", "")).strip()
            if text:
                options.append((letter, text))
        return options

    def _build_single_choice(self, q, global_num):
        options_frame = tk.Frame(self.q_answer_area, bg="white")
        options_frame.pack(fill=tk.X, pady=(0, 10))

        options = self._get_options_for_question(q)
        for opt_letter, opt_text in options:
            tk.Label(options_frame, text=f"({opt_letter}){opt_text}",
                     font=("微软雅黑", 11), bg="white",
                     anchor="w").pack(fill=tk.X, pady=5)

        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG, bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar, text="在下面选择答案（点击选项字母进行选择）",
                 bg=Theme.HINT_BG, fg="#0066cc",
                 font=("微软雅黑", 9), anchor="w").pack(fill=tk.X, padx=10, pady=5)

        buttons_frame = tk.Frame(self.q_answer_area, bg="white")
        buttons_frame.pack(fill=tk.X, pady=10)

        current_answer = self.answer_var.get()
        for opt_letter, _ in options:
            selected = (current_answer == opt_letter)
            btn = tk.Button(
                buttons_frame, text=opt_letter,
                font=("微软雅黑", 16, "bold"),
                width=3, height=1, bd=1, relief=tk.SOLID,
                bg=Theme.PRIMARY if selected else Theme.BG,
                fg="white" if selected else "black",
                cursor="hand2",
                command=lambda o=opt_letter: self._choose(global_num, o))
            btn.pack(side=tk.LEFT, padx=20)
            self._choice_buttons[opt_letter] = btn

    def _build_judge(self, q, global_num):
        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG, bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar, text="在下面选择答案（点击对或错进行选择）",
                 bg=Theme.HINT_BG, fg="#0066cc",
                 font=("微软雅黑", 9), anchor="w").pack(fill=tk.X, padx=10, pady=5)

        buttons_frame = tk.Frame(self.q_answer_area, bg="white")
        buttons_frame.pack(fill=tk.X, pady=15)

        current_answer = self.answer_var.get()
        for label, value in [("对", "A"), ("错", "B")]:
            selected = (current_answer == value)
            btn = tk.Button(
                buttons_frame, text=label,
                font=("微软雅黑", 16, "bold"),
                width=5, height=1, bd=1, relief=tk.SOLID,
                bg=Theme.PRIMARY if selected else Theme.BG,
                fg="white" if selected else "black",
                cursor="hand2",
                command=lambda v=value: self._choose(global_num, v))
            btn.pack(side=tk.LEFT, padx=20)
            self._choice_buttons[value] = btn

    def _build_text_input(self, q, global_num, q_type):
        hint_bar = tk.Frame(self.q_answer_area, bg=Theme.HINT_BG, bd=1, relief=tk.SOLID)
        hint_bar.pack(fill=tk.X, pady=(0, 10))
        tk.Label(hint_bar, text="在下面输入答案",
                 bg=Theme.HINT_BG, fg="#0066cc",
                 font=("微软雅黑", 9), anchor="w").pack(fill=tk.X, padx=10, pady=5)

        input_frame = tk.Frame(self.q_answer_area, bg="white")
        input_frame.pack(fill=tk.X, pady=10)
        tk.Label(input_frame, text="答案：",
                 font=("微软雅黑", 11), bg="white").pack(side=tk.LEFT)

        if q_type in ("程序填空", "程序改错", "程序设计"):
            self.answer_text = tk.Text(input_frame,
                                       font=("Consolas", 11),
                                       width=60, height=5)
            self.answer_text.pack(side=tk.LEFT, padx=10)
            if global_num in self.user_answers:
                self.answer_text.insert(tk.END, self.user_answers[global_num])
            self.answer_text.bind("<KeyRelease>",
                                  lambda e: self._save_text(global_num))
            self.answer_text.bind("<FocusOut>",
                                  lambda e: self._save_text(global_num))
        else:
            self.answer_entry = ttk.Entry(input_frame,
                                          font=("微软雅黑", 11),
                                          width=40,
                                          textvariable=self.answer_var)
            self.answer_entry.pack(side=tk.LEFT, padx=10)
            self.answer_entry.bind("<KeyRelease>",
                                   lambda e: self._save_var(global_num))
            self.answer_entry.bind("<FocusOut>",
                                   lambda e: self._save_var(global_num))

    def _choose(self, global_num, option):
        self.user_answers[global_num] = option
        self.answer_var.set(option)

        current_q = self.questions[self.current_index]
        correct = self._normalize_answer(current_q["正确答案"], current_q["题型"])
        chosen = self._normalize_answer(option, current_q["题型"])

        if hasattr(self, '_choice_buttons'):
            for opt_letter, btn in self._choice_buttons.items():
                try:
                    normalized = self._normalize_answer(opt_letter, current_q["题型"])
                    if normalized == correct:
                        btn.config(bg=Theme.SUCCESS, fg="white")
                    elif opt_letter == option and chosen != correct:
                        btn.config(bg=Theme.DANGER, fg="white")
                    else:
                        btn.config(bg=Theme.BG, fg="black")
                except tk.TclError:
                    pass

        try:
            if chosen == correct:
                self.answer_label.config(text="✅ 回答正确！", fg=Theme.SUCCESS)
            else:
                self.answer_label.config(
                    text=f"❌ 回答错误，正确答案是：{current_q['正确答案']}",
                    fg=Theme.DANGER)
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

    def _update_nav_status(self):
        answered_count = len(self.user_answers)
        marked_count = len(self.marked_questions)

        used_type_order = self.active_type_order

        for q_type in used_type_order:
            if q_type not in self.question_groups:
                continue
            for type_idx, q in enumerate(self.question_groups[q_type]):
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
                    bg_color = Theme.NAV_ANSWERED
                    fg_color = "white"
                    font = ("微软雅黑", 9)
                else:
                    bg_color = "white"
                    fg_color = Theme.TEXT
                    font = ("微软雅黑", 9)

                btn_text = f"  第{type_idx + 1}题（{global_num}）"
                if global_idx in self.marked_questions:
                    btn_text = "🚩" + btn_text.strip()
                    if global_idx != self.current_index and global_num not in self.user_answers:
                        fg_color = Theme.NAV_MARKED

                btn.config(text=btn_text, bg=bg_color, fg=fg_color, font=font)

        if hasattr(self, 'mark_btn'):
            if self.current_index in self.marked_questions:
                self.mark_btn.config(text="取消标记", fg=Theme.NAV_MARKED)
            else:
                self.mark_btn.config(text="标记试题", fg=Theme.TEXT)

        self.status_label.config(
            text=f"未答 {len(self.questions) - answered_count}，已答 {answered_count}，标记 {marked_count}")
        self._update_progress()

    def _update_progress(self):
        total = len(self.questions)
        if total == 0:
            return
        answered = len(self.user_answers)
        ratio = answered / total
        self.progress_bar.place(x=0, y=0, relheight=1.0, relwidth=ratio)

    def _auto_save_current(self):
        if self.answer_text is not None:
            try:
                answer = self.answer_text.get("1.0", tk.END).strip()
                global_num = self.questions[self.current_index]["题号"]
                if answer:
                    self.user_answers[global_num] = answer
                else:
                    self.user_answers.pop(global_num, None)
            except tk.TclError:
                pass
        elif self.answer_entry is not None:
            try:
                answer = self.answer_var.get().strip()
                global_num = self.questions[self.current_index]["题号"]
                if answer:
                    self.user_answers[global_num] = answer
                else:
                    self.user_answers.pop(global_num, None)
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
            font=("微软雅黑", 11, "bold"), bg="white", fg="#0000ff",
            justify="left", anchor="w")
        self.answer_label.pack(fill=tk.X)

        answer_frame = tk.Frame(self.q_feedback, bg="#e6ffe6", bd=1, relief=tk.SOLID)
        answer_frame.pack(fill=tk.X, pady=5)

        tk.Label(answer_frame, text="📋 标准答案",
                 font=("微软雅黑", 10, "bold"), bg="#e6ffe6",
                 fg=Theme.SUCCESS, anchor="w").pack(fill=tk.X, padx=10, pady=(8, 2))
        tk.Label(answer_frame, text=q["正确答案"],
                 font=("Consolas", 12, "bold"), bg="#e6ffe6",
                 fg=Theme.SUCCESS, anchor="w",
                 wraplength=800, justify="left").pack(fill=tk.X, padx=10, pady=(0, 10))

    def start_timer(self):
        self.timer_running = True
        self._timer_thread = Thread(target=self._tick, daemon=True)
        self._timer_thread.start()

    def _tick(self):
        while self.timer_running and self.remaining_time > 0:
            time.sleep(1)
            if not self.timer_running:
                return
            self.remaining_time -= 1
            h = self.remaining_time // 3600
            m = (self.remaining_time % 3600) // 60
            s = self.remaining_time % 60
            time_str = f"{h:02d}:{m:02d}:{s:02d}"

            if self.remaining_time <= 300:
                fg_color = "#ff0000"
            elif self.remaining_time <= 600:
                fg_color = "#ff6600"
            else:
                fg_color = Theme.TEXT

            try:
                self.root.after(0, lambda ts=time_str, fc=fg_color:
                (self.time_label.config(text=ts, fg=fc)))
            except tk.TclError:
                return

        if self.remaining_time <= 0 and self.timer_running:
            try:
                self.root.after(0, lambda: self._force_submit())
            except tk.TclError:
                pass

    def _force_submit(self):
        if self.exam_submitted:
            return
        self.timer_running = False
        self.exam_submitted = True
        try:
            self.time_label.config(text="00:00:00")
        except tk.TclError:
            pass
        messagebox.showinfo("提示", "考试时间到！系统将自动交卷。")
        self._do_score_and_show_result()

    def _normalize_answer(self, ans, q_type):
        ans = str(ans).strip().lower()
        if q_type == "判断":
            mapping = {
                "对": "a", "错": "b",
                "t": "a", "f": "b",
                "true": "a", "false": "b",
                "1": "a", "0": "b",
                "y": "a", "n": "b",
                "yes": "a", "no": "b",
                "正确": "a", "错误": "b",
                "√": "a", "×": "b",
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
        lines = [line.strip() for line in str(code).strip().splitlines()]
        lines = [line for line in lines if line]
        return "\n".join(lines)

    def submit_exam(self):
        if self.exam_submitted:
            return

        answered = len(self.user_answers)
        total = len(self.questions)
        unanswered = total - answered

        unanswered_list = []
        for q in self.questions:
            if q["题号"] not in self.user_answers:
                unanswered_list.append(q)

        preview_win = tk.Toplevel(self.root)
        preview_win.title("交卷确认")
        preview_win.configure(bg="white")
        preview_win.transient(self.root)
        preview_win.grab_set()

        preview_win.geometry("420x500")
        preview_win.minsize(420, 300)
        preview_win.resizable(True, True)

        preview_win.update_idletasks()
        screen_w = preview_win.winfo_screenwidth()
        screen_h = preview_win.winfo_screenheight()
        x = (screen_w - 420) // 2
        y = (screen_h - 500) // 2
        preview_win.geometry(f"+{x}+{y}")

        container = tk.Frame(preview_win, bg="white")
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(container, text="📋 交卷前检查",
                 font=Theme.FONT_TITLE, bg="white").pack(pady=(10, 5))

        stats_frame = tk.Frame(container, bg="white", bd=1, relief=tk.SOLID)
        stats_frame.pack(fill=tk.X, padx=15, pady=5)

        stats_data = [
            ("总题数", str(total), Theme.TEXT),
            ("已作答", str(answered), Theme.SUCCESS),
            ("未作答", str(unanswered), Theme.DANGER if unanswered > 0 else Theme.TEXT),
            ("标记数", str(len(self.marked_questions)),
             Theme.NAV_MARKED if len(self.marked_questions) > 0 else Theme.TEXT),
        ]

        for label, value, color in stats_data:
            row = tk.Frame(stats_frame, bg="white")
            row.pack(fill=tk.X, padx=10, pady=2)
            tk.Label(row, text=label, font=Theme.FONT, bg="white",
                     fg=Theme.TEXT, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=value, font=Theme.FONT_BOLD, bg="white",
                     fg=color, anchor="e").pack(side=tk.RIGHT)

        if unanswered > 0:
            tk.Label(container, text="⚠️ 以下题目尚未作答：",
                     font=Theme.FONT_BOLD, bg="white",
                     fg=Theme.DANGER).pack(anchor="w", padx=15, pady=(10, 2))

            list_outer = tk.Frame(container, bg="white", bd=1, relief=tk.SOLID)
            list_outer.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

            list_canvas = tk.Canvas(list_outer, bg="white", highlightthickness=0)
            list_scrollbar = ttk.Scrollbar(list_outer, orient="vertical",
                                           command=list_canvas.yview)
            list_inner = tk.Frame(list_canvas, bg="white")

            list_canvas.create_window((0, 0), window=list_inner, anchor="nw")
            list_canvas.configure(yscrollcommand=list_scrollbar.set)
            list_inner.bind("<Configure>",
                            lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all")))
            list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            bind_mousewheel(list_inner, list_canvas)
            bind_mousewheel(list_canvas, list_canvas)

            for q in unanswered_list:
                tk.Label(list_inner,
                         text=f"  ❌ 第{q['题号']}题 · {q['题型']} · {q['分值']}分",
                         font=Theme.FONT_SMALL, bg="white", fg=Theme.DANGER,
                         anchor="w").pack(fill=tk.X, padx=5, pady=1)

        tk.Label(container, text="注意：请勿修改程序中的其它任何内容。",
                 font=Theme.FONT_TINY, bg="white", fg=Theme.MUTED).pack(pady=(5, 10))

        btn_frame = tk.Frame(container, bg="white")
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15), side=tk.BOTTOM)

        def _confirm_submit():
            self._auto_save_current()
            preview_win.destroy()
            self.timer_running = False
            self.exam_submitted = True
            self._cleanup_backup()
            self._do_score_and_show_result()

        tk.Button(btn_frame, text="返回继续答题",
                  font=Theme.FONT_BOLD, bg=Theme.BG, bd=1, relief=tk.SOLID,
                  padx=15, pady=6, cursor="hand2",
                  command=preview_win.destroy).pack(side=tk.LEFT)

        tk.Button(btn_frame, text="确认交卷",
                  font=Theme.FONT_BOLD, bg=Theme.DANGER, fg="white", bd=0,
                  padx=20, pady=6, cursor="hand2",
                  command=_confirm_submit).pack(side=tk.RIGHT)

        preview_win.update_idletasks()
        req_h = container.winfo_reqheight() + 40
        max_h = int(screen_h * 0.8)
        final_h = min(req_h, max_h)
        final_w = max(420, container.winfo_reqwidth() + 20)
        preview_win.geometry(f"{final_w}x{final_h}+{x}+{y}")

    def _do_score_and_show_result(self):
        self.score = 0
        total_score = 0
        results = []

        for q in self.questions:
            global_num = q["题号"]
            q_type = q["题型"]
            try:
                question_score = int(float(q["分值"]))
            except (ValueError, TypeError):
                question_score = 0
            total_score += question_score

            is_correct = False
            if global_num in self.user_answers:
                user_ans = self._normalize_answer(self.user_answers[global_num], q_type)
                correct_ans = self._normalize_answer(q["正确答案"], q_type)
                if q_type in ("填空", "程序填空"):
                    is_correct = self._check_fill_in(user_ans, correct_ans)
                elif q_type == "程序设计":
                    is_correct = (self._normalize_code(user_ans) == self._normalize_code(correct_ans))
                else:
                    is_correct = (user_ans == correct_ans)
                if is_correct:
                    self.score += question_score

            results.append({
                "题号": global_num, "题型": q_type,
                "分值": question_score,
                "用户答案": self.user_answers.get(global_num, "未作答"),
                "正确答案": q["正确答案"],
                "正确": is_correct
            })

        if total_score == 0:
            total_score = 1

        self._clear_window()

        top_bar = tk.Frame(self.root, bg=Theme.PRIMARY, height=50)
        top_bar.pack(fill=tk.X)
        tk.Label(top_bar, text="📊 考试结果",
                 bg=Theme.PRIMARY, fg="white",
                 font=("微软雅黑", 16, "bold")).pack(padx=20, pady=10)

        score_frame = tk.Frame(self.root, bg="white", bd=1, relief=tk.SOLID)
        score_frame.pack(fill=tk.X, padx=40, pady=20)

        pct = self.score / total_score * 100
        if pct >= 90:
            grade_color, grade_text = Theme.SUCCESS, "优秀 🎉"
        elif pct >= 60:
            grade_color, grade_text = Theme.PRIMARY, "及格"
        else:
            grade_color, grade_text = Theme.DANGER, "不及格 😢"

        tk.Label(score_frame, text=f"{self.score} / {total_score}",
                 font=("微软雅黑", 36, "bold"), fg=grade_color,
                 bg="white").pack(pady=(15, 0))
        tk.Label(score_frame, text=f"正确率 {pct:.1f}%  ·  {grade_text}",
                 font=Theme.FONT_BOLD, fg=grade_color,
                 bg="white").pack(pady=(0, 15))

        detail_frame = tk.Frame(self.root, bg="white", bd=1, relief=tk.SOLID)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 10))

        canvas = tk.Canvas(detail_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(detail_frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        bind_mousewheel(inner, canvas)
        bind_mousewheel(canvas, canvas)

        for r in results:
            row = tk.Frame(inner, bg="white")
            row.pack(fill=tk.X, padx=10, pady=2)

            icon = "✅" if r["正确"] else "❌"
            color = Theme.SUCCESS if r["正确"] else Theme.DANGER

            tk.Label(row, text=icon, font=Theme.FONT, bg="white").pack(side=tk.LEFT, padx=(5, 10))
            tk.Label(row, text=f"{r['题号']} · {r['题型']} · {r['分值']}分",
                     font=Theme.FONT_SMALL, bg="white", fg=Theme.TEXT,
                     width=20, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=f"你的答案: {r['用户答案']}",
                     font=Theme.FONT_SMALL, bg="white", fg=Theme.MUTED,
                     width=20, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=f"正确答案: {r['正确答案']}",
                     font=Theme.FONT_SMALL, bg="white", fg=color,
                     anchor="w").pack(side=tk.LEFT, padx=(10, 0))

        btn_frame = tk.Frame(self.root, bg=Theme.BG)
        btn_frame.pack(fill=tk.X, padx=40, pady=(0, 20))

        tk.Button(btn_frame, text="返回选卷",
                  font=Theme.FONT_BOLD, bg=Theme.PRIMARY, fg="white", bd=0,
                  padx=30, pady=10, cursor="hand2",
                  command=self.create_select_window).pack(side=tk.RIGHT)


if __name__ == "__main__":
    try:
        if sys.platform == "win32":
            import ctypes
            app_id = "HNUST.ExamSystem.V1.0"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        root = tk.Tk()

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(default=icon_path)
                root.iconbitmap(icon_path)
            except Exception as e:
                print(f"设置图标失败: {e}")

        app = HNUSTExamSystem(root)
        root.mainloop()
    except Exception as e:
        import traceback

        log_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "error.log")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"启动错误: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(traceback.format_exc())
            try:
                root_err = tk.Tk()
                root_err.withdraw()
                messagebox.showerror(
                    "启动失败",
                    f"程序启动时发生错误：\n\n{str(e)}\n\n"
                    f"详细日志已保存到：\n{log_path}")
                root_err.destroy()
            except Exception:
                pass
        except Exception:
            pass
