import tkinter as tk
from tkinter import messagebox
import requests
import webbrowser
import sys
import os


def check_for_updates():
    """检查GitHub上是否有新版本"""
    try:
        # ====================== 配置信息（已填好，直接用） ======================
        GITHUB_USERNAME = "RyanTanC"
        GITHUB_REPO_NAME = "HNUST-Exam-System"
        CURRENT_VERSION = "v1.0.0-beta.2"  # 每次发布新版本时，把这里改成新版本号
        # =========================================================================

        repo_api_url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}/releases/latest"
        response = requests.get(repo_api_url, timeout=5)
        response.raise_for_status()

        latest_release = response.json()
        latest_version = latest_release["tag_name"]

        # 版本比较逻辑：正式版 > 测试版
        def version_key(v):
            parts = v.replace("v", "").split("-")
            main = [int(x) for x in parts[0].split(".")]
            pre = parts[1] if len(parts) > 1 else "z"  # z确保正式版排在后面
            return (main, pre)

        if version_key(latest_version) > version_key(CURRENT_VERSION):
            download_url = latest_release["assets"][0]["browser_download_url"]
            release_notes = latest_release["body"]

            root = tk.Tk()
            root.withdraw()

            result = messagebox.askyesno(
                "发现新版本",
                f"当前版本：{CURRENT_VERSION}\n"
                f"最新版本：{latest_version}\n\n"
                f"更新内容：\n{release_notes}\n\n"
                "是否立即下载更新？"
            )

            if result:
                webbrowser.open(download_url)
                sys.exit(0)

    except Exception as e:
        # 检查更新失败不影响程序运行
        print(f"检查更新失败：{e}")


# 在程序最开始调用（必须放在所有代码的最前面）
if __name__ == "__main__":
    check_for_updates()

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
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


# ★★★ 新增：查找系统中真实的 Python 路径 ★★★
def find_system_python():
    """
    在打包环境中找到系统安装的 Python 解释器路径。
    返回找到的 python.exe 路径，找不到返回 None。
    """
    if sys.platform != "win32":
        return None

    candidates = []

    # 方法1：从注册表查找
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
                                    if os.path.isfile(exe):
                                        candidates.append(exe)
                            except Exception:
                                continue
                except Exception:
                    continue
    except Exception:
        pass

    # 方法2：用 where 命令查找 PATH 中的 python
    try:
        result = subprocess.run(
            ["where", "python"],
            capture_output=True, text=True, timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        if result.returncode == 0:
            for line in result.stdout.strip().splitlines():
                line = line.strip()
                if line and os.path.isfile(line) and line not in candidates:
                    candidates.append(line)
    except Exception:
        pass

    # 方法3：常见安装路径
    local_app = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
    common_paths = [
        os.path.join(local_app, r"Programs\Python\Python313\python.exe"),
        os.path.join(local_app, r"Programs\Python\Python312\python.exe"),
        os.path.join(local_app, r"Programs\Python\Python311\python.exe"),
        os.path.join(local_app, r"Programs\Python\Python310\python.exe"),
        os.path.join(local_app, r"Programs\Python\Python39\python.exe"),
        r"C:\Python313\python.exe",
        r"C:\Python312\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python310\python.exe",
        r"C:\Python39\python.exe",
        os.path.join(program_files, r"Python313\python.exe"),
        os.path.join(program_files, r"Python312\python.exe"),
        os.path.join(program_files, r"Python311\python.exe"),
    ]
    for p in common_paths:
        if os.path.isfile(p) and p not in candidates:
            candidates.append(p)

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

        self._backup_dir = None

        self.create_welcome_window()

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
        tk.Label(title_bar, text="v1.0 内测版",
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
            ("📝 内测反馈", [
                "当前版本为**内测版**，可能存在一些bug和不完善的地方",
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

        # ★★★ 底部按钮区 — 带10秒倒计时 ★★★
        bottom_frame = tk.Frame(self.root, bg=Theme.BG)
        bottom_frame.pack(fill=tk.X, padx=100, pady=(0, 40))

        self.agree_var = tk.BooleanVar(value=False)
        agree_check = tk.Checkbutton(bottom_frame, text="我已阅读并同意以上所有条款",
                                     variable=self.agree_var, font=("微软雅黑", 11),
                                     bg=Theme.BG, fg=Theme.TEXT, cursor="hand2")
        agree_check.pack(side=tk.LEFT, padx=100)

        # 进入系统按钮（初始禁用）
        def enter_system():
            if not self.agree_var.get():
                messagebox.showwarning("提示", "请先阅读并同意以上条款")
                return
            self.create_select_window()

        self.enter_btn = tk.Button(
            bottom_frame, text="进入系统（请等待 10 秒）",
            font=("微软雅黑", 14, "bold"), command=enter_system,
            bg="#cccccc", fg="white",  # 灰色背景表示不可用
            padx=40, pady=10, bd=0, cursor="arrow",  # 箭头光标表示禁用
            state=tk.DISABLED)
        self.enter_btn.pack(side=tk.RIGHT, padx=20)

        # 启动倒计时
        self._welcome_countdown(10)

    def _welcome_countdown(self, remaining):
        """欢迎页倒计时，倒计时结束后启用进入按钮"""
        if remaining > 0:
            self.enter_btn.config(text=f"进入系统（请等待 {remaining} 秒）")
            # 每秒调用一次自身
            self.root.after(1000, lambda: self._welcome_countdown(remaining - 1))
        else:
            # 倒计时结束，恢复按钮
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

            self.user_answers = {}
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

        self.left_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.right_frame = tk.Frame(main_frame, bg="white", bd=1, relief=tk.SOLID, width=250)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)
        self.right_frame.pack_propagate(False)

        self._build_nav_panels()

        self.q_title_bar = tk.Frame(self.left_frame, bg=Theme.PRIMARY)
        self.q_title_bar.pack(fill=tk.X, padx=20, pady=(20, 10))

        self.q_instruction = tk.Frame(self.left_frame, bg="white")
        self.q_instruction.pack(fill=tk.X)

        self.q_content = tk.Frame(self.left_frame, bg="white")
        self.q_content.pack(fill=tk.X, padx=20, pady=(10, 5))

        self.q_answer_area = tk.Frame(self.left_frame, bg="white")
        self.q_answer_area.pack(fill=tk.X, padx=20, pady=10)

        self.q_feedback = tk.Frame(self.left_frame, bg="white")
        self.q_feedback.pack(fill=tk.X, padx=20, pady=5)

        bottom_bar = tk.Frame(self.root, bg=Theme.BG, height=60)
        bottom_bar.pack(fill=tk.X, padx=10, pady=(0, 10))

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
                  command=self.reset_program_file).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_bar, text="标记试题", **btn_style, bg=Theme.BG,
                  command=self.not_implemented).pack(side=tk.LEFT, padx=5)
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

    # ★★★ 核心修改1：重写 open_program_file ★★★
    def open_program_file(self):
        current_q = self.questions[self.current_index]
        program_file = current_q.get("程序文件", "").strip()

        if not program_file:
            messagebox.showinfo("提示", "该题目没有对应的程序文件")
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
                    # ★ 关键修复：用 CREATE_NO_WINDOW 避免弹出控制台
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

    # ★★★ 核心修改2：完全重写 _open_with_idle ★★★
    def _open_with_idle(self, file_path):
        """
        在打包环境中找到系统Python并用IDLE打开文件。
        不使用 sys.executable（打包后它是你的exe，不是python.exe）。
        """
        abs_path = os.path.abspath(file_path)
        NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        # ---- 非打包环境：直接用 sys.executable ----
        if not hasattr(sys, '_MEIPASS'):
            try:
                subprocess.Popen(
                    [sys.executable, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        # ---- 打包环境：查找系统中真实的 Python ----
        python_exe = find_system_python()

        if python_exe is None:
            return False

        python_dir = os.path.dirname(python_exe)

        # 尝试方式1：python.exe -m idlelib file.py
        try:
            subprocess.Popen(
                [python_exe, "-m", "idlelib", abs_path],
                creationflags=NO_WINDOW)
            return True
        except Exception:
            pass

        # 尝试方式2：直接运行 idle.pyw
        idle_pyw = os.path.join(python_dir, "Lib", "idlelib", "idle.pyw")
        if os.path.isfile(idle_pyw):
            try:
                subprocess.Popen(
                    [python_exe, idle_pyw, abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        # 尝试方式3：Scripts/idle.exe
        idle_exe = os.path.join(python_dir, "Scripts", "idle.exe")
        if os.path.isfile(idle_exe):
            try:
                subprocess.Popen(
                    [idle_exe, abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        # 尝试方式4：pythonw.exe -m idlelib（无控制台窗口）
        pythonw_exe = os.path.join(python_dir, "pythonw.exe")
        if os.path.isfile(pythonw_exe):
            try:
                subprocess.Popen(
                    [pythonw_exe, "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW)
                return True
            except Exception:
                pass

        return False

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

    def _build_nav_panels(self):
        tk.Label(self.right_frame, text="题目导航",
                 font=Theme.FONT_TITLE, bg="white").pack(pady=(10, 5))

        canvas = tk.Canvas(self.right_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.right_frame, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        nav_inner = tk.Frame(canvas, bg="white")
        canvas_window = canvas.create_window((0, 0), window=nav_inner, anchor="nw")

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        nav_inner.bind("<Configure>", _on_inner_configure)

        bind_mousewheel(nav_inner, canvas)
        bind_mousewheel(canvas, canvas)

        self.nav_panels = {}
        self.nav_q_buttons = {}

        if self.is_pure_program_exam:
            used_type_order = self.pure_program_type_order
        else:
            used_type_order = self.question_type_order

        for idx, q_type in enumerate(used_type_order):
            count = len(self.question_groups.get(q_type, []))
            if count == 0:
                continue

            if idx > 0:
                tk.Frame(nav_inner, bg=Theme.BORDER, height=1).pack(fill=tk.X, pady=4)

            header = tk.Frame(nav_inner, bg="#e8f0fe")
            header.pack(fill=tk.X, padx=4, pady=1)

            arrow_var = tk.StringVar(value="▼" if idx == 0 else "▶")
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

            def _toggle(qt=q_type, al=arrow_var, bd=body, hd=header, cv=canvas):
                if bd.winfo_ismapped():
                    bd.pack_forget()
                    al.set("▶")
                else:
                    bd.pack(fill=tk.X, after=hd)
                    al.set("▼")
                    bind_mousewheel(bd, cv)
                cv.after(50, lambda: cv.configure(scrollregion=cv.bbox("all")))

            for widget in (header, arrow_lbl, title_lbl):
                widget.bind("<Button-1>", lambda e, fn=_toggle: fn())

            if idx == 0:
                body.pack(fill=tk.X, after=header)

        self.status_label = tk.Label(
            self.right_frame,
            text=f"未答 {len(self.questions)}，已答 0，标记 0",
            font=("微软雅黑", 9), bg="white", fg="#666")
        self.status_label.pack(side=tk.BOTTOM, pady=10)

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

        used_type_order = self.pure_program_type_order if self.is_pure_program_exam else self.question_type_order

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
                    btn.config(bg=Theme.NAV_CURRENT, fg="white",
                               font=("微软雅黑", 9, "bold"))
                elif global_num in self.user_answers:
                    btn.config(bg=Theme.NAV_ANSWERED, fg="white",
                               font=("微软雅黑", 9))
                else:
                    btn.config(bg="white", fg=Theme.TEXT,
                               font=("微软雅黑", 9))

        self.status_label.config(
            text=f"未答 {len(self.questions) - answered_count}，已答 {answered_count}，标记 0")
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


# ★★★ 核心修改3：主入口增加全局异常捕获 ★★★
if __name__ == "__main__":
    try:
        # ★★★ 必须在创建Tk()之前设置AppUserModelID ★★★
        if sys.platform == "win32":
            import ctypes

            # 这个ID必须唯一，建议用"公司名.产品名.版本号"格式
            app_id = "HNUST.ExamSystem.V1.0"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        root = tk.Tk()

        # ★★★ 立即设置窗口图标（在任何其他操作之前）★★★
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(default=icon_path)  # 注意加了default参数
                root.iconbitmap(icon_path)
            except Exception as e:
                print(f"设置图标失败: {e}")

        app = HNUSTExamSystem(root)
        root.mainloop()
    except Exception as e:
        # 全局异常捕获（保持不变）
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