"""Python 环境查找与 IDLE 启动服务."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys


def find_system_python() -> str | None:
    """查找系统中可用的 Python 解释器路径."""
    if sys.platform != "win32":
        return None

    NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    candidates: list[str] = []

    # 优先 py -3
    try:
        py_path = shutil.which("py")
        if py_path:
            result = subprocess.run(
                [py_path, "-3", "-c", "import sys; print(sys.executable)"],
                capture_output=True, text=True, timeout=5,
                creationflags=NO_WINDOW,
            )
            if result.returncode == 0:
                exe = result.stdout.strip()
                if exe and os.path.isfile(exe) and exe not in candidates:
                    candidates.append(exe)
    except Exception:
        pass

    # PATH 中的 python/python3
    for name in ["python", "python3"]:
        try:
            p = shutil.which(name)
            if p and os.path.isfile(p) and p not in candidates:
                candidates.append(p)
        except Exception:
            pass

    # where python
    try:
        result = subprocess.run(
            ["where", "python"],
            capture_output=True, text=True, timeout=5,
            creationflags=NO_WINDOW,
        )
        if result.returncode == 0:
            for line in result.stdout.strip().splitlines():
                line = line.strip()
                if line and os.path.isfile(line) and line not in candidates:
                    candidates.append(line)
    except Exception:
        pass

    # 注册表查找
    try:
        import winreg
        for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            for key_path in [
                r"SOFTWARE\Python\PythonCore",
                r"SOFTWARE\WOW6432Node\Python\PythonCore",
            ]:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        for i in range(winreg.QueryInfoKey(key)[0]):
                            version = winreg.EnumKey(key, i)
                            try:
                                with winreg.OpenKey(
                                    key, f"{version}\\InstallPath"
                                ) as ik:
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

    # 常见安装路径
    local_app = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
    program_files_x86 = os.environ.get(
        "ProgramFiles(x86)", "C:\\Program Files (x86)"
    )
    home = os.path.expanduser("~")

    common_paths: list[str] = []
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

    # PATH 目录扫描
    for dir_path in os.environ.get("PATH", "").split(os.pathsep):
        dir_path = dir_path.strip()
        if not dir_path:
            continue
        exe = os.path.join(dir_path, "python.exe")
        if os.path.isfile(exe) and exe not in candidates:
            candidates.append(exe)

    return candidates[0] if candidates else None


def open_with_idle(file_path: str, python_exe: str | None = None) -> bool:
    """尝试用 IDLE 打开文件，返回是否成功."""
    abs_path = os.path.abspath(file_path)
    NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    # 用户指定的 Python
    if python_exe and os.path.isfile(python_exe):
        try:
            subprocess.Popen(
                [python_exe, "-m", "idlelib", abs_path],
                creationflags=NO_WINDOW,
            )
            return True
        except Exception:
            pass

    # 当前 Python
    if not hasattr(sys, "_MEIPASS"):
        try:
            subprocess.Popen(
                [sys.executable, "-m", "idlelib", abs_path],
                creationflags=NO_WINDOW,
            )
            return True
        except Exception:
            pass

    # py -3
    try:
        if shutil.which("py"):
            check = subprocess.run(
                ["py", "-3", "--version"],
                capture_output=True, text=True, timeout=5,
                creationflags=NO_WINDOW,
            )
            if check.returncode == 0 and "Python" in (check.stdout + check.stderr):
                subprocess.Popen(
                    ["py", "-3", "-m", "idlelib", abs_path],
                    creationflags=NO_WINDOW,
                )
                return True
    except Exception:
        pass

    # python / python3
    for cmd_name in ["python", "python3"]:
        python_path = shutil.which(cmd_name)
        if python_path:
            try:
                check = subprocess.run(
                    [python_path, "--version"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=NO_WINDOW,
                )
                if check.returncode == 0 and "Python" in (check.stdout + check.stderr):
                    subprocess.Popen(
                        [python_path, "-m", "idlelib", abs_path],
                        creationflags=NO_WINDOW,
                    )
                    return True
            except Exception:
                continue

    # 系统 Python
    python_exe = find_system_python()
    if python_exe:
        try:
            subprocess.Popen(
                [python_exe, "-m", "idlelib", abs_path],
                creationflags=NO_WINDOW,
            )
            return True
        except Exception:
            pass
        # idle.pyw 回退
        python_dir = os.path.dirname(python_exe)
        idle_pyw = os.path.join(python_dir, "Lib", "idlelib", "idle.pyw")
        if os.path.isfile(idle_pyw):
            try:
                subprocess.Popen(
                    [python_exe, idle_pyw, abs_path],
                    creationflags=NO_WINDOW,
                )
                return True
            except Exception:
                pass

    return False
