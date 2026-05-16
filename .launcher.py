import os, sys, shutil, subprocess, tempfile, signal, atexit
from tkinter import messagebox, ttk
from pathlib import Path
import tkinter as tk


def install_frozen_executable_icon(root, retry_ms: int = 250) -> None:
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return

    _set_frozen_executable_icon(root)
    try:
        root.after(retry_ms, lambda: _set_frozen_executable_icon(root))
    except Exception:
        pass


def _set_frozen_executable_icon(root) -> None:
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return

    try:
        import ctypes
        from ctypes import wintypes

        wm_seticon = 0x0080
        icon_small = 0
        icon_big = 1
        gclp_hicon = -14
        gclp_hiconsm = -34

        shell32 = ctypes.WinDLL("shell32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        long_ptr = (
            ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long
        )

        shell32.ExtractIconExW.argtypes = (
            wintypes.LPCWSTR,
            ctypes.c_int,
            ctypes.POINTER(wintypes.HICON),
            ctypes.POINTER(wintypes.HICON),
            wintypes.UINT,
        )
        shell32.ExtractIconExW.restype = wintypes.UINT
        user32.SendMessageW.argtypes = (
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        )
        user32.SendMessageW.restype = long_ptr
        user32.GetParent.argtypes = (wintypes.HWND,)
        user32.GetParent.restype = wintypes.HWND

        try:
            set_class_icon = user32.SetClassLongPtrW
        except AttributeError:
            set_class_icon = user32.SetClassLongW
        set_class_icon.argtypes = (wintypes.HWND, ctypes.c_int, long_ptr)
        set_class_icon.restype = long_ptr

        icon_handles = getattr(root, "_frozen_executable_icon_handles", None)
        if icon_handles is None:
            large_icons = (wintypes.HICON * 1)()
            small_icons = (wintypes.HICON * 1)()
            icon_count = shell32.ExtractIconExW(
                sys.executable,
                0,
                large_icons,
                small_icons,
                1,
            )
            if icon_count <= 0:
                return
            icon_handles = (small_icons[0], large_icons[0])
            root._frozen_executable_icon_handles = icon_handles
        small_icon, large_icon = icon_handles

        root.update_idletasks()
        hwnds = []
        for raw_hwnd in (root.winfo_id(), root.wm_frame()):
            try:
                hwnd = int(str(raw_hwnd), 0)
            except (TypeError, ValueError):
                continue
            while hwnd and hwnd not in hwnds:
                hwnds.append(hwnd)
                hwnd = user32.GetParent(hwnd)

        for hwnd in hwnds:
            if small_icon:
                user32.SendMessageW(hwnd, wm_seticon, icon_small, small_icon)
                set_class_icon(hwnd, gclp_hiconsm, small_icon)
            if large_icon:
                user32.SendMessageW(hwnd, wm_seticon, icon_big, large_icon)
                set_class_icon(hwnd, gclp_hicon, large_icon)
    except Exception:
        pass


def configure_launch_styles(root):
    colors = {
        "window": "#F4F7FB",
        "panel": "#FFFFFF",
        "border": "#D6DEE8",
        "text": "#15202B",
        "muted": "#617084",
        "accent": "#0F766E",
        "accent_dark": "#0B5D56",
        "header": "#103C3A",
        "header_badge": "#0F766E",
    }

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    try:
        style.configure("LaunchRoot.TFrame", background=colors["window"])
        style.configure("LaunchPanel.TFrame", background=colors["panel"])
        style.configure(
            "LaunchSection.TLabel",
            background=colors["panel"],
            foreground=colors["text"],
            font=("Segoe UI", 12, "bold"),
        )
        style.configure(
            "LaunchField.TLabel",
            background=colors["panel"],
            foreground=colors["text"],
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "LaunchMuted.TLabel",
            background=colors["panel"],
            foreground=colors["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "LaunchStatus.TLabel",
            background=colors["window"],
            foreground=colors["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "Launch.Vertical.TScrollbar",
            gripcount=0,
            background="#E8EEF5",
            darkcolor="#E8EEF5",
            lightcolor="#E8EEF5",
            troughcolor=colors["panel"],
            bordercolor=colors["panel"],
            arrowcolor=colors["muted"],
            relief="flat",
            width=12,
        )
        style.configure(
            "LaunchPrimary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(18, 10),
            foreground="#FFFFFF",
            background="#053D39",
            bordercolor=colors["accent_dark"],
            lightcolor=colors["accent_dark"],
            darkcolor=colors["accent_dark"],
            focuscolor=colors["accent_dark"],
        )
        style.map(
            "LaunchPrimary.TButton",
            background=[
                ("pressed", colors["accent_dark"]),
                ("active", colors["accent_dark"]),
                ("disabled", "#B8C2CC"),
            ],
            foreground=[("disabled", "#EEF2F7")],
            bordercolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
            lightcolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
            darkcolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
        )
        style.configure(
            "LaunchSecondary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            foreground=colors["text"],
            background="#E8EEF5",
            bordercolor="#D7E0EA",
            lightcolor="#D7E0EA",
            darkcolor="#D7E0EA",
            focuscolor="#E8EEF5",
        )
        style.map(
            "LaunchSecondary.TButton",
            background=[("pressed", "#D7E0EA"), ("active", "#EDF3F8")],
            foreground=[("pressed", colors["text"]), ("active", colors["text"])],
            bordercolor=[("pressed", "#C8D3DF"), ("active", "#C8D3DF")],
            lightcolor=[("pressed", "#C8D3DF"), ("active", "#C8D3DF")],
            darkcolor=[("pressed", "#C8D3DF"), ("active", "#C8D3DF")],
        )
    except Exception:
        pass

    return colors

# ================= НАЛАШТУВАННЯ =================

def _resolve_launcher_path() -> Path:
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve()
    # У compiled-режимі намагаємось визначити шлях за фактичним іменем запуску.
    argv0 = Path(sys.argv[0])
    if argv0.is_absolute() and argv0.exists():
        return argv0.resolve()
    exe = Path(sys.executable)
    if exe.exists():
        return exe.resolve()
    return (Path.cwd() / argv0).resolve()

LAUNCHER_PATH = _resolve_launcher_path()
ROOT_DIR = LAUNCHER_PATH.parent

# Папка зі скриптами (підпапка Codex відносно кореня)
SCRIPTS_DIR = (ROOT_DIR).resolve()

# Які розширення показувати у лаунчері
SCRIPT_EXTENSIONS = {".py", ".exe", ".bat", ".cmd", ".reg", ".ps1"}

# Імена файлів, які НЕ показуємо в списку
_self_stem = LAUNCHER_PATH.stem.lower()
EXCLUDE_FILES = {
    "__init__.py",
    "launcher.py",
    "launcher.exe",
    LAUNCHER_PATH.name.lower(),
    f"{_self_stem}.py",
    f"{_self_stem}.exe",
}

# Папки й файли з такими префіксами не скануємо і не показуємо.
EXCLUDE_NAME_PREFIXES = ("#", "_", ".")

# Розмір головного вікна лаунчера.
WINDOW_WIDTH = 470
WINDOW_HEIGHT = 320

# Файл для одиночного екземпляра лаунчера (у temp, щоб не кидалося в очі)
LAUNCHER_LOCKFILE = Path(tempfile.gettempdir()) / "codex_launcher.pid"

# =================================================

# Глобальний стан для одного активного запуску.
_current_proc = None


def _resolve_python_cmd_for_py_scripts() -> list[str] | None:
    # У звичайному режимі використовуємо той самий інтерпретатор.
    if not getattr(sys, "frozen", False):
        py_exe = Path(sys.executable)
        if py_exe.name.lower() == "pythonw.exe":
            py_exe = py_exe.with_name("python.exe")
        return [str(py_exe)]

    # У frozen-режимі launcher.exe не є інтерпретатором Python.
    # Пріоритет: python поруч -> версія, що відповідає версії збірки -> PATH -> py launcher.
    major = sys.version_info.major
    minor = sys.version_info.minor
    preferred_version = f"Python{major}{minor}"

    def _is_bad_windows_alias(p: Path) -> bool:
        # WindowsApps\python.exe часто є alias-заглушкою.
        return "windowsapps" in str(p).lower() and p.name.lower() == "python.exe"

    candidates: list[list[str]] = []

    local_python = (ROOT_DIR / "python.exe").resolve()
    if local_python.exists():
        candidates.append([str(local_python)])

    if os.name == "nt":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            base = Path(local_app_data) / "Programs" / "Python"
            preferred = (base / preferred_version / "python.exe").resolve()
            if preferred.exists():
                candidates.append([str(preferred)])
            for p in sorted(base.glob("Python*/python.exe"), reverse=True):
                rp = p.resolve()
                if rp == preferred:
                    continue
                candidates.append([str(rp)])

    path_python = shutil.which("python")
    if path_python:
        pp = Path(path_python)
        if pp.exists() and not _is_bad_windows_alias(pp):
            candidates.append([str(pp)])

    py_launcher = shutil.which("py")
    if py_launcher:
        candidates.append([str(Path(py_launcher)), f"-{major}.{minor}"])
        candidates.append([str(Path(py_launcher)), "-3"])

    for cmd in candidates:
        exe = Path(cmd[0])
        if exe.exists() and not _is_bad_windows_alias(exe):
            return cmd
    return None


def apply_dark_title_bar(root: tk.Tk):
    # For Windows 10/11: ask DWM to use dark caption/title bar.
    if os.name != "nt":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        value = ctypes.c_int(1)
        # 20 - modern builds, 19 - some older Windows 10 builds.
        for attr in (20, 19):
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                attr,
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
    except Exception:
        pass


def _normalize_path_entry(path_value: str) -> str:
    path_value = path_value.strip().strip('"')
    if not path_value:
        return ""
    expanded = os.path.expandvars(path_value)
    return os.path.normcase(os.path.abspath(expanded))


def _path_value_contains(path_value: str, target_dir: Path) -> bool:
    target = _normalize_path_entry(str(target_dir))
    if not target:
        return False
    for entry in path_value.split(os.pathsep):
        if _normalize_path_entry(entry) == target:
            return True
    return False


def _broadcast_environment_change():
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.user32.SendMessageTimeoutW(
            0xFFFF,  # HWND_BROADCAST
            0x001A,  # WM_SETTINGCHANGE
            0,
            "Environment",
            0x0002,  # SMTO_ABORTIFHUNG
            5000,
            None,
        )
    except Exception:
        pass


def ensure_launcher_dir_in_user_path():
    if os.name != "nt":
        return

    launcher_dir = ROOT_DIR.resolve()
    launcher_dir_text = str(launcher_dir)

    current_path = os.environ.get("PATH", "")
    if not _path_value_contains(current_path, launcher_dir):
        os.environ["PATH"] = (
            f"{current_path}{os.pathsep}{launcher_dir_text}"
            if current_path
            else launcher_dir_text
        )

    try:
        import winreg

        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            r"Environment",
            0,
            winreg.KEY_READ | winreg.KEY_SET_VALUE,
        ) as key:
            try:
                user_path, value_type = winreg.QueryValueEx(key, "Path")
                user_path = str(user_path)
            except FileNotFoundError:
                user_path = ""
                value_type = winreg.REG_EXPAND_SZ

            if _path_value_contains(user_path, launcher_dir):
                return

            new_user_path = (
                f"{user_path.rstrip(os.pathsep)}{os.pathsep}{launcher_dir_text}"
                if user_path
                else launcher_dir_text
            )
            if value_type not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ):
                value_type = winreg.REG_EXPAND_SZ
            winreg.SetValueEx(key, "Path", 0, value_type, new_user_path)
        _broadcast_environment_change()
    except Exception:
        pass


def _get_process_executable_path(pid: int) -> Path | None:
    if pid <= 0:
        return None

    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            process_query_limited_information = 0x1000
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
            kernel32.OpenProcess.restype = wintypes.HANDLE
            kernel32.QueryFullProcessImageNameW.argtypes = (
                wintypes.HANDLE,
                wintypes.DWORD,
                wintypes.LPWSTR,
                ctypes.POINTER(wintypes.DWORD),
            )
            kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
            kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
            kernel32.CloseHandle.restype = wintypes.BOOL

            handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
            if not handle:
                return None

            try:
                size = wintypes.DWORD(32768)
                buffer = ctypes.create_unicode_buffer(size.value)
                if not kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                    return None
                return Path(buffer.value)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return None

    proc_exe = Path(f"/proc/{pid}/exe")
    try:
        return proc_exe.resolve(strict=True)
    except Exception:
        return None


def _pid_matches_launcher(pid: int) -> bool:
    if not getattr(sys, "frozen", False):
        return False

    process_path = _get_process_executable_path(pid)
    if process_path is None:
        return False

    try:
        return process_path.resolve() == LAUNCHER_PATH.resolve()
    except Exception:
        return os.path.normcase(str(process_path)) == os.path.normcase(str(LAUNCHER_PATH))


def _terminate_pid(pid: int):
    try:
        if pid <= 0:
            return False
        os.kill(pid, signal.SIGTERM)
        return True
    except Exception:
        return False


def ensure_single_launcher_instance():
    # Якщо попередній лаунчер ще висить у процесах, прибираємо його.
    try:
        if LAUNCHER_LOCKFILE.exists():
            data = LAUNCHER_LOCKFILE.read_text(encoding="utf-8").strip()
            if data.isdigit():
                old_pid = int(data)
                if old_pid != os.getpid() and _pid_matches_launcher(old_pid):
                    _terminate_pid(old_pid)
    except Exception:
        pass

    try:
        LAUNCHER_LOCKFILE.write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        return

    def _cleanup():
        try:
            if LAUNCHER_LOCKFILE.exists():
                data = LAUNCHER_LOCKFILE.read_text(encoding="utf-8").strip()
                if data == str(os.getpid()):
                    LAUNCHER_LOCKFILE.unlink(missing_ok=True)
        except Exception:
            pass

    atexit.register(_cleanup)


def _is_visible_scripts_dir(dir_name: str) -> bool:
    dir_name = dir_name.strip()
    return bool(dir_name) and not dir_name.startswith(EXCLUDE_NAME_PREFIXES)


def _is_visible_script_file_name(file_name: str) -> bool:
    file_name = file_name.strip()
    return bool(file_name) and not file_name.startswith(EXCLUDE_NAME_PREFIXES)


def _is_launcher_file(path: Path) -> bool:
    if path.name.lower() in EXCLUDE_FILES:
        return True
    try:
        return path.resolve() == LAUNCHER_PATH.resolve()
    except Exception:
        return os.path.normcase(str(path)) == os.path.normcase(str(LAUNCHER_PATH))


def _is_short_name_in_launcher_dir(path: Path) -> bool:
    try:
        is_in_launcher_dir = path.parent.resolve() == ROOT_DIR.resolve()
    except Exception:
        is_in_launcher_dir = (
            os.path.normcase(str(path.parent)) == os.path.normcase(str(ROOT_DIR))
        )
    return is_in_launcher_dir and len(path.stem.strip()) < 3


def find_scripts():
    if not SCRIPTS_DIR.exists():
        messagebox.showerror(
            "Помилка",
            f"Папку зі скриптами не знайдено:\n{SCRIPTS_DIR}",
        )
        return []

    # Ключ: відносний шлях без розширення (case-insensitive), значення: вибраний файл.
    chosen_by_stem = {}
    for dirpath, dirnames, filenames in os.walk(SCRIPTS_DIR):
        current_dir = Path(dirpath)
        dirnames[:] = sorted(
            (d for d in dirnames if _is_visible_scripts_dir(d)),
            key=str.lower,
        )

        for filename in sorted(filenames, key=str.lower):
            if not _is_visible_script_file_name(filename):
                continue
            path = Path(dirpath) / filename
            if _is_launcher_file(path):
                continue
            if _is_short_name_in_launcher_dir(path):
                continue
            if path.suffix.lower() not in SCRIPT_EXTENSIONS:
                continue
            try:
                stem_key = path.relative_to(SCRIPTS_DIR).with_suffix("").as_posix().lower()
            except ValueError:
                stem_key = path.with_suffix("").as_posix().lower()
            current = chosen_by_stem.get(stem_key)
            if current is None:
                chosen_by_stem[stem_key] = path
                continue
            try:
                current_mtime = current.stat().st_mtime
            except Exception:
                current_mtime = float("inf")
            try:
                path_mtime = path.stat().st_mtime
            except Exception:
                path_mtime = float("inf")
            # З двох однакових назв залишаємо файл з новішою датою модифікації.
            if path_mtime > current_mtime:
                chosen_by_stem[stem_key] = path
    return sorted(
        chosen_by_stem.values(),
        key=lambda p: p.relative_to(SCRIPTS_DIR).as_posix().lower(),
    )

def get_script_display_name(path: Path) -> str:
    try:
        resolved = path.resolve()
    except Exception:
        resolved = path
    drive_prefix = (resolved.drive or resolved.anchor.rstrip("\\/")).replace("\\", "/")
    if not drive_prefix:
        drive_prefix = "."
    folder_name = resolved.parent.name or drive_prefix
    return f"{folder_name}/{resolved.name}"


def _build_launch_command(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix == ".py":
        py_cmd = _resolve_python_cmd_for_py_scripts()
        if not py_cmd:
            raise RuntimeError(
                "Не знайдено Python для запуску .py (перевір PATH або покладіть python.exe поруч з launcher)."
            )
        # Щоб скрипт не буферизував вивід, додаємо -u.
        return [*py_cmd, "-u", str(path)]
    if suffix == ".exe":
        return [str(path)]
    if suffix in {".bat", ".cmd"}:
        return ["cmd.exe", "/c", str(path)]
    if suffix == ".reg":
        return ["reg.exe", "import", str(path)]
    if suffix == ".ps1":
        return [
            "cmd.exe",
            "/c",
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(path),
        ]
    raise RuntimeError(f"Непідтримуване розширення: {path.suffix}")


def run_script(parent_root: tk.Tk, path: Path):
    global _current_proc
    try:
        def close_previous_run():
            global _current_proc
            if _current_proc is not None and _current_proc.poll() is None:
                try:
                    _current_proc.terminate()
                except Exception:
                    pass
            _current_proc = None

        close_previous_run()

        cmd = _build_launch_command(path)

        creationflags = 0
        if os.name == "nt":
            creationflags = subprocess.CREATE_NO_WINDOW  # щоб не мигало консольне вікно

        proc = subprocess.Popen(
            cmd,
            cwd=str(path.parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
        _current_proc = proc

    except Exception as e:
        try:
            parent_root.deiconify()
            parent_root.lift()
            parent_root.attributes("-topmost", True)
            parent_root.update_idletasks()
        except Exception:
            pass
        messagebox.showerror(
            "Помилка запуску",
            f"Не вдалося запустити скрипт:\n{path}\n\n{e}",
            parent=parent_root,
        )

def main():
    ensure_single_launcher_instance()
    scripts = find_scripts()
    if not scripts:
        return

    root = tk.Tk()
    root.withdraw()
    install_frozen_executable_icon(root)
    root.title("CODEX LAUNCHER")
    root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
    root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
    root.resizable(False, False)
    root.attributes("-topmost", True)
    colors = configure_launch_styles(root)
    root.configure(bg=colors["window"])
    root.update_idletasks()
    apply_dark_title_bar(root)
    root.after(80, lambda: apply_dark_title_bar(root))

    surface = ttk.Frame(root, style="LaunchRoot.TFrame")
    surface.pack(fill="both", expand=True)

    header = tk.Frame(surface, bg=colors["header"], padx=22, pady=7)
    header.pack(fill="x")
    header.grid_columnconfigure(0, weight=1)
    title = tk.Label(
        header,
        text="CODEX PyScripts",
        bg=colors["header"],
        fg="#FFFFFF",
        font=("Segoe UI Semibold", 16),
        anchor="w",
    )
    title.grid(row=0, column=0, sticky="w")
    subtitle = tk.Label(
        header,
        text="Швидкий запуск локальних виконавчих файлів...",
        bg=colors["header"],
        fg="#D7FBF5",
        font=("Segoe UI", 9),
        anchor="w",
    )
    subtitle.grid(row=1, column=0, sticky="w", pady=(0, 4))

    body = ttk.Frame(surface, style="LaunchRoot.TFrame", padding=(18, 16, 18, 18))
    body.pack(fill="both", expand=True)

    list_row = tk.Frame(body, bg=colors["border"], padx=1, pady=1)
    list_row.pack(fill="both", expand=True)

    list_wrap = tk.Frame(list_row, bg=colors["panel"], padx=10, pady=8)
    list_wrap.pack(fill="both", expand=True)

    listbox = tk.Listbox(
        list_wrap,
        activestyle="none",
        font=("Ink Free", 11, "bold"),
        height=7,
        highlightthickness=0,
        bd=0,
        bg=colors["panel"],
        fg=colors["text"],
        selectbackground="#D7FBF5",
        selectforeground=colors["text"],
        selectborderwidth=0,
        relief="flat",
        exportselection=False,
    )
    scrollbar = ttk.Scrollbar(
        list_wrap,
        orient="vertical",
        command=listbox.yview,
        style="Launch.Vertical.TScrollbar",
    )
    listbox.configure(yscrollcommand=scrollbar.set)

    listbox.pack(side="left", fill="both", expand=True, padx=(0, 8))
    scrollbar.pack(side="right", fill="y")

    idx_to_path = {}

    for idx, script in enumerate(scripts):
        listbox.insert(tk.END, get_script_display_name(script))
        idx_to_path[idx] = script

    def on_run():
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning(
                "Не вибрано",
                "Будь ласка, виберіть скрипт зі списку.",
                parent=root,
            )
            return
        idx = selection[0]
        script_path = idx_to_path[idx]
        root.iconify()
        run_script(root, script_path)

    def on_double_click(event):
        on_run()

    def on_enter(event=None):
        on_run()
        return "break"

    def open_work_folder():
        try:
            if os.name == "nt":
                os.startfile(str(SCRIPTS_DIR))
            else:
                subprocess.Popen(["xdg-open", str(SCRIPTS_DIR)])
        except Exception as e:
            messagebox.showerror(
                "Помилка",
                f"Не вдалося відкрити папку:\n{SCRIPTS_DIR}\n\n{e}",
                parent=root,
            )

    def on_title_double_click(event=None):
        ensure_launcher_dir_in_user_path()
        open_work_folder()
        return "break"

    def move_selection(delta: int):
        size = listbox.size()
        if size == 0:
            return
        current = listbox.curselection()
        idx = current[0] if current else 0
        new_idx = max(0, min(size - 1, idx + delta))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(new_idx)
        listbox.activate(new_idx)
        listbox.see(new_idx)

    def on_scrollbar_click(event):
        element = scrollbar.identify(event.x, event.y)
        if element in ("arrow1", "uparrow"):
            move_selection(-1)
            return
        if element in ("arrow2", "downarrow"):
            move_selection(1)
            return
        if element in ("trough1", "trough2", "slider", "trough", "thumb"):
            root.after(1, on_run)
            return

    root.bind("<Up>", lambda event: (move_selection(-1), "break")[1])
    root.bind("<Down>", lambda event: (move_selection(1), "break")[1])
    root.bind("<Return>", on_enter)
    root.bind("<KP_Enter>", on_enter)
    title.bind("<Double-Button-1>", on_title_double_click)
    title.configure(cursor="hand2")
    listbox.bind("<Double-Button-1>", on_double_click)
    scrollbar.bind("<Button-1>", on_scrollbar_click)
    listbox.bind("<Return>", on_enter)
    listbox.bind("<KP_Enter>", on_enter)
    listbox.selection_set(0)
    listbox.focus_set()
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.focus_force()

    root.mainloop()


if __name__ == "__main__":
    main()
