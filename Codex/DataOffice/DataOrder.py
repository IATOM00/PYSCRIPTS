from tkinter import Tk, ttk, filedialog, messagebox, StringVar
import os, sys, time, ctypes, logging, tempfile, webbrowser, stat
import shutil, queue, subprocess, threading
from tkinter import font as tkfont
from datetime import datetime
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


def install_dark_title_bar(window, retry_ms: int = 80) -> None:
    if sys.platform != "win32":
        return

    def _apply() -> None:
        try:
            import ctypes

            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            value = ctypes.c_int(1)
            for attr in (20, 19):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    attr,
                    ctypes.byref(value),
                    ctypes.sizeof(value),
                )
        except Exception:
            pass

    try:
        window.update_idletasks()
        _apply()
        window.after(retry_ms, _apply)
    except Exception:
        pass
from PIL import Image, ImageOps
import img2pdf
from send2trash import send2trash
from pypdf import PdfReader, PdfWriter

try:
    import win32com.client as win32
except ImportError:
    win32 = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

# ================== НАЛАШТУВАННЯ ==================
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
PDF_EXT = ".pdf"
OCR_LANG = "ukr+eng"
TESSERACT_DOWNLOAD_URL = "https://github.com/UB-Mannheim/tesseract/wiki"
GHOSTSCRIPT_DOWNLOAD_URL = "https://ghostscript.com/releases/gsdnld.html"
WORD_COM_PROG_ID = "Word.Application"
EXCEL_COM_PROG_ID = "Excel.Application"
WORD_FORMAT_DOCX = 16
EXCEL_FORMAT_XLSX = 51
WORD_ALERTS_NONE = 0
WORD_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Максимальні розміри зображень перед PDF
MAX_W = 900
MAX_H = 1600
JPEG_QUALITY = 85


class UserCancelled(Exception):
    pass


# ================== ДОПОМІЖНЕ ==================
def is_image_file(p: Path) -> bool:
    return p.suffix.lower() in IMAGE_EXTS

def is_pdf_file(p: Path) -> bool:
    return p.suffix.lower() == PDF_EXT

def iter_files_recursive(root: Path):
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            yield Path(dirpath) / fn


def _program_files_dirs() -> list[Path]:
    dirs = []
    for env_name in ("ProgramFiles", "ProgramW6432", "ProgramFiles(x86)", "LOCALAPPDATA"):
        raw = os.environ.get(env_name)
        if raw:
            p = Path(raw)
            if p.exists() and p not in dirs:
                dirs.append(p)
    return dirs


def _prepend_to_path(folder: Path):
    if not folder.exists():
        return

    current_paths = [
        Path(part)
        for part in os.environ.get("PATH", "").split(os.pathsep)
        if part
    ]
    if any(str(p).lower() == str(folder).lower() for p in current_paths):
        return

    os.environ["PATH"] = str(folder) + os.pathsep + os.environ.get("PATH", "")


def find_tesseract_exe() -> str | None:
    for name in ("tesseract", "tesseract.exe"):
        exe = shutil.which(name)
        if exe:
            return exe

    candidates = []
    for base in _program_files_dirs():
        candidates.append(base / "Tesseract-OCR" / "tesseract.exe")
        candidates.append(base / "Programs" / "Tesseract-OCR" / "tesseract.exe")

    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def _ghostscript_version_key(path: Path) -> tuple[int, ...]:
    version_dir = path.parents[1].name.lower().removeprefix("gs")
    parts = []
    for part in version_dir.split("."):
        try:
            parts.append(int(part))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def find_ghostscript_exe() -> str | None:
    for name in ("gswin64c", "gswin32c", "gs"):
        exe = shutil.which(name)
        if exe:
            return exe

    candidates = []
    for base in _program_files_dirs():
        gs_root = base / "gs"
        if not gs_root.exists():
            continue
        for exe_name in ("gswin64c.exe", "gswin32c.exe", "gs.exe"):
            candidates.extend(gs_root.glob(f"gs*/bin/{exe_name}"))

    candidates = [p for p in candidates if p.is_file()]
    candidates.sort(key=_ghostscript_version_key, reverse=True)
    if candidates:
        return str(candidates[0])
    return None


def configure_external_tool_paths():
    for exe in (find_tesseract_exe(), find_ghostscript_exe()):
        if exe:
            _prepend_to_path(Path(exe).parent)

    tesseract_exe = find_tesseract_exe()
    if tesseract_exe and "TESSDATA_PREFIX" not in os.environ:
        tessdata_dir = Path(tesseract_exe).parent / "tessdata"
        if tessdata_dir.exists():
            os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)


def _required_tesseract_languages() -> list[str]:
    return [part.strip() for part in OCR_LANG.split("+") if part.strip()]


def _tesseract_languages(tesseract_exe: str) -> set[str] | None:
    try:
        result = subprocess.run(
            [tesseract_exe, "--list-langs"],
            check=False,
            capture_output=True,
            text=True,
            timeout=10,
            **_hidden_subprocess_kwargs(quiet=False),
        )
    except Exception:
        return None

    output = (result.stdout or "") + "\n" + (result.stderr or "")
    langs = set()
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line or line.lower().startswith("list of"):
            continue
        if line.replace("_", "").replace("-", "").isalnum():
            langs.add(line)
    return langs


def _missing_external_dependencies() -> tuple[list[str], set[str]]:
    configure_external_tool_paths()
    missing = []
    download_keys = set()

    tesseract_exe = find_tesseract_exe()
    if not tesseract_exe:
        missing.append("Tesseract OCR (tesseract.exe)")
        download_keys.add("tesseract")
    else:
        langs = _tesseract_languages(tesseract_exe)
        if langs is not None:
            missing_langs = [
                lang for lang in _required_tesseract_languages() if lang not in langs
            ]
            if missing_langs:
                missing.append(
                    "мовні дані Tesseract: " + ", ".join(missing_langs)
                )
                download_keys.add("tesseract")

    if not find_ghostscript_exe():
        missing.append("Ghostscript (gswin64c.exe / gswin32c.exe)")
        download_keys.add("ghostscript")

    return missing, download_keys


def _open_dependency_download_pages(download_keys: set[str]):
    urls = []
    if "tesseract" in download_keys:
        urls.append(TESSERACT_DOWNLOAD_URL)
    if "ghostscript" in download_keys:
        urls.append(GHOSTSCRIPT_DOWNLOAD_URL)

    for url in urls:
        try:
            webbrowser.open_new_tab(url)
        except Exception:
            pass


def ensure_external_dependencies_available(owner: Tk) -> bool:
    missing, download_keys = _missing_external_dependencies()
    if not missing:
        return True

    message = (
        "Для етапу \"PDF -> PDF with OCR\" потрібні зовнішні компоненти:\n\n"
        + "\n".join(f"- {item}" for item in missing)
        + "\n\nНатисніть \"Так\", щоб відкрити сторінки завантаження."
        + "\nПісля встановлення перезапустіть програму."
    )
    should_open = messagebox.askyesno(
        "Потрібні компоненти OCR",
        message,
        parent=owner,
    )
    if should_open:
        _open_dependency_download_pages(download_keys)
    return False


def ensure_word_com_available(owner: Tk) -> bool:
    if sys.platform != "win32":
        messagebox.showerror(
            "Потрібен Microsoft Word",
            "Конвертація PDF у DOCX через Microsoft Word COM доступна тільки у Windows.",
            parent=owner,
        )
        return False

    if win32 is None:
        messagebox.showerror(
            "Потрібен Microsoft Word",
            "Для конвертації PDF у DOCX потрібні встановлений Microsoft Word та Python-пакет pywin32.",
            parent=owner,
        )
        return False

    word_app = None
    com_initialized = False
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
            com_initialized = True
        word_app = win32.DispatchEx(WORD_COM_PROG_ID)
        word_app.Visible = False
        word_app.DisplayAlerts = WORD_ALERTS_NONE
        try:
            word_app.AutomationSecurity = WORD_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        return True
    except Exception as exc:
        details = str(exc).strip()
        messagebox.showerror(
            "Потрібен Microsoft Word",
            "Не вдалося запустити Microsoft Word через COM.\n\n"
            "Перевірте, що Word встановлений і активований."
            + (f"\n\nДеталі: {details}" if details else ""),
            parent=owner,
        )
        return False
    finally:
        if word_app is not None:
            try:
                word_app.Quit()
            except Exception:
                pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
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
        style.configure("LaunchSection.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 12, "bold"))
        style.configure("LaunchField.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10, "bold"))
        style.configure("LaunchMuted.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("LaunchStatus.TLabel", background=colors["window"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("ProgressHeader.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 12, "bold"))
        style.configure("ProgressBody.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
        style.configure("ProgressFile.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("ProgressCount.TLabel", background=colors["panel"], foreground=colors["accent_dark"], font=("Segoe UI", 10, "bold"))
        style.configure(
            "Launch.TEntry",
            fieldbackground="#FFFFFF",
            foreground=colors["text"],
            bordercolor=colors["border"],
            lightcolor=colors["border"],
            darkcolor=colors["border"],
            padding=(8, 6),
        )
        style.map(
            "Launch.TEntry",
            fieldbackground=[("readonly", "#FFFFFF")],
            foreground=[("readonly", colors["text"])],
        )
        style.configure(
            "LaunchBrowse.TButton",
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
            "LaunchBrowse.TButton",
            background=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
            foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
        )
        style.configure(
            "LaunchPrimary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(20, 14),
            foreground="#FFFFFF",
            background="#053D39",
            bordercolor=colors["accent_dark"],
            lightcolor=colors["accent_dark"],
            darkcolor=colors["accent_dark"],
            focuscolor=colors["accent_dark"],
        )
        style.configure(
            "LaunchPrimaryHover.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(20, 14),
            foreground="#FFFFFF",
            background=colors["accent_dark"],
            bordercolor=colors["accent_dark"],
            lightcolor=colors["accent_dark"],
            darkcolor=colors["accent_dark"],
            focuscolor=colors["accent_dark"],
        )
        style.map(
            "LaunchPrimary.TButton",
            background=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
            bordercolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            lightcolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            darkcolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
        )
        style.map(
            "LaunchPrimaryHover.TButton",
            background=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
            bordercolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            lightcolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
            darkcolor=[("pressed", "#042F2C"), ("active", colors["accent_dark"])],
        )
        style.configure(
            "Launch.Horizontal.TProgressbar",
            background=colors["header_badge"],
            troughcolor="#E8EEF5",
            bordercolor=colors["border"],
            lightcolor=colors["header_badge"],
            darkcolor=colors["header_badge"],
            thickness=12,
        )
    except Exception:
        pass

    return colors


def _widget_exists(widget) -> bool:
    try:
        return bool(widget.winfo_exists())
    except Exception:
        return False


def _prepare_dialog_parent(root) -> None:
    if not _widget_exists(root):
        return
    try:
        root.lift()
        root.attributes("-topmost", True)
        root.update()
    except Exception:
        pass


def format_path_for_entry_display(path: str) -> str:
    path = path.strip()
    if not path:
        return ""

    normalized = Path(path)
    drive_prefix = ""
    if normalized.drive and len(normalized.drive) == 2 and normalized.drive[1] == ":":
        drive_prefix = f"{normalized.drive}/"

    name = normalized.name
    parent_name = normalized.parent.name
    if parent_name and name:
        return f"{drive_prefix}.../{parent_name}/{name}"
    if name:
        return f"{drive_prefix}{name}" if drive_prefix else f".../{name}"
    if drive_prefix:
        return drive_prefix
    return path.replace("\\", "/")


def fit_path_for_entry_display(path: str, entry) -> str:
    display_path = format_path_for_entry_display(path)
    if not display_path:
        return ""

    try:
        entry.update_idletasks()
        font_name = entry.cget("font") or "TkDefaultFont"
        try:
            font = tkfont.nametofont(font_name)
        except Exception:
            font = tkfont.nametofont("TkDefaultFont")

        available_width = max(120, entry.winfo_width() - 14)
        if font.measure(display_path) <= available_width:
            return display_path

        parts = display_path.split("/")
        if len(parts) >= 3:
            prefix = "/".join(parts[:-1]) + "/..."
            shortened_name = parts[-1]
            while len(shortened_name) > 1 and font.measure(prefix + shortened_name) > available_width:
                shortened_name = shortened_name[1:]
            return prefix + shortened_name

        ellipsis = "..."
        shortened_path = display_path
        while len(shortened_path) > 1 and font.measure(ellipsis + shortened_path) > available_width:
            shortened_path = shortened_path[1:]
        return ellipsis + shortened_path
    except Exception:
        return display_path


def refresh_path_entry_display(source_var, display_var, entry) -> None:
    if not _widget_exists(entry):
        return
    display_var.set(fit_path_for_entry_display(source_var.get(), entry))


def bind_primary_button_hover(button) -> None:
    def on_enter(_event) -> None:
        button.configure(style="LaunchPrimaryHover.TButton")
        try:
            button.focus_force()
        except Exception:
            pass

    def on_leave(_event) -> None:
        button.configure(style="LaunchPrimary.TButton")
        try:
            button.winfo_toplevel().focus_force()
        except Exception:
            pass

    button.bind("<Enter>", on_enter, add="+")
    button.bind("<Leave>", on_leave, add="+")


class ToggleSwitch(ttk.Frame):
    def __init__(self, owner, variable):
        super().__init__(owner)
        self.variable = variable
        self.state_var = tk.StringVar()

        style = ttk.Style(self)
        try:
            parent_style = owner.cget("style")
            if parent_style:
                self.configure(style=parent_style)
                canvas_bg = style.lookup(parent_style, "background")
            else:
                canvas_bg = ""
        except Exception:
            canvas_bg = ""
        canvas_bg = canvas_bg or style.lookup("LaunchPanel.TFrame", "background") or self.winfo_toplevel().cget("bg")
        state_label_style = "ToggleState.TLabel"
        try:
            style.configure(
                state_label_style,
                background="#ECECFA",
                foreground="#15202B",
                font=("Segoe UI", 9),
            )
        except Exception:
            state_label_style = ""

        self.canvas = tk.Canvas(
            self,
            width=52,
            height=28,
            bd=0,
            highlightthickness=0,
            bg=canvas_bg,
            cursor="hand2",
        )
        self.canvas.grid(row=0, column=0, padx=(0, 8), sticky="w")

        self.grid_columnconfigure(1, minsize=78)
        self.state_label = ttk.Label(
            self,
            textvariable=self.state_var,
            style=state_label_style,
            width=10,
            anchor="w",
        )
        self.state_label.grid(row=0, column=1, sticky="w")

        for widget in (self, self.canvas, self.state_label):
            widget.bind("<Button-1>", self.toggle)

        self.variable.trace_add("write", self._redraw)
        self._redraw()

    def toggle(self, _event=None) -> None:
        self.variable.set(not self.variable.get())

    def _redraw(self, *_args) -> None:
        is_enabled = bool(self.variable.get())
        track_color = "#0F766E" if is_enabled else "#CBD5E1"
        knob_color = "#FFFFFF"
        self.state_var.set("Увімкнено" if is_enabled else "Вимкнено")

        self.canvas.delete("all")

        left = 2
        top = 2
        right = 50
        bottom = 26
        radius = (bottom - top) // 2

        self.canvas.create_oval(left, top, left + 2 * radius, bottom, fill=track_color, outline=track_color)
        self.canvas.create_oval(right - 2 * radius, top, right, bottom, fill=track_color, outline=track_color)
        self.canvas.create_rectangle(left + radius, top, right - radius, bottom, fill=track_color, outline=track_color)

        knob_diameter = (2 * radius) - 2
        knob_left = right - knob_diameter - 1 if is_enabled else left + 1
        self.canvas.create_oval(
            knob_left,
            top + 1,
            knob_left + knob_diameter,
            bottom - 1,
            fill=knob_color,
            outline=knob_color,
        )


def choose_folder(root, source_var, display_var, entry, title: str) -> None:
    _prepare_dialog_parent(root)
    path = filedialog.askdirectory(title=title, parent=root)
    if path:
        source_var.set(path)
        refresh_path_entry_display(source_var, display_var, entry)
    _prepare_dialog_parent(root)


def ask_folder_settings_window(
    parent,
    *,
    app_name: str,
    subtitle: str,
    hint: str,
    browse_title: str,
) -> tuple[Path, bool, bool, bool, bool, bool, bool] | None:
    dialog = parent
    created_dialog = False
    if dialog is None:
        dialog = tk.Tk()
        created_dialog = True

    selected = {
        "path": None,
        "cleanup_temp_locks": False,
        "convert_office_formats": False,
        "convert_images_to_pdf": False,
        "perform_ocr": False,
        "process_pdf": False,
        "convert_to_word": False,
    }
    dialog.withdraw()
    install_frozen_executable_icon(dialog)
    dialog.title(f"НАЛАШТУВАННЯ: FDataOCR")
    dialog.attributes("-topmost", True)
    dialog.resizable(False, False)

    colors = configure_launch_styles(dialog)
    dialog.configure(bg=colors["window"])
    install_dark_title_bar(dialog)

    path_var = tk.StringVar(value="")
    path_display_var = tk.StringVar(value="")
    cleanup_var = tk.BooleanVar(master=dialog, value=False)
    office_format_var = tk.BooleanVar(master=dialog, value=False)
    image_pdf_var = tk.BooleanVar(master=dialog, value=False)
    ocr_var = tk.BooleanVar(master=dialog, value=False)
    process_pdf_var = tk.BooleanVar(master=dialog, value=False)
    word_var = tk.BooleanVar(master=dialog, value=False)
    done_var = tk.BooleanVar(master=dialog, value=False)
    start_button = None

    def update_launch_state(*_args) -> None:
        pass

    def on_start() -> None:
        folder_text = path_var.get().strip()
        if not folder_text:
            _prepare_dialog_parent(dialog)
            messagebox.showerror("Помилка", "Потрібно вказати папку для обробки.", parent=dialog)
            return

        folder = Path(folder_text).expanduser().resolve()
        if not folder.exists() or not folder.is_dir():
            _prepare_dialog_parent(dialog)
            messagebox.showerror("Помилка", f"Це не папка:\n{folder}", parent=dialog)
            return

        if not (
            cleanup_var.get()
            or office_format_var.get()
            or image_pdf_var.get()
            or ocr_var.get()
            or process_pdf_var.get()
            or word_var.get()
        ):
            _prepare_dialog_parent(dialog)
            messagebox.showerror("Помилка", "Увімкніть хоча б одну операцію.", parent=dialog)
            return

        selected["path"] = folder
        selected["cleanup_temp_locks"] = bool(cleanup_var.get())
        selected["convert_office_formats"] = bool(office_format_var.get())
        selected["convert_images_to_pdf"] = bool(image_pdf_var.get())
        selected["perform_ocr"] = bool(ocr_var.get())
        selected["process_pdf"] = bool(process_pdf_var.get())
        selected["convert_to_word"] = bool(word_var.get())
        dialog.withdraw()
        done_var.set(True)

    def on_cancel() -> None:
        selected["path"] = None
        selected["cleanup_temp_locks"] = False
        selected["convert_office_formats"] = False
        selected["convert_images_to_pdf"] = False
        selected["perform_ocr"] = False
        selected["process_pdf"] = False
        selected["convert_to_word"] = False
        dialog.withdraw()
        done_var.set(True)

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    dialog.bind("<Escape>", lambda _event: on_cancel())
    dialog.bind("<Return>", lambda _event: on_start())

    surface = ttk.Frame(dialog, style="LaunchRoot.TFrame")
    surface.grid(row=0, column=0, sticky="nsew")
    surface.grid_columnconfigure(0, weight=1)

    header = tk.Frame(surface, bg=colors["header"], padx=22, pady=12)
    header.grid(row=0, column=0, sticky="we")
    header.grid_columnconfigure(0, weight=1)
    tk.Label(
        header,
        text=app_name,
        bg=colors["header"],
        fg="#FFFFFF",
        font=("Segoe UI Semibold", 16),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        header,
        text=subtitle,
        bg=colors["header"],
        fg="#D7FBF5",
        font=("Segoe UI", 9),
        anchor="w",
    ).grid(row=1, column=0, sticky="w", pady=(2, 6))
    badge_shell = tk.Frame(header, bg=colors["accent_dark"], padx=1, pady=1)
    badge = tk.Label(
        badge_shell,
        text="НАЛАШТУВАННЯ",
        bg="#053D39",
        fg="#FFFFFF",
        font=("Segoe UI", 12, "bold"),
        padx=10,
        pady=5,
    )
    badge.pack()
    badge_shell.place(relx=1.0, x=5, y=5, anchor="ne")

    body = ttk.Frame(surface, style="LaunchRoot.TFrame", padding=(18, 16, 18, 18))
    body.grid(row=1, column=0, sticky="nsew")
    body.grid_columnconfigure(0, weight=1)

    panel_shell = tk.Frame(body, bg=colors["border"], padx=1, pady=1)
    panel_shell.grid(row=0, column=0, sticky="we")
    settings = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 12))
    settings.pack(fill="both", expand=True)
    settings.grid_columnconfigure(1, weight=1)

    ttk.Label(settings, text="Папка для обробки", style="LaunchSection.TLabel").grid(
        row=0, column=0, columnspan=3, sticky="w"
    )
    ttk.Label(settings, text=hint, style="LaunchMuted.TLabel", wraplength=520).grid(
        row=1, column=0, columnspan=3, pady=(4, 10), sticky="w"
    )
    ttk.Label(settings, text="ПАПКА:", style="LaunchField.TLabel").grid(
        row=2, column=0, padx=(0, 8), pady=7, sticky="e"
    )
    path_entry = ttk.Entry(settings, textvariable=path_display_var, width=48, state="readonly", style="Launch.TEntry")
    path_entry.grid(row=2, column=1, padx=(0, 14), pady=7, sticky="we")
    ttk.Button(
        settings,
        text="ОБРАТИ",
        width=14,
        style="LaunchBrowse.TButton",
        command=lambda: choose_folder(dialog, path_var, path_display_var, path_entry, browse_title),
    ).grid(row=2, column=2, pady=7, sticky="nsew")

    ttk.Separator(settings, orient="horizontal").grid(
        row=3, column=0, columnspan=3, sticky="we", pady=(10, 12)
    )
    ttk.Label(settings, text="Етапи обробки", style="LaunchSection.TLabel").grid(
        row=4, column=0, columnspan=3, sticky="w"
    )
    toggle_grid = ttk.Frame(settings, style="LaunchPanel.TFrame")
    toggle_grid.grid(row=5, column=0, columnspan=3, pady=(8, 0), sticky="we")
    toggle_grid.grid_columnconfigure(0, weight=1, uniform="toggle_columns", minsize=250)
    toggle_grid.grid_columnconfigure(1, weight=1, uniform="toggle_columns", minsize=250)

    def add_toggle_cell(row: int, column: int, label: str, variable) -> None:
        cell = ttk.Frame(toggle_grid, style="LaunchPanel.TFrame")
        cell.grid(
            row=row,
            column=column,
            padx=(0, 14) if column == 0 else (14, 0),
            pady=6,
            sticky="we",
        )
        cell.grid_columnconfigure(0, minsize=150)
        cell.grid_columnconfigure(1, weight=1)
        ttk.Label(cell, text=label, style="LaunchField.TLabel").grid(
            row=0, column=0, padx=(0, 10), sticky="w"
        )
        ToggleSwitch(cell, variable).grid(row=0, column=1, sticky="w")

    add_toggle_cell(0, 0, "Очищення temp/lock", cleanup_var)
    add_toggle_cell(0, 1, "Новіший формат", office_format_var)
    add_toggle_cell(1, 0, "Зображення -> PDF", image_pdf_var)
    add_toggle_cell(1, 1, "PDF -> PDF with OCR", ocr_var)
    add_toggle_cell(2, 0, "Обробка PDF", process_pdf_var)
    add_toggle_cell(2, 1, "PDF -> Word", word_var)

    footer = ttk.Frame(body, style="LaunchRoot.TFrame")
    footer.grid(row=1, column=0, pady=(10, 0), sticky="we")
    footer.grid_columnconfigure(0, weight=1)
    start_button = ttk.Button(
        footer,
        text="ЗАПУСТИТИ",
        style="LaunchPrimary.TButton",
        command=on_start,
        padding=(18, 18),
    )
    start_button.grid(row=2, column=0, sticky="we")
    bind_primary_button_hover(start_button)

    path_var.trace_add("write", lambda *_args: (refresh_path_entry_display(path_var, path_display_var, path_entry), update_launch_state()))
    path_entry.bind("<Configure>", lambda _event: refresh_path_entry_display(path_var, path_display_var, path_entry))
    update_launch_state()

    dialog.update_idletasks()
    dialog.geometry(f"{dialog.winfo_reqwidth()}x{dialog.winfo_reqheight()}")
    dialog.deiconify()
    dialog.lift()
    dialog.grab_set()
    dialog.focus_force()
    dialog.wait_variable(done_var)

    try:
        dialog.grab_release()
    except Exception:
        pass

    if created_dialog and _widget_exists(dialog):
        dialog.destroy()

    if selected["path"] is None:
        return None
    return (
        selected["path"],
        bool(selected["cleanup_temp_locks"]),
        bool(selected["convert_office_formats"]),
        bool(selected["convert_images_to_pdf"]),
        bool(selected["perform_ocr"]),
        bool(selected["process_pdf"]),
        bool(selected["convert_to_word"]),
    )


# ================== ПРОГРЕС-ВІКНО ==================
class ProgressWindow:
    def __init__(self, owner: Tk):
        self._closed = False
        self.dialog = tk.Toplevel(owner)
        self.dialog.withdraw()
        install_frozen_executable_icon(self.dialog)
        self.dialog.title("FDataOCR - Обробка")
        self.dialog.resizable(False, False)
        self.dialog.attributes("-topmost", True)
        self.cancel_requested = False
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close_requested)
        colors = configure_launch_styles(self.dialog)
        self.dialog.configure(bg=colors["window"])
        install_dark_title_bar(self.dialog)

        container = ttk.Frame(self.dialog, style="LaunchRoot.TFrame")
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=colors["header"], padx=22, pady=10)
        header.pack(fill="x")
        header.grid_columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="FDataOCR",
            bg=colors["header"],
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Обробка файлів виконується. Закриття вікна зупинить процес.",
            bg=colors["header"],
            fg="#D7FBF5",
            font=("Segoe UI", 9),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(3, 4))
        badge_shell = tk.Frame(header, bg=colors["accent_dark"], padx=1, pady=1)
        badge = tk.Label(
            badge_shell,
            text="ОБРОБКА",
            bg="#053D39",
            fg="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
            padx=10,
            pady=5,
        )
        badge.pack()
        badge_shell.place(relx=1.0, x=5, y=5, anchor="ne")

        body = ttk.Frame(container, style="LaunchRoot.TFrame", padding=(18, 16, 18, 18))
        body.pack(fill="both", expand=True)
        panel_shell = tk.Frame(body, bg=colors["border"], padx=1, pady=1)
        panel_shell.pack(fill="both", expand=True)
        frame = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 12))
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)

        self.header_var = StringVar(value="Підготовка...")
        self.detail_var = StringVar(value="Будь ласка, зачекайте")
        self.file_var = StringVar(value="")
        self.count_var = StringVar(value="0 / 0")

        ttk.Label(frame, textvariable=self.header_var, style="ProgressHeader.TLabel").grid(
            row=0, column=0, padx=(0, 12), sticky="w"
        )
        ttk.Label(frame, textvariable=self.count_var, style="ProgressCount.TLabel").grid(
            row=0, column=1, sticky="e"
        )
        ttk.Label(frame, textvariable=self.detail_var, style="ProgressBody.TLabel", wraplength=500).grid(
            row=1, column=0, columnspan=2, pady=(8, 6), sticky="w"
        )
        ttk.Label(frame, textvariable=self.file_var, style="ProgressFile.TLabel", wraplength=500).grid(
            row=2, column=0, columnspan=2, pady=(0, 12), sticky="w"
        )
        self.progress = ttk.Progressbar(
            frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            style="Launch.Horizontal.TProgressbar",
        )
        self.progress.grid(row=3, column=0, columnspan=2, sticky="we")

        self.dialog.update_idletasks()
        width = max(560, self.dialog.winfo_reqwidth())
        height = max(248, self.dialog.winfo_reqheight())
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.deiconify()
        self.dialog.lift()
        self.refresh()

    def _is_alive(self) -> bool:
        if self._closed:
            return False
        try:
            return bool(self.dialog.winfo_exists())
        except Exception:
            return False

    def update(self, header=None, detail=None, current=None, total=None, file_name=None):
        if not self._is_alive():
            return
        try:
            if header is not None:
                self.header_var.set(header)
            if detail is not None:
                self.detail_var.set(detail)
            if file_name is not None:
                self.file_var.set(file_name)
            if current is not None and total is not None:
                safe_total = total if total > 0 else 1
                pct = max(0.0, min(100.0, (current / safe_total) * 100.0))
                self.progress["value"] = pct
                self.count_var.set(f"{current} / {total}")
            self.refresh()
        except tk.TclError:
            return

    def refresh(self):
        try:
            if not self._is_alive():
                return
            self.dialog.update_idletasks()
            self.dialog.update()
        except Exception:
            pass

    def _on_close_requested(self):
        self.cancel_requested = True
        self.close()

    def is_cancelled(self) -> bool:
        return self.cancel_requested

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            if self.dialog.winfo_exists():
                self.dialog.destroy()
        except Exception:
            pass


class ProgressUpdateProxy:
    def __init__(self, updates, should_cancel=None):
        self.updates = updates
        self.should_cancel = should_cancel

    def update(self, header=None, detail=None, current=None, total=None, file_name=None):
        self.updates.put(
            {
                "header": header,
                "detail": detail,
                "current": current,
                "total": total,
                "file_name": file_name,
            }
        )

    def refresh(self):
        return

    def is_cancelled(self) -> bool:
        if self.should_cancel is None:
            return False
        try:
            return bool(self.should_cancel())
        except Exception:
            return False


def apply_pending_progress_updates(progress_window: ProgressWindow, updates) -> None:
    while True:
        try:
            payload = updates.get_nowait()
        except queue.Empty:
            return
        progress_window.update(**payload)

# ================== ЛОГЕР ==================
def setup_logger(log_path: Path) -> logging.Logger:
    logger = logging.getLogger("processor")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(message)s")

    fh = logging.FileHandler(str(log_path), encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    return logger

def close_logger(logger: logging.Logger):
    for h in list(logger.handlers):
        try:
            h.flush()
            h.close()
        except Exception:
            pass
        try:
            logger.removeHandler(h)
        except Exception:
            pass


# ================== OFFICE CLEANUP / FORMAT ==================
def _log_to(logger: logging.Logger | None, message: str) -> None:
    if logger:
        logger.info(message)


def make_writable(path: Path) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
    except Exception:
        pass


class TrashStager:
    def __init__(self, source_root: Path, logger: logging.Logger | None = None):
        self.source_root = source_root
        self.logger = logger
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H.%M.%S")
        self.root: Path | None = None
        self.staged_count = 0
        self.trash_destination = ""

    def stage_file(self, path: Path) -> bool:
        if not path.exists() or not path.is_file():
            return False

        try:
            staging_root = self._ensure_root()
            try:
                rel = path.relative_to(self.source_root)
            except ValueError:
                rel = Path(path.name)

            destination = self._unique_path(staging_root / rel)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, destination)
            self.staged_count += 1
            _log_to(self.logger, f"[TRASH] STAGE {path} -> {destination}")
            return True
        except Exception as exc:
            _log_to(self.logger, f"[TRASH] ERROR staging {path}: {exc}")
            return False

    def finalize(self) -> None:
        if self.root is None or not self.root.exists():
            return
        if not any(self.root.iterdir()):
            self.root.rmdir()
            self.root = None
            return

        destination_name = self.root.name
        send2trash(str(self.root))
        self.trash_destination = destination_name
        _log_to(self.logger, f"[TRASH] STAGED FOLDER -> BIN {destination_name}")
        self.root = None

    def _ensure_root(self) -> Path:
        if self.root is not None:
            return self.root

        base = Path(tempfile.gettempdir()) / f"FDataOCR - deleted files - {self.timestamp}"
        self.root = self._unique_path(base)
        self.root.mkdir(parents=True, exist_ok=False)
        return self.root

    @staticmethod
    def _unique_path(path: Path) -> Path:
        if not path.exists():
            return path

        for idx in range(1, 1000):
            candidate = path.with_name(f"{path.name} ({idx})")
            if not candidate.exists():
                return candidate

        raise FileExistsError(f"Не вдалося підібрати унікальний шлях для backup: {path}")


_active_trash_stager: TrashStager | None = None


def remove_file_permanently(path: Path) -> None:
    try:
        path.unlink()
    except PermissionError:
        make_writable(path)
        path.unlink()


def trash_file(path: Path, logger: logging.Logger | None = None) -> bool:
    try:
        make_writable(path)
        if _active_trash_stager is not None:
            if not _active_trash_stager.stage_file(path):
                return False
            remove_file_permanently(path)
        else:
            send2trash(str(path))
        return True
    except Exception as exc:
        _log_to(logger, f"[TRASH] ERROR {path}: {exc}")
        return False


def is_office_temp_file(path: Path) -> bool:
    name = path.name
    lower = name.lower()
    return (
        name.startswith("~$")
        or name.startswith(".~lock.")
        or (lower.endswith(".tmp") and (name.startswith("~") or name.startswith("~$")))
    )


def cleanup_office_temp_files(
    folder: Path,
    logger: logging.Logger | None = None,
    should_cancel=None,
) -> int:
    deleted = 0
    for pattern in ("~$*", ".~lock.*", "*.tmp"):
        for path in folder.rglob(pattern):
            if should_cancel and should_cancel():
                raise UserCancelled()
            if not path.is_file():
                continue
            if pattern == "*.tmp" and not is_office_temp_file(path):
                continue
            if trash_file(path, logger=logger):
                deleted += 1
                _log_to(logger, f"[TMP] TRASH {path}")
    return deleted


def _is_legacy_doc(path: Path) -> bool:
    return path.suffix.lower() == ".doc" and not path.name.startswith("~$")


def _is_legacy_xls(path: Path) -> bool:
    return path.suffix.lower() == ".xls" and not path.name.startswith("~$")


def create_excel_app_for_conversion(logger: logging.Logger | None = None):
    if win32 is None:
        raise RuntimeError("pywin32 недоступний: не можна запустити Microsoft Excel через COM.")

    excel_app = win32.DispatchEx(EXCEL_COM_PROG_ID)
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    try:
        excel_app.AutomationSecurity = WORD_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    _log_to(logger, "[EXCEL] Microsoft Excel COM started")
    return excel_app


def convert_doc_to_docx(
    doc_path: Path,
    word_app,
    logger: logging.Logger | None = None,
    ui_pump=None,
    should_cancel=None,
) -> tuple[bool, bool, bool]:
    if not _is_legacy_doc(doc_path):
        return False, False, False

    docx_path = doc_path.with_suffix(".docx")
    converted = False
    doc = None
    try:
        if should_cancel and should_cancel():
            raise UserCancelled()

        if not docx_path.exists():
            _log_to(logger, f"[DOC] CONVERT {doc_path.name} -> {docx_path.name}")
            doc = word_app.Documents.Open(
                FileName=str(doc_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            try:
                doc.SaveAs2(
                    FileName=str(docx_path),
                    FileFormat=WORD_FORMAT_DOCX,
                    AddToRecentFiles=False,
                )
            except AttributeError:
                doc.SaveAs(
                    FileName=str(docx_path),
                    FileFormat=WORD_FORMAT_DOCX,
                    AddToRecentFiles=False,
                )
            doc.Close(False)
            doc = None
            if not docx_path.exists() or docx_path.stat().st_size <= 0:
                _log_to(logger, f"[DOC] ERROR {doc_path.name}: output was not created")
                return False, False, True
            converted = True
        else:
            _log_to(logger, f"[DOC] SKIP CONVERT {doc_path.name}: {docx_path.name} already exists")

        if should_cancel and should_cancel():
            raise UserCancelled()
        if ui_pump:
            ui_pump()

        deleted = trash_file(doc_path, logger=logger)
        if deleted:
            _log_to(logger, f"[DOC] TRASH OLD {doc_path.name}")
        return converted, deleted, False
    except UserCancelled:
        raise
    except Exception as exc:
        _log_to(logger, f"[DOC] ERROR {doc_path.name}: {exc}")
        return False, False, True
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def convert_xls_to_xlsx(
    xls_path: Path,
    excel_app,
    logger: logging.Logger | None = None,
    ui_pump=None,
    should_cancel=None,
) -> tuple[bool, bool, bool]:
    if not _is_legacy_xls(xls_path):
        return False, False, False

    xlsx_path = xls_path.with_suffix(".xlsx")
    converted = False
    workbook = None
    try:
        if should_cancel and should_cancel():
            raise UserCancelled()

        if not xlsx_path.exists():
            _log_to(logger, f"[XLS] CONVERT {xls_path.name} -> {xlsx_path.name}")
            workbook = excel_app.Workbooks.Open(
                Filename=str(xls_path),
                ReadOnly=True,
                AddToMru=False,
            )
            workbook.SaveAs(Filename=str(xlsx_path), FileFormat=EXCEL_FORMAT_XLSX)
            workbook.Close(False)
            workbook = None
            if not xlsx_path.exists() or xlsx_path.stat().st_size <= 0:
                _log_to(logger, f"[XLS] ERROR {xls_path.name}: output was not created")
                return False, False, True
            converted = True
        else:
            _log_to(logger, f"[XLS] SKIP CONVERT {xls_path.name}: {xlsx_path.name} already exists")

        if should_cancel and should_cancel():
            raise UserCancelled()
        if ui_pump:
            ui_pump()

        deleted = trash_file(xls_path, logger=logger)
        if deleted:
            _log_to(logger, f"[XLS] TRASH OLD {xls_path.name}")
        return converted, deleted, False
    except UserCancelled:
        raise
    except Exception as exc:
        _log_to(logger, f"[XLS] ERROR {xls_path.name}: {exc}")
        return False, False, True
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass


# ================== IMAGE -> PDF ==================
def prepare_image_for_pdf(src: Path, temp_dir: Path) -> Path:
    with Image.open(src) as im:
        im = ImageOps.exif_transpose(im)
        w, h = im.size
        scale = min(MAX_W / w, MAX_H / h, 1.0)

        if scale >= 1.0:
            return src

        new_w = max(1, int(w * scale))
        new_h = max(1, int(h * scale))

        im2 = im.resize((new_w, new_h), Image.LANCZOS)
        out = temp_dir / f"{src.stem}__{new_w}x{new_h}{src.suffix.lower()}"

        suf = src.suffix.lower()
        if suf in (".jpg", ".jpeg"):
            im2 = im2.convert("RGB")
            im2.save(out, "JPEG", quality=JPEG_QUALITY, optimize=True)
        elif suf == ".png":
            im2.save(out, "PNG", optimize=True, compress_level=6)
        else:
            out = out.with_suffix(".png")
            im2.save(out, "PNG", optimize=True, compress_level=6)

        return out

def image_to_pdf(img_path: Path, pdf_path: Path):
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    with open(img_path, "rb") as f:
        pdf_bytes = img2pdf.convert(f.read())
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)

# ================== OCR ==================
def _hidden_subprocess_kwargs(quiet: bool = True) -> dict:
    kwargs = {}
    if quiet:
        kwargs.update({
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
        })
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs["startupinfo"] = si
    return kwargs


def kill_process_tree(pid: int):
    if pid <= 0:
        return

    if os.name == "nt":
        # /T = діти процесу теж, /F = форсоване завершення
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            check=False,
            **_hidden_subprocess_kwargs(),
        )
        return

    try:
        os.kill(pid, 15)
    except Exception:
        pass


def kill_lingering_tesseract_stack():
    if os.name != "nt":
        return

    # Fallback при скасуванні: прибираємо завислі OCR/GS процеси
    for image in ("tesseract.exe", "ocrmypdf.exe", "gswin64c.exe", "gswin32c.exe"):
        subprocess.run(
            ["taskkill", "/IM", image, "/T", "/F"],
            check=False,
            **_hidden_subprocess_kwargs(),
        )


def run_subprocess_hidden(cmd: list[str], check: bool = False, ui_pump=None, should_cancel=None) -> int:
    proc = subprocess.Popen(cmd, **_hidden_subprocess_kwargs())
    while True:
        code = proc.poll()
        if code is not None:
            break
        if should_cancel and should_cancel():
            kill_process_tree(proc.pid)
            return 130
        if ui_pump:
            try:
                ui_pump()
            except Exception:
                pass
        time.sleep(0.05)

    if check and code != 0:
        raise subprocess.CalledProcessError(code, cmd)
    return code


def run_ocrmypdf(input_pdf: Path, output_pdf: Path, ui_pump=None, should_cancel=None) -> int:
    # У зібраному .exe sys.executable вказує на цей же exe.
    # Якщо запускати "sys.executable -m ocrmypdf", він перезапускає цей скрипт.
    if getattr(sys, "frozen", False):
        ocrmypdf_cmd = shutil.which("ocrmypdf")
        if not ocrmypdf_cmd:
            return 127
        cmd = [ocrmypdf_cmd]
    else:
        cmd = [sys.executable, "-m", "ocrmypdf"]

    cmd.extend([
        "--jobs", "1",
        "--language", OCR_LANG,
        "--force-ocr",
        "--output-type", "pdf",
        str(input_pdf),
        str(output_pdf),
    ])

    return run_subprocess_hidden(cmd, check=False, ui_pump=ui_pump, should_cancel=should_cancel)


def safe_replace(src: Path, dst: Path):
    os.replace(str(src), str(dst))


def ensure_hidden_console_windows():
    """Створює приховану консоль, щоб tesseract/gs не відкривали свої вікна."""
    if os.name != "nt":
        return

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    user32 = ctypes.WinDLL("user32", use_last_error=True)

    SW_HIDE = 0

    hwnd = kernel32.GetConsoleWindow()
    if hwnd == 0:
        # У pythonw консолі нема -> створюємо
        kernel32.AllocConsole()
        hwnd = kernel32.GetConsoleWindow()

    if hwnd:
        user32.ShowWindow(hwnd, SW_HIDE)

# ================== PDF META: REMOVE TITLE ==================
def compressPDF(
    pdf_path: Path,
    logger: logging.Logger | None = None,
    ui_pump=None,
    should_cancel=None,
) -> tuple[bool, Path]:
    old_size = pdf_path.stat().st_size
    tmp_compressed = pdf_path.with_name(pdf_path.stem + "._comp_tmp.pdf")

    def _log(msg: str):
        if logger:
            logger.info(msg)

    def _cleanup_tmp():
        try:
            tmp_compressed.unlink(missing_ok=True)
        except Exception:
            pass

    _log(f"[PDF][COMPRESS] START {pdf_path.name} ({old_size // 1024} KB)")

    gs_exe = find_ghostscript_exe()
    if gs_exe:
        _log(f"[PDF][COMPRESS] {pdf_path.name}: Ghostscript -> /ebook")
        try:
            cmd = [
                str(gs_exe),
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/ebook",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                "-dAutoRotatePages=/None",
                f"-sOutputFile={tmp_compressed}",
                str(pdf_path),
            ]
            rc = run_subprocess_hidden(cmd, check=True, ui_pump=ui_pump, should_cancel=should_cancel)
            if rc == 130:
                raise UserCancelled()

            if tmp_compressed.exists():
                new_size = tmp_compressed.stat().st_size
                if 0 < new_size < old_size:
                    _log(
                        f"[PDF][COMPRESS] OK {pdf_path.name}: "
                        f"{old_size // 1024} KB -> {new_size // 1024} KB"
                    )
                    return True, tmp_compressed
        except Exception:
            _log(f"[PDF][COMPRESS] {pdf_path.name}: Ghostscript failed, fallback to pypdf")
        _cleanup_tmp()
    else:
        _log(f"[PDF][COMPRESS] {pdf_path.name}: Ghostscript not found, fallback to pypdf")

    try:
        _log(f"[PDF][COMPRESS] {pdf_path.name}: pypdf compression")
        reader = PdfReader(str(pdf_path))
        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                _cleanup_tmp()
                _log(f"[PDF][COMPRESS] SKIP {pdf_path.name}: encrypted")
                return False, pdf_path

        writer = PdfWriter()
        for i, page in enumerate(reader.pages, 1):
            writer.add_page(page)
            if should_cancel and should_cancel():
                raise UserCancelled()
            if ui_pump and i % 5 == 0:
                ui_pump()
        for i, page in enumerate(writer.pages, 1):
            page.compress_content_streams()
            if should_cancel and should_cancel():
                raise UserCancelled()
            if ui_pump and i % 5 == 0:
                ui_pump()

        meta = reader.metadata or {}
        new_meta = {str(k): "" if v is None else str(v) for k, v in meta.items()}
        if new_meta:
            writer.add_metadata(new_meta)

        with tmp_compressed.open("wb") as f:
            writer.write(f)

        if tmp_compressed.exists():
            new_size = tmp_compressed.stat().st_size
            if 0 < new_size < old_size:
                _log(
                    f"[PDF][COMPRESS] OK {pdf_path.name}: "
                    f"{old_size // 1024} KB -> {new_size // 1024} KB"
                )
                return True, tmp_compressed
    except Exception:
        pass

    _cleanup_tmp()
    _log(f"[PDF][COMPRESS] SKIP {pdf_path.name}: no size reduction")
    return False, pdf_path


def remove_pdf_title_inplace(
    pdf_path: Path,
    logger: logging.Logger | None = None,
    ui_pump=None,
    should_cancel=None,
) -> str:
    compressed = False
    compressed_tmp = pdf_path
    no_title_tmp = pdf_path.with_name(pdf_path.name + "._notitle_tmp")
    try:
        compressed, compressed_tmp = compressPDF(
            pdf_path,
            logger=logger,
            ui_pump=ui_pump,
            should_cancel=should_cancel,
        )
        source_pdf = compressed_tmp
        reader = PdfReader(str(source_pdf))

        if getattr(reader, "is_encrypted", False):
            if compressed and compressed_tmp != pdf_path:
                safe_replace(compressed_tmp, pdf_path)
                return "OK_COMPRESS_ONLY"
            return "SKIP_ENCRYPTED"

        md = {}
        if reader.metadata:
            for k, v in reader.metadata.items():
                if v is not None:
                    md[str(k)] = str(v)

        if "/Title" not in md:
            if compressed and compressed_tmp != pdf_path:
                safe_replace(compressed_tmp, pdf_path)
                return "OK_COMPRESS_ONLY"
            return "SKIP_NO_TITLE"

        md.pop("/Title", None)

        writer = PdfWriter()
        for i, page in enumerate(reader.pages, 1):
            writer.add_page(page)
            if should_cancel and should_cancel():
                raise UserCancelled()
            if ui_pump and i % 5 == 0:
                ui_pump()

        # Записуємо метадані назад (без /Title). Якщо md пустий, просто не додаємо.
        if md:
            writer.add_metadata(md)

        with no_title_tmp.open("wb") as f:
            writer.write(f)

        safe_replace(no_title_tmp, pdf_path)
        return "OK"
    finally:
        if compressed_tmp != pdf_path:
            try:
                compressed_tmp.unlink(missing_ok=True)
            except Exception:
                pass
        try:
            no_title_tmp.unlink(missing_ok=True)
        except Exception:
            pass


# ================== PDF -> WORD ==================
def is_pipeline_temp_pdf(pdf_path: Path) -> bool:
    return pdf_path.name.endswith(
        (
            "._ocr_tmp.pdf",
            "._comp_tmp.pdf",
        )
    )


def create_word_app_for_pdf_conversion(logger: logging.Logger | None = None):
    if win32 is None:
        raise RuntimeError("pywin32 недоступний: не можна запустити Microsoft Word через COM.")

    word_app = win32.DispatchEx(WORD_COM_PROG_ID)
    word_app.Visible = False
    word_app.DisplayAlerts = WORD_ALERTS_NONE
    try:
        word_app.AutomationSecurity = WORD_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception:
        pass
    if logger:
        logger.info("[WORD] Microsoft Word COM started")
    return word_app


def convert_pdf_to_word(
    pdf_path: Path,
    word_app,
    logger: logging.Logger | None = None,
    ui_pump=None,
    should_cancel=None,
) -> str:
    def _log(message: str) -> None:
        if logger:
            logger.info(message)

    if pdf_path.suffix.lower() != PDF_EXT:
        return "SKIP_NOT_PDF"
    if is_pipeline_temp_pdf(pdf_path):
        return "SKIP_TEMP"

    docx_path = pdf_path.with_suffix(".docx")
    if docx_path.exists():
        _log(f"[WORD] SKIP {pdf_path.name}: {docx_path.name} already exists")
        return "SKIP_EXISTS"

    if should_cancel and should_cancel():
        raise UserCancelled()

    tmp_docx = pdf_path.with_name(pdf_path.stem + "._word_tmp.docx")
    doc = None

    try:
        try:
            tmp_docx.unlink(missing_ok=True)
        except Exception:
            pass

        _log(f"[WORD] CONVERT {pdf_path.name} -> {docx_path.name}")
        if ui_pump:
            ui_pump()

        doc = word_app.Documents.Open(
            FileName=str(pdf_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        try:
            doc.SaveAs2(
                FileName=str(tmp_docx),
                FileFormat=WORD_FORMAT_DOCX,
                AddToRecentFiles=False,
            )
        except AttributeError:
            doc.SaveAs(
                FileName=str(tmp_docx),
                FileFormat=WORD_FORMAT_DOCX,
                AddToRecentFiles=False,
            )
        doc.Close(False)
        doc = None

        if should_cancel and should_cancel():
            raise UserCancelled()
        if ui_pump:
            ui_pump()

        if not tmp_docx.exists() or tmp_docx.stat().st_size <= 0:
            _log(f"[WORD] ERROR {pdf_path.name}: output file was not created")
            return "FAILED"
        if docx_path.exists():
            _log(f"[WORD] SKIP {pdf_path.name}: {docx_path.name} already exists")
            return "SKIP_EXISTS"

        safe_replace(tmp_docx, docx_path)
        return "OK"
    except UserCancelled:
        raise
    except Exception as exc:
        _log(f"[WORD] ERROR {pdf_path.name}: {exc}")
        return "FAILED"
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            tmp_docx.unlink(missing_ok=True)
        except Exception:
            pass


# ================== MAIN ==================
def process_folder_pipeline(
    base: Path,
    progress_window,
    convert_to_word: bool = False,
    convert_images_to_pdf: bool = False,
    perform_ocr: bool = True,
    cleanup_temp_locks: bool = True,
    convert_office_formats: bool = True,
    process_pdf: bool = True,
) -> dict:
    global _active_trash_stager

    log_path = base / "process.log"
    logger = setup_logger(log_path)
    ensure_hidden_console_windows()
    temp_dir = Path(tempfile.mkdtemp(prefix="img_resize_"))
    trash_stager = TrashStager(base, logger=logger)
    previous_trash_stager = _active_trash_stager
    _active_trash_stager = trash_stager

    office_temp_deleted = 0
    doc_converted = 0
    doc_deleted = 0
    doc_failed = 0
    xls_converted = 0
    xls_deleted = 0
    xls_failed = 0
    converted = 0
    convert_failed = 0
    ocr_ok = 0
    ocr_failed = 0
    title_ok = 0
    compress_only_ok = 0
    title_skipped = 0
    title_encrypted = 0
    title_failed = 0
    word_ok = 0
    word_skipped = 0
    word_failed = 0
    trashed_files = 0
    trash_destination = ""
    cancelled = False
    total_stages = (
        (1 if cleanup_temp_locks else 0)
        + (1 if convert_office_formats else 0)
        + (1 if convert_images_to_pdf else 0)
        + (1 if perform_ocr else 0)
        + (1 if process_pdf else 0)
        + (1 if convert_to_word else 0)
    )
    total_stages = max(1, total_stages)
    stage_no = 1

    try:
        if cleanup_temp_locks:
            logger.info(f"ЕТАП {stage_no}/{total_stages}: CLEANUP Office temp/lock")
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: Очищення temp/lock",
                detail="Пошук службових Office файлів...",
                current=0,
                total=1,
                file_name="",
            )
            office_temp_deleted += cleanup_office_temp_files(
                base,
                logger=logger,
                should_cancel=progress_window.is_cancelled,
            )
            progress_window.update(
                current=1,
                total=1,
                detail="Очищення temp/lock завершено",
            )
            stage_no += 1
            logger.info("")
        else:
            logger.info("CLEANUP Office temp/lock: skipped by user setting")
            logger.info("")

        if convert_office_formats:
            logger.info(f"ЕТАП {stage_no}/{total_stages}: NEWER OFFICE FORMAT")
            doc_files = [p for p in iter_files_recursive(base) if _is_legacy_doc(p)]
            xls_files = [p for p in iter_files_recursive(base) if _is_legacy_xls(p)]
            office_total = len(doc_files) + len(xls_files)
            office_done = 0
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: Новіший формат",
                detail="Конвертація DOC -> DOCX та XLS -> XLSX...",
                current=0,
                total=office_total,
                file_name="",
            )

            if win32 is None:
                doc_failed += len(doc_files)
                xls_failed += len(xls_files)
                logger.info("[OFFICE] pywin32 недоступний. Пропуск конвертації DOC/XLS.")
                progress_window.update(current=office_total, total=office_total, file_name="")
            else:
                com_initialized = False
                try:
                    if pythoncom is not None:
                        pythoncom.CoInitialize()
                        com_initialized = True

                    if doc_files:
                        word_app = None
                        doc_done = 0
                        try:
                            word_app = create_word_app_for_pdf_conversion(logger=logger)
                            for p in doc_files:
                                if progress_window.is_cancelled():
                                    raise UserCancelled()
                                progress_window.update(
                                    detail="Конвертація DOC -> DOCX...",
                                    current=office_done,
                                    total=office_total,
                                    file_name=p.name,
                                )
                                conv, deleted, failed = convert_doc_to_docx(
                                    p,
                                    word_app,
                                    logger=logger,
                                    ui_pump=progress_window.refresh,
                                    should_cancel=progress_window.is_cancelled,
                                )
                                doc_converted += int(conv)
                                doc_deleted += int(deleted)
                                doc_failed += int(failed)
                                doc_done += 1
                                office_done += 1
                                progress_window.update(
                                    current=office_done,
                                    total=office_total,
                                    file_name=p.name,
                                )
                                if cleanup_temp_locks and office_done % 50 == 0:
                                    office_temp_deleted += cleanup_office_temp_files(
                                        base,
                                        logger=logger,
                                        should_cancel=progress_window.is_cancelled,
                                    )
                        except UserCancelled:
                            raise
                        except Exception as exc:
                            remaining = len(doc_files) - doc_done
                            doc_failed += remaining
                            office_done += remaining
                            logger.info(f"[DOC] ERROR Microsoft Word COM unavailable: {exc}")
                            progress_window.update(current=office_done, total=office_total)
                        finally:
                            if word_app is not None:
                                try:
                                    word_app.Quit()
                                except Exception:
                                    pass

                    if xls_files:
                        excel_app = None
                        xls_done = 0
                        try:
                            excel_app = create_excel_app_for_conversion(logger=logger)
                            for p in xls_files:
                                if progress_window.is_cancelled():
                                    raise UserCancelled()
                                progress_window.update(
                                    detail="Конвертація XLS -> XLSX...",
                                    current=office_done,
                                    total=office_total,
                                    file_name=p.name,
                                )
                                conv, deleted, failed = convert_xls_to_xlsx(
                                    p,
                                    excel_app,
                                    logger=logger,
                                    ui_pump=progress_window.refresh,
                                    should_cancel=progress_window.is_cancelled,
                                )
                                xls_converted += int(conv)
                                xls_deleted += int(deleted)
                                xls_failed += int(failed)
                                xls_done += 1
                                office_done += 1
                                progress_window.update(
                                    current=office_done,
                                    total=office_total,
                                    file_name=p.name,
                                )
                                if cleanup_temp_locks and office_done % 50 == 0:
                                    office_temp_deleted += cleanup_office_temp_files(
                                        base,
                                        logger=logger,
                                        should_cancel=progress_window.is_cancelled,
                                    )
                        except UserCancelled:
                            raise
                        except Exception as exc:
                            remaining = len(xls_files) - xls_done
                            xls_failed += remaining
                            office_done += remaining
                            logger.info(f"[XLS] ERROR Microsoft Excel COM unavailable: {exc}")
                            progress_window.update(current=office_done, total=office_total)
                        finally:
                            if excel_app is not None:
                                try:
                                    excel_app.Quit()
                                except Exception:
                                    pass
                finally:
                    if com_initialized:
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass

            stage_no += 1
            logger.info("")
        else:
            logger.info("NEWER OFFICE FORMAT: skipped by user setting")
            logger.info("")

        if convert_images_to_pdf:
            logger.info(f"ЕТАП {stage_no}/{total_stages}: CONVERT (images -> pdf)")
            image_files = [p for p in iter_files_recursive(base) if is_image_file(p)]
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: Image -> PDF",
                detail="Конвертація зображень у PDF...",
                current=0,
                total=len(image_files),
                file_name="",
            )
            for i, p in enumerate(image_files, 1):
                if progress_window.is_cancelled():
                    raise UserCancelled()
                progress_window.update(current=i - 1, total=len(image_files), file_name=p.name)

                out_pdf = p.with_suffix(".pdf")
                if out_pdf.exists():
                    progress_window.update(current=i, total=len(image_files), file_name=p.name)
                    continue

                logger.info(f"[CONVERT] {p.relative_to(base)}")
                try:
                    prep = prepare_image_for_pdf(p, temp_dir)
                    image_to_pdf(prep, out_pdf)
                    send2trash(str(p))
                    converted += 1
                except Exception:
                    convert_failed += 1
                progress_window.update(current=i, total=len(image_files), file_name=p.name)

            stage_no += 1
            logger.info("")
        else:
            logger.info("CONVERT (images -> pdf): skipped by user setting")
            logger.info("")

        if perform_ocr:
            logger.info(f"ЕТАП {stage_no}/{total_stages}: OCR (pdf -> searchable pdf)")

            pdf_files_stage2 = [p for p in iter_files_recursive(base) if is_pdf_file(p)]
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: OCR",
                detail="Розпізнавання тексту в PDF...",
                current=0,
                total=len(pdf_files_stage2),
                file_name="",
            )
            for i, p in enumerate(pdf_files_stage2, 1):
                if progress_window.is_cancelled():
                    raise UserCancelled()
                progress_window.update(current=i - 1, total=len(pdf_files_stage2), file_name=p.name)

                logger.info(f"[OCR] {p.relative_to(base)}")
                tmp_out = p.with_name(p.stem + "._ocr_tmp.pdf")

                try:
                    code = run_ocrmypdf(
                        p,
                        tmp_out,
                        ui_pump=progress_window.refresh,
                        should_cancel=progress_window.is_cancelled,
                    )
                    if code == 130:
                        raise UserCancelled()
                    if code == 0 and tmp_out.exists():
                        safe_replace(tmp_out, p)
                        ocr_ok += 1
                    else:
                        ocr_failed += 1
                        if tmp_out.exists():
                            tmp_out.unlink()
                except Exception:
                    ocr_failed += 1
                    if tmp_out.exists():
                        tmp_out.unlink()
                progress_window.update(current=i, total=len(pdf_files_stage2), file_name=p.name)

            stage_no += 1
            logger.info("")
        else:
            logger.info("OCR (pdf -> searchable pdf): skipped by user setting")
            logger.info("")

        if process_pdf:
            logger.info(f"ЕТАП {stage_no}/{total_stages}: PDF PROCESSING (compress + remove Title)")

            pdf_files_stage3 = [p for p in iter_files_recursive(base) if is_pdf_file(p)]
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: Обробка PDF",
                detail="Стиснення PDF + видалення /Title із метаданих...",
                current=0,
                total=len(pdf_files_stage3),
                file_name="",
            )
            for i, p in enumerate(pdf_files_stage3, 1):
                if progress_window.is_cancelled():
                    raise UserCancelled()
                progress_window.update(current=i - 1, total=len(pdf_files_stage3), file_name=p.name)

                if (
                    p.name.endswith("._ocr_tmp.pdf")
                    or p.name.endswith("._notitle_tmp")
                    or p.name.endswith("._comp_tmp.pdf")
                ):
                    progress_window.update(current=i, total=len(pdf_files_stage3), file_name=p.name)
                    continue

                try:
                    res = remove_pdf_title_inplace(
                        p,
                        logger=logger,
                        ui_pump=progress_window.refresh,
                        should_cancel=progress_window.is_cancelled,
                    )
                    rel = p.relative_to(base)

                    if res == "OK":
                        title_ok += 1
                    elif res == "OK_COMPRESS_ONLY":
                        compress_only_ok += 1
                        logger.info(f"[TITLE] {rel} | skip (no /Title) | compressed")
                    elif res == "SKIP_NO_TITLE":
                        title_skipped += 1
                        logger.info(f"[TITLE] {rel} | skip (no /Title)")
                    elif res == "SKIP_ENCRYPTED":
                        title_encrypted += 1
                        logger.info(f"[TITLE] {rel} | skip (encrypted)")
                    else:
                        title_failed += 1
                        logger.info(f"[TITLE] {rel} | skip ({res})")
                except Exception:
                    title_failed += 1
                    logger.info(f"[TITLE] {p.relative_to(base)} | ERROR")
                progress_window.update(current=i, total=len(pdf_files_stage3), file_name=p.name)

                if cleanup_temp_locks and i % 100 == 0:
                    office_temp_deleted += cleanup_office_temp_files(
                        base,
                        logger=logger,
                        should_cancel=progress_window.is_cancelled,
                    )

            stage_no += 1
        else:
            logger.info("PDF PROCESSING: skipped by user setting")

        if convert_to_word:
            logger.info("")
            logger.info(f"ЕТАП {stage_no}/{total_stages}: PDF -> WORD")

            pdf_files_stage4 = [
                p
                for p in iter_files_recursive(base)
                if is_pdf_file(p) and not is_pipeline_temp_pdf(p)
            ]
            progress_window.update(
                header=f"Етап {stage_no}/{total_stages}: PDF -> Word",
                detail="Конвертація PDF у DOCX через Microsoft Word...",
                current=0,
                total=len(pdf_files_stage4),
                file_name="",
            )
            if pdf_files_stage4:
                word_app = None
                com_initialized = False
                word_processed = 0
                try:
                    if pythoncom is not None:
                        pythoncom.CoInitialize()
                        com_initialized = True
                    word_app = create_word_app_for_pdf_conversion(logger=logger)

                    for i, p in enumerate(pdf_files_stage4, 1):
                        if progress_window.is_cancelled():
                            raise UserCancelled()
                        progress_window.update(current=i - 1, total=len(pdf_files_stage4), file_name=p.name)

                        result = convert_pdf_to_word(
                            p,
                            word_app,
                            logger=logger,
                            ui_pump=progress_window.refresh,
                            should_cancel=progress_window.is_cancelled,
                        )
                        if result == "OK":
                            word_ok += 1
                        elif result == "SKIP_EXISTS":
                            word_skipped += 1
                        else:
                            word_failed += 1
                        word_processed += 1
                        progress_window.update(current=i, total=len(pdf_files_stage4), file_name=p.name)
                except UserCancelled:
                    raise
                except Exception as exc:
                    word_failed += len(pdf_files_stage4) - word_processed
                    logger.info(f"[WORD] ERROR Microsoft Word COM unavailable: {exc}")
                finally:
                    if word_app is not None:
                        try:
                            word_app.Quit()
                        except Exception:
                            pass
                    if com_initialized:
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass

        if cleanup_temp_locks:
            logger.info("")
            logger.info("FINAL: CLEANUP Office temp/lock")
            progress_window.update(
                header="Завершення",
                detail="Фінальне очищення temp/lock файлів...",
                current=0,
                total=1,
                file_name="",
            )
            office_temp_deleted += cleanup_office_temp_files(
                base,
                logger=logger,
                should_cancel=progress_window.is_cancelled,
            )
            progress_window.update(current=1, total=1, detail="Готово", file_name="")

    except UserCancelled:
        cancelled = True
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        try:
            trash_stager.finalize()
        except Exception as exc:
            logger.info(f"[TRASH] ERROR finalizing staged files: {exc}")
        trashed_files = trash_stager.staged_count
        trash_destination = trash_stager.trash_destination
        _active_trash_stager = previous_trash_stager
        close_logger(logger)

    if not cancelled:
        for _ in range(5):
            try:
                send2trash(str(log_path))
                break
            except Exception:
                time.sleep(0.25)

    return {
        "cancelled": cancelled,
        "office_temp_deleted": office_temp_deleted,
        "doc_converted": doc_converted,
        "doc_deleted": doc_deleted,
        "doc_failed": doc_failed,
        "xls_converted": xls_converted,
        "xls_deleted": xls_deleted,
        "xls_failed": xls_failed,
        "converted": converted,
        "convert_failed": convert_failed,
        "ocr_ok": ocr_ok,
        "ocr_failed": ocr_failed,
        "title_ok": title_ok,
        "compress_only_ok": compress_only_ok,
        "title_skipped": title_skipped,
        "title_encrypted": title_encrypted,
        "title_failed": title_failed,
        "word_ok": word_ok,
        "word_skipped": word_skipped,
        "word_failed": word_failed,
        "trashed_files": trashed_files,
        "trash_destination": trash_destination,
    }


def run_folder_pipeline_with_progress(
    base: Path,
    progress_window: ProgressWindow,
    convert_to_word: bool = False,
    convert_images_to_pdf: bool = False,
    perform_ocr: bool = True,
    cleanup_temp_locks: bool = True,
    convert_office_formats: bool = True,
    process_pdf: bool = True,
) -> dict:
    updates = queue.Queue()
    result = {"payload": None, "error": None}
    progress_proxy = ProgressUpdateProxy(updates, should_cancel=progress_window.is_cancelled)

    def worker() -> None:
        try:
            result["payload"] = process_folder_pipeline(
                base,
                progress_proxy,
                convert_to_word=convert_to_word,
                convert_images_to_pdf=convert_images_to_pdf,
                perform_ocr=perform_ocr,
                cleanup_temp_locks=cleanup_temp_locks,
                convert_office_formats=convert_office_formats,
                process_pdf=process_pdf,
            )
        except Exception as exc:
            result["error"] = exc

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    while thread.is_alive():
        apply_pending_progress_updates(progress_window, updates)
        progress_window.refresh()
        time.sleep(0.03)

    thread.join()
    apply_pending_progress_updates(progress_window, updates)
    progress_window.refresh()

    if result["error"] is not None:
        raise result["error"]
    if result["payload"] is None:
        raise RuntimeError("PDF-обробка завершилася без результату.")
    return result["payload"]


def main():
    root = Tk()
    root.withdraw()
    install_frozen_executable_icon(root)
    install_dark_title_bar(root)
    root.attributes("-topmost", True)

    selection = ask_folder_settings_window(
        root,
        app_name="FDataOCR",
        subtitle="Підготовка обробки файлів у вибраній директорії...",
        hint="Вкажіть папку та оберіть потрібні етапи обробки.",
        browse_title="Обери папку для обробки",
    )
    if selection is None:
        root.destroy()
        return
    (
        base,
        cleanup_temp_locks,
        convert_office_formats,
        convert_images_to_pdf,
        perform_ocr,
        process_pdf,
        convert_to_word,
    ) = selection

    if perform_ocr and not ensure_external_dependencies_available(root):
        root.destroy()
        return
    if convert_to_word and not ensure_word_com_available(root):
        root.destroy()
        return

    progress_window = ProgressWindow(root)
    try:
        result = run_folder_pipeline_with_progress(
            base,
            progress_window,
            convert_to_word=convert_to_word,
            convert_images_to_pdf=convert_images_to_pdf,
            perform_ocr=perform_ocr,
            cleanup_temp_locks=cleanup_temp_locks,
            convert_office_formats=convert_office_formats,
            process_pdf=process_pdf,
        )
    finally:
        progress_window.close()

    if result["cancelled"]:
        kill_lingering_tesseract_stack()
        root.destroy()
        return

    root.lift()
    root.attributes("-topmost", True)
    root.update_idletasks()

    cleanup_summary = (
        f"Office temp/lock у кошику: {result['office_temp_deleted']}\n\n"
        if cleanup_temp_locks
        else "Очищення temp/lock: вимкнено\n\n"
    )
    office_summary = (
        f"DOC -> DOCX: {result['doc_converted']}\n"
        f"XLS -> XLSX: {result['xls_converted']}\n"
        f"Старі DOC у кошику: {result['doc_deleted']}\n"
        f"Старі XLS у кошику: {result['xls_deleted']}\n"
        f"DOC помилок: {result['doc_failed']}\n"
        f"XLS помилок: {result['xls_failed']}\n\n"
        if convert_office_formats
        else "Новіший формат: вимкнено\n\n"
    )
    image_summary = (
        f"Сконвертовано зображень: {result['converted']}\n"
        f"Невдало конвертовано: {result['convert_failed']}\n\n"
        if convert_images_to_pdf
        else "Конвертація зображень у PDF: вимкнено\n\n"
    )
    ocr_summary = (
        f"OCR успішно: {result['ocr_ok']}\n"
        f"OCR помилок: {result['ocr_failed']}\n\n"
        if perform_ocr
        else "OCR PDF: вимкнено\n\n"
    )
    pdf_summary = (
        f"PDF: Title прибрано: {result['title_ok']}\n"
        f"PDF: тільки стиснено (без /Title): {result['compress_only_ok']}\n"
        f"PDF: пропущено без /Title: {result['title_skipped']}\n"
        f"PDF: encrypted: {result['title_encrypted']}\n"
        f"PDF: помилки: {result['title_failed']}\n\n"
        if process_pdf
        else "Обробка PDF: вимкнено\n\n"
    )
    done_message = (
        "Завершено.\n\n"
        + cleanup_summary
        + office_summary
        + image_summary
        + ocr_summary
        + pdf_summary
    )
    if convert_to_word:
        done_message += (
            "\n"
            f"Word: створено DOCX: {result['word_ok']}\n"
            f"Word: пропущено існуючих DOCX: {result['word_skipped']}\n"
            f"Word: помилки: {result['word_failed']}\n"
        )

    if result["trashed_files"]:
        done_message += (
            "\n"
            f"Backup-файлів для кошика: {result['trashed_files']}\n"
            f"Backup-папка у кошику: {result['trash_destination'] or '-'}\n"
        )

    messagebox.showinfo(
        "Готово",
        done_message,
        parent=root,
    )

    root.destroy()


if __name__ == "__main__":
    main()
