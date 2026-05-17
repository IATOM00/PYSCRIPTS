from __future__ import annotations

import sys, hmac, hashlib, argparse, unicodedata
from dataclasses import dataclass

# Titles
APP_TITLE = "ГЕНЕРАТОР КЛЮЧІВ & ПАРОЛІВ"
INACTIVE_APP_TITLE = " "
DEFAULT_DIFFICULTY = "maximum"

# Argon2id parameters are part of the deterministic generator contract.
# Changing them will change every generated key for the same secret.
ARGON2_TIME_COST = 3
ARGON2_MEMORY_COST_KIB = 64 * 1024
ARGON2_PARALLELISM = 2
ARGON2_HASH_LEN = 64

# Сontext values 
MINIMUM_DIFFICULTY_SALT = b"BeatPass.v1.normal.10"
MEDIUM_DIFFICULTY_SALT = b"BeatPass.v1.medium.18"
MAXIMUM_DIFFICULTY_SALT = b"BeatPass.v1.maximum.28"
PASSWORD_CONTEXT_PREFIX = b"BeatPass.password.v1."
MINIMUM_DIFFICULTY_CONTEXT_KEY = "min"
MEDIUM_DIFFICULTY_CONTEXT_KEY = "medium"
MAXIMUM_DIFFICULTY_CONTEXT_KEY = "max"

MIN_SECRET_LENGTH = 5
MIN_KEYBOARD_SEQUENCE_LENGTH = 3
MIN_REPEAT_RUN_LENGTH = 3

# Keep this list separate so weak corporate/project words can be added later.
WEAK_DICTIONARY_WORDS = {
    "admin",
    "administrator",
    "archive",
    "backup",
    "codex",
    "default",
    "guest",
    "login",
    "password",
    "secret",
    "welcome",
    "архів",
    "гость",
    "дефолт",
    "пароль",
    "секрет",
}

KEYBOARD_ROWS = (
    "`1234567890-=",
    "qwertyuiop[]\\",
    "asdfghjkl;'",
    "zxcvbnm,./",
    "йцукенгшщзхї",
    "фівапролджє",
    "ячсмитьбю.",
)

LOWERCASE = "abcdefghijkmnopqrstuvwxyz"
UPPERCASE = "ABCDEFGHJKLMNPQRSTUVWXYZ"
DIGITS = "1234567890"
SPECIAL_CHARS = "!#$%&*+-=?@_"
PASSWORD_ALPHABET = LOWERCASE + UPPERCASE + DIGITS + SPECIAL_CHARS
REQUIRED_CATEGORIES = (LOWERCASE, UPPERCASE, DIGITS, SPECIAL_CHARS)


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


try:
    import tkinter as tk
    from tkinter import ttk
except Exception:
    tk = None
    ttk = None

try:
    from argon2.low_level import ARGON2_VERSION, Type, hash_secret_raw
except Exception:
    ARGON2_VERSION = 19
    Type = None
    hash_secret_raw = None


@dataclass(frozen=True)
class DifficultyProfile:
    title: str
    length: int
    salt: bytes
    context_key: str


DIFFICULTY_PROFILES = {
    "minimum": DifficultyProfile(
        title="minimum",
        length=10,
        salt=MINIMUM_DIFFICULTY_SALT,
        context_key=MINIMUM_DIFFICULTY_CONTEXT_KEY,
    ),
    "medium": DifficultyProfile(
        title="medium",
        length=18,
        salt=MEDIUM_DIFFICULTY_SALT,
        context_key=MEDIUM_DIFFICULTY_CONTEXT_KEY,
    ),
    "maximum": DifficultyProfile(
        title="maximum",
        length=28,
        salt=MAXIMUM_DIFFICULTY_SALT,
        context_key=MAXIMUM_DIFFICULTY_CONTEXT_KEY,
    ),
}

DIFFICULTY_DISPLAY_ORDER = ("maximum", "medium", "minimum")
DIFFICULTY_ALIASES = {"normal": "minimum"}
DIFFICULTY_CLI_CHOICES = DIFFICULTY_DISPLAY_ORDER + tuple(DIFFICULTY_ALIASES)
SECRET_HINT = (
    "Мінімум 5 символів, без вразливих послідовностей."
)
COPY_CONFIRMATION = "Готово! Ключ скопійовано в буфер обміну!"
UI_BG = "#F4F7FB"
UI_SURFACE = "#FFFFFF"
UI_FIELD_BG = "#FFFFFF"
UI_FG = "#15202B"
UI_MUTED_FG = "#617084"
UI_WARNING_FG = "#8f1f1f"
UI_BORDER = "#D6DEE8"
UI_BORDER_ACTIVE = "#0F766E"
UI_SEPARATOR = "#D6DEE8"
UI_CONTROL_BG = "#E8EEF5"
UI_CONTROL_HOVER_BG = "#D7E0EA"
UI_ACCENT = "#0F766E"
UI_ACCENT_DARK = "#0B5D56"
UI_HEADER = "#103C3A"
UI_ACTION_BG = "#053D39"
UI_ACTION_HOVER_BG = "#0B5D56"
UI_ACTION_PRESSED_BG = "#042F2C"
UI_ACTION_FG = "#FFFFFF"
UI_SUCCESS_FG = "#0F766E"
UI_FONT = ("Segoe UI", 10)
UI_MONO_FONT = ("Cascadia Mono", 10)
UI_WINDOW_PAD_X = 18
UI_WINDOW_PAD_Y = 16
UI_PANEL_PAD_X = 14
UI_PANEL_PAD_Y = 14
UI_FIELD_PAD_X = 10
UI_FIELD_PAD_Y = 9
UI_ROW_GAP = 9
UI_CONTROL_PAD_Y = 5
UI_ACTION_PAD_X = 12
UI_ACTION_PAD_Y = 12
DIFFICULTY_MENU_ARROW = "▾"
DIFFICULTY_BUTTON_LABELS = {
    "minimum": "min",
    "medium": "med",
    "maximum": "max",
}


class ValidationError(ValueError):
    pass


def normalize_secret(secret: str) -> str:
    return unicodedata.normalize("NFKC", secret)


def normalize_for_checks(text: str) -> str:
    return normalize_secret(text).casefold()


def has_adjacent_repeat(text: str) -> tuple[bool, str]:
    if not text:
        return False, ""

    run_char = text[0]
    run_length = 1
    longest_repeat = ""
    for char in text[1:]:
        if char == run_char:
            run_length += 1
        else:
            if run_length >= MIN_REPEAT_RUN_LENGTH:
                candidate = run_char * run_length
                if len(candidate) > len(longest_repeat):
                    longest_repeat = candidate
            run_char = char
            run_length = 1
    if run_length >= MIN_REPEAT_RUN_LENGTH:
        candidate = run_char * run_length
        if len(candidate) > len(longest_repeat):
            longest_repeat = candidate
    return bool(longest_repeat), longest_repeat


def iter_keyboard_sequences():
    for row in KEYBOARD_ROWS:
        for size in range(len(row), MIN_KEYBOARD_SEQUENCE_LENGTH - 1, -1):
            for start in range(0, len(row) - size + 1):
                sequence = row[start : start + size]
                yield sequence
                yield sequence[::-1]


def find_keyboard_sequence(text: str) -> str:
    compact = "".join(char for char in text if not char.isspace())
    match = ""
    match_index = len(compact)
    for sequence in iter_keyboard_sequences():
        index = compact.find(sequence)
        if index == -1:
            continue
        if len(sequence) > len(match) or (
            len(sequence) == len(match) and index < match_index
        ):
            match = sequence
            match_index = index
    return match


def find_dictionary_word(text: str) -> str:
    compact = "".join(char for char in text if not char.isspace())
    for word in sorted(WEAK_DICTIONARY_WORDS, key=len, reverse=True):
        prepared_word = normalize_for_checks(word)
        if prepared_word and prepared_word in compact:
            return word
    return ""


def validate_secret(secret: str) -> str:
    if secret != secret.strip():
        raise ValidationError("Секрет не повинен починатися або завершуватися пробілом.")

    normalized = normalize_secret(secret)
    if len(normalized) < MIN_SECRET_LENGTH:
        raise ValidationError(f"Коротка послідовність: < {MIN_SECRET_LENGTH} символів.")

    checked = normalized.casefold()
    has_repeat, repeat = has_adjacent_repeat(checked)
    if has_repeat:
        raise ValidationError(f"Символьна послідовність: '{repeat}'.")

    keyboard_sequence = find_keyboard_sequence(checked)
    if keyboard_sequence:
        raise ValidationError(
            f"Клавіатурна послідовність: '{keyboard_sequence}'."
        )

    weak_word = find_dictionary_word(checked)
    if weak_word:
        raise ValidationError(f"Словникова послідовність: '{weak_word}'.")

    return normalized


def derive_argon2id_seed(secret: str, salt: bytes) -> bytes:
    if hash_secret_raw is None or Type is None:
        raise RuntimeError(
            "Не знайдено залежність 'argon2-cffi'. Встановіть її командою: pip install argon2-cffi"
        )

    return hash_secret_raw(
        secret=secret.encode("utf-8"),
        salt=salt,
        time_cost=ARGON2_TIME_COST,
        memory_cost=ARGON2_MEMORY_COST_KIB,
        parallelism=ARGON2_PARALLELISM,
        hash_len=ARGON2_HASH_LEN,
        type=Type.ID,
        version=ARGON2_VERSION,
    )


def expand_bytes(seed: bytes, context: bytes, length: int) -> bytes:
    result = bytearray()
    counter = 1
    while len(result) < length:
        result.extend(
            hmac.new(seed, context + counter.to_bytes(4, "big"), hashlib.sha256).digest()
        )
        counter += 1
    return bytes(result[:length])


def pick_char(source_byte: int, alphabet: str) -> str:
    return alphabet[source_byte % len(alphabet)]


def pick_distinct_positions(stream: bytes, count: int, length: int) -> list[int]:
    positions: list[int] = []
    cursor = 0
    while len(positions) < count:
        position = stream[cursor % len(stream)] % length
        cursor += 1
        if position not in positions:
            positions.append(position)
    return positions


def generate_key(secret: str, difficulty: str = DEFAULT_DIFFICULTY) -> str:
    difficulty_key = DIFFICULTY_ALIASES.get(difficulty, difficulty)
    if difficulty_key not in DIFFICULTY_PROFILES:
        choices = ", ".join(DIFFICULTY_DISPLAY_ORDER)
        raise ValueError(f"Невідома складність '{difficulty}'. Доступно: {choices}.")

    normalized_secret = validate_secret(secret)
    profile = DIFFICULTY_PROFILES[difficulty_key]
    seed = derive_argon2id_seed(normalized_secret, profile.salt)

    stream = expand_bytes(
        seed,
        PASSWORD_CONTEXT_PREFIX + profile.context_key.encode("ascii"),
        profile.length * 4,
    )
    chars = [
        pick_char(stream[index], PASSWORD_ALPHABET)
        for index in range(profile.length)
    ]

    required_positions = pick_distinct_positions(
        stream[profile.length : profile.length + 32],
        len(REQUIRED_CATEGORIES),
        profile.length,
    )
    cursor = profile.length + 32
    for position, alphabet in zip(required_positions, REQUIRED_CATEGORIES):
        chars[position] = pick_char(stream[cursor % len(stream)], alphabet)
        cursor += 1

    return "".join(chars)


def center_window(window) -> None:
    try:
        window.update_idletasks()
        width = window.winfo_reqwidth()
        height = window.winfo_reqheight()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
    except Exception:
        pass


class BeatPassApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.resizable(False, False)
        self.root.configure(bg=UI_BG)
        install_dark_title_bar(self.root)
        self._bind_title_visibility()
        try:
            self.root.attributes("-topmost", True)
        except tk.TclError:
            pass

        self.secret_var = tk.StringVar()
        self.difficulty_var = tk.StringVar(value=DEFAULT_DIFFICULTY)
        self.generated_var = tk.StringVar(value=SECRET_HINT)
        self.generated_key = ""
        self.secret_notice_var = tk.StringVar()

        self._configure_style()
        self._build_ui()
        self.secret_var.trace_add("write", self._on_input_changed)
        self.difficulty_var.trace_add("write", self._on_input_changed)
        try:
            self.root.grab_set()
            self.root.focus_force()
        except tk.TclError:
            pass

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        try:
            style.configure("TFrame", background=UI_BG)
            style.configure("TLabel", background=UI_BG, foreground=UI_FG, font=UI_FONT)
            style.configure(
                "TEntry",
                fieldbackground=UI_FIELD_BG,
                foreground=UI_FG,
                font=UI_FONT,
            )
        except Exception:
            pass

    def _bind_title_visibility(self) -> None:
        def restore_title(_event=None):
            self.root.title(APP_TITLE)

        def hide_title(_event=None):
            self.root.title(INACTIVE_APP_TITLE)

        try:
            self.root.bind("<Activate>", restore_title, add="+")
            self.root.bind("<Deactivate>", hide_title, add="+")
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        entry_width = 45

        surface = tk.Frame(
            self.root,
            bg=UI_BG,
            padx=UI_WINDOW_PAD_X,
            pady=UI_WINDOW_PAD_Y,
        )
        surface.grid(row=0, column=0, sticky="nsew")
        surface.grid_columnconfigure(0, weight=1)

        panel_shell = tk.Frame(surface, bg=UI_BORDER, bd=0, padx=1, pady=1)
        panel_shell.grid(row=0, column=0, sticky="nsew")
        panel_shell.grid_columnconfigure(0, weight=1)

        panel = tk.Frame(
            panel_shell,
            bg=UI_SURFACE,
            bd=0,
            padx=UI_PANEL_PAD_X,
            pady=UI_PANEL_PAD_Y,
        )
        panel.grid(row=0, column=0, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)

        accent_line = tk.Frame(panel, bg=UI_HEADER, height=3, bd=0)
        accent_line.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        container = tk.Frame(panel, bg=UI_SURFACE, bd=0)
        container.grid(row=1, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)

        secret_frame = tk.Frame(container, bg=UI_SURFACE)
        secret_frame.grid(row=0, column=0, sticky="we")
        secret_frame.grid_columnconfigure(0, weight=1)

        self.secret_box = tk.Frame(secret_frame, bg=UI_BORDER, bd=0, padx=1, pady=1)
        self.secret_box.grid(row=0, column=0, sticky="ew")
        self.secret_box.grid_columnconfigure(0, weight=1)

        secret_panel = tk.Frame(self.secret_box, bg=UI_FIELD_BG, bd=0)
        secret_panel.grid(row=0, column=0, sticky="ew")
        secret_panel.grid_columnconfigure(0, weight=1)

        self.secret_entry = tk.Entry(
            secret_panel,
            bg=UI_FIELD_BG,
            bd=0,
            fg=UI_FG,
            font=UI_FONT,
            highlightthickness=0,
            insertbackground=UI_FG,
            justify="center",
            relief="flat",
            textvariable=self.secret_var,
            width=entry_width,
        )
        self.secret_entry.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(UI_FIELD_PAD_X, 6),
            pady=UI_FIELD_PAD_Y,
        )
        self.secret_entry.bind("<Return>", lambda _event: self.generate())
        self.secret_entry.bind("<FocusIn>", self._on_secret_focus_in, add="+")
        self.secret_entry.bind(
            "<FocusOut>",
            lambda _event: self._set_secret_focus(False),
            add="+",
        )

        self.secret_notice_label = tk.Label(
            secret_panel,
            anchor="center",
            bg=UI_FIELD_BG,
            fg=UI_SUCCESS_FG,
            font=("Segoe UI", 10, "bold"),
            justify="center",
            padx=0,
            textvariable=self.secret_notice_var,
        )
        self.secret_notice_label.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(UI_FIELD_PAD_X, 6),
            pady=UI_FIELD_PAD_Y,
        )
        self.secret_notice_label.grid_remove()
        self.secret_notice_label.bind(
            "<Button-1>", lambda _event: self._hide_secret_notice(focus=True)
        )

        self.secret_separator = tk.Frame(secret_panel, bg=UI_SEPARATOR, width=1)
        self.secret_separator.grid(row=0, column=1, sticky="ns")

        self.difficulty_button = tk.Menubutton(
            secret_panel,
            activebackground=UI_CONTROL_HOVER_BG,
            activeforeground=UI_FG,
            bg=UI_CONTROL_BG,
            bd=0,
            cursor="hand2",
            fg=UI_FG,
            font=("Segoe UI", 9, "bold"),
            highlightthickness=0,
            padx=8,
            pady=UI_CONTROL_PAD_Y,
            relief="flat",
            text=DIFFICULTY_MENU_ARROW,
            width=6,
        )
        self.difficulty_menu = tk.Menu(
            self.difficulty_button,
            activebackground=UI_ACCENT_DARK,
            activeborderwidth=0,
            activeforeground=UI_ACTION_FG,
            bg=UI_SURFACE,
            bd=0,
            fg=UI_FG,
            font=UI_FONT,
            relief="flat",
            tearoff=False,
        )
        self.difficulty_button.configure(menu=self.difficulty_menu)
        self.difficulty_button.grid(row=0, column=2, sticky="ns")
        self._bind_hover(self.difficulty_button, UI_CONTROL_BG, UI_CONTROL_HOVER_BG)
        self._refresh_difficulty_menu()

        result_frame = tk.Frame(container, bg=UI_SURFACE)
        result_frame.grid(row=1, column=0, sticky="we", pady=(UI_ROW_GAP, 0))
        result_frame.grid_columnconfigure(0, weight=1)

        self.result_box = tk.Frame(result_frame, bg=UI_BORDER, bd=0, padx=1, pady=1)
        self.result_box.grid(row=0, column=0, sticky="ew")
        self.result_box.grid_columnconfigure(0, weight=1)

        result_panel = tk.Frame(self.result_box, bg=UI_FIELD_BG, bd=0)
        result_panel.grid(row=0, column=0, sticky="ew")
        result_panel.grid_columnconfigure(0, weight=1)

        self.result_entry = tk.Entry(
            result_panel,
            bg=UI_FIELD_BG,
            bd=0,
            fg=UI_MUTED_FG,
            font=UI_MONO_FONT,
            highlightthickness=0,
            insertbackground=UI_FG,
            justify="center",
            readonlybackground=UI_FIELD_BG,
            relief="flat",
            state="readonly",
            textvariable=self.generated_var,
            width=entry_width,
        )
        self.result_entry.grid(
            row=0,
            column=0,
            sticky="we",
            padx=UI_FIELD_PAD_X,
            pady=UI_FIELD_PAD_Y,
        )
        self.result_entry.bind(
            "<FocusIn>",
            lambda _event: self._set_result_focus(True),
            add="+",
        )
        self.result_entry.bind(
            "<FocusOut>",
            lambda _event: self._set_result_focus(False),
            add="+",
        )

        self.generate_box = tk.Frame(container, bg=UI_ACCENT_DARK, bd=0, padx=1, pady=1)
        self.generate_box.grid(row=0, column=1, rowspan=2, padx=(10, 0), sticky="ns")
        self.generate_box.grid_rowconfigure(0, weight=1)
        self.generate_box.grid_columnconfigure(0, weight=1)

        self.generate_button = tk.Button(
            self.generate_box,
            activebackground=UI_ACTION_PRESSED_BG,
            activeforeground=UI_ACTION_FG,
            bg=UI_ACTION_BG,
            bd=0,
            command=self.generate,
            cursor="hand2",
            fg=UI_ACTION_FG,
            font=("Segoe UI Symbol", 10, "bold"),
            highlightthickness=0,
            padx=UI_ACTION_PAD_X,
            pady=UI_ACTION_PAD_Y,
            relief="flat",
            text="▶",
            width=5,
        )
        self.generate_button.grid(row=0, column=0, sticky="nsew")
        self._bind_hover(
            self.generate_button,
            UI_ACTION_BG,
            UI_ACTION_HOVER_BG,
            UI_ACTION_FG,
            UI_ACTION_FG,
        )

        self.secret_entry.focus_set()
        self._set_secret_focus(True)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.update_idletasks()

    def _bind_hover(
        self,
        widget,
        normal_bg: str,
        hover_bg: str,
        normal_fg: str | None = None,
        hover_fg: str | None = None,
    ) -> None:
        def on_enter(_event=None):
            options = {"bg": hover_bg}
            if hover_fg is not None:
                options["fg"] = hover_fg
            widget.configure(**options)

        def on_leave(_event=None):
            options = {"bg": normal_bg}
            if normal_fg is not None:
                options["fg"] = normal_fg
            widget.configure(**options)

        widget.bind("<Enter>", on_enter, add="+")
        widget.bind("<Leave>", on_leave, add="+")

    def _set_secret_focus(self, focused: bool) -> None:
        if hasattr(self, "secret_box"):
            self.secret_box.configure(bg=UI_BORDER_ACTIVE if focused else UI_BORDER)

    def _set_result_focus(self, focused: bool) -> None:
        if hasattr(self, "result_box"):
            self.result_box.configure(bg=UI_BORDER_ACTIVE if focused else UI_BORDER)

    def _on_secret_focus_in(self, _event=None) -> None:
        self._hide_secret_notice()
        self._set_secret_focus(True)

    def _set_result_message(self, text: str = SECRET_HINT, warning: bool = False) -> None:
        self.generated_key = ""
        self.generated_var.set(text)
        if hasattr(self, "result_entry"):
            self.result_entry.configure(fg=UI_WARNING_FG if warning else UI_MUTED_FG)
            self.result_entry.xview_moveto(0)

    def _set_generated(self, value: str) -> None:
        self.generated_key = value
        self.generated_var.set(value)
        if hasattr(self, "result_entry"):
            self.result_entry.configure(fg=UI_FG)
            self.result_entry.xview_moveto(0)

    def _hide_secret_notice(self, focus: bool = False) -> None:
        if hasattr(self, "secret_notice_label"):
            self.secret_notice_label.grid_remove()
        if hasattr(self, "secret_separator"):
            self.secret_separator.grid()
        if hasattr(self, "difficulty_button"):
            self.difficulty_button.grid()
        self.secret_notice_var.set("")
        if focus:
            self.secret_entry.focus_set()

    def _show_secret_notice(self, text: str) -> None:
        self.secret_notice_var.set(text)
        if hasattr(self, "secret_separator"):
            self.secret_separator.grid_remove()
        if hasattr(self, "difficulty_button"):
            self.difficulty_button.grid_remove()
        self.secret_notice_label.grid()
        self.secret_notice_label.lift()

    def _on_input_changed(self, *_args) -> None:
        if self.generated_key or self.generated_var.get() != SECRET_HINT:
            self._set_result_message()
        if hasattr(self, "difficulty_menu"):
            self._refresh_difficulty_menu()

    def _refresh_difficulty_menu(self) -> None:
        self.difficulty_menu.delete(0, "end")
        current = self.difficulty_var.get()
        label = DIFFICULTY_BUTTON_LABELS.get(current, current)
        self.difficulty_button.configure(text=f"{label} {DIFFICULTY_MENU_ARROW}")
        for difficulty in DIFFICULTY_DISPLAY_ORDER:
            self.difficulty_menu.add_checkbutton(
                label=difficulty,
                offvalue=current,
                onvalue=difficulty,
                selectcolor=UI_FG,
                variable=self.difficulty_var,
            )

    def generate(self) -> None:
        try:
            generated = generate_key(self.secret_var.get(), self.difficulty_var.get())
        except Exception as exc:
            self._set_result_message(str(exc), warning=True)
            self.secret_entry.focus_set()
            return

        self._set_generated(generated)
        self.root.clipboard_clear()
        self.root.clipboard_append(generated)
        self._show_secret_notice(COPY_CONFIRMATION)


def run_gui() -> int:
    if tk is None or ttk is None:
        print("Tkinter недоступний. Запустіть із --secret для консольної генерації.", file=sys.stderr)
        return 1
    root = tk.Tk()
    root.withdraw()
    app = BeatPassApp(root)
    install_frozen_executable_icon(root)
    root.update_idletasks()
    center_window(root)
    root.deiconify()
    root.lift()
    try:
        root.grab_set()
        app.secret_entry.focus_set()
        root.focus_force()
    except tk.TclError:
        pass
    root.mainloop()
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="BeatPass deterministic password generator")
    parser.add_argument("--secret", help="секрет для консольної генерації")
    parser.add_argument(
        "--difficulty",
        choices=DIFFICULTY_CLI_CHOICES,
        default=DEFAULT_DIFFICULTY,
        help="складність/довжина ключа",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.secret is None:
        return run_gui()

    try:
        print(generate_key(args.secret, args.difficulty))
    except Exception as exc:
        print(f"Помилка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
