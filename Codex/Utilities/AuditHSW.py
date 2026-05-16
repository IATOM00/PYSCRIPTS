from __future__ import annotations

import os, sys, re, platform, shutil, threading, subprocess
import ctypes, json, socket, struct, winreg
from typing import Any, Callable, Iterable
from dataclasses import dataclass, field
from queue import Empty, Queue
from ctypes import wintypes
from pathlib import Path
import datetime as dt


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
    from tkinter import messagebox, ttk
except Exception:
    tk = None
    messagebox = None
    ttk = None


REPORT_BASENAME = "Log"
ESET_TRACE_LOG = Path(
    r"C:\ProgramData\ESET\RemoteAdministrator\Agent\EraAgentApplicationData\Logs\trace.log"
)
PROFILE_NOT_FOUND_NOTE = "Профіль користувача не знайдено в ProfileList або C:\\Users."
APPX_PROGRAM_SOURCE = "Microsoft Store/Appx"
APPX_DEFAULT_OS_PUBLISHERS = {
    "microsoft corporation",
    "8wekyb3d8bbwe",
    "cw5n1h2txyewy",
}
DISPLAY_ADAPTER_CLASS_KEY = r"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
NPU_REGISTRY_ROOTS = (
    r"SYSTEM\CurrentControlSet\Enum\PCI",
    r"SYSTEM\CurrentControlSet\Enum\ACPI",
    r"SYSTEM\CurrentControlSet\Enum\HID",
    r"SYSTEM\CurrentControlSet\Enum\USB",
)
SYSTEM_ACCOUNTS = {
    "defaultaccount",
    "defaultuser0",
    "defaultuser1",
    "guest",
    "wdagutilityaccount",
    "wsiaccount",
    "codexsandboxoffline",
    "codexsandboxonline",
    "administrator",
    "dis-acc-a",
    "dis-acc-g",
}
LICENSE_STATUS_MAP = {
    0: "Unlicensed",
    1: "Licensed",
    2: "OOB Grace",
    3: "OOT Grace",
    4: "Non-Genuine Grace",
    5: "Notification",
    6: "Extended Grace",
}
BUS_TYPE_NAMES = {
    0: "Unknown",
    1: "SCSI",
    2: "ATAPI",
    3: "ATA",
    4: "1394",
    5: "SSA",
    6: "Fibre",
    7: "USB",
    8: "RAID",
    9: "iSCSI",
    10: "SAS",
    11: "SATA",
    12: "SD",
    13: "MMC",
    14: "Virtual",
    15: "File-Backed Virtual",
    16: "Storage Spaces",
    17: "NVMe",
    18: "SCM",
    19: "UFS",
}
MEMORY_TYPE_NAMES = {
    0x12: "DDR",
    0x13: "DDR2",
    0x18: "DDR3",
    0x1A: "DDR4",
    0x1B: "LPDDR",
    0x1C: "LPDDR2",
    0x1D: "LPDDR3",
    0x1E: "LPDDR4",
    0x20: "HBM",
    0x21: "HBM2",
    0x22: "DDR5",
    0x23: "LPDDR5",
}
WOW64_64KEY = getattr(winreg, "KEY_WOW64_64KEY", 0)
WOW64_32KEY = getattr(winreg, "KEY_WOW64_32KEY", 0)


@dataclass
class SMBiosRecord:
    type_id: int
    handle: int
    formatted: bytes
    strings: list[str]
    raw_strings: list[str] = field(default_factory=list)

    def get_string(self, offset: int) -> str:
        if offset >= len(self.formatted):
            return ""
        index = self.formatted[offset]
        if index <= 0 or index > len(self.strings):
            return ""
        return clean_text(self.strings[index - 1])

    def get_raw_string(self, offset: int) -> str:
        if offset >= len(self.formatted):
            return ""
        index = self.formatted[offset]
        source = self.raw_strings if self.raw_strings else self.strings
        if index <= 0 or index > len(source):
            return ""
        return clean_text(source[index - 1])

    def get_string_index_hex(self, offset: int) -> str:
        if offset >= len(self.formatted):
            return ""
        return f"{self.formatted[offset]:02X}"


@dataclass
class MemoryModule:
    size_bytes: int
    manufacturer: str = ""
    part_number: str = ""
    serial_number: str = ""
    memory_type: str = ""
    speed_mhz: int | None = None
    locator: str = ""
    bank_locator: str = ""


@dataclass
class DiskDevice:
    index: int
    path: str
    model: str
    vendor: str
    serial_number: str
    size_bytes: int
    bus_type: str
    removable: bool
    incurs_seek_penalty: bool | None
    drive_letters: list[str] = field(default_factory=list)

    @property
    def media_type(self) -> str:
        upper_model = self.model.upper()
        if self.bus_type == "NVMe":
            return "SSD (NVMe)"
        if self.bus_type == "MMC":
            return "eMMC"
        if "SSD" in upper_model or "NVME" in upper_model:
            return "SSD"
        if self.removable and self.bus_type == "USB":
            return "USB"
        if self.incurs_seek_penalty is False:
            return "SSD"
        if self.incurs_seek_penalty is True:
            return "HDD"
        return self.bus_type if self.bus_type != "Unknown" else "Unknown"


@dataclass
class GraphicsDevice:
    name: str
    vendor: str = ""
    video_processor: str = ""
    adapter_ram_bytes: int = 0
    pnp_device_id: str = ""
    driver_version: str = ""
    current_resolution: str = ""
    status: str = ""


@dataclass
class NpuDevice:
    name: str
    manufacturer: str = ""
    pnp_class: str = ""
    pnp_device_id: str = ""
    status: str = ""


@dataclass
class NetworkAdapter:
    name: str
    description: str
    mac: str
    ipv4: list[str]
    ipv6: list[str]
    gateways: list[str]
    is_up: bool
    if_type: int
    transmit_speed: int
    receive_speed: int

    @property
    def is_loopback(self) -> bool:
        return self.if_type == 24 or "loopback" in f"{self.name} {self.description}".lower()

    @property
    def is_virtual(self) -> bool:
        text = f"{self.name} {self.description}".lower()
        markers = ("virtual", "hyper-v", "vpn", "loopback", "container", "tunnel")
        return any(marker in text for marker in markers)


@dataclass
class LocalUser:
    name: str
    sid: str = ""
    enabled: bool | None = None
    last_logon: str = ""
    is_admin: bool = False

    @property
    def is_system(self) -> bool:
        return self.name.casefold() in SYSTEM_ACCOUNTS


@dataclass
class UserProfileInfo:
    username: str = ""
    path: str = ""


@dataclass
class ProgramEntry:
    name: str
    version: str = ""
    install_date: str = ""
    publisher: str = ""
    source: str = ""


@dataclass
class LicenseInfo:
    type_guess: str = "Невизначено"
    status: str = ""
    channel: str = ""
    description: str = ""
    partial_product_key: str = ""
    backup_product_key: str = ""
    oa3_product_key: str = ""
    backup_key_tail: str = ""
    oa3_key_tail: str = ""
    kms_host: str = ""
    kms_port: str = ""
    notes: list[str] = field(default_factory=list)


class AuditContext:
    def __init__(self) -> None:
        self.warnings: list[str] = []

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)


ProgressCallback = Callable[[str, str, str, int, int], None]


class AuditCancelled(Exception):
    pass


def decode_bytes_text(value: bytes) -> str:
    if not value:
        return ""

    candidates: list[str] = []
    if value.startswith((b"\xff\xfe", b"\xfe\xff")) or b"\x00" in value:
        candidates.append(value.decode("utf-16-le", errors="ignore"))
    for encoding in ("utf-8-sig", "cp1251", "latin-1"):
        try:
            candidates.append(value.decode(encoding))
        except UnicodeDecodeError:
            continue

    for candidate in candidates:
        text = candidate.replace("\x00", " ").replace("\r", " ").replace("\n", " ")
        text = re.sub(r"\s+", " ", text).strip()
        if text and any(char.isalnum() for char in text):
            return text
    return ""


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        text = decode_bytes_text(value)
    else:
        text = str(value)
    text = text.replace("\x00", " ").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def unique_preserve(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for item in items:
        text = clean_text(item)
        if not text:
            continue
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(text)
    return ordered


def safe_get(mapping: dict[str, Any], key: str, default: Any = "") -> Any:
    return mapping.get(key, default) if isinstance(mapping, dict) else default


def safe_path_exists(path: Path) -> bool:
    try:
        path.stat()
        return True
    except OSError:
        return False


def int_or_zero(value: Any) -> int:
    try:
        if value is None or value == "":
            return 0
        return int(value)
    except (TypeError, ValueError, OverflowError):
        return 0


def registry_binary_to_int(value: Any) -> int:
    if isinstance(value, bytes):
        return int.from_bytes(value, "little", signed=False)
    return int_or_zero(value)


def clean_device_description(value: Any) -> str:
    text = clean_text(value)
    if ";" in text and text.startswith("@"):
        text = text.rsplit(";", 1)[-1]
    return clean_text(text)


def read_registry_value(
    root: int,
    subkey: str,
    value_name: str,
    access: int = 0,
) -> Any | None:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | access) as key:
            value, _ = winreg.QueryValueEx(key, value_name)
            return value
    except OSError:
        return None


def read_registry_values(root: int, subkey: str, access: int = 0) -> dict[str, Any]:
    values: dict[str, Any] = {}
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | access) as key:
            index = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, index)
                except OSError:
                    break
                values[name] = value
                index += 1
    except OSError:
        return {}
    return values


def iter_registry_subkeys(root: int, subkey: str, access: int = 0) -> Iterable[str]:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | access) as key:
            index = 0
            while True:
                try:
                    yield winreg.EnumKey(key, index)
                except OSError:
                    break
                index += 1
    except OSError:
        return


def ensure_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def parse_install_date(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d{8}", text):
        try:
            parsed = dt.datetime.strptime(text, "%Y%m%d")
            return parsed.strftime("%d.%m.%Y")
        except ValueError:
            return text
    return text


def format_memory_gb(size_bytes: int | None) -> str:
    if not size_bytes or size_bytes < 0:
        return "н/д"
    value = round(size_bytes / (1024 ** 3))
    return f"{value}GB"


def format_memory_gb_spaced(size_bytes: int | None) -> str:
    if not size_bytes or size_bytes < 0:
        return "н/д"
    value = round(size_bytes / (1024 ** 3))
    return f"{value} GB"


def format_bytes_gb(size_bytes: int | None) -> str:
    if not size_bytes or size_bytes < 0:
        return "н/д"
    value = round(size_bytes / (1000 ** 3))
    return f"{value}GB"


def format_bytes_gb_spaced(size_bytes: int | None) -> str:
    if not size_bytes or size_bytes < 0:
        return "н/д"
    value = round(size_bytes / (1000 ** 3))
    return f"{value} GB"


def format_binary_capacity(size_bytes: int | None) -> str:
    if not size_bytes or size_bytes < 0:
        return "н/д"
    if size_bytes >= 1024 ** 3:
        value = size_bytes / (1024 ** 3)
        text = f"{value:.1f}".rstrip("0").rstrip(".")
        return f"{text} GB"
    if size_bytes >= 1024 ** 2:
        return f"{round(size_bytes / (1024 ** 2))} MB"
    if size_bytes >= 1024:
        return f"{round(size_bytes / 1024)} KB"
    return f"{size_bytes} B"


def format_speed_mbps(speed_bps: int) -> str:
    if not speed_bps or speed_bps > 100_000_000_000_000:
        return ""
    if speed_bps >= 1_000_000_000:
        return f"{speed_bps / 1_000_000_000:.1f} Gbps"
    if speed_bps >= 1_000_000:
        return f"{speed_bps / 1_000_000:.0f} Mbps"
    return f"{speed_bps} bps"


def format_key_tail(key: str) -> str:
    text = clean_text(key).replace(" ", "")
    if len(text) >= 5:
        return text[-5:]
    return text


def summarize_cpu_model(cpu_name: str) -> str:
    text = clean_text(cpu_name)
    patterns = [
        re.compile(r"\b(i[3579]-\d{4,5}[A-Z0-9]*)\b", re.IGNORECASE),
        re.compile(r"\b(Ultra)\s+(\d+)\s+(\d+[A-Z]{0,2})\b", re.IGNORECASE),
        re.compile(r"\b(Ryzen(?:\s+AI)?)\s+(\d+)\s+(\d+[A-Z]{0,2})\b", re.IGNORECASE),
        re.compile(r"\b(Celeron)\s+([A-Z]?\d{3,4})\b", re.IGNORECASE),
        re.compile(r"\b(Pentium)\s+([A-Z]?\d{3,4})\b", re.IGNORECASE),
    ]
    for pattern in patterns:
        match = pattern.search(text)
        if not match:
            continue
        groups = [clean_text(group) for group in match.groups()]
        if len(groups) == 1:
            return groups[0]
        return "-".join(groups)
    tail = text.split("@", 1)[0].strip()
    return tail[:48] if tail else "CPU"


def sanitize_filename_component(value: str, fallback: str = "n-d") -> str:
    text = pick_first_meaningful(value) or fallback
    text = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{index}" for index in range(1, 10)),
        *(f"LPT{index}" for index in range(1, 10)),
    }
    if not text or text.upper() in reserved_names:
        return fallback
    return text


def make_report_filename(serial_number: str) -> str:
    serial = sanitize_filename_component(serial_number)
    rights_suffix = "A" if is_process_elevated() is True else "U"
    return f"{REPORT_BASENAME}{rights_suffix} - {serial}.log"


def utc_seconds_to_local_date(value: int | None) -> str:
    if value is None:
        return ""
    try:
        dt_value = dt.datetime.fromtimestamp(int(value))
    except (OverflowError, OSError, ValueError):
        return ""
    return dt_value.strftime("%d.%m.%Y")


def normalize_date_text(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    for pattern in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            parsed = dt.datetime.strptime(text, pattern)
            return parsed.strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def pick_first_meaningful(*values: str) -> str:
    placeholders = {
        "",
        "to be filled by o.e.m.",
        "to be filled by oem",
        "not specified",
        "unknown",
        "none",
        "default string",
    }
    for value in values:
        text = clean_text(value)
        if text.casefold() in placeholders:
            continue
        if text:
            return text
    return ""


def windows_edition_label(edition_id: str) -> str:
    labels = {
        "core": "Home",
        "corecountryspecific": "Home China",
        "coren": "Home N",
        "professional": "Pro",
        "professionaln": "Pro N",
        "enterprise": "Enterprise",
        "enterprisen": "Enterprise N",
        "education": "Education",
        "educationn": "Education N",
        "serverstandard": "Server Standard",
        "serverdatacenter": "Server Datacenter",
    }
    key = clean_text(edition_id).casefold()
    return labels.get(key, clean_text(edition_id))


def normalize_windows_product_name(
    *candidates: str,
    build_number: str = "",
    edition_id: str = "",
) -> str:
    build = int_or_zero(build_number)

    for candidate in candidates:
        text = clean_text(candidate)
        if not text:
            continue
        text = text.replace("Майкрософт", "Microsoft")
        match = re.search(r"\bWindows\s+(?:\d+|Server)\b.*", text, flags=re.IGNORECASE)
        if match:
            text = clean_text(match.group(0))
        if build >= 22000:
            text = re.sub(r"\bWindows\s+10\b", "Windows 11", text, count=1, flags=re.IGNORECASE)
        if re.match(r"^Microsoft\s+Windows\b", text, flags=re.IGNORECASE):
            return "Microsoft " + re.sub(r"^Microsoft\s+", "", text, count=1, flags=re.IGNORECASE)
        if re.match(r"^Windows\b", text, flags=re.IGNORECASE):
            return f"Microsoft {text}"

    if build >= 22000:
        base = "Windows 11"
    elif build >= 10240:
        base = "Windows 10"
    else:
        base = "Windows"

    edition = windows_edition_label(edition_id)
    return f"Microsoft {base}" + (f" {edition}" if edition else "")


def normalized_path_key(value: str | Path | None) -> str:
    text = clean_text(value)
    if not text:
        return ""
    return os.path.normcase(os.path.abspath(os.path.expandvars(text))).rstrip("\\/")


def same_path(left: str | Path | None, right: str | Path | None) -> bool:
    left_key = normalized_path_key(left)
    right_key = normalized_path_key(right)
    return bool(left_key and right_key and left_key == right_key)


def is_process_elevated() -> bool | None:
    if os.name != "nt":
        return None
    try:
        advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)

        class TOKEN_ELEVATION(ctypes.Structure):
            _fields_ = [("TokenIsElevated", wintypes.DWORD)]

        token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(kernel.GetCurrentProcess(), 0x0008, ctypes.byref(token)):
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        try:
            elevation = TOKEN_ELEVATION()
            returned = wintypes.DWORD()
            ok = advapi32.GetTokenInformation(
                token,
                20,
                ctypes.byref(elevation),
                ctypes.sizeof(elevation),
                ctypes.byref(returned),
            )
            if not ok:
                return bool(ctypes.windll.shell32.IsUserAnAdmin())
            return bool(elevation.TokenIsElevated)
        finally:
            kernel.CloseHandle(token)
    except Exception:
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return None


def current_user_display_name() -> str:
    username = clean_text(os.environ.get("USERNAME") or os.environ.get("USER"))
    domain = clean_text(os.environ.get("USERDOMAIN"))
    if username and domain and "\\" not in username:
        return f"{domain}\\{username}"
    return username or "н/д"


def format_launch_context_lines() -> list[str]:
    elevated = is_process_elevated()
    if elevated is True:
        rights = "Адміністратор (elevated)"
    elif elevated is False:
        rights = "Користувацькі / без elevation"
    else:
        rights = "н/д"
    return [
        f"- Користувач запуску: {current_user_display_name()}",
        f"- Права процесу: {rights}",
    ]


def maybe_show_message(title: str, message: str, is_error: bool = False) -> None:
    frozen_without_console = getattr(sys, "frozen", False) and (
        sys.stdout is None or not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty()
    )
    if not frozen_without_console and not is_error:
        return
    if tk is None or messagebox is None:
        return
    try:
        root = tk.Tk()
        root.withdraw()
        install_frozen_executable_icon(root)
        install_dark_title_bar(root)
        root.attributes("-topmost", True)
        if is_error:
            messagebox.showerror(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)
        root.destroy()
    except Exception:
        return


def widget_exists(widget: Any | None) -> bool:
    if widget is None or tk is None:
        return False
    try:
        return bool(widget.winfo_exists())
    except (tk.TclError, RuntimeError):
        return False


def destroy_widget(widget: Any | None) -> None:
    if not widget_exists(widget):
        return
    try:
        widget.destroy()
    except (tk.TclError, RuntimeError):
        pass


def center_window(
    window: Any,
    parent: Any | None = None,
    width: int | None = None,
    height: int | None = None,
) -> None:
    if not widget_exists(window):
        return
    window.update_idletasks()

    if width is None:
        width = window.winfo_width()
        if width <= 1:
            width = window.winfo_reqwidth()

    if height is None:
        height = window.winfo_height()
        if height <= 1:
            height = window.winfo_reqheight()

    window.geometry(f"{width}x{height}")


def configure_launch_styles(root: Any) -> dict[str, str]:
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
        style.configure("LaunchMuted.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure(
            "LaunchPrimary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(18, 10),
            foreground="#FFFFFF",
            background=colors["accent_dark"],
            bordercolor=colors["accent_dark"],
            lightcolor=colors["accent_dark"],
            darkcolor=colors["accent_dark"],
            focuscolor=colors["accent_dark"],
        )
        style.map(
            "LaunchPrimary.TButton",
            background=[("pressed", "#053D39"), ("active", "#053D39")],
            foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
            bordercolor=[("pressed", "#053D39"), ("active", "#053D39")],
            lightcolor=[("pressed", "#053D39"), ("active", "#053D39")],
            darkcolor=[("pressed", "#053D39"), ("active", "#053D39")],
        )
        style.configure("ProgressHeader.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 12, "bold"))
        style.configure("ProgressBody.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
        style.configure("ProgressFile.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("ProgressCount.TLabel", background=colors["panel"], foreground=colors["accent_dark"], font=("Segoe UI", 10, "bold"))
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


class ProgressWindow:
    def __init__(self, owner: Any, on_close: Callable[[], None] | None = None):
        if tk is None or ttk is None:
            raise RuntimeError("Tkinter недоступний.")

        self.on_close = on_close
        self.close_requested = False
        self.dialog = tk.Toplevel(owner)
        self.dialog.withdraw()
        install_frozen_executable_icon(self.dialog)
        self.dialog.title("AUDIT PROCESSING")
        self.dialog.resizable(False, False)
        self.dialog.protocol("WM_DELETE_WINDOW", self.request_close)
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
            text="AuditHSW",
            bg=colors["header"],
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Аудит пристрою виконується. Будь ласка, зачекайте...",
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

        self.header_var = tk.StringVar(value="Підготовка...")
        self.detail_var = tk.StringVar(value="Будь ласка, зачекайте")
        self.file_var = tk.StringVar(value="")
        self.count_var = tk.StringVar(value="0 / 0")

        ttk.Label(frame, textvariable=self.header_var, style="ProgressHeader.TLabel").grid(
            row=0, column=0, padx=(0, 12), sticky="w"
        )
        ttk.Label(frame, textvariable=self.count_var, style="ProgressCount.TLabel").grid(
            row=0, column=1, sticky="e"
        )
        ttk.Label(frame, textvariable=self.detail_var, style="ProgressBody.TLabel", wraplength=500).grid(
            row=1, column=0, columnspan=2, pady=(8, 12), sticky="w"
        )

        self.progress = ttk.Progressbar(
            frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            style="Launch.Horizontal.TProgressbar",
        )
        self.progress.grid(row=2, column=0, columnspan=2, sticky="we")
        center_window(self.dialog, parent=owner, width=560, height=232)
        self._show_in_front()
        self.refresh()

    def _show_in_front(self) -> None:
        try:
            self.dialog.deiconify()
            self.dialog.lift()
            self.dialog.attributes("-topmost", True)
            self.dialog.update_idletasks()
            self.dialog.focus_force()
        except Exception:
            pass

    def update(
        self,
        header: str | None = None,
        detail: str | None = None,
        file_name: str | None = None,
        current: int | None = None,
        total: int | None = None,
    ) -> None:
        try:
            if not self.dialog.winfo_exists():
                return
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
        except Exception:
            return

    def refresh(self) -> None:
        try:
            self.dialog.update_idletasks()
        except Exception:
            pass

    def close(self) -> None:
        destroy_widget(self.dialog)

    def request_close(self) -> None:
        if self.close_requested:
            return
        self.close_requested = True
        try:
            if self.on_close is not None:
                self.on_close()
            else:
                self.close()
        except Exception:
            self.close()


class AuditStartWindow:
    def __init__(
        self,
        owner: Any,
        on_start: Callable[[], None] | None = None,
        on_close: Callable[[], None] | None = None,
    ):
        if tk is None or ttk is None:
            raise RuntimeError("Tkinter недоступний.")

        self.on_start = on_start
        self.on_close = on_close
        self.action_requested = False
        self.dialog = tk.Toplevel(owner)
        self.dialog.withdraw()
        install_frozen_executable_icon(self.dialog)
        self.dialog.title("AUDIT - AuditHSW")
        self.dialog.resizable(False, False)
        self.dialog.protocol("WM_DELETE_WINDOW", self.request_close)
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
            text="AuditHSW",
            bg=colors["header"],
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Збір даних з ПК, під керуванням Windows...",
            bg=colors["header"],
            fg="#D7FBF5",
            font=("Segoe UI", 9),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(2, 4))
        badge_shell = tk.Frame(header, bg=colors["accent_dark"], padx=1, pady=1)
        badge = tk.Label(
            badge_shell,
            text="АУДИТ WINDOWS",
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
        frame = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 14))
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        content_frame = ttk.Frame(frame, style="LaunchPanel.TFrame")
        content_frame.grid(row=0, column=0, sticky="nsew")
        content_frame.grid_columnconfigure(0, weight=1)

        ttk.Label(content_frame, text="Підготовка аудиту ПК", style="LaunchSection.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            content_frame,
            text=(
                "Після натискання утиліта розпочне збирати системні відомості, диски, мережеві параметри, "
                "локальних користувачів, встановлені програми та стан ліцензії Windows, а потім "
                "сформує текстовий звіт у поточній директорії."
            ),
            style="ProgressBody.TLabel",
            wraplength=570,
        ).grid(row=1, column=0, pady=(8, 5), sticky="w")

        footer_frame = ttk.Frame(frame, style="LaunchPanel.TFrame")
        footer_frame.grid(row=1, column=0, sticky="nsew")
        footer_frame.grid_columnconfigure(0, weight=1)
        footer_frame.grid_rowconfigure(1, weight=1)

        ttk.Separator(footer_frame, orient="horizontal").grid(row=0, column=0, pady=(7, 5), sticky="ew")

        self.start_button = ttk.Button(
            footer_frame,
            text="ПОЧАТИ АУДИТ WINDOWS",
            style="LaunchPrimary.TButton",
            command=self.request_start,
            padding=(18, 18),
        )
        self.start_button.grid(row=2, column=0, sticky="ew")

        self.dialog.bind("<Return>", lambda _event: self.request_start())
        self.dialog.bind("<Escape>", lambda _event: self.request_close())

        self.dialog.update_idletasks()
        launch_width = max(620, self.dialog.winfo_reqwidth())
        launch_height = max(320, self.dialog.winfo_reqheight())
        center_window(self.dialog, parent=owner, width=launch_width, height=launch_height)
        self._show_in_front()
        self.refresh()

    def _show_in_front(self) -> None:
        try:
            self.dialog.deiconify()
            self.dialog.lift()
            self.dialog.attributes("-topmost", True)
            self.dialog.update_idletasks()
            self.start_button.focus_force()
        except Exception:
            pass

    def refresh(self) -> None:
        try:
            self.dialog.update_idletasks()
        except Exception:
            pass

    def close(self) -> None:
        destroy_widget(self.dialog)

    def request_start(self) -> None:
        if self.action_requested:
            return
        self.action_requested = True
        try:
            if self.on_start is not None:
                self.on_start()
            else:
                self.close()
        except Exception:
            self.close()

    def request_close(self) -> None:
        if self.action_requested:
            return
        self.action_requested = True
        try:
            if self.on_close is not None:
                self.on_close()
            else:
                self.close()
        except Exception:
            self.close()


def find_powershell() -> str:
    for candidate in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        path = shutil.which(candidate)
        if path:
            return path
    return "powershell.exe"


def hidden_subprocess_kwargs() -> dict[str, Any]:
    if os.name != "nt":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    # CreateProcess inherits the current process token, so an elevated exe
    # keeps these helper tools elevated; CREATE_NO_WINDOW only hides their UI.
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


def run_hidden_subprocess(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        **hidden_subprocess_kwargs(),
        **kwargs,
    )


def run_powershell(
    script: str,
    ctx: AuditContext,
    purpose: str,
    timeout: int = 20,
) -> subprocess.CompletedProcess[str] | None:
    command = [
        find_powershell(),
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]
    try:
        completed = run_hidden_subprocess(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except Exception as exc:
        ctx.warn(f"{purpose}: PowerShell не запустився: {exc}")
        return None
    if completed.returncode != 0:
        stderr = clean_text(completed.stderr) or clean_text(completed.stdout) or "невідома помилка"
        ctx.warn(f"{purpose}: PowerShell повернув помилку: {stderr}")
    return completed


def powershell_json(script: str, ctx: AuditContext, purpose: str, timeout: int = 20) -> Any | None:
    completed = run_powershell(script, ctx, purpose, timeout=timeout)
    if completed is None:
        return None
    stdout = completed.stdout.strip()
    if completed.returncode != 0 or not stdout:
        return None
    try:
        return json.loads(stdout)
    except json.JSONDecodeError as exc:
        ctx.warn(f"{purpose}: не вдалося розібрати JSON: {exc}")
        return None


def parse_smbios_records(raw_data: bytes, ctx: AuditContext) -> list[SMBiosRecord]:
    if len(raw_data) < 8:
        ctx.warn("SMBIOS: значення SMBiosData занадто коротке.")
        return []
    length = struct.unpack_from("<I", raw_data, 4)[0]
    table = raw_data[8 : 8 + length]
    records: list[SMBiosRecord] = []
    index = 0
    while index + 4 <= len(table):
        type_id = table[index]
        record_length = table[index + 1]
        handle = struct.unpack_from("<H", table, index + 2)[0]
        if record_length < 4 or index + record_length > len(table):
            break
        formatted = table[index : index + record_length]
        string_area_start = index + record_length
        strings: list[str] = []
        raw_strings: list[str] = []
        cursor = string_area_start
        string_start = string_area_start
        while cursor < len(table):
            if table[cursor] == 0 and cursor + 1 < len(table) and table[cursor + 1] == 0:
                if cursor > string_start:
                    raw_string = table[string_start:cursor].decode("latin-1", errors="ignore")
                    raw_strings.append(raw_string)
                    strings.append(clean_text(raw_string))
                cursor += 2
                break
            if table[cursor] == 0:
                raw_string = table[string_start:cursor].decode("latin-1", errors="ignore")
                raw_strings.append(raw_string)
                strings.append(clean_text(raw_string))
                cursor += 1
                string_start = cursor
                continue
            cursor += 1
        records.append(SMBiosRecord(type_id, handle, formatted, strings, raw_strings))
        index = cursor
        if type_id == 127:
            break
    return records


def load_smbios_records(ctx: AuditContext) -> list[SMBiosRecord]:
    data = read_registry_value(
        winreg.HKEY_LOCAL_MACHINE,
        r"SYSTEM\CurrentControlSet\Services\mssmbios\Data",
        "SMBiosData",
    )
    if not isinstance(data, bytes):
        ctx.warn("SMBIOS: не вдалося прочитати SMBiosData з реєстру.")
        return []
    return parse_smbios_records(data, ctx)


def load_bios_registry() -> dict[str, Any]:
    return read_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        r"HARDWARE\DESCRIPTION\System\BIOS",
    )


def load_cpu_registry() -> dict[str, Any]:
    return read_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        r"HARDWARE\DESCRIPTION\System\CentralProcessor\0",
    )


def extract_system_info(smbios_records: list[SMBiosRecord]) -> dict[str, str]:
    bios_values = load_bios_registry()
    type1 = next((record for record in smbios_records if record.type_id == 1), None)
    type2 = next((record for record in smbios_records if record.type_id == 2), None)
    type3 = next((record for record in smbios_records if record.type_id == 3), None)

    manufacturer = pick_first_meaningful(
        safe_get(bios_values, "SystemManufacturer"),
        type1.get_string(0x04) if type1 else "",
        safe_get(bios_values, "BaseBoardManufacturer"),
        type2.get_string(0x04) if type2 else "",
    )
    product_name = pick_first_meaningful(
        safe_get(bios_values, "SystemProductName"),
        type1.get_string(0x05) if type1 else "",
    )
    system_version = pick_first_meaningful(
        safe_get(bios_values, "SystemVersion"),
        type1.get_string(0x06) if type1 else "",
    )
    system_family = pick_first_meaningful(
        safe_get(bios_values, "SystemFamily"),
        type1.get_string(0x1A) if type1 and len(type1.formatted) > 0x1A else "",
    )
    system_sku = pick_first_meaningful(
        safe_get(bios_values, "SystemSKU"),
        type1.get_string(0x19) if type1 and len(type1.formatted) > 0x19 else "",
    )
    serial_number = pick_first_meaningful(
        type1.get_string(0x07) if type1 else "",
        type2.get_string(0x07) if type2 else "",
        type3.get_string(0x07) if type3 else "",
    )
    return {
        "manufacturer": manufacturer,
        "product_name": product_name,
        "system_version": system_version,
        "system_family": system_family,
        "system_sku": system_sku,
        "serial_number": serial_number,
        "bios_version": clean_text(safe_get(bios_values, "BIOSVersion")),
        "bios_date": normalize_date_text(clean_text(safe_get(bios_values, "BIOSReleaseDate"))),
        "baseboard_product": clean_text(safe_get(bios_values, "BaseBoardProduct")),
    }


def extract_processor_info(smbios_records: list[SMBiosRecord]) -> dict[str, str]:
    cpu_values = load_cpu_registry()
    type4 = next((record for record in smbios_records if record.type_id == 4), None)
    processor_id = ""
    if type4 and len(type4.formatted) >= 0x10:
        raw_processor_id = type4.formatted[0x08:0x10]
        processor_id = raw_processor_id.hex().upper()
    serial_string = type4.get_string(0x20) if type4 else ""
    part_number = type4.get_string(0x22) if type4 else ""
    return {
        "architecture": clean_text(os.environ.get("PROCESSOR_ARCHITECTURE") or platform.machine()),
        "name": pick_first_meaningful(
            safe_get(cpu_values, "ProcessorNameString"),
            type4.get_string(0x10) if type4 else "",
        ),
        "identifier": clean_text(safe_get(cpu_values, "Identifier")),
        "vendor": clean_text(safe_get(cpu_values, "VendorIdentifier")),
        "serial_string": clean_text(serial_string),
        "part_number": clean_text(part_number),
        "processor_id": processor_id,
    }


def decode_memory_size(record: SMBiosRecord) -> int:
    if len(record.formatted) < 0x0E:
        return 0
    size_raw = struct.unpack_from("<H", record.formatted, 0x0C)[0]
    if size_raw in (0, 0xFFFF):
        return 0
    if size_raw == 0x7FFF and len(record.formatted) >= 0x20:
        extended_mb = struct.unpack_from("<I", record.formatted, 0x1C)[0]
        return extended_mb * 1024 * 1024
    if size_raw & 0x8000:
        return (size_raw & 0x7FFF) * 1024
    return size_raw * 1024 * 1024


def is_empty_or_zero_serial(value: str) -> bool:
    text = clean_text(value)
    return not text or bool(re.fullmatch(r"0+", text))


def memory_module_serial_number(record: SMBiosRecord) -> str:
    serial = record.get_string(0x18)
    if not is_empty_or_zero_serial(serial):
        return serial

    raw_serial = record.get_raw_string(0x18)
    if raw_serial:
        return raw_serial

    return record.get_string_index_hex(0x18)


def extract_memory_modules(smbios_records: list[SMBiosRecord]) -> list[MemoryModule]:
    modules: list[MemoryModule] = []
    for record in smbios_records:
        if record.type_id != 17:
            continue
        size_bytes = decode_memory_size(record)
        if size_bytes <= 0:
            continue
        memory_type_code = record.formatted[0x12] if len(record.formatted) > 0x12 else 0
        speed_mhz = struct.unpack_from("<H", record.formatted, 0x15)[0] if len(record.formatted) > 0x17 else 0
        modules.append(
            MemoryModule(
                size_bytes=size_bytes,
                manufacturer=record.get_string(0x17),
                part_number=record.get_string(0x1A),
                serial_number=memory_module_serial_number(record),
                memory_type=MEMORY_TYPE_NAMES.get(memory_type_code, f"Type {memory_type_code}" if memory_type_code else ""),
                speed_mhz=speed_mhz or None,
                locator=record.get_string(0x10),
                bank_locator=record.get_string(0x11),
            )
        )
    return modules


def get_total_physical_memory() -> int:
    class MEMORYSTATUSEX(ctypes.Structure):
        _fields_ = [
            ("dwLength", wintypes.DWORD),
            ("dwMemoryLoad", wintypes.DWORD),
            ("ullTotalPhys", ctypes.c_ulonglong),
            ("ullAvailPhys", ctypes.c_ulonglong),
            ("ullTotalPageFile", ctypes.c_ulonglong),
            ("ullAvailPageFile", ctypes.c_ulonglong),
            ("ullTotalVirtual", ctypes.c_ulonglong),
            ("ullAvailVirtual", ctypes.c_ulonglong),
            ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
        ]

    memory_status = MEMORYSTATUSEX()
    memory_status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
    if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory_status)):
        return int(memory_status.ullTotalPhys)
    return 0


kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
OPEN_EXISTING = 3
FILE_SHARE_READ = 1
FILE_SHARE_WRITE = 2
IOCTL_STORAGE_QUERY_PROPERTY = 0x002D1400
IOCTL_DISK_GET_DRIVE_GEOMETRY_EX = 0x000700A0
IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x002D1080
StorageDeviceProperty = 0
StorageStandardQuery = 0
StorageDeviceSeekPenaltyProperty = 7


class STORAGE_PROPERTY_QUERY(ctypes.Structure):
    _fields_ = [
        ("PropertyId", wintypes.DWORD),
        ("QueryType", wintypes.DWORD),
        ("AdditionalParameters", ctypes.c_byte * 1),
    ]


class STORAGE_DEVICE_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ("Version", wintypes.DWORD),
        ("Size", wintypes.DWORD),
        ("DeviceType", ctypes.c_byte),
        ("DeviceTypeModifier", ctypes.c_byte),
        ("RemovableMedia", ctypes.c_byte),
        ("CommandQueueing", ctypes.c_byte),
        ("VendorIdOffset", wintypes.DWORD),
        ("ProductIdOffset", wintypes.DWORD),
        ("ProductRevisionOffset", wintypes.DWORD),
        ("SerialNumberOffset", wintypes.DWORD),
        ("BusType", ctypes.c_byte),
        ("RawPropertiesLength", wintypes.DWORD),
    ]


class DISK_GEOMETRY(ctypes.Structure):
    _fields_ = [
        ("Cylinders", ctypes.c_longlong),
        ("MediaType", wintypes.DWORD),
        ("TracksPerCylinder", wintypes.DWORD),
        ("SectorsPerTrack", wintypes.DWORD),
        ("BytesPerSector", wintypes.DWORD),
    ]


class DISK_GEOMETRY_EX(ctypes.Structure):
    _fields_ = [
        ("Geometry", DISK_GEOMETRY),
        ("DiskSize", ctypes.c_longlong),
        ("Data", ctypes.c_byte * 1),
    ]


class DEVICE_SEEK_PENALTY_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ("Version", wintypes.DWORD),
        ("Size", wintypes.DWORD),
        ("IncursSeekPenalty", wintypes.BOOLEAN),
    ]


class STORAGE_DEVICE_NUMBER(ctypes.Structure):
    _fields_ = [
        ("DeviceType", wintypes.DWORD),
        ("DeviceNumber", wintypes.DWORD),
        ("PartitionNumber", wintypes.DWORD),
    ]


CreateFileW = kernel32.CreateFileW
CreateFileW.argtypes = [
    wintypes.LPCWSTR,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.LPVOID,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.HANDLE,
]
CreateFileW.restype = wintypes.HANDLE
DeviceIoControl = kernel32.DeviceIoControl
DeviceIoControl.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPVOID,
    wintypes.DWORD,
    wintypes.LPVOID,
    wintypes.DWORD,
    ctypes.POINTER(wintypes.DWORD),
    wintypes.LPVOID,
]
DeviceIoControl.restype = wintypes.BOOL
CloseHandle = kernel32.CloseHandle


def open_device(path: str) -> wintypes.HANDLE | None:
    handle = CreateFileW(path, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, None, OPEN_EXISTING, 0, None)
    if handle == INVALID_HANDLE_VALUE:
        return None
    return handle


def read_null_terminated_ascii(buffer: bytes, offset: int) -> str:
    if not offset or offset >= len(buffer):
        return ""
    end = buffer.find(b"\x00", offset)
    if end == -1:
        end = len(buffer)
    return clean_text(buffer[offset:end].decode("ascii", errors="ignore"))


def query_disk_descriptor(handle: wintypes.HANDLE) -> dict[str, Any]:
    query = STORAGE_PROPERTY_QUERY(StorageDeviceProperty, StorageStandardQuery, (0,))
    returned = wintypes.DWORD()
    buffer = ctypes.create_string_buffer(1024)
    ok = DeviceIoControl(
        handle,
        IOCTL_STORAGE_QUERY_PROPERTY,
        ctypes.byref(query),
        ctypes.sizeof(query),
        buffer,
        ctypes.sizeof(buffer),
        ctypes.byref(returned),
        None,
    )
    if not ok:
        raise OSError(ctypes.get_last_error(), "IOCTL_STORAGE_QUERY_PROPERTY failed")
    descriptor = STORAGE_DEVICE_DESCRIPTOR.from_buffer_copy(
        buffer.raw[: ctypes.sizeof(STORAGE_DEVICE_DESCRIPTOR)]
    )
    return {
        "vendor": read_null_terminated_ascii(buffer.raw, descriptor.VendorIdOffset),
        "model": read_null_terminated_ascii(buffer.raw, descriptor.ProductIdOffset),
        "serial": read_null_terminated_ascii(buffer.raw, descriptor.SerialNumberOffset).rstrip("."),
        "bus_type": BUS_TYPE_NAMES.get(int(descriptor.BusType), f"Bus {descriptor.BusType}"),
        "removable": bool(descriptor.RemovableMedia),
    }


def query_disk_size(handle: wintypes.HANDLE) -> int:
    returned = wintypes.DWORD()
    buffer = ctypes.create_string_buffer(256)
    ok = DeviceIoControl(
        handle,
        IOCTL_DISK_GET_DRIVE_GEOMETRY_EX,
        None,
        0,
        buffer,
        ctypes.sizeof(buffer),
        ctypes.byref(returned),
        None,
    )
    if not ok:
        raise OSError(ctypes.get_last_error(), "IOCTL_DISK_GET_DRIVE_GEOMETRY_EX failed")
    geometry = DISK_GEOMETRY_EX.from_buffer_copy(buffer.raw[: ctypes.sizeof(DISK_GEOMETRY_EX)])
    return int(geometry.DiskSize)


def query_seek_penalty(handle: wintypes.HANDLE) -> bool | None:
    query = STORAGE_PROPERTY_QUERY(StorageDeviceSeekPenaltyProperty, StorageStandardQuery, (0,))
    descriptor = DEVICE_SEEK_PENALTY_DESCRIPTOR()
    returned = wintypes.DWORD()
    ok = DeviceIoControl(
        handle,
        IOCTL_STORAGE_QUERY_PROPERTY,
        ctypes.byref(query),
        ctypes.sizeof(query),
        ctypes.byref(descriptor),
        ctypes.sizeof(descriptor),
        ctypes.byref(returned),
        None,
    )
    if not ok:
        return None
    return bool(descriptor.IncursSeekPenalty)


def get_drive_letter_map(ctx: AuditContext) -> dict[int, list[str]]:
    mapping: dict[int, list[str]] = {}
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        volume_path = rf"\\.\{letter}:"
        handle = open_device(volume_path)
        if handle is None:
            continue
        try:
            device_number = STORAGE_DEVICE_NUMBER()
            returned = wintypes.DWORD()
            ok = DeviceIoControl(
                handle,
                IOCTL_STORAGE_GET_DEVICE_NUMBER,
                None,
                0,
                ctypes.byref(device_number),
                ctypes.sizeof(device_number),
                ctypes.byref(returned),
                None,
            )
            if ok:
                mapping.setdefault(int(device_number.DeviceNumber), []).append(f"{letter}:")
        except Exception as exc:
            ctx.warn(f"Том {letter}: не вдалося зіставити з фізичним диском: {exc}")
        finally:
            CloseHandle(handle)
    for drive_letters in mapping.values():
        drive_letters.sort()
    return mapping


def collect_disks(ctx: AuditContext) -> list[DiskDevice]:
    drive_letter_map = get_drive_letter_map(ctx)
    disks: list[DiskDevice] = []
    for index in range(0, 32):
        path = rf"\\.\PhysicalDrive{index}"
        handle = open_device(path)
        if handle is None:
            continue
        try:
            descriptor = query_disk_descriptor(handle)
            size_bytes = query_disk_size(handle)
            seeks = query_seek_penalty(handle)
            disks.append(
                DiskDevice(
                    index=index,
                    path=path,
                    model=pick_first_meaningful(descriptor["model"], descriptor["vendor"], path),
                    vendor=descriptor["vendor"],
                    serial_number=descriptor["serial"],
                    size_bytes=size_bytes,
                    bus_type=descriptor["bus_type"],
                    removable=descriptor["removable"],
                    incurs_seek_penalty=seeks,
                    drive_letters=drive_letter_map.get(index, []),
                )
            )
        except Exception as exc:
            ctx.warn(f"Disk {index}: не вдалося повністю зчитати дані: {exc}")
        finally:
            CloseHandle(handle)
    return disks


iphlpapi = ctypes.WinDLL("iphlpapi", use_last_error=True)
AF_UNSPEC = 0
AF_INET = socket.AF_INET
AF_INET6 = socket.AF_INET6
ERROR_BUFFER_OVERFLOW = 111
NO_ERROR = 0
GAA_FLAG_SKIP_ANYCAST = 0x0002
GAA_FLAG_SKIP_MULTICAST = 0x0004
GAA_FLAG_SKIP_DNS_SERVER = 0x0008
GAA_FLAG_INCLUDE_GATEWAYS = 0x0080


class SOCKET_ADDRESS(ctypes.Structure):
    _fields_ = [
        ("lpSockaddr", ctypes.c_void_p),
        ("iSockaddrLength", ctypes.c_int),
    ]


class SOCKADDR(ctypes.Structure):
    _fields_ = [
        ("sa_family", wintypes.USHORT),
        ("sa_data", ctypes.c_ubyte * 14),
    ]


class IN_ADDR(ctypes.Structure):
    _fields_ = [("S_addr", ctypes.c_ubyte * 4)]


class SOCKADDR_IN(ctypes.Structure):
    _fields_ = [
        ("sin_family", wintypes.USHORT),
        ("sin_port", wintypes.USHORT),
        ("sin_addr", IN_ADDR),
        ("sin_zero", ctypes.c_ubyte * 8),
    ]


class IN6_ADDR(ctypes.Structure):
    _fields_ = [("Byte", ctypes.c_ubyte * 16)]


class SOCKADDR_IN6(ctypes.Structure):
    _fields_ = [
        ("sin6_family", wintypes.USHORT),
        ("sin6_port", wintypes.USHORT),
        ("sin6_flowinfo", wintypes.ULONG),
        ("sin6_addr", IN6_ADDR),
        ("sin6_scope_id", wintypes.ULONG),
    ]


class IP_ADAPTER_UNICAST_ADDRESS(ctypes.Structure):
    pass


class IP_ADAPTER_GATEWAY_ADDRESS(ctypes.Structure):
    pass


PIP_ADAPTER_UNICAST_ADDRESS = ctypes.POINTER(IP_ADAPTER_UNICAST_ADDRESS)
PIP_ADAPTER_GATEWAY_ADDRESS = ctypes.POINTER(IP_ADAPTER_GATEWAY_ADDRESS)

IP_ADAPTER_UNICAST_ADDRESS._fields_ = [
    ("Length", wintypes.ULONG),
    ("Flags", wintypes.DWORD),
    ("Next", PIP_ADAPTER_UNICAST_ADDRESS),
    ("Address", SOCKET_ADDRESS),
    ("PrefixOrigin", wintypes.ULONG),
    ("SuffixOrigin", wintypes.ULONG),
    ("DadState", wintypes.ULONG),
    ("ValidLifetime", wintypes.ULONG),
    ("PreferredLifetime", wintypes.ULONG),
    ("LeaseLifetime", wintypes.ULONG),
    ("OnLinkPrefixLength", ctypes.c_ubyte),
]

IP_ADAPTER_GATEWAY_ADDRESS._fields_ = [
    ("Length", wintypes.ULONG),
    ("Reserved", wintypes.DWORD),
    ("Next", PIP_ADAPTER_GATEWAY_ADDRESS),
    ("Address", SOCKET_ADDRESS),
]


class IP_ADAPTER_ADDRESSES(ctypes.Structure):
    pass


PIP_ADAPTER_ADDRESSES = ctypes.POINTER(IP_ADAPTER_ADDRESSES)

IP_ADAPTER_ADDRESSES._fields_ = [
    ("Length", wintypes.ULONG),
    ("IfIndex", wintypes.DWORD),
    ("Next", PIP_ADAPTER_ADDRESSES),
    ("AdapterName", ctypes.c_char_p),
    ("FirstUnicastAddress", PIP_ADAPTER_UNICAST_ADDRESS),
    ("FirstAnycastAddress", ctypes.c_void_p),
    ("FirstMulticastAddress", ctypes.c_void_p),
    ("FirstDnsServerAddress", ctypes.c_void_p),
    ("DnsSuffix", wintypes.LPWSTR),
    ("Description", wintypes.LPWSTR),
    ("FriendlyName", wintypes.LPWSTR),
    ("PhysicalAddress", ctypes.c_ubyte * 8),
    ("PhysicalAddressLength", wintypes.DWORD),
    ("Flags", wintypes.DWORD),
    ("Mtu", wintypes.DWORD),
    ("IfType", wintypes.DWORD),
    ("OperStatus", wintypes.DWORD),
    ("Ipv6IfIndex", wintypes.DWORD),
    ("ZoneIndices", wintypes.DWORD * 16),
    ("FirstPrefix", ctypes.c_void_p),
    ("TransmitLinkSpeed", ctypes.c_ulonglong),
    ("ReceiveLinkSpeed", ctypes.c_ulonglong),
    ("FirstWinsServerAddress", ctypes.c_void_p),
    ("FirstGatewayAddress", PIP_ADAPTER_GATEWAY_ADDRESS),
    ("Ipv4Metric", wintypes.ULONG),
    ("Ipv6Metric", wintypes.ULONG),
    ("Luid", ctypes.c_ulonglong),
    ("Dhcpv4Server", SOCKET_ADDRESS),
    ("CompartmentId", wintypes.ULONG),
    ("NetworkGuid", ctypes.c_ubyte * 16),
    ("ConnectionType", wintypes.ULONG),
    ("TunnelType", wintypes.ULONG),
    ("Dhcpv6Server", SOCKET_ADDRESS),
    ("Dhcpv6ClientDuid", ctypes.c_ubyte * 130),
    ("Dhcpv6ClientDuidLength", wintypes.ULONG),
    ("Dhcpv6Iaid", wintypes.ULONG),
    ("FirstDnsSuffix", ctypes.c_void_p),
]


GetAdaptersAddresses = iphlpapi.GetAdaptersAddresses
GetAdaptersAddresses.argtypes = [
    wintypes.ULONG,
    wintypes.ULONG,
    ctypes.c_void_p,
    PIP_ADAPTER_ADDRESSES,
    ctypes.POINTER(wintypes.ULONG),
]
GetAdaptersAddresses.restype = wintypes.ULONG


def socket_address_to_text(address: SOCKET_ADDRESS) -> str:
    if not address.lpSockaddr or address.iSockaddrLength <= 0:
        return ""
    family = ctypes.cast(address.lpSockaddr, ctypes.POINTER(SOCKADDR)).contents.sa_family
    if family == AF_INET:
        sockaddr = ctypes.cast(address.lpSockaddr, ctypes.POINTER(SOCKADDR_IN)).contents
        return socket.inet_ntop(AF_INET, bytes(sockaddr.sin_addr.S_addr))
    if family == AF_INET6:
        sockaddr = ctypes.cast(address.lpSockaddr, ctypes.POINTER(SOCKADDR_IN6)).contents
        return socket.inet_ntop(AF_INET6, bytes(sockaddr.sin6_addr.Byte))
    return ""


def iter_pointer_chain(pointer: Any, attr_name: str) -> Iterable[Any]:
    current = pointer
    while current:
        obj = current.contents
        yield obj
        current = getattr(obj, attr_name)


def collect_network_adapters(ctx: AuditContext) -> list[NetworkAdapter]:
    size = wintypes.ULONG(15_000)
    buffer = ctypes.create_string_buffer(size.value)
    flags = (
        GAA_FLAG_SKIP_ANYCAST
        | GAA_FLAG_SKIP_MULTICAST
        | GAA_FLAG_SKIP_DNS_SERVER
        | GAA_FLAG_INCLUDE_GATEWAYS
    )
    result = GetAdaptersAddresses(AF_UNSPEC, flags, None, ctypes.cast(buffer, PIP_ADAPTER_ADDRESSES), ctypes.byref(size))
    if result == ERROR_BUFFER_OVERFLOW:
        buffer = ctypes.create_string_buffer(size.value)
        result = GetAdaptersAddresses(AF_UNSPEC, flags, None, ctypes.cast(buffer, PIP_ADAPTER_ADDRESSES), ctypes.byref(size))
    if result != NO_ERROR:
        ctx.warn(f"Мережа: GetAdaptersAddresses завершився кодом {result}.")
        return []

    adapters: list[NetworkAdapter] = []
    current = ctypes.cast(buffer, PIP_ADAPTER_ADDRESSES)
    while current:
        adapter = current.contents
        mac_length = int(adapter.PhysicalAddressLength)
        mac = "-".join(f"{adapter.PhysicalAddress[i]:02X}" for i in range(mac_length)) if mac_length else ""
        unicast_values = list(iter_pointer_chain(adapter.FirstUnicastAddress, "Next"))
        ipv4 = unique_preserve(
            socket_address_to_text(item.Address)
            for item in unicast_values
            if socket_address_to_text(item.Address) and "." in socket_address_to_text(item.Address)
        )
        ipv6 = unique_preserve(
            socket_address_to_text(item.Address)
            for item in unicast_values
            if ":" in socket_address_to_text(item.Address)
            and not socket_address_to_text(item.Address).startswith("fe80:")
        )
        gateways = unique_preserve(
            socket_address_to_text(item.Address)
            for item in iter_pointer_chain(adapter.FirstGatewayAddress, "Next")
            if socket_address_to_text(item.Address)
        )
        network_adapter = NetworkAdapter(
            name=clean_text(adapter.FriendlyName),
            description=clean_text(adapter.Description),
            mac=mac,
            ipv4=ipv4,
            ipv6=ipv6,
            gateways=gateways,
            is_up=int(adapter.OperStatus) == 1,
            if_type=int(adapter.IfType),
            transmit_speed=int(adapter.TransmitLinkSpeed),
            receive_speed=int(adapter.ReceiveLinkSpeed),
        )
        if network_adapter.mac and not network_adapter.is_loopback:
            adapters.append(network_adapter)
        current = adapter.Next

    adapters.sort(
        key=lambda item: (
            0 if item.is_up else 1,
            0 if item.gateways and not item.is_virtual else 1,
            0 if not item.is_virtual else 1,
            item.name.casefold(),
        )
    )
    return adapters


def collect_graphics_from_registry(ctx: AuditContext) -> list[GraphicsDevice]:
    devices: list[GraphicsDevice] = []
    for child in iter_registry_subkeys(winreg.HKEY_LOCAL_MACHINE, DISPLAY_ADAPTER_CLASS_KEY, WOW64_64KEY):
        if not re.fullmatch(r"\d{4}", child):
            continue
        values = read_registry_values(
            winreg.HKEY_LOCAL_MACHINE,
            fr"{DISPLAY_ADAPTER_CLASS_KEY}\{child}",
            WOW64_64KEY,
        )
        if not values:
            continue

        name = pick_first_meaningful(
            clean_device_description(values.get("HardwareInformation.AdapterString")),
            clean_device_description(values.get("DriverDesc")),
            clean_device_description(values.get("DeviceDesc")),
            clean_device_description(values.get("DriverTargetSegment")),
        )
        if not name:
            continue

        memory_value = values.get("HardwareInformation.qwMemorySize")
        if memory_value is None and "intel" not in name.casefold():
            memory_value = values.get("HardwareInformation.MemorySize")

        devices.append(
            GraphicsDevice(
                name=name,
                vendor=pick_first_meaningful(
                    clean_text(values.get("ProviderName")),
                    clean_text(values.get("Manufacturer")),
                ),
                video_processor=clean_device_description(values.get("HardwareInformation.ChipType")),
                adapter_ram_bytes=registry_binary_to_int(memory_value),
                pnp_device_id=clean_text(values.get("MatchingDeviceId") or values.get("Driver")),
                driver_version=clean_text(values.get("DriverVersion")),
                status=registry_status_from_values(values),
            )
        )
    return devices


def registry_status_from_values(values: dict[str, Any]) -> str:
    problem = registry_binary_to_int(values.get("Problem"))
    config_flags = registry_binary_to_int(values.get("ConfigFlags"))
    if problem:
        return f"Problem {problem}"
    if config_flags:
        return f"ConfigFlags {config_flags}"
    return "OK"


def dedupe_npu_devices(npus: list[NpuDevice]) -> list[NpuDevice]:
    deduped: dict[tuple[str, str, str], NpuDevice] = {}
    for npu in npus:
        key = (npu.name.casefold(), npu.manufacturer.casefold(), npu.pnp_class.casefold())
        if not key[0]:
            continue
        existing = deduped.get(key)
        if existing is None:
            deduped[key] = npu
            continue
        if not existing.status and npu.status:
            existing.status = npu.status
        if not existing.pnp_device_id and npu.pnp_device_id:
            existing.pnp_device_id = npu.pnp_device_id
    return list(deduped.values())


def registry_value_search_text(values: dict[str, Any]) -> str:
    parts: list[str] = []
    for value in values.values():
        if isinstance(value, (list, tuple)):
            parts.extend(clean_device_description(item) for item in value)
        else:
            parts.append(clean_device_description(value))
    return " ".join(part for part in parts if part)


def collect_npu_from_registry(ctx: AuditContext) -> list[NpuDevice]:
    markers = re.compile(
        r"(\bnpu\b|neural|ai boost|ryzen ai|hexagon|inference|\bvpu\b|\bipu\b|input configuration device)",
        re.IGNORECASE,
    )
    found: list[NpuDevice] = []
    seen: set[tuple[str, str]] = set()

    def scan_key(path: str, depth: int) -> None:
        values = read_registry_values(winreg.HKEY_LOCAL_MACHINE, path, WOW64_64KEY)
        if values:
            search_text = registry_value_search_text(values)
            service = clean_text(values.get("Service"))
            class_name = clean_text(values.get("Class"))
            if markers.search(search_text) or service.casefold() == "npu":
                name = pick_first_meaningful(
                    clean_device_description(values.get("FriendlyName")),
                    clean_device_description(values.get("DeviceDesc")),
                    clean_device_description(values.get("Desc")),
                )
                if name:
                    key = (name.casefold(), path.casefold())
                    if key not in seen:
                        pnp_path = path
                        enum_prefix = r"SYSTEM\CurrentControlSet\Enum" + "\\"
                        if pnp_path.startswith(enum_prefix):
                            pnp_path = pnp_path[len(enum_prefix):]
                        seen.add(key)
                        found.append(
                            NpuDevice(
                                name=name,
                                manufacturer=pick_first_meaningful(
                                    clean_device_description(values.get("Mfg")),
                                    clean_device_description(values.get("Manufacturer")),
                                ),
                                pnp_class=class_name or (f"Service: {service}" if service else ""),
                                pnp_device_id=pnp_path,
                                status=registry_status_from_values(values),
                            )
                        )
        if depth >= 2:
            return
        for child in iter_registry_subkeys(winreg.HKEY_LOCAL_MACHINE, path, WOW64_64KEY):
            scan_key(fr"{path}\{child}", depth + 1)

    for root_path in NPU_REGISTRY_ROOTS:
        scan_key(root_path, 0)
    return dedupe_npu_devices(found)


def collect_graphics_and_npu(ctx: AuditContext) -> tuple[list[GraphicsDevice], list[NpuDevice]]:
    script = """
    $ErrorActionPreference = 'SilentlyContinue'

    $gpus = @()
    try {
        foreach ($gpu in @(Get-CimInstance Win32_VideoController)) {
            $resolution = $null
            if ($gpu.CurrentHorizontalResolution -and $gpu.CurrentVerticalResolution) {
                $resolution = "$($gpu.CurrentHorizontalResolution)x$($gpu.CurrentVerticalResolution)"
            }
            $gpus += @{
                Name = "$($gpu.Name)"
                AdapterCompatibility = "$($gpu.AdapterCompatibility)"
                VideoProcessor = "$($gpu.VideoProcessor)"
                AdapterRAM = "$($gpu.AdapterRAM)"
                PNPDeviceID = "$($gpu.PNPDeviceID)"
                DriverVersion = "$($gpu.DriverVersion)"
                CurrentResolution = "$resolution"
                Status = "$($gpu.Status)"
            }
        }
    } catch {}

    $patterns = '\\bNPU\\b|Neural|AI Boost|Ryzen AI|Hexagon|Inference|\\bVPU\\b|\\bIPU\\b|Input Configuration Device'
    $npus = @()
    try {
        foreach ($npu in @(
            Get-CimInstance Win32_PnPEntity |
                Where-Object {
                    (
                        ($_.Name -match $patterns) -or
                        ($_.Description -match $patterns) -or
                        ($_.DeviceID -match $patterns) -or
                        ($_.Manufacturer -match $patterns)
                    ) -and ($_.PNPClass -notmatch '^Display$')
                }
        )) {
            $npus += @{
                Name = "$($npu.Name)"
                Manufacturer = "$($npu.Manufacturer)"
                PNPClass = "$($npu.PNPClass)"
                DeviceID = "$($npu.DeviceID)"
                Status = "$($npu.Status)"
            }
        }
    } catch {}

    @{
        GPUs = $gpus
        NPUs = $npus
    } | ConvertTo-Json -Depth 5 -Compress
    """
    raw = powershell_json(script, ctx, "GPU/NPU")
    graphics: list[GraphicsDevice] = []
    npus: list[NpuDevice] = []

    if isinstance(raw, dict):
        for item in ensure_list(safe_get(raw, "GPUs", [])):
            name = clean_text(safe_get(item, "Name"))
            if not name:
                continue
            graphics.append(
                GraphicsDevice(
                    name=name,
                    vendor=clean_text(safe_get(item, "AdapterCompatibility")),
                    video_processor=clean_text(safe_get(item, "VideoProcessor")),
                    adapter_ram_bytes=int_or_zero(safe_get(item, "AdapterRAM")),
                    pnp_device_id=clean_text(safe_get(item, "PNPDeviceID")),
                    driver_version=clean_text(safe_get(item, "DriverVersion")),
                    current_resolution=clean_text(safe_get(item, "CurrentResolution")),
                    status=clean_text(safe_get(item, "Status")),
                )
            )

        for item in ensure_list(safe_get(raw, "NPUs", [])):
            name = clean_text(safe_get(item, "Name"))
            if not name:
                continue
            npus.append(
                NpuDevice(
                    name=name,
                    manufacturer=clean_text(safe_get(item, "Manufacturer")),
                    pnp_class=clean_text(safe_get(item, "PNPClass")),
                    pnp_device_id=clean_text(safe_get(item, "DeviceID")),
                    status=clean_text(safe_get(item, "Status")),
                )
            )

    if not graphics:
        graphics = collect_graphics_from_registry(ctx)
    if not npus:
        npus = collect_npu_from_registry(ctx)

    return graphics, dedupe_npu_devices(npus)


def collect_eset_product_instance_lines(ctx: AuditContext) -> list[str]:
    if not ESET_TRACE_LOG.exists():
        return []

    needle = b"ProductInstanceID"
    needle_utf16le = "ProductInstanceID".encode("utf-16-le")
    try:
        with ESET_TRACE_LOG.open("rb") as handle:
            for raw_line in handle:
                if needle not in raw_line and needle_utf16le not in raw_line:
                    continue
                text = ""
                for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "cp1251"):
                    try:
                        text = raw_line.decode(encoding)
                    except UnicodeDecodeError:
                        continue
                    if "ProductInstanceID" in text:
                        break
                if "ProductInstanceID" not in text:
                    text = raw_line.decode("utf-8", errors="replace")
                text = clean_text(text)
                if text:
                    return [text]
    except OSError as exc:
        ctx.warn(f"ESET ProductInstanceID: не вдалося прочитати {ESET_TRACE_LOG}: {exc}")
        return []

    return []


def collect_local_users(ctx: AuditContext) -> list[LocalUser]:
    script = (
        "Get-LocalUser | "
        "Select-Object Name,Enabled,LastLogon,SID | "
        "ConvertTo-Json -Depth 3"
    )
    raw = powershell_json(script, ctx, "Користувачі")
    users = []
    for item in ensure_list(raw):
        name = clean_text(safe_get(item, "Name"))
        if not name:
            continue
        users.append(
            LocalUser(
                name=name,
                sid=clean_text(safe_get(item, "SID")),
                enabled=safe_get(item, "Enabled"),
                last_logon=clean_text(safe_get(item, "LastLogon")),
            )
        )
    if users:
        return users

    completed = run_hidden_subprocess(
        ["net", "user"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode != 0:
        ctx.warn("Користувачі: не вдалося зчитати навіть через net user.")
        return []
    for line in completed.stdout.splitlines():
        if "----" in line or "command completed" in line.lower() or "\\" in line:
            continue
        for token in line.split():
            users.append(LocalUser(name=token))
    return users


def mark_administrators(users: list[LocalUser], ctx: AuditContext) -> None:
    script = (
        "Get-LocalGroupMember -SID 'S-1-5-32-544' | "
        "Select-Object Name,SID | ConvertTo-Json -Depth 3"
    )
    raw = powershell_json(script, ctx, "Адміністратори")
    admin_sids = {clean_text(safe_get(item, "SID")) for item in ensure_list(raw)}
    admin_names = {
        clean_text(safe_get(item, "Name")).split("\\")[-1].casefold()
        for item in ensure_list(raw)
        if clean_text(safe_get(item, "Name"))
    }
    if not admin_sids and not admin_names:
        completed = run_hidden_subprocess(
            ["net", "localgroup", "Administrators"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode == 0:
            capture = False
            for line in completed.stdout.splitlines():
                stripped = clean_text(line)
                if "----" in stripped:
                    capture = True
                    continue
                if not capture or not stripped or stripped.lower().startswith("the command"):
                    continue
                admin_names.add(stripped.split("\\")[-1].casefold())
    for user in users:
        if user.sid and user.sid in admin_sids:
            user.is_admin = True
            continue
        if user.name.casefold() in admin_names:
            user.is_admin = True


def collect_profile_info_map() -> dict[str, UserProfileInfo]:
    base_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    profile_info: dict[str, UserProfileInfo] = {}
    for sid in iter_registry_subkeys(winreg.HKEY_LOCAL_MACHINE, base_path, WOW64_64KEY):
        profile_path = read_registry_value(
            winreg.HKEY_LOCAL_MACHINE,
            fr"{base_path}\{sid}",
            "ProfileImagePath",
            WOW64_64KEY,
        )
        expanded_path = clean_text(os.path.expandvars(clean_text(profile_path)))
        username = Path(expanded_path).name if expanded_path else ""
        if username or expanded_path:
            profile_info[sid] = UserProfileInfo(username=username, path=expanded_path)
    return profile_info


def default_user_profile_roots() -> list[Path]:
    roots: list[Path] = []
    public_path = clean_text(os.environ.get("PUBLIC"))
    if public_path:
        roots.append(Path(public_path).parent)
    system_drive = clean_text(os.environ.get("SystemDrive")) or "C:"
    roots.append(Path(f"{system_drive}\\Users"))

    ordered: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        key = str(root).casefold()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(root)
    return ordered


def resolve_user_profile_path(user: LocalUser, profile_info: dict[str, UserProfileInfo]) -> Path | None:
    if user.sid:
        info = profile_info.get(user.sid)
        if info and info.path:
            return Path(info.path)

    username = clean_text(user.name)
    if not username:
        return None
    for root in default_user_profile_roots():
        candidate = root / username
        if safe_path_exists(candidate):
            return candidate
    return None


def should_skip_program(values: dict[str, Any]) -> bool:
    display_name = clean_text(values.get("DisplayName") or values.get("QuietDisplayName"))
    if not display_name:
        return True
    if values.get("SystemComponent") == 1:
        return True
    release_type = clean_text(values.get("ReleaseType")).casefold()
    if release_type in {"security update", "update", "hotfix"}:
        return True
    if clean_text(values.get("ParentKeyName")):
        return True
    return False


def read_uninstall_entries(
    root: int,
    subkey: str,
    access: int,
    ctx: AuditContext,
    label: str,
) -> list[ProgramEntry]:
    entries: list[ProgramEntry] = []
    for child in iter_registry_subkeys(root, subkey, access):
        values = read_registry_values(root, fr"{subkey}\{child}", access)
        if not values or should_skip_program(values):
            continue
        entry = ProgramEntry(
            name=clean_text(values.get("DisplayName") or values.get("QuietDisplayName")),
            version=clean_text(values.get("DisplayVersion")),
            install_date=parse_install_date(clean_text(values.get("InstallDate"))),
            publisher=clean_text(values.get("Publisher")),
        )
        entries.append(entry)
    if not entries:
        return []
    deduped: dict[tuple[str, str, str, str, str], ProgramEntry] = {}
    for entry in entries:
        deduped[program_entry_key(entry)] = entry
    ordered = list(deduped.values())
    ordered.sort(key=program_entry_sort_key)
    if not ordered:
        ctx.warn(f"{label}: не знайдено жодної програми.")
    return ordered


def program_entry_key(entry: ProgramEntry) -> tuple[str, str, str, str, str]:
    return (
        entry.name.casefold(),
        entry.version.casefold(),
        entry.install_date.casefold(),
        entry.publisher.casefold(),
        entry.source.casefold(),
    )


def program_entry_sort_key(entry: ProgramEntry) -> tuple[str, str, str, str]:
    return (
        entry.name.casefold(),
        entry.version.casefold(),
        entry.install_date.casefold(),
        entry.source.casefold(),
    )


def read_uninstall_entries_from_list(entries: list[ProgramEntry]) -> list[ProgramEntry]:
    deduped: dict[tuple[str, str, str, str, str], ProgramEntry] = {}
    for entry in entries:
        deduped[program_entry_key(entry)] = entry
    ordered = list(deduped.values())
    ordered.sort(key=program_entry_sort_key)
    return ordered


def powershell_single_quote(value: str) -> str:
    return "'" + clean_text(value).replace("'", "''") + "'"


def appx_package_inventory_script(command: str) -> str:
    return f"""
$ErrorActionPreference = 'Stop'

function Convert-AuditAppxPackage {{
    param($Package)

    if ($null -eq $Package) {{ return }}
    if ($Package.PSObject.Properties['IsFramework'] -and [bool]$Package.IsFramework) {{ return }}
    if ($Package.PSObject.Properties['IsResourcePackage'] -and [bool]$Package.IsResourcePackage) {{ return }}

    $name = ''
    if ($Package.PSObject.Properties['Name']) {{ $name = [string]$Package.Name }}
    if ([string]::IsNullOrWhiteSpace($name) -and $Package.PSObject.Properties['PackageFullName']) {{
        $name = [string]$Package.PackageFullName
    }}
    if ([string]::IsNullOrWhiteSpace($name)) {{ return }}

    $signatureKind = ''
    if ($Package.PSObject.Properties['SignatureKind'] -and $null -ne $Package.SignatureKind) {{
        $signatureKind = [string]$Package.SignatureKind
    }}
    if ($signatureKind -eq 'System') {{ return }}

    $installDate = ''
    if ($Package.PSObject.Properties['InstallDate'] -and $Package.InstallDate) {{
        try {{ $installDate = ([datetime]$Package.InstallDate).ToString('dd.MM.yyyy') }}
        catch {{ $installDate = [string]$Package.InstallDate }}
    }}

    [pscustomobject]@{{
        Name = $name
        Version = if ($Package.PSObject.Properties['Version']) {{ [string]$Package.Version }} else {{ '' }}
        InstallDate = $installDate
        Publisher = if ($Package.PSObject.Properties['Publisher']) {{ [string]$Package.Publisher }} else {{ '' }}
        PublisherId = if ($Package.PSObject.Properties['PublisherId']) {{ [string]$Package.PublisherId }} else {{ '' }}
        PackageFullName = if ($Package.PSObject.Properties['PackageFullName']) {{ [string]$Package.PackageFullName }} else {{ '' }}
        PackageFamilyName = if ($Package.PSObject.Properties['PackageFamilyName']) {{ [string]$Package.PackageFamilyName }} else {{ '' }}
        SignatureKind = $signatureKind
    }}
}}

$entries = @({command} | ForEach-Object {{ Convert-AuditAppxPackage $_ }})
ConvertTo-Json -InputObject $entries -Depth 4
"""


def appx_provisioned_inventory_script() -> str:
    return """
$ErrorActionPreference = 'Stop'

$entries = @(Get-AppxProvisionedPackage -Online | ForEach-Object {
    $name = [string]$_.DisplayName
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$_.PackageName }
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        [pscustomobject]@{
            Name = $name
            Version = if ($_.PSObject.Properties['Version']) { [string]$_.Version } else { '' }
            InstallDate = ''
            Publisher = ''
            PublisherId = if ($_.PSObject.Properties['PublisherId']) { [string]$_.PublisherId } else { '' }
            PackageFullName = if ($_.PSObject.Properties['PackageName']) { [string]$_.PackageName } else { '' }
            PackageFamilyName = ''
            SignatureKind = ''
        }
    }
})
ConvertTo-Json -InputObject $entries -Depth 4
"""


def should_skip_appx_program_name(name: str) -> bool:
    text = clean_text(name).casefold()
    if not text:
        return True
    if re.fullmatch(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", text):
        return True
    dependency_prefixes = (
        "microsoft.net.native.",
        "microsoft.vclibs.",
        "microsoft.ui.xaml.",
        "microsoft.windowsappruntime.",
    )
    system_prefixes = (
        "microsoft.aad.brokerplugin",
        "microsoft.accountscontrol",
        "microsoft.asynctextservice",
        "microsoft.bioenrollment",
        "microsoft.creddialoghost",
        "microsoft.ecapp",
        "microsoft.lockapp",
        "microsoft.microsoftedgedevtoolsclient",
        "microsoft.win32webviewhost",
        "microsoft.windows.apprep.",
        "microsoft.windows.assignedaccesslockapp",
        "microsoft.windows.capturepicker",
        "microsoft.windows.cloudexperiencehost",
        "microsoft.windows.contentdeliverymanager",
        "microsoft.windows.oobenetwork",
        "microsoft.windows.parentalcontrols",
        "microsoft.windows.peopleexperiencehost",
        "microsoft.windows.pinningconfirmationdialog",
        "microsoft.windows.printqueueactioncenter",
        "microsoft.windows.secureassessmentbrowser",
        "microsoft.windows.shellexperiencehost",
        "microsoft.windows.startmenuexperiencehost",
        "microsoft.windows.xgpuejectdialog",
        "microsoft.xboxgamecallableui",
        "microsoftwindows.",
        "windows.cbspreview",
        "windows.immersivecontrolpanel",
        "windows.printdialog",
    )
    return any(text.startswith(prefix) for prefix in (*dependency_prefixes, *system_prefixes))


def appx_publisher_name(raw_publisher: str, publisher_id: str = "") -> str:
    publisher = clean_text(raw_publisher)
    if publisher:
        match = re.search(r"(?:^|,\s*)CN=([^,]+)", publisher, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1))
        return publisher
    return clean_text(publisher_id)


def should_skip_appx_publisher(publisher: str) -> bool:
    return clean_text(publisher).casefold() in APPX_DEFAULT_OS_PUBLISHERS


def appx_entries_from_json(raw: Any | None) -> list[ProgramEntry]:
    entries: list[ProgramEntry] = []
    for item in ensure_list(raw):
        if not isinstance(item, dict):
            continue
        name = clean_text(
            safe_get(item, "Name")
            or safe_get(item, "PackageFullName")
            or safe_get(item, "PackageFamilyName")
        )
        if should_skip_appx_program_name(name):
            continue
        publisher = appx_publisher_name(
            clean_text(safe_get(item, "Publisher")),
            clean_text(safe_get(item, "PublisherId")),
        )
        if should_skip_appx_publisher(publisher):
            continue
        entries.append(
            ProgramEntry(
                name=name,
                version=clean_text(safe_get(item, "Version")),
                install_date=clean_text(safe_get(item, "InstallDate")),
                publisher=publisher,
                source=APPX_PROGRAM_SOURCE,
            )
        )
    return read_appx_entries_from_list(entries)


def parse_appx_package_full_name(package_full_name: str) -> ProgramEntry | None:
    text = clean_text(package_full_name)
    if not text:
        return None
    parts = text.rsplit("_", 4)
    if len(parts) != 5:
        name = text
        version = ""
        publisher_id = ""
    else:
        name, version, _architecture, _resource_id, publisher_id = parts
    if should_skip_appx_program_name(name):
        return None
    if should_skip_appx_publisher(publisher_id):
        return None
    return ProgramEntry(
        name=name,
        version=version,
        publisher=publisher_id,
        source=APPX_PROGRAM_SOURCE,
    )


def appx_version_sort_key(version: str) -> tuple[Any, ...]:
    text = clean_text(version)
    if not text:
        return ()
    parts: list[Any] = []
    for part in re.split(r"[.\-+]", text):
        if part.isdigit():
            parts.append((0, int(part)))
        else:
            parts.append((1, part.casefold()))
    return tuple(parts)


def read_appx_entries_from_list(entries: list[ProgramEntry]) -> list[ProgramEntry]:
    newest: dict[tuple[str, str, str], ProgramEntry] = {}
    for entry in entries:
        key = (entry.name.casefold(), entry.publisher.casefold(), entry.source.casefold())
        current = newest.get(key)
        if current is None or appx_version_sort_key(entry.version) > appx_version_sort_key(current.version):
            newest[key] = entry
    return read_uninstall_entries_from_list(list(newest.values()))


def read_appx_registry_entries(root: int, subkey: str) -> list[ProgramEntry]:
    entries: list[ProgramEntry] = []
    for child in iter_registry_subkeys(root, subkey, WOW64_64KEY):
        entry = parse_appx_package_full_name(child)
        if entry is not None:
            entries.append(entry)
    return read_appx_entries_from_list(entries)


def collect_all_users_appx_registry_programs() -> list[ProgramEntry]:
    return read_appx_registry_entries(
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications",
    )


def collect_user_appx_registry_programs(user: LocalUser) -> list[ProgramEntry]:
    user_sid = clean_text(user.sid)
    if not user_sid:
        return []
    return read_appx_registry_entries(
        winreg.HKEY_LOCAL_MACHINE,
        fr"SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\{user_sid}",
    )


def collect_appx_package_programs(ctx: AuditContext, command: str, label: str) -> list[ProgramEntry]:
    raw = powershell_json(appx_package_inventory_script(command), ctx, label, timeout=60)
    return appx_entries_from_json(raw)


def collect_all_users_appx_programs(ctx: AuditContext) -> list[ProgramEntry]:
    if is_process_elevated() is not True:
        return collect_all_users_appx_registry_programs()

    entries = collect_appx_package_programs(
        ctx,
        "Get-AppxPackage -AllUsers",
        "Microsoft Store/Appx для всіх користувачів",
    )
    if entries:
        return entries

    raw = powershell_json(
        appx_provisioned_inventory_script(),
        ctx,
        "Microsoft Store/Appx provisioned packages",
        timeout=60,
    )
    entries = appx_entries_from_json(raw)
    if entries:
        return entries
    return collect_all_users_appx_registry_programs()


def collect_user_appx_programs(
    user: LocalUser,
    ctx: AuditContext,
    is_current_profile: bool = False,
) -> list[ProgramEntry]:
    if is_current_profile:
        entries = collect_appx_package_programs(ctx, "Get-AppxPackage", f"Microsoft Store/Appx {user.name}")
        if entries:
            return entries
        return collect_user_appx_registry_programs(user)

    entries = collect_user_appx_registry_programs(user)
    if entries or is_process_elevated() is not True:
        return entries

    user_key = clean_text(user.sid) or clean_text(user.name)
    if not user_key:
        return []
    command = f"Get-AppxPackage -User {powershell_single_quote(user_key)}"

    label_user = clean_text(user.name) or clean_text(user.sid) or "профіль"
    return collect_appx_package_programs(
        ctx,
        command,
        f"Microsoft Store/Appx {label_user}",
    )


def compact_subprocess_error(completed: subprocess.CompletedProcess[str]) -> str:
    text = clean_text(completed.stderr) or clean_text(completed.stdout)
    if len(text) > 180:
        return f"{text[:177]}..."
    return text


def compact_exception_text(exc: BaseException) -> str:
    text = clean_text(exc)
    if len(text) > 180:
        return f"{text[:177]}..."
    return text


def enable_windows_privilege(privilege_name: str) -> bool:
    if os.name != "nt":
        return False
    try:
        advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)

        class LUID(ctypes.Structure):
            _fields_ = [
                ("LowPart", wintypes.DWORD),
                ("HighPart", wintypes.LONG),
            ]

        class TOKEN_PRIVILEGES(ctypes.Structure):
            _fields_ = [
                ("PrivilegeCount", wintypes.DWORD),
                ("Luid", LUID),
                ("Attributes", wintypes.DWORD),
            ]

        token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(
            kernel.GetCurrentProcess(),
            0x0020 | 0x0008,
            ctypes.byref(token),
        ):
            return False
        try:
            luid = LUID()
            if not advapi32.LookupPrivilegeValueW(None, privilege_name, ctypes.byref(luid)):
                return False
            privileges = TOKEN_PRIVILEGES(1, luid, 0x00000002)
            ctypes.set_last_error(0)
            if not advapi32.AdjustTokenPrivileges(
                token,
                False,
                ctypes.byref(privileges),
                0,
                None,
                None,
            ):
                return False
            return ctypes.get_last_error() == 0
        finally:
            kernel.CloseHandle(token)
    except Exception:
        return False


def load_user_hive_native(hive_name: str, ntuser_path: Path) -> None:
    enable_windows_privilege("SeBackupPrivilege")
    enable_windows_privilege("SeRestorePrivilege")
    winreg.LoadKey(winreg.HKEY_USERS, hive_name, str(ntuser_path))


def temporary_hive_name(user: LocalUser) -> str:
    source = clean_text(user.sid) or clean_text(user.name) or "USER"
    token = re.sub(r"[^A-Za-z0-9_]", "_", source).strip("_")[:80] or "USER"
    return f"AUDIT_HSW_{token}_{os.getpid()}"


def read_unloaded_profile_programs(
    user: LocalUser,
    profile_info: dict[str, UserProfileInfo],
    ctx: AuditContext,
) -> tuple[list[ProgramEntry], str]:
    profile_path = resolve_user_profile_path(user, profile_info)
    if profile_path is None:
        return [], PROFILE_NOT_FOUND_NOTE

    ntuser_path = profile_path / "NTUSER.DAT"
    if not safe_path_exists(ntuser_path):
        return [], f"Профіль знайдено, але NTUSER.DAT недоступний для читання: {profile_path}"

    hive_name = temporary_hive_name(user)
    hive_root = fr"HKU\{hive_name}"
    loaded_with_native_api = False
    try:
        load_user_hive_native(hive_name, ntuser_path)
        loaded_with_native_api = True
    except Exception as native_exc:
        native_detail = compact_exception_text(native_exc)
        try:
            load_completed = run_hidden_subprocess(
                ["reg.exe", "load", hive_root, str(ntuser_path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
                check=False,
            )
        except Exception as exc:
            detail = compact_exception_text(exc) or native_detail
            return [], f"Registry hive профілю не завантажений у HKEY_USERS; NTUSER.DAT не вдалося підключити: {detail}"
        if load_completed.returncode != 0:
            detail = compact_subprocess_error(load_completed) or native_detail
            suffix = f": {detail}" if detail else "."
            return [], f"Registry hive профілю не завантажений у HKEY_USERS; NTUSER.DAT не вдалося підключити{suffix}"

    try:
        programs: list[ProgramEntry] = []
        programs.extend(
            read_uninstall_entries(
                winreg.HKEY_USERS,
                fr"{hive_name}\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                0,
                ctx,
                f"Програми {user.name} NTUSER.DAT",
            )
        )
        programs.extend(
            read_uninstall_entries(
                winreg.HKEY_USERS,
                fr"{hive_name}\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                0,
                ctx,
                f"Програми {user.name} NTUSER.DAT WOW6432Node",
            )
        )
        ordered = read_uninstall_entries_from_list(programs)
        if not ordered:
            return [], "Персональні інсталяції профілю не знайдено."
        return ordered, ""
    finally:
        if loaded_with_native_api:
            try:
                winreg.UnloadKey(winreg.HKEY_USERS, hive_name)
            except Exception as exc:
                ctx.warn(f"Профіль {user.name}: тимчасовий hive {hive_root} не вдалося відключити: {exc}")
        else:
            unload_completed = None
            try:
                unload_completed = run_hidden_subprocess(
                    ["reg.exe", "unload", hive_root],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=30,
                    check=False,
                )
            except Exception as exc:
                ctx.warn(f"Профіль {user.name}: тимчасовий hive {hive_root} не вдалося відключити: {exc}")
            if unload_completed is not None and unload_completed.returncode != 0:
                detail = compact_subprocess_error(unload_completed)
                suffix = f": {detail}" if detail else "."
                ctx.warn(f"Профіль {user.name}: тимчасовий hive {hive_root} не вдалося відключити{suffix}")


def collect_programs(ctx: AuditContext, users: list[LocalUser]) -> tuple[list[ProgramEntry], dict[str, list[ProgramEntry]], dict[str, str]]:
    common_programs = []
    common_programs.extend(
        read_uninstall_entries(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            WOW64_64KEY,
            ctx,
            "Програми HKLM x64",
        )
    )
    common_programs.extend(
        read_uninstall_entries(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            WOW64_64KEY,
            ctx,
            "Програми HKLM x86",
        )
    )
    common_programs.extend(collect_all_users_appx_programs(ctx))
    common_programs = read_uninstall_entries_from_list(common_programs)

    profile_info = collect_profile_info_map()
    loaded_hives = {sid for sid in iter_registry_subkeys(winreg.HKEY_USERS, "", 0) if sid and not sid.endswith("_Classes")}

    per_user: dict[str, list[ProgramEntry]] = {}
    user_notes: dict[str, str] = {}
    current_username = clean_text(os.environ.get("USERNAME")).casefold()
    current_profile_path = clean_text(os.environ.get("USERPROFILE"))

    for user in users:
        if user.is_system:
            continue
        if user.enabled is False and not user.is_admin:
            continue
        username = user.name
        user_sid = user.sid
        profile_path = resolve_user_profile_path(user, profile_info)
        is_current_profile = username.casefold() == current_username or same_path(profile_path, current_profile_path)
        if is_current_profile:
            current_entries: list[ProgramEntry] = []
            current_entries.extend(
                read_uninstall_entries(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
                    0,
                    ctx,
                    f"Програми {username} HKCU",
                )
            )
            current_entries.extend(
                read_uninstall_entries(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                    0,
                    ctx,
                    f"Програми {username} HKCU WOW6432Node",
                )
            )
            current_entries.extend(collect_user_appx_programs(user, ctx, is_current_profile=True))
            per_user[username] = read_uninstall_entries_from_list(current_entries)
            if not per_user[username]:
                user_notes[username] = "Персональні інсталяції поточного профілю не знайдено."
            continue
        if user_sid and user_sid in loaded_hives:
            programs: list[ProgramEntry] = []
            programs.extend(
                read_uninstall_entries(
                    winreg.HKEY_USERS,
                    fr"{user_sid}\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                    0,
                    ctx,
                    f"Програми {username} HKU",
                )
            )
            programs.extend(
                read_uninstall_entries(
                    winreg.HKEY_USERS,
                    fr"{user_sid}\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                    0,
                    ctx,
                    f"Програми {username} HKU WOW6432Node",
                )
            )
            programs.extend(collect_user_appx_programs(user, ctx))
            per_user[username] = read_uninstall_entries_from_list(programs)
            if not per_user[username]:
                user_notes[username] = "Персональні інсталяції профілю не знайдено."
            continue

        programs, note = read_unloaded_profile_programs(user, profile_info, ctx)
        programs = read_uninstall_entries_from_list(
            [*programs, *collect_user_appx_programs(user, ctx)]
        )
        if programs:
            per_user[username] = programs
        else:
            user_notes[username] = note or "Персональні інсталяції не знайдено або профіль недоступний."

    return common_programs, per_user, user_notes


def get_runtime_windows_version() -> dict[str, str]:
    class OSVERSIONINFOEXW(ctypes.Structure):
        _fields_ = [
            ("dwOSVersionInfoSize", wintypes.DWORD),
            ("dwMajorVersion", wintypes.DWORD),
            ("dwMinorVersion", wintypes.DWORD),
            ("dwBuildNumber", wintypes.DWORD),
            ("dwPlatformId", wintypes.DWORD),
            ("szCSDVersion", wintypes.WCHAR * 128),
            ("wServicePackMajor", wintypes.WORD),
            ("wServicePackMinor", wintypes.WORD),
            ("wSuiteMask", wintypes.WORD),
            ("wProductType", wintypes.BYTE),
            ("wReserved", wintypes.BYTE),
        ]

    info = OSVERSIONINFOEXW()
    info.dwOSVersionInfoSize = ctypes.sizeof(OSVERSIONINFOEXW)
    try:
        status = ctypes.WinDLL("ntdll").RtlGetVersion(ctypes.byref(info))
    except Exception:
        return {}
    if status != 0:
        return {}
    version_number = f"{info.dwMajorVersion}.{info.dwMinorVersion}.{info.dwBuildNumber}"
    return {
        "major": str(info.dwMajorVersion),
        "minor": str(info.dwMinorVersion),
        "build_number": str(info.dwBuildNumber),
        "version_number": version_number,
    }


def collect_os_info(ctx: AuditContext) -> dict[str, str]:
    values = read_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
        WOW64_64KEY,
    )
    install_date_raw = values.get("InstallDate")
    runtime = get_runtime_windows_version()

    script = """
    $os = $null
    try {
        $os = Get-CimInstance Win32_OperatingSystem |
            Select-Object Caption,Version,BuildNumber,OSArchitecture,InstallDate
    } catch {}

    $current = $null
    try {
        $current = Get-ItemProperty -LiteralPath 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion'
    } catch {}

    $installDate = $null
    if ($os -and $os.InstallDate) {
        try { $installDate = ([datetime]$os.InstallDate).ToString('dd.MM.yyyy') } catch {}
    }
    $caption = $null
    $version = $null
    $buildNumber = $null
    $osArchitecture = $null
    if ($os) {
        $caption = $os.Caption
        $version = $os.Version
        $buildNumber = $os.BuildNumber
        $osArchitecture = $os.OSArchitecture
    }
    $productName = $null
    $displayVersion = $null
    $releaseId = $null
    $currentBuildNumber = $null
    $ubr = $null
    $editionId = $null
    if ($current) {
        $productName = $current.ProductName
        $displayVersion = $current.DisplayVersion
        $releaseId = $current.ReleaseId
        $currentBuildNumber = $current.CurrentBuildNumber
        $ubr = $current.UBR
        $editionId = $current.EditionID
    }

    @{
        Caption = $caption
        Version = $version
        BuildNumber = $buildNumber
        OSArchitecture = $osArchitecture
        InstallDate = $installDate
        ProductName = $productName
        DisplayVersion = $displayVersion
        ReleaseId = $releaseId
        CurrentBuildNumber = $currentBuildNumber
        UBR = $ubr
        EditionID = $editionId
    } | ConvertTo-Json -Depth 3 -Compress
    """
    current_raw = powershell_json(script, ctx, "Windows")

    build_number = pick_first_meaningful(
        runtime.get("build_number", ""),
        clean_text(safe_get(current_raw, "BuildNumber")) if isinstance(current_raw, dict) else "",
        clean_text(values.get("CurrentBuildNumber")),
        clean_text(values.get("CurrentBuild")),
    )
    edition_id = pick_first_meaningful(
        clean_text(safe_get(current_raw, "EditionID")) if isinstance(current_raw, dict) else "",
        clean_text(values.get("EditionID")),
    )
    product_name = normalize_windows_product_name(
        clean_text(values.get("ProductName")),
        clean_text(safe_get(current_raw, "ProductName")) if isinstance(current_raw, dict) else "",
        clean_text(safe_get(current_raw, "Caption")) if isinstance(current_raw, dict) else "",
        build_number=build_number,
        edition_id=edition_id,
    )
    display_version = pick_first_meaningful(
        clean_text(safe_get(current_raw, "DisplayVersion")) if isinstance(current_raw, dict) else "",
        clean_text(values.get("DisplayVersion")),
        clean_text(safe_get(current_raw, "ReleaseId")) if isinstance(current_raw, dict) else "",
        clean_text(values.get("ReleaseId")),
    )
    version_number = pick_first_meaningful(
        runtime.get("version_number", ""),
        clean_text(safe_get(current_raw, "Version")) if isinstance(current_raw, dict) else "",
    )
    ubr = pick_first_meaningful(
        clean_text(safe_get(current_raw, "UBR")) if isinstance(current_raw, dict) else "",
        clean_text(values.get("UBR")),
    )
    install_date = pick_first_meaningful(
        clean_text(safe_get(current_raw, "InstallDate")) if isinstance(current_raw, dict) else "",
        utc_seconds_to_local_date(int(install_date_raw)) if install_date_raw is not None else "",
    )
    architecture = pick_first_meaningful(
        clean_text(safe_get(current_raw, "OSArchitecture")) if isinstance(current_raw, dict) else "",
        clean_text(os.environ.get("PROCESSOR_ARCHITECTURE") or platform.machine()),
    )

    return {
        "product_name": product_name,
        "display_version": display_version,
        "version_number": version_number,
        "build_number": build_number,
        "ubr": ubr,
        "edition_id": edition_id,
        "install_date": install_date,
        "architecture": architecture,
    }


def collect_license_info(ctx: AuditContext) -> LicenseInfo:
    info = LicenseInfo()
    spp_values = read_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
        WOW64_64KEY,
    )
    info.kms_host = clean_text(spp_values.get("KeyManagementServiceName"))
    info.kms_port = clean_text(spp_values.get("KeyManagementServicePort"))
    info.backup_product_key = clean_text(spp_values.get("BackupProductKeyDefault"))
    info.backup_key_tail = format_key_tail(info.backup_product_key)

    script = """
    $svcKey = $null
    $prod = $null
    try {
        $svc = Get-CimInstance SoftwareLicensingService
        if ($svc) {
            $svcKey = $svc.OA3xOriginalProductKey
        }
    } catch {}
    try {
        $candidate = Get-CimInstance SoftwareLicensingProduct |
            Where-Object {
                $_.ApplicationID -eq '55c92734-d682-4d71-983e-d6ec3f16059f' -and
                $_.PartialProductKey -and
                $_.Name -like 'Windows*'
            } |
            Sort-Object -Property LicenseStatus -Descending |
            Select-Object -First 1
        if ($candidate) {
            $prod = @{
                Name = "$($candidate.Name)"
                Description = "$($candidate.Description)"
                LicenseStatus = "$($candidate.LicenseStatus)"
                PartialProductKey = "$($candidate.PartialProductKey)"
                ProductKeyChannel = "$($candidate.ProductKeyChannel)"
            }
        }
    } catch {}
    @{
        Service = @{
            OA3xOriginalProductKey = "$svcKey"
        }
        Product = $prod
    } | ConvertTo-Json -Depth 4 -Compress
    """
    raw = powershell_json(script, ctx, "Ліцензія", timeout=12)
    service = safe_get(raw, "Service", {}) if isinstance(raw, dict) else {}
    product = safe_get(raw, "Product", {}) if isinstance(raw, dict) else {}

    info.oa3_product_key = clean_text(safe_get(service, "OA3xOriginalProductKey"))
    info.oa3_key_tail = format_key_tail(info.oa3_product_key)
    info.channel = clean_text(safe_get(product, "ProductKeyChannel"))
    info.description = clean_text(safe_get(product, "Description"))
    info.partial_product_key = clean_text(safe_get(product, "PartialProductKey"))
    status_code = safe_get(product, "LicenseStatus")
    if status_code is not None and str(status_code).strip() != "":
        try:
            info.status = LICENSE_STATUS_MAP.get(int(status_code), str(status_code))
        except (TypeError, ValueError):
            info.status = clean_text(status_code)

    channel_upper = f"{info.channel} {info.description}".upper()
    if info.oa3_product_key or "OEM" in channel_upper:
        info.type_guess = "OEM"
    elif "MAK" in channel_upper:
        info.type_guess = "MAK / Volume"
    elif "KMS" in channel_upper or info.kms_host:
        info.type_guess = "KMS / Volume"
    elif "RETAIL" in channel_upper:
        info.type_guess = "Retail (FPP/ESD)"
    elif info.backup_key_tail:
        info.type_guess = "Невизначено (ключ знайдено)"

    if info.kms_host:
        suffix = f":{info.kms_port}" if info.kms_port else ""
        info.notes.append(f"Налаштовано KMS host: {info.kms_host}{suffix}")
    if not raw or (not info.channel and not info.partial_product_key and not info.status and not info.oa3_product_key):
        info.notes.append("Детальний канал ліцензії через WMI/CIM зчитати не вдалося.")
    if info.type_guess.startswith("Retail"):
        info.notes.append("FPP та ESD автоматично не розрізняються, система бачить лише Retail-канал.")
    if info.status == "Licensed" and not info.kms_host and "KMS" not in channel_upper and "MAK" not in channel_upper:
        info.notes.append("Явних ознак KMS/MAK не виявлено.")
    return info


def build_device_title(system_info: dict[str, str], processor_info: dict[str, str], total_memory: int, disks: list[DiskDevice]) -> str:
    manufacturer = system_info.get("manufacturer") or "Невідомий виробник"
    model_parts = unique_preserve(
        [
            system_info.get("system_family", ""),
            system_info.get("product_name", ""),
            system_info.get("system_version", ""),
        ]
    )
    if model_parts and model_parts[0].casefold() == manufacturer.casefold():
        model_parts = model_parts[1:]
    model = " ".join(model_parts) if model_parts else "Невідома модель"

    cpu_short = summarize_cpu_model(processor_info.get("name", ""))
    storage_bytes = 0
    if disks:
        system_drive = os.environ.get("SystemDrive", "C:").upper()
        system_disk = next(
            (disk for disk in disks if system_drive in (drive.upper() for drive in disk.drive_letters)),
            None,
        )
        chosen_disk = system_disk or max(disks, key=lambda item: item.size_bytes, default=None)
        if chosen_disk:
            storage_bytes = chosen_disk.size_bytes
    summary = "/".join(
        item
        for item in [
            cpu_short,
            format_memory_gb(total_memory),
            format_bytes_gb(storage_bytes) if storage_bytes else "",
        ]
        if item and item != "н/д"
    )
    serial = system_info.get("serial_number") or "н/д"
    return f"{manufacturer} {model} ({summary}) зав. № {serial}".strip()


def format_memory_lines(modules: list[MemoryModule]) -> list[str]:
    if not modules:
        return ["- Модулі пам'яті не вдалося зчитати."]

    lines: list[str] = []
    sorted_modules = sorted(
        modules,
        key=lambda item: (
            clean_text(item.locator or item.bank_locator).casefold(),
            item.manufacturer.casefold(),
            item.part_number.casefold(),
            item.serial_number.casefold(),
        ),
    )
    for index, module in enumerate(sorted_modules, start=1):
        locators = unique_preserve(
            [module.locator, module.bank_locator]
        )
        details = [
            f"Memory {index}",
            format_memory_gb_spaced(module.size_bytes),
            module.manufacturer or "Manufacturer: н/д",
            f"PartNumber: {module.part_number or 'н/д'}",
            f"Serial: {module.serial_number or 'н/д'}",
        ]
        if module.memory_type:
            details.append(module.memory_type)
        if module.speed_mhz:
            details.append(f"{module.speed_mhz} MT/s")
        if locators:
            details.append("Slots: " + ", ".join(locators))
        lines.append("- " + " | ".join(details))
    return lines


def format_disk_lines(disks: list[DiskDevice]) -> list[str]:
    if not disks:
        return ["- Фізичні диски не вдалося зчитати."]
    lines = []
    for disk in sorted(disks, key=lambda item: item.index):
        details = [
            f"Disk {disk.index}",
            disk.media_type,
            format_bytes_gb_spaced(disk.size_bytes),
            disk.model,
            f"Serial: {disk.serial_number or 'н/д'}",
        ]
        if disk.drive_letters:
            details.append("Volumes: " + ", ".join(disk.drive_letters))
        lines.append("- " + " | ".join(details))
    return lines


def gpu_discrete_label(gpu: GraphicsDevice) -> str:
    text = f"{gpu.name} {gpu.vendor} {gpu.video_processor}".casefold()
    if not text.strip():
        return "Discrete: н/д"
    if any(marker in text for marker in ("microsoft basic", "remote desktop", "parsec", "virtual")):
        return "Discrete: н/д (віртуальний/базовий адаптер)"
    if any(marker in text for marker in ("nvidia", "geforce", "quadro", "rtx", "gtx", "tesla")):
        return "Discrete: так"
    if any(marker in text for marker in ("radeon rx", "radeon pro", "firepro")):
        return "Discrete: так"
    if re.search(r"\barc\s*(pro\s*)?(a|b)\d", text) or re.search(r"\barc\(tm\)\s*(pro\s*)?(a|b)\d", text):
        return "Discrete: так"
    if "intel" in text and "arc" in text:
        return "Discrete: ні (інтегрований)"
    if any(marker in text for marker in ("intel uhd", "intel iris", "intel hd graphics", "iris xe")):
        return "Discrete: ні (інтегрований)"
    if any(marker in text for marker in ("radeon(tm) graphics", "radeon graphics", "amd radeon graphics")):
        return "Discrete: ні (інтегрований)"
    return "Discrete: н/д"


def format_graphics_lines(graphics: list[GraphicsDevice]) -> list[str]:
    if not graphics:
        return ["- GPU не вдалося зчитати."]

    lines: list[str] = []
    for index, gpu in enumerate(graphics, start=1):
        details = [
            f"GPU {index}",
            gpu.name or "н/д",
            gpu_discrete_label(gpu),
        ]
        if gpu.vendor and gpu.vendor.casefold() not in gpu.name.casefold():
            details.append(gpu.vendor)
        if gpu.video_processor and gpu.video_processor.casefold() not in gpu.name.casefold():
            details.append(f"Processor: {gpu.video_processor}")
        if gpu.adapter_ram_bytes:
            details.append(f"VRAM: {format_binary_capacity(gpu.adapter_ram_bytes)}")
        if gpu.current_resolution:
            details.append(f"Resolution: {gpu.current_resolution}")
        if gpu.driver_version:
            details.append(f"Driver: {gpu.driver_version}")
        if gpu.status:
            details.append(f"Status: {gpu.status}")
        lines.append("- " + " | ".join(details))
    return lines


def format_npu_lines(npus: list[NpuDevice]) -> list[str]:
    if not npus:
        return ["- NPU не знайдено або не вдалося зчитати."]

    lines: list[str] = []
    for index, npu in enumerate(npus, start=1):
        details = [f"NPU {index}", npu.name or "н/д"]
        if npu.manufacturer and npu.manufacturer.casefold() not in npu.name.casefold():
            details.append(npu.manufacturer)
        if npu.pnp_class:
            details.append(npu.pnp_class if npu.pnp_class.startswith("Service:") else f"Class: {npu.pnp_class}")
        if npu.status:
            details.append(f"Status: {npu.status}")
        if npu.pnp_device_id:
            details.append(f"DeviceID: {npu.pnp_device_id}")
        lines.append("- " + " | ".join(details))
    return lines


def format_eset_product_instance_lines(product_instance_lines: list[str]) -> list[str]:
    if not product_instance_lines:
        return []
    return [f"- {line}" for line in product_instance_lines]


def format_network_lines(adapters: list[NetworkAdapter]) -> list[str]:
    if not adapters:
        return ["- Активні мережеві адаптери не знайдено."]
    lines = []
    for adapter in adapters:
        meaningful_ipv4 = [ip for ip in adapter.ipv4 if not ip.startswith("169.254.")]
        if not adapter.is_up and not adapter.gateways and not meaningful_ipv4:
            continue
        details = [
            adapter.name or adapter.description or "Adapter",
            f"MAC {adapter.mac or 'н/д'}",
        ]
        if meaningful_ipv4:
            details.append("IPv4 " + ", ".join(meaningful_ipv4))
        elif adapter.ipv4 and adapter.is_up:
            details.append("IPv4 " + ", ".join(adapter.ipv4))
        if adapter.gateways:
            details.append("GW " + ", ".join(adapter.gateways))
        if adapter.description and adapter.description.casefold() != adapter.name.casefold():
            details.append(adapter.description)
        speed = format_speed_mbps(max(adapter.transmit_speed, adapter.receive_speed)) if adapter.is_up else ""
        if speed:
            details.append(speed)
        if adapter.is_virtual:
            details.append("virtual")
        lines.append("- " + " | ".join(details))
    return lines or ["- Активні мережеві адаптери не знайдено."]


def format_program_lines(entries: list[ProgramEntry]) -> list[str]:
    if not entries:
        return ["- Немає даних."]
    lines = []
    for entry in entries:
        details = [entry.name]
        if entry.version:
            details.append(f"v{entry.version}")
        if entry.install_date:
            details.append(f"інст. {entry.install_date}")
        if entry.publisher:
            details.append(entry.publisher)
        if entry.source:
            details.append(entry.source)
        lines.append("- " + " | ".join(details))
    return lines


def format_users(users: list[LocalUser]) -> tuple[str, str, str]:
    admins = [user.name for user in users if user.is_admin and not user.is_system]
    regular = [user.name for user in users if not user.is_admin and not user.is_system]
    system = [user.name for user in users if user.is_system]
    return (
        ", ".join(admins) if admins else "не знайдено",
        ", ".join(regular) if regular else "не знайдено",
        ", ".join(system) if system else "не знайдено",
    )


def format_section_header(title: str, width: int = 60) -> str:
    label = clean_text(title).upper()
    if not label:
        return "-" * width
    padded = f" {label} "
    if len(padded) >= width:
        return f"---- {label}"
    left = (width - len(padded)) // 2
    right = width - len(padded) - left
    return f"{'-' * left}{padded}{'-' * right}"


def format_license_lines(license_info: LicenseInfo) -> list[str]:
    return [
        (
            f"- Тип: {license_info.type_guess} | "
            f"Статус: {license_info.status or 'н/д'} | "
            f"Канал: {license_info.channel or 'н/д'} | "
            f"PartialProductKey: {license_info.partial_product_key or 'н/д'}"
        ),
        (
            f"- BackupProductKeyDefault: {license_info.backup_product_key or 'н/д'} | "
            f"OA3/OEM ProductKey: {license_info.oa3_product_key or 'н/д'}"
        ),
        (
            f"- BackupKey tail: {license_info.backup_key_tail or 'н/д'} | "
            f"OA3/OEM tail: {license_info.oa3_key_tail or 'н/д'}"
        ),
    ]


def build_report(
    system_info: dict[str, str],
    os_info: dict[str, str],
    processor_info: dict[str, str],
    total_memory: int,
    memory_modules: list[MemoryModule],
    disks: list[DiskDevice],
    graphics_devices: list[GraphicsDevice],
    npu_devices: list[NpuDevice],
    adapters: list[NetworkAdapter],
    eset_product_instance_lines: list[str],
    users: list[LocalUser],
    common_programs: list[ProgramEntry],
    per_user_programs: dict[str, list[ProgramEntry]],
    user_program_notes: dict[str, str],
    license_info: LicenseInfo,
    ctx: AuditContext,
    report_path: Path,
) -> str:
    admins_text, regular_text, system_text = format_users(users)
    run_time = dt.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    title = build_device_title(system_info, processor_info, total_memory, disks)

    processor_serial = pick_first_meaningful(
        processor_info.get("serial_string", ""),
        processor_info.get("processor_id", ""),
        processor_info.get("identifier", ""),
    )
    eset_lines = format_eset_product_instance_lines(eset_product_instance_lines)

    lines = [
        title,
        format_section_header("Запуск скрипта"),
        *format_launch_context_lines(),
        f"- Звіт згенеровано: {run_time}",
        f"- Файл звіту: {report_path}",
        format_section_header("Система"),
        f"SerialNumber: {system_info.get('serial_number') or 'н/д'}",
        f"SystemName: {os.environ.get('COMPUTERNAME') or platform.node() or 'н/д'}",
        (
            "Windows: "
            f"{os_info.get('product_name') or 'н/д'}, "
            f"{os_info.get('display_version') or 'н/д'}, "
            f"{os_info.get('install_date') or 'дата інсталяції н/д'}"
        ),
        (
            "Build: "
            f"{os_info.get('build_number') or 'н/д'}"
            + (f".{os_info.get('ubr')}" if os_info.get("ubr") else "")
            + f" | Version: {os_info.get('version_number') or 'н/д'}"
            + f" | Edition: {os_info.get('edition_id') or 'н/д'}"
            + f" | Arch: {os_info.get('architecture') or 'н/д'}"
        ),
        f"BIOS: {system_info.get('bios_version') or 'н/д'} | {system_info.get('bios_date') or 'дата н/д'}",
        format_section_header("Мережа та ESET" if eset_lines else "Мережа"),
        "MAC та IP:",
        *format_network_lines(adapters),
        *(["ProductInstanceID (ESET):", *eset_lines] if eset_lines else []),
        format_section_header("Апаратна частина"),
        (
            "Processor: "
            f"{processor_info.get('architecture') or 'н/д'} | "
            f"{processor_info.get('name') or 'н/д'} | "
            f"Serial/ID: {processor_serial or 'н/д'}"
        ),
        f"Memory: Разом {format_memory_gb_spaced(total_memory) if total_memory else 'н/д'}",
        *format_memory_lines(memory_modules),
        "Disk:",
        *format_disk_lines(disks),
        "GPU:",
        *format_graphics_lines(graphics_devices),
        "NPU:",
        *format_npu_lines(npu_devices),
        format_section_header("Ліцензія Windows"),
        *format_license_lines(license_info),
    ]

    lines.extend(
        [
            format_section_header("Користувачі"),
            f"- Адміністратори: {admins_text}",
            f"- Звичайні: {regular_text}",
            f"- Системні: {system_text}",
            format_section_header("Програми для всіх користувачів"),
            f"- Всього: {len(common_programs)}",
            *format_program_lines(common_programs),
            format_section_header("Персональні інсталяції по профілях"),
        ]
    )

    target_profiles: list[LocalUser] = []
    mindefence = next((user for user in users if user.name.casefold() == "mindefence"), None)
    if mindefence:
        target_profiles.append(mindefence)

    regular_targets = [
        user
        for user in users
        if not user.is_admin and not user.is_system and (user.enabled is None or user.enabled)
    ]
    for user in regular_targets:
        if mindefence and user.name.casefold() == mindefence.name.casefold():
            continue
        target_profiles.append(user)

    rendered_profiles = 0
    for user in target_profiles:
        entries = per_user_programs.get(user.name, [])
        note = user_program_notes.get(user.name)
        if not entries and note == PROFILE_NOT_FOUND_NOTE:
            continue
        lines.append(f"[{user.name}]")
        if entries:
            lines.extend(format_program_lines(entries))
        else:
            lines.append(f"- {note or 'Персональні інсталяції не знайдено або профіль недоступний.'}")
        rendered_profiles += 1

    if rendered_profiles == 0:
        lines.append("- Користувацькі профілі для детального скану не знайдено.")

    return "\n".join(lines) + "\n"


def emit_progress(
    progress_callback: ProgressCallback | None,
    header: str,
    detail: str,
    file_name: str = "",
    current: int = 0,
    total: int = 1,
) -> None:
    if progress_callback is None:
        return
    try:
        progress_callback(header, detail, file_name, current, total)
    except AuditCancelled:
        raise
    except Exception:
        return


def check_cancelled(cancel_event: threading.Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise AuditCancelled("Аудит зупинено користувачем.")


def run_audit(
    progress_callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> Path:
    total_steps = 10
    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 1/10: Система",
        "Зчитуємо SMBIOS, BIOS, процесор та модулі пам'яті.",
        "",
        0,
        total_steps,
    )
    ctx = AuditContext()
    smbios_records = load_smbios_records(ctx)
    system_info = extract_system_info(smbios_records)
    processor_info = extract_processor_info(smbios_records)
    memory_modules = extract_memory_modules(smbios_records)
    total_memory = get_total_physical_memory() or sum(module.size_bytes for module in memory_modules)

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 2/10: Диски",
        "Зчитуємо фізичні накопичувачі, серійники та томи.",
        "",
        1,
        total_steps,
    )
    disks = collect_disks(ctx)

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 3/10: GPU та NPU",
        "Збираємо графічні адаптери, дискретність GPU та NPU.",
        "",
        2,
        total_steps,
    )
    graphics_devices, npu_devices = collect_graphics_and_npu(ctx)

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 4/10: Мережа та ESET",
        "Зчитуємо MAC/IP і шукаємо ProductInstanceID у лозі ESET.",
        "",
        3,
        total_steps,
    )
    adapters = collect_network_adapters(ctx)
    eset_product_instance_lines = collect_eset_product_instance_lines(ctx)

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 5/10: Користувачі",
        "Збираємо локальні облікові записи та групу адміністраторів.",
        "",
        4,
        total_steps,
    )
    users = collect_local_users(ctx)
    mark_administrators(users, ctx)
    users.sort(key=lambda item: (0 if item.is_admin else 1, item.name.casefold()))

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 6/10: Програми",
        "Збираємо встановлені програми для пристрою та профілів.",
        "",
        5,
        total_steps,
    )
    common_programs, per_user_programs, user_program_notes = collect_programs(ctx, users)

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 7/10: Windows та ліцензія",
        "Збираємо дані Windows, build та стан ліцензії.",
        "",
        6,
        total_steps,
    )
    os_info = collect_os_info(ctx)
    license_info = collect_license_info(ctx)

    report_name = make_report_filename(system_info.get("serial_number", ""))
    report_path = Path.cwd() / report_name

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 8/10: Формування звіту",
        "Складаємо текст лог-файлу з усіх зібраних секцій.",
        "",
        7,
        total_steps,
    )
    report_text = build_report(
        system_info=system_info,
        os_info=os_info,
        processor_info=processor_info,
        total_memory=total_memory,
        memory_modules=memory_modules,
        disks=disks,
        graphics_devices=graphics_devices,
        npu_devices=npu_devices,
        adapters=adapters,
        eset_product_instance_lines=eset_product_instance_lines,
        users=users,
        common_programs=common_programs,
        per_user_programs=per_user_programs,
        user_program_notes=user_program_notes,
        license_info=license_info,
        ctx=ctx,
        report_path=report_path,
    )

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 9/10: Запис файлу",
        "Зберігаємо лог-файл на диск.",
        "",
        8,
        total_steps,
    )
    report_path.write_text(report_text, encoding="utf-8", errors="replace")

    check_cancelled(cancel_event)
    emit_progress(
        progress_callback,
        "Етап 10/10: Готово",
        "Звіт сформовано та збережено.",
        "",
        total_steps,
        total_steps,
    )
    return report_path


def finish_success(report_path: Path) -> int:
    message = f"Звіт збережено:\n{report_path}"
    print(message)
    return 0


def finish_error(exc: Exception | str) -> int:
    message = f"Не вдалося сформувати або зберегти звіт:\n{exc}"
    print(message, file=sys.stderr)
    maybe_show_message("AuditHSW - помилка", message, is_error=True)
    return 1


def finish_cancelled() -> int:
    stream = getattr(sys, "stderr", None) or getattr(sys, "stdout", None)
    if stream is not None:
        try:
            print("Аудит зупинено користувачем.", file=stream)
        except Exception:
            pass
    return 130


def should_show_progress_window() -> bool:
    if tk is None or ttk is None:
        return False
    if "--no-progress-window" in sys.argv:
        return False
    return True


def run_audit_with_progress_window() -> int:
    if tk is None or ttk is None:
        try:
            return finish_success(run_audit())
        except Exception as exc:
            return finish_error(exc)

    root = None
    start_window = None
    progress_window = None
    result_queue: Queue = Queue()
    cancel_event = threading.Event()
    terminal_result: dict[str, tuple[Any, ...] | None] = {"value": None}
    pending_progress: list[tuple[Any, ...]] = []
    progress_step_active = {"value": False}
    audit_started = {"value": False}
    min_progress_step_ms = 1000
    success_message_ms = 1500

    def close_progress_loop() -> None:
        try:
            if start_window is not None:
                start_window.close()
        except Exception:
            pass
        try:
            if progress_window is not None:
                progress_window.close()
        except Exception:
            pass
        if widget_exists(root):
            try:
                root.quit()
            except Exception:
                pass

    def cancel_run() -> None:
        if terminal_result["value"] is not None:
            close_progress_loop()
            return
        cancel_event.set()
        terminal_result["value"] = ("cancelled",)
        close_progress_loop()

    def schedule_after(delay_ms: int, callback: Callable[[], None]) -> None:
        if not widget_exists(root):
            return
        try:
            root.after(delay_ms, callback)
        except Exception:
            pass

    def show_success_then_close() -> None:
        if progress_window is not None:
            progress_window.update(
                header="Готово",
                detail="Звіт сформовано та збережено.",
                file_name="",
                current=10,
                total=10,
            )
            schedule_after(success_message_ms, close_progress_loop)
            return
        close_progress_loop()

    def handle_terminal_when_ready() -> None:
        result = terminal_result["value"]
        if result is None:
            return
        if result[0] == "success":
            show_success_then_close()
        else:
            close_progress_loop()

    def show_next_progress_step() -> None:
        if terminal_result["value"] is not None and terminal_result["value"][0] != "success":
            close_progress_loop()
            return
        if not pending_progress:
            progress_step_active["value"] = False
            handle_terminal_when_ready()
            return

        progress_step_active["value"] = True
        _, header, detail, file_name, current, total = pending_progress.pop(0)
        if progress_window is not None:
            progress_window.update(
                header=header,
                detail=detail,
                file_name=file_name,
                current=current,
                total=total,
            )
        schedule_after(min_progress_step_ms, show_next_progress_step)

    try:
        root = tk.Tk()
        root.withdraw()
        install_frozen_executable_icon(root)
        install_dark_title_bar(root)
        start_window = AuditStartWindow(root, on_start=None, on_close=cancel_run)
    except Exception:
        destroy_widget(start_window.dialog if start_window is not None else None)
        destroy_widget(progress_window.dialog if progress_window is not None else None)
        destroy_widget(root)
        try:
            return finish_success(run_audit())
        except Exception as exc:
            return finish_error(exc)

    def push_progress(header: str, detail: str, file_name: str, current: int, total: int) -> None:
        if cancel_event.is_set():
            raise AuditCancelled("Аудит зупинено користувачем.")
        result_queue.put(("progress", header, detail, file_name, current, total))

    def worker() -> None:
        try:
            result_queue.put(("success", run_audit(push_progress, cancel_event)))
        except AuditCancelled:
            result_queue.put(("cancelled",))
        except Exception as exc:
            result_queue.put(("error", str(exc)))

    def poll_queue() -> None:
        terminal = None
        while True:
            try:
                item = result_queue.get_nowait()
            except Empty:
                break
            if item[0] == "progress":
                pending_progress.append(item)
                continue
            terminal = item
            break

        if terminal is not None:
            terminal_result["value"] = terminal
            if terminal[0] != "success":
                close_progress_loop()
                return

        if pending_progress and not progress_step_active["value"]:
            show_next_progress_step()
        elif terminal_result["value"] is not None and not progress_step_active["value"]:
            handle_terminal_when_ready()

        if terminal_result["value"] is None and widget_exists(root):
            schedule_after(100, poll_queue)

    def start_audit() -> None:
        nonlocal start_window, progress_window
        if terminal_result["value"] is not None or audit_started["value"]:
            return
        audit_started["value"] = True
        if start_window is not None:
            start_window.close()
            start_window = None
        try:
            progress_window = ProgressWindow(root, on_close=cancel_run)
            progress_window.update(
                header="Підготовка аудиту",
                detail="Ініціалізуємо перевірку пристрою.",
                file_name="",
                current=0,
                total=10,
            )
        except Exception:
            destroy_widget(progress_window.dialog if progress_window is not None else None)
            progress_window = None
        threading.Thread(target=worker, daemon=True).start()
        schedule_after(100, poll_queue)

    if start_window is not None:
        start_window.on_start = start_audit

    root.mainloop()
    if start_window is not None:
        start_window.close()
    if progress_window is not None:
        progress_window.close()
    destroy_widget(root)

    result = terminal_result["value"]
    if result is None:
        return finish_error("Вікно прогресу закрилося до завершення аудиту.")
    if result[0] == "cancelled":
        return finish_cancelled()
    if result[0] == "success":
        return finish_success(result[1])
    return finish_error(result[1])


def main() -> int:
    if should_show_progress_window():
        return run_audit_with_progress_window()

    try:
        return finish_success(run_audit())
    except Exception as exc:
        return finish_error(exc)


if __name__ == "__main__":
    raise SystemExit(main())
