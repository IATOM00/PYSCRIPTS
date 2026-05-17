from __future__ import annotations

import os, sys, time, queue, shutil, stat, logging, threading, tempfile, subprocess, errno
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple
import ctypes, difflib, unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

try:
    from send2trash import send2trash
except ImportError:
    send2trash = None


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
        from ctypes import wintypes

        wm_seticon = 0x0080
        icon_small = 0
        icon_big = 1
        gclp_hicon = -14
        gclp_hiconsm = -34

        shell32 = ctypes.WinDLL("shell32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        long_ptr = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long

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
            icon_count = shell32.ExtractIconExW(sys.executable, 0, large_icons, small_icons, 1)
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
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    tk = None
    tkfont = None
    filedialog = None
    messagebox = None
    ttk = None


APP_NAME = "DataSync"
MODE_ONE_WAY = "one_way"
MODE_TWO_WAY = "two_way"
MODE_LABELS = {
    MODE_ONE_WAY: "One-way",
    MODE_TWO_WAY: "Two-way",
}
MTIME_TOLERANCE_NS = 2_000_000_000
USB_SEARCH_MAX_DEPTH = 3
USB_SEARCH_MAX_DIRS_PER_DRIVE = 300
LOCAL_SEARCH_MAX_DEPTH = 5
LOCAL_SEARCH_MAX_DIRS_PER_ROOT = 2500
LOCAL_SEARCH_EXCLUDED_DRIVE_LETTERS = {"C"}
LOCAL_SEARCH_SNAPSHOT_TIMEOUT_SECONDS = 30
MATCH_SCORE_THRESHOLD = 0.88
COPY_BUFFER_SIZE = 4 * 1024 * 1024
COPY_UI_PUMP_INTERVAL = 0.05

SKIPPED_AUTODETECT_DIRS = {
    "$recycle.bin",
    "system volume information",
    "windows",
    "program files",
    "program files (x86)",
    "programdata",
    "appdata",
    "node_modules",
    "__pycache__",
    ".git",
    ".idea",
    ".vscode",
    "windowsapps",
}

SKIPPED_SYNC_DIRS = {
    ".dropbox.cache",
}

SKIPPABLE_SYNC_ERROR_WINERRORS = {87}
SKIPPABLE_SYNC_ERROR_ERRNOS = {errno.EINVAL}


@dataclass
class FolderSuggestion:
    source: Path
    target: Path
    score: float
    source_root: Path
    usb_root: Path


@dataclass
class PathInfo:
    key: str
    rel: Path
    path: Path
    kind: str
    size: int = 0
    mtime_ns: int = 0


@dataclass
class SyncAction:
    kind: str
    rel: Path
    description: str
    dst_path: Path
    dst_root: Path
    dst_side: str
    src_path: Optional[Path] = None
    src_side: str = ""
    replace_existing: bool = False
    conflict: bool = False


@dataclass
class SyncPlan:
    actions: List[SyncAction]
    skipped_files: int = 0
    skipped_dirs: int = 0


@dataclass
class SyncStats:
    created_dirs: int = 0
    copied_files: int = 0
    updated_files: int = 0
    deleted_paths: int = 0
    conflict_paths: int = 0
    skipped_files: int = 0
    skipped_dirs: int = 0
    staged_paths: int = 0
    trash_destination: str = ""
    log_file_name: str = ""
    errors: List[str] = field(default_factory=list)
    skipped_log: List[str] = field(default_factory=list)
    action_log: List[str] = field(default_factory=list)
    cancelled: bool = False

    @property
    def changed_count(self) -> int:
        return self.created_dirs + self.copied_files + self.updated_files + self.deleted_paths


class TrashStager:
    def __init__(self, enabled: bool, timestamp: str):
        if send2trash is None:
            raise RuntimeError(
                "Для обов'язкового відправлення log-файлу у кошик потрібен пакет send2trash. "
                "Встановіть його перед запуском DataSync."
            )
        self.enabled = enabled
        self.timestamp = timestamp
        self.root: Optional[Path] = None
        self.log_path: Optional[Path] = None
        self.log_file_name: str = ""
        self.staged_count = 0
        self.trash_destination = ""

    def trash_path(
        self,
        path: Path,
        sync_root: Path,
        side_label: str,
        ui_pump: Optional[Callable[[], None]] = None,
        should_cancel: Optional[Callable[[], bool]] = None,
    ) -> bool:
        if not self.enabled or not path.exists():
            return True

        ensure_not_cancelled(should_cancel)
        if not is_path_on_removable_drive(path):
            send_path_to_trash(path, should_cancel=should_cancel)
            self.staged_count += 1
            logging.info("Відправлено у кошик напряму: %s", path)
            return False

        staging_root = self._ensure_root()
        try:
            rel = path.relative_to(sync_root)
        except ValueError:
            rel = Path(path.name)

        destination = self._unique_path(staging_root / side_label / rel)
        destination.parent.mkdir(parents=True, exist_ok=True)

        try:
            if path.is_dir() and not path.is_symlink():
                copy_directory_tree(path, destination, ui_pump=ui_pump, should_cancel=should_cancel)
            else:
                copy_file_with_metadata(path, destination, ui_pump=ui_pump, should_cancel=should_cancel)
        except Exception:
            remove_path_if_exists(destination)
            raise
        self.staged_count += 1
        logging.info("Підготовлено USB backup до кошика: %s -> %s", path, destination)
        return True

    def finalize(self) -> None:
        if self.log_path is not None and self.log_path.exists():
            desktop_log_path = self._move_log_to_desktop()
            send2trash(str(desktop_log_path))
            logging.info("Log-файл відправлено в кошик окремим файлом з Desktop: %s", desktop_log_path.name)
            self.log_file_name = desktop_log_path.name
            self.log_path = None

        if not self.enabled:
            return

        if self.root is None or not self.root.exists():
            return
        if not any(self.root.iterdir()):
            self.root.rmdir()
            self.root = None
            return

        destination_name = self.root.name
        send2trash(str(self.root))
        self.trash_destination = destination_name
        logging.info("USB staging-папку відправлено в кошик: %s", destination_name)
        self.root = None

    def write_log(self, file_name: str, content: str) -> Path:
        # The run log must stay outside the USB staging folder so Recycle Bin receives it as a separate item.
        destination = self._unique_path(Path(tempfile.gettempdir()) / file_name)
        destination.write_text(content, encoding="utf-8")
        self.log_path = destination
        logging.info("Збережено log-файл запуску: %s", destination)
        return destination

    def _move_log_to_desktop(self) -> Path:
        if self.log_path is None or not self.log_path.exists():
            raise FileNotFoundError("Log-файл для кошика не знайдено.")

        desktop_dir = self._desktop_directory()
        desktop_dir.mkdir(parents=True, exist_ok=True)

        if self.log_path.parent.resolve() == desktop_dir.resolve():
            return self.log_path

        destination = self._unique_path(desktop_dir / self.log_path.name)
        shutil.move(str(self.log_path), str(destination))
        self.log_path = destination
        logging.info("Log-файл перенесено на Desktop перед кошиком: %s", destination)
        return destination

    def _ensure_root(self) -> Path:
        if self.root is not None:
            return self.root

        base = Path(tempfile.gettempdir()) / f"{APP_NAME} - USB - {self.timestamp}"
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

    @staticmethod
    def _desktop_directory() -> Path:
        user_profile = os.environ.get("USERPROFILE")
        if user_profile:
            return Path(user_profile) / "Desktop"
        return Path.home() / "Desktop"


def make_run_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H.%M")


def make_run_log_filename(timestamp: str) -> str:
    return f"{APP_NAME} - USB - {timestamp}.log"


def rel_key(rel: Path) -> str:
    return rel.as_posix().casefold()


def key_depth(key: str) -> int:
    if not key:
        return 0
    return key.count("/") + 1


def is_key_under(child_key: str, parent_key: str) -> bool:
    return child_key == parent_key or child_key.startswith(parent_key + "/")


def sort_info_top_down(info: PathInfo) -> Tuple[int, str]:
    return key_depth(info.key), info.key


def sort_info_bottom_up(info: PathInfo) -> Tuple[int, str]:
    return -key_depth(info.key), info.key


def normalize_folder_name(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).casefold()
    value = "".join(ch if ch.isalnum() else " " for ch in value)
    return " ".join(value.split())


def compact_folder_name(value: str) -> str:
    return normalize_folder_name(value).replace(" ", "")


def folder_match_score(left: str, right: str) -> float:
    left_norm = normalize_folder_name(left)
    right_norm = normalize_folder_name(right)
    if not left_norm or not right_norm:
        return 0.0
    if left_norm == right_norm:
        return 1.0

    left_compact = left_norm.replace(" ", "")
    right_compact = right_norm.replace(" ", "")
    if left_compact == right_compact:
        return 0.98

    shorter, longer = sorted((left_compact, right_compact), key=len)
    if len(shorter) >= 6 and shorter in longer:
        return 0.92

    return difflib.SequenceMatcher(None, left_compact, right_compact).ratio()


def is_skipped_autodetect_dir(path: Path) -> bool:
    name = path.name.casefold()
    if name in SKIPPED_AUTODETECT_DIRS:
        return True
    return name.startswith(".") and name not in {".", ".."}


def iter_folders_top_down(root: Path, max_depth: int, max_dirs: int) -> Iterable[Path]:
    if not root.exists() or not root.is_dir():
        return

    queue: List[Tuple[Path, int]] = [(root, 0)]
    emitted = 0
    while queue and emitted < max_dirs:
        current, depth = queue.pop(0)
        if depth >= max_depth:
            continue

        try:
            children = sorted(
                [Path(entry.path) for entry in os.scandir(current) if entry.is_dir(follow_symlinks=False)],
                key=lambda item: item.name.casefold(),
            )
        except OSError:
            continue

        for child in children:
            if is_skipped_autodetect_dir(child):
                continue
            emitted += 1
            yield child
            if emitted >= max_dirs:
                break
            queue.append((child, depth + 1))


def detect_removable_drives() -> List[Path]:
    if sys.platform != "win32":
        return []

    drive_removable = 2
    kernel32 = ctypes.windll.kernel32
    bitmask = kernel32.GetLogicalDrives()
    drives: List[Path] = []
    for idx in range(26):
        if not bitmask & (1 << idx):
            continue
        letter = chr(ord("A") + idx)
        root = Path(f"{letter}:\\")
        try:
            drive_type = kernel32.GetDriveTypeW(str(root))
        except Exception:
            continue
        if drive_type == drive_removable and root.exists():
            drives.append(root)
    return drives


def drive_root_key(root: Path) -> str:
    drive = root.drive
    if not drive:
        drive = str(root)[:2]
    if len(drive) == 2 and drive[1] == ":":
        return f"{drive[0].upper()}:\\"
    return str(root).casefold()


def local_search_roots(excluded_roots: Iterable[Path] = ()) -> List[Path]:
    if sys.platform != "win32":
        return []

    excluded = {drive_root_key(root) for root in excluded_roots}
    roots: List[Path] = []
    kernel32 = ctypes.windll.kernel32
    bitmask = kernel32.GetLogicalDrives()
    for idx in range(26):
        if not bitmask & (1 << idx):
            continue
        letter = chr(ord("A") + idx)
        if letter in LOCAL_SEARCH_EXCLUDED_DRIVE_LETTERS:
            continue
        root = Path(f"{letter}:\\")
        if drive_root_key(root) in excluded:
            continue
        try:
            if root.exists():
                roots.append(root)
        except OSError:
            continue
    return roots


def powershell_folder_snapshot(root: Path, max_depth: int, max_dirs: int) -> Optional[List[Path]]:
    if sys.platform != "win32":
        return None

    skip_names = ", ".join(
        "'{}'".format(name.replace("'", "''")) for name in sorted(SKIPPED_AUTODETECT_DIRS)
    )
    script = r"""
$ErrorActionPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$rootPath = $env:DataSync_AUTODETECT_ROOT
$maxDepth = [int]$env:DataSync_AUTODETECT_MAX_DEPTH
$maxDirs = [int]$env:DataSync_AUTODETECT_MAX_DIRS
$skip = @{}
foreach ($skipName in @(__SKIP_NAMES__)) {
    $skip[$skipName.ToLowerInvariant()] = $true
}
$queue = New-Object 'System.Collections.Generic.Queue[object]'
$queue.Enqueue([pscustomobject]@{ Path = $rootPath; Depth = 0 })
$emitted = 0
while ($queue.Count -gt 0 -and $emitted -lt $maxDirs) {
    $node = $queue.Dequeue()
    if ($node.Depth -ge $maxDepth) { continue }
    $children = Get-ChildItem -LiteralPath $node.Path -Directory -Force -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($child in $children) {
        $name = $child.Name.ToLowerInvariant()
        if ($skip.ContainsKey($name)) { continue }
        if ($name.StartsWith(".") -and $name -ne "." -and $name -ne "..") { continue }
        $emitted++
        [Console]::Out.WriteLine($child.FullName)
        if ($emitted -ge $maxDirs) { break }
        $queue.Enqueue([pscustomobject]@{ Path = $child.FullName; Depth = $node.Depth + 1 })
    }
}
""".replace("__SKIP_NAMES__", skip_names)
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    command = [
        "powershell.exe",
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        script,
    ]
    env = os.environ.copy()
    env["DataSync_AUTODETECT_ROOT"] = str(root)
    env["DataSync_AUTODETECT_MAX_DEPTH"] = str(max_depth)
    env["DataSync_AUTODETECT_MAX_DIRS"] = str(max_dirs)
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=LOCAL_SEARCH_SNAPSHOT_TIMEOUT_SECONDS,
            creationflags=creationflags,
            env=env,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        logging.debug("PowerShell-знімок папок для автодобору не вдався: %s", exc)
        return None

    if completed.returncode != 0:
        logging.debug(
            "PowerShell-знімок папок для автодобору повернув код %s: %s",
            completed.returncode,
            completed.stderr.strip(),
        )
        return None

    folders: List[Path] = []
    for line in completed.stdout.splitlines():
        if not line:
            continue
        folder = Path(line)
        if not is_skipped_autodetect_dir(folder):
            folders.append(folder)
    return folders


def iter_local_folders_top_down(root: Path, max_depth: int, max_dirs: int) -> Iterable[Path]:
    snapshot = powershell_folder_snapshot(root, max_depth, max_dirs)
    if snapshot is not None:
        yield from snapshot[:max_dirs]
        return

    yield from iter_folders_top_down(root, max_depth, max_dirs)


def find_autodetect_suggestion() -> Optional[FolderSuggestion]:
    usb_drives = detect_removable_drives()
    roots = local_search_roots(excluded_roots=usb_drives)
    if not usb_drives or not roots:
        return None

    # `local_search_roots()` returns available non-C: drive roots in A..Z priority order.
    for root in roots:
        local_candidates: List[Tuple[Path, Path, int]] = []
        for folder in iter_local_folders_top_down(root, LOCAL_SEARCH_MAX_DEPTH, LOCAL_SEARCH_MAX_DIRS_PER_ROOT):
            try:
                depth = len(folder.relative_to(root).parts)
            except ValueError:
                depth = 99
            local_candidates.append((folder, root, depth))

        if not local_candidates:
            continue

        best: Optional[FolderSuggestion] = None
        best_rank: Optional[Tuple[int, int, float, str]] = None

        for usb_root in usb_drives:
            for usb_folder in iter_folders_top_down(usb_root, USB_SEARCH_MAX_DEPTH, USB_SEARCH_MAX_DIRS_PER_DRIVE):
                try:
                    usb_depth = len(usb_folder.relative_to(usb_root).parts)
                except ValueError:
                    usb_depth = 99

                for local_folder, local_root, local_depth in local_candidates:
                    score = folder_match_score(usb_folder.name, local_folder.name)
                    if score < MATCH_SCORE_THRESHOLD:
                        continue
                    rank = (usb_depth, local_depth, -score, str(local_folder).casefold())
                    if best_rank is None or rank < best_rank:
                        best_rank = rank
                        best = FolderSuggestion(
                            source=local_folder,
                            target=usb_folder,
                            score=score,
                            source_root=local_root,
                            usb_root=usb_root,
                        )

        if best is not None:
            return best

    return None


def build_snapshot(root: Path) -> Dict[str, PathInfo]:
    snapshot: Dict[str, PathInfo] = {}

    for path in root.rglob("*"):
        try:
            rel_for_skip = path.relative_to(root)
        except ValueError:
            rel_for_skip = path
        if has_skipped_sync_parent_dir(rel_for_skip):
            logging.info("Пропущено службовий об'єкт синхронізації: %s", path)
            continue
        try:
            if path.is_dir() and not path.is_symlink():
                kind = "dir"
                size = 0
            elif path.is_file():
                kind = "file"
                size = path.stat().st_size
            else:
                logging.info("Пропущено нестандартний об'єкт: %s", path)
                continue

            stat_result = path.stat()
            rel = path.relative_to(root)
            if kind == "dir" and is_skipped_sync_dir_name(rel.name):
                logging.info("РџСЂРѕРїСѓС‰РµРЅРѕ СЃР»СѓР¶Р±РѕРІСѓ РїР°РїРєСѓ СЃРёРЅС…СЂРѕРЅС–Р·Р°С†С–С—: %s", path)
                continue
        except OSError as exc:
            logging.warning("Не вдалося прочитати '%s': %s", path, exc)
            continue

        key = rel_key(rel)
        snapshot[key] = PathInfo(
            key=key,
            rel=rel,
            path=path,
            kind=kind,
            size=size,
            mtime_ns=stat_result.st_mtime_ns,
        )

    return snapshot


def files_equivalent(source: PathInfo, target: PathInfo) -> bool:
    if source.kind != "file" or target.kind != "file":
        return False
    if source.size != target.size:
        return False
    return abs(source.mtime_ns - target.mtime_ns) <= MTIME_TOLERANCE_NS


def choose_newer_side(source: PathInfo, target: PathInfo) -> str:
    if source.mtime_ns > target.mtime_ns + MTIME_TOLERANCE_NS:
        return "SOURCE"
    if target.mtime_ns > source.mtime_ns + MTIME_TOLERANCE_NS:
        return "TARGET"
    return "SOURCE"


def inspect_relative_path(root: Path, rel: Path) -> Optional[PathInfo]:
    path = root / rel
    if has_skipped_sync_parent_dir(rel):
        logging.info("Пропущено службовий об'єкт синхронізації: %s", path)
        return None
    try:
        stat_result = path.stat()
        if path.is_dir() and not path.is_symlink():
            if is_skipped_sync_dir_name(path.name):
                logging.info("РџСЂРѕРїСѓС‰РµРЅРѕ СЃР»СѓР¶Р±РѕРІСѓ РїР°РїРєСѓ СЃРёРЅС…СЂРѕРЅС–Р·Р°С†С–С—: %s", path)
                return None
            kind = "dir"
            size = 0
        elif path.is_file():
            kind = "file"
            size = stat_result.st_size
        else:
            logging.info("Пропущено нестандартний об'єкт: %s", path)
            return None
    except OSError as exc:
        logging.warning("Не вдалося прочитати '%s': %s", path, exc)
        return None

    return PathInfo(
        key=rel_key(rel),
        rel=rel,
        path=path,
        kind=kind,
        size=size,
        mtime_ns=stat_result.st_mtime_ns,
    )


def list_directory_entries(root: Path, rel: Path = Path()) -> Dict[str, PathInfo]:
    directory = root / rel
    try:
        entries = sorted(os.scandir(directory), key=lambda entry: entry.name.casefold())
    except OSError as exc:
        logging.warning("Не вдалося прочитати папку '%s': %s", directory, exc)
        return {}

    children: Dict[str, PathInfo] = {}
    for entry in entries:
        try:
            is_dir = entry.is_dir(follow_symlinks=False)
        except OSError as exc:
            logging.warning("Пропущено недоступний об'єкт '%s': %s", Path(entry.path), exc)
            continue
        if is_dir and is_skipped_sync_dir_name(entry.name):
            logging.info("Пропущено службову папку синхронізації: %s", Path(entry.path))
            continue
        child_rel = rel / entry.name if rel.parts else Path(entry.name)
        info = inspect_relative_path(root, child_rel)
        if info is not None:
            children[entry.name.casefold()] = info
    return children


def push_progress(progress: Any | None, header: str, detail: str, file_name: str = "") -> None:
    if progress:
        progress.update(header=header, detail=detail, file_name=format_progress_path_text(file_name))


def build_action_status(action: SyncAction) -> str:
    if action.kind == "copy" and action.src_path is not None:
        return f"{action.src_path} -> {action.dst_path}"
    return str(action.dst_path)


def build_action_log_entry(action: SyncAction) -> str:
    prefix = "[CONFLICT] " if action.conflict else ""
    if action.kind == "copy" and action.src_path is not None:
        return f"{prefix}{action.description}: {action.rel} | {action.src_side} -> {action.dst_side}"
    if action.kind == "delete":
        return f"{prefix}{action.description}: {action.rel} | {action.dst_side}"
    return f"{prefix}{action.description}: {action.rel}"


def is_skipped_sync_dir_name(name: str) -> bool:
    return name.startswith(".") or name.casefold() in SKIPPED_SYNC_DIRS


def has_skipped_sync_parent_dir(rel: Path) -> bool:
    return any(is_skipped_sync_dir_name(part) for part in rel.parts[:-1])


def is_skippable_sync_error(action: SyncAction, exc: BaseException) -> bool:
    if not isinstance(exc, OSError):
        return False

    winerror = getattr(exc, "winerror", None)
    errno_value = getattr(exc, "errno", None)
    return winerror in SKIPPABLE_SYNC_ERROR_WINERRORS or errno_value in SKIPPABLE_SYNC_ERROR_ERRNOS


def action_looks_like_directory(action: SyncAction) -> bool:
    if action.kind == "mkdir":
        return True
    for path in (action.src_path, action.dst_path):
        if path is None:
            continue
        try:
            if path.exists() and path.is_dir() and not path.is_symlink():
                return True
        except OSError:
            continue
    return False


def record_skipped_sync_error(action: SyncAction, stats: SyncStats, exc: BaseException) -> None:
    if action_looks_like_directory(action):
        stats.skipped_dirs += 1
    else:
        stats.skipped_files += 1

    message = f"{action.rel}: {exc}"
    stats.skipped_log.append(message)
    logging.warning("Пропущено недоступний службовий/некоректний об'єкт '%s': %s", action.rel, exc)


def run_sync_action(
    action: SyncAction,
    trash: TrashStager,
    stats: SyncStats,
    header: str,
    progress: Any | None = None,
) -> None:
    ensure_not_cancelled(progress.is_cancelled if progress else None)
    push_progress(progress, header, action.description, build_action_status(action))
    try:
        execute_action(
            action,
            trash,
            stats,
            ui_pump=progress.refresh if progress else None,
            should_cancel=progress.is_cancelled if progress else None,
        )
        stats.action_log.append(build_action_log_entry(action))
    except SyncCancelledError:
        raise
    except Exception as exc:
        if is_skippable_sync_error(action, exc):
            record_skipped_sync_error(action, stats, exc)
            return
        message = f"{action.rel}: {exc}"
        stats.errors.append(message)
        logging.exception("Помилка дії '%s': %s", action.description, message)


def sync_one_way_directory(
    source_root: Path,
    target_root: Path,
    rel: Path,
    trash: TrashStager,
    stats: SyncStats,
    header: str,
    progress: Any | None = None,
) -> None:
    ensure_not_cancelled(progress.is_cancelled if progress else None)
    current_source = source_root / rel
    push_progress(progress, header, "Сканування папки", f"SOURCE: {current_source}")

    source_children = list_directory_entries(source_root, rel)
    target_children = list_directory_entries(target_root, rel)

    for key in sorted(source_children):
        ensure_not_cancelled(progress.is_cancelled if progress else None)
        source_info = source_children[key]
        target_info = target_children.get(key)
        push_progress(progress, header, "Перевірка об'єкта", str(source_info.path))

        if source_info.kind == "dir":
            if target_info is None:
                run_sync_action(
                    SyncAction(
                        kind="mkdir",
                        rel=source_info.rel,
                        description="Створення папки у TARGET",
                        dst_path=target_root / source_info.rel,
                        dst_root=target_root,
                        dst_side="TARGET",
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )
            elif target_info.kind != "dir":
                run_sync_action(
                    SyncAction(
                        kind="mkdir",
                        rel=source_info.rel,
                        description="Заміна файла папкою у TARGET",
                        dst_path=target_root / source_info.rel,
                        dst_root=target_root,
                        dst_side="TARGET",
                        replace_existing=True,
                        conflict=True,
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )
            else:
                stats.skipped_dirs += 1

            sync_one_way_directory(
                source_root,
                target_root,
                source_info.rel,
                trash,
                stats,
                header,
                progress,
            )
            continue

        if target_info is None:
            run_sync_action(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Копіювання SOURCE -> TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target_root / source_info.rel,
                    dst_root=target_root,
                    dst_side="TARGET",
                ),
                trash,
                stats,
                header,
                progress,
            )
        elif target_info.kind != "file":
            run_sync_action(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Заміна папки файлом у TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target_root / source_info.rel,
                    dst_root=target_root,
                    dst_side="TARGET",
                    replace_existing=True,
                    conflict=True,
                ),
                trash,
                stats,
                header,
                progress,
            )
        elif files_equivalent(source_info, target_info):
            stats.skipped_files += 1
        else:
            run_sync_action(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Оновлення SOURCE -> TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target_root / source_info.rel,
                    dst_root=target_root,
                    dst_side="TARGET",
                    replace_existing=True,
                ),
                trash,
                stats,
                header,
                progress,
            )

    for key in sorted(target_children):
        if key in source_children:
            continue
        target_info = target_children[key]
        run_sync_action(
            SyncAction(
                kind="delete",
                rel=target_info.rel,
                description="Видалення зайвого з TARGET",
                dst_path=target_info.path,
                dst_root=target_root,
                dst_side="TARGET",
            ),
            trash,
            stats,
            header,
            progress,
        )


def sync_two_way_directory(
    source_root: Path,
    target_root: Path,
    rel: Path,
    trash: TrashStager,
    stats: SyncStats,
    header: str,
    progress: Any | None = None,
) -> None:
    ensure_not_cancelled(progress.is_cancelled if progress else None)
    current_source = source_root / rel
    current_target = target_root / rel
    push_progress(progress, header, "Сканування папки", f"SOURCE: {current_source} | TARGET: {current_target}")

    source_children = list_directory_entries(source_root, rel)
    target_children = list_directory_entries(target_root, rel)

    for key in sorted(set(source_children) | set(target_children)):
        ensure_not_cancelled(progress.is_cancelled if progress else None)
        source_info = source_children.get(key)
        target_info = target_children.get(key)
        current_path = (
            str(source_info.path)
            if source_info is not None
            else str(target_info.path)
            if target_info is not None
            else str(current_source)
        )
        push_progress(progress, header, "Перевірка об'єкта", current_path)

        if source_info is not None and target_info is not None:
            if source_info.kind == "dir" and target_info.kind == "dir":
                stats.skipped_dirs += 1
                sync_two_way_directory(
                    source_root,
                    target_root,
                    source_info.rel,
                    trash,
                    stats,
                    header,
                    progress,
                )
                continue

            if source_info.kind == "file" and target_info.kind == "file":
                if files_equivalent(source_info, target_info):
                    stats.skipped_files += 1
                    continue

                winner = choose_newer_side(source_info, target_info)
                if winner == "SOURCE":
                    run_sync_action(
                        SyncAction(
                            kind="copy",
                            rel=source_info.rel,
                            description="Оновлення SOURCE -> TARGET",
                            src_path=source_info.path,
                            src_side="SOURCE",
                            dst_path=target_root / source_info.rel,
                            dst_root=target_root,
                            dst_side="TARGET",
                            replace_existing=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
                else:
                    run_sync_action(
                        SyncAction(
                            kind="copy",
                            rel=target_info.rel,
                            description="Оновлення TARGET -> SOURCE",
                            src_path=target_info.path,
                            src_side="TARGET",
                            dst_path=source_root / target_info.rel,
                            dst_root=source_root,
                            dst_side="SOURCE",
                            replace_existing=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
                continue

            winner = choose_newer_side(source_info, target_info)
            if winner == "SOURCE":
                if source_info.kind == "dir":
                    run_sync_action(
                        SyncAction(
                            kind="mkdir",
                            rel=source_info.rel,
                            description="Конфлікт: SOURCE має папку, TARGET має файл",
                            dst_path=target_root / source_info.rel,
                            dst_root=target_root,
                            dst_side="TARGET",
                            replace_existing=True,
                            conflict=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
                    sync_two_way_directory(
                        source_root,
                        target_root,
                        source_info.rel,
                        trash,
                        stats,
                        header,
                        progress,
                    )
                else:
                    run_sync_action(
                        SyncAction(
                            kind="copy",
                            rel=source_info.rel,
                            description="Конфлікт: SOURCE перезаписує TARGET",
                            src_path=source_info.path,
                            src_side="SOURCE",
                            dst_path=target_root / source_info.rel,
                            dst_root=target_root,
                            dst_side="TARGET",
                            replace_existing=True,
                            conflict=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
            else:
                if target_info.kind == "dir":
                    run_sync_action(
                        SyncAction(
                            kind="mkdir",
                            rel=target_info.rel,
                            description="Конфлікт: TARGET має папку, SOURCE має файл",
                            dst_path=source_root / target_info.rel,
                            dst_root=source_root,
                            dst_side="SOURCE",
                            replace_existing=True,
                            conflict=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
                    sync_two_way_directory(
                        source_root,
                        target_root,
                        target_info.rel,
                        trash,
                        stats,
                        header,
                        progress,
                    )
                else:
                    run_sync_action(
                        SyncAction(
                            kind="copy",
                            rel=target_info.rel,
                            description="Конфлікт: TARGET перезаписує SOURCE",
                            src_path=target_info.path,
                            src_side="TARGET",
                            dst_path=source_root / target_info.rel,
                            dst_root=source_root,
                            dst_side="SOURCE",
                            replace_existing=True,
                            conflict=True,
                        ),
                        trash,
                        stats,
                        header,
                        progress,
                    )
            continue

        if source_info is not None:
            if source_info.kind == "dir":
                run_sync_action(
                    SyncAction(
                        kind="mkdir",
                        rel=source_info.rel,
                        description="Створення папки SOURCE -> TARGET",
                        dst_path=target_root / source_info.rel,
                        dst_root=target_root,
                        dst_side="TARGET",
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )
                sync_two_way_directory(
                    source_root,
                    target_root,
                    source_info.rel,
                    trash,
                    stats,
                    header,
                    progress,
                )
            else:
                run_sync_action(
                    SyncAction(
                        kind="copy",
                        rel=source_info.rel,
                        description="Копіювання SOURCE -> TARGET",
                        src_path=source_info.path,
                        src_side="SOURCE",
                        dst_path=target_root / source_info.rel,
                        dst_root=target_root,
                        dst_side="TARGET",
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )
            continue

        if target_info is not None:
            if target_info.kind == "dir":
                run_sync_action(
                    SyncAction(
                        kind="mkdir",
                        rel=target_info.rel,
                        description="Створення папки TARGET -> SOURCE",
                        dst_path=source_root / target_info.rel,
                        dst_root=source_root,
                        dst_side="SOURCE",
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )
                sync_two_way_directory(
                    source_root,
                    target_root,
                    target_info.rel,
                    trash,
                    stats,
                    header,
                    progress,
                )
            else:
                run_sync_action(
                    SyncAction(
                        kind="copy",
                        rel=target_info.rel,
                        description="Копіювання TARGET -> SOURCE",
                        src_path=target_info.path,
                        src_side="TARGET",
                        dst_path=source_root / target_info.rel,
                        dst_root=source_root,
                        dst_side="SOURCE",
                    ),
                    trash,
                    stats,
                    header,
                    progress,
                )


def top_level_infos(
    snapshot: Dict[str, PathInfo],
    keys: Iterable[str],
    skip_under_keys: Iterable[str] = (),
) -> List[PathInfo]:
    selected: List[PathInfo] = []
    selected_keys: List[str] = []
    skip_keys = list(skip_under_keys)

    for key in sorted(keys, key=lambda item: (key_depth(item), item)):
        if any(is_key_under(key, skip_key) for skip_key in skip_keys):
            continue
        if any(is_key_under(key, selected_key) for selected_key in selected_keys):
            continue
        info = snapshot.get(key)
        if info is None:
            continue
        selected.append(info)
        selected_keys.append(key)

    return selected


def build_one_way_plan(source: Path, target: Path) -> SyncPlan:
    source_snapshot = build_snapshot(source)
    target_snapshot = build_snapshot(target)
    actions: List[SyncAction] = []
    skipped_files = 0
    skipped_dirs = 0

    conflict_keys = {
        key
        for key in source_snapshot.keys() & target_snapshot.keys()
        if source_snapshot[key].kind != target_snapshot[key].kind
    }

    for source_info in sorted(
        [info for info in source_snapshot.values() if info.kind == "dir"],
        key=sort_info_top_down,
    ):
        target_info = target_snapshot.get(source_info.key)
        if target_info is None:
            actions.append(
                SyncAction(
                    kind="mkdir",
                    rel=source_info.rel,
                    description="Створення папки у TARGET",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                )
            )
        elif target_info.kind != "dir":
            actions.append(
                SyncAction(
                    kind="mkdir",
                    rel=source_info.rel,
                    description="Заміна файла папкою у TARGET",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                    replace_existing=True,
                    conflict=True,
                )
            )
        else:
            skipped_dirs += 1

    for source_info in sorted(
        [info for info in source_snapshot.values() if info.kind == "file"],
        key=sort_info_top_down,
    ):
        target_info = target_snapshot.get(source_info.key)
        if target_info is None:
            actions.append(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Копіювання SOURCE -> TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                )
            )
        elif target_info.kind != "file":
            actions.append(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Заміна папки файлом у TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                    replace_existing=True,
                    conflict=True,
                )
            )
        elif files_equivalent(source_info, target_info):
            skipped_files += 1
        else:
            actions.append(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Оновлення SOURCE -> TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                    replace_existing=True,
                )
            )

    target_extra_keys = target_snapshot.keys() - source_snapshot.keys()
    for target_info in top_level_infos(target_snapshot, target_extra_keys, skip_under_keys=conflict_keys):
        actions.append(
            SyncAction(
                kind="delete",
                rel=target_info.rel,
                description="Видалення зайвого з TARGET",
                dst_path=target_info.path,
                dst_root=target,
                dst_side="TARGET",
            )
        )

    return SyncPlan(actions=actions, skipped_files=skipped_files, skipped_dirs=skipped_dirs)


def build_two_way_plan(source: Path, target: Path) -> SyncPlan:
    source_snapshot = build_snapshot(source)
    target_snapshot = build_snapshot(target)
    actions: List[SyncAction] = []
    skipped_files = 0
    skipped_dirs = 0
    blocked_source_roots: List[str] = []
    blocked_target_roots: List[str] = []

    shared_keys = source_snapshot.keys() & target_snapshot.keys()
    conflict_keys = [
        key for key in shared_keys if source_snapshot[key].kind != target_snapshot[key].kind
    ]
    resolved_conflict_keys = set(conflict_keys)

    for key in sorted(conflict_keys, key=lambda item: (key_depth(item), item)):
        source_info = source_snapshot[key]
        target_info = target_snapshot[key]
        winner = choose_newer_side(source_info, target_info)

        if winner == "SOURCE":
            if source_info.kind == "dir":
                actions.append(
                    SyncAction(
                        kind="mkdir",
                        rel=source_info.rel,
                        description="Конфлікт: SOURCE має папку, TARGET має файл",
                        dst_path=target / source_info.rel,
                        dst_root=target,
                        dst_side="TARGET",
                        replace_existing=True,
                        conflict=True,
                    )
                )
            else:
                actions.append(
                    SyncAction(
                        kind="copy",
                        rel=source_info.rel,
                        description="Конфлікт: SOURCE перезаписує TARGET",
                        src_path=source_info.path,
                        src_side="SOURCE",
                        dst_path=target / source_info.rel,
                        dst_root=target,
                        dst_side="TARGET",
                        replace_existing=True,
                        conflict=True,
                    )
                )
            blocked_target_roots.append(key)
        else:
            if target_info.kind == "dir":
                actions.append(
                    SyncAction(
                        kind="mkdir",
                        rel=target_info.rel,
                        description="Конфлікт: TARGET має папку, SOURCE має файл",
                        dst_path=source / target_info.rel,
                        dst_root=source,
                        dst_side="SOURCE",
                        replace_existing=True,
                        conflict=True,
                    )
                )
            else:
                actions.append(
                    SyncAction(
                        kind="copy",
                        rel=target_info.rel,
                        description="Конфлікт: TARGET перезаписує SOURCE",
                        src_path=target_info.path,
                        src_side="TARGET",
                        dst_path=source / target_info.rel,
                        dst_root=source,
                        dst_side="SOURCE",
                        replace_existing=True,
                        conflict=True,
                    )
                )
            blocked_source_roots.append(key)

    for source_info in sorted(
        [info for info in source_snapshot.values() if info.kind == "dir"],
        key=sort_info_top_down,
    ):
        if any(is_key_under(source_info.key, key) for key in blocked_source_roots):
            continue
        target_info = target_snapshot.get(source_info.key)
        if target_info is None:
            actions.append(
                SyncAction(
                    kind="mkdir",
                    rel=source_info.rel,
                    description="Створення папки SOURCE -> TARGET",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                )
            )
        elif target_info.kind == "dir":
            skipped_dirs += 1

    for target_info in sorted(
        [info for info in target_snapshot.values() if info.kind == "dir"],
        key=sort_info_top_down,
    ):
        if any(is_key_under(target_info.key, key) for key in blocked_target_roots):
            continue
        source_info = source_snapshot.get(target_info.key)
        if source_info is None:
            actions.append(
                SyncAction(
                    kind="mkdir",
                    rel=target_info.rel,
                    description="Створення папки TARGET -> SOURCE",
                    dst_path=source / target_info.rel,
                    dst_root=source,
                    dst_side="SOURCE",
                )
            )
        elif source_info.kind == "dir":
            skipped_dirs += 1

    all_file_keys = {
        key
        for key, info in source_snapshot.items()
        if info.kind == "file"
    } | {
        key
        for key, info in target_snapshot.items()
        if info.kind == "file"
    }

    for key in sorted(all_file_keys, key=lambda item: (key_depth(item), item)):
        if key in resolved_conflict_keys:
            continue

        source_info = source_snapshot.get(key)
        target_info = target_snapshot.get(key)

        source_blocked = any(is_key_under(key, blocked_key) for blocked_key in blocked_source_roots)
        target_blocked = any(is_key_under(key, blocked_key) for blocked_key in blocked_target_roots)

        if source_blocked:
            source_info = None
        if target_blocked:
            target_info = None

        if source_info is not None and source_info.kind != "file":
            source_info = None
        if target_info is not None and target_info.kind != "file":
            target_info = None

        if source_info is not None and target_info is not None:
            if files_equivalent(source_info, target_info):
                skipped_files += 1
                continue
            winner = choose_newer_side(source_info, target_info)
            if winner == "SOURCE":
                actions.append(
                    SyncAction(
                        kind="copy",
                        rel=source_info.rel,
                        description="Оновлення SOURCE -> TARGET",
                        src_path=source_info.path,
                        src_side="SOURCE",
                        dst_path=target / source_info.rel,
                        dst_root=target,
                        dst_side="TARGET",
                        replace_existing=True,
                    )
                )
            else:
                actions.append(
                    SyncAction(
                        kind="copy",
                        rel=target_info.rel,
                        description="Оновлення TARGET -> SOURCE",
                        src_path=target_info.path,
                        src_side="TARGET",
                        dst_path=source / target_info.rel,
                        dst_root=source,
                        dst_side="SOURCE",
                        replace_existing=True,
                    )
                )
            continue

        if source_info is not None:
            actions.append(
                SyncAction(
                    kind="copy",
                    rel=source_info.rel,
                    description="Копіювання SOURCE -> TARGET",
                    src_path=source_info.path,
                    src_side="SOURCE",
                    dst_path=target / source_info.rel,
                    dst_root=target,
                    dst_side="TARGET",
                )
            )
        elif target_info is not None:
            actions.append(
                SyncAction(
                    kind="copy",
                    rel=target_info.rel,
                    description="Копіювання TARGET -> SOURCE",
                    src_path=target_info.path,
                    src_side="TARGET",
                    dst_path=source / target_info.rel,
                    dst_root=source,
                    dst_side="SOURCE",
                )
            )

    return SyncPlan(actions=actions, skipped_files=skipped_files, skipped_dirs=skipped_dirs)


def remove_readonly(func, path, exc_info) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        raise exc_info[1]


def remove_path(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path, onerror=remove_readonly)
    else:
        try:
            path.unlink()
        except PermissionError:
            os.chmod(path, stat.S_IWRITE)
            path.unlink()


def temporary_copy_path(destination: Path) -> Path:
    parent = destination.parent
    stem = destination.name
    for idx in range(1000):
        suffix = f".DataSync_tmp_{os.getpid()}_{idx}"
        candidate = parent / f"{stem}{suffix}"
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"Не вдалося створити тимчасовий шлях для копіювання: {destination}")


def pump_ui(ui_pump: Optional[Callable[[], None]]) -> None:
    if ui_pump is None:
        return
    try:
        ui_pump()
    except Exception:
        pass


def copy_file_with_metadata(
    source: Path,
    destination: Path,
    ui_pump: Optional[Callable[[], None]] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> None:
    ensure_not_cancelled(should_cancel)
    if ui_pump is None:
        ensure_not_cancelled(should_cancel)
        shutil.copy2(source, destination)
        ensure_not_cancelled(should_cancel)
        return

    destination.parent.mkdir(parents=True, exist_ok=True)
    last_pump = 0.0
    try:
        with source.open("rb") as src_handle, destination.open("wb") as dst_handle:
            while True:
                ensure_not_cancelled(should_cancel)
                chunk = src_handle.read(COPY_BUFFER_SIZE)
                if not chunk:
                    break
                dst_handle.write(chunk)
                now = time.monotonic()
                if now - last_pump >= COPY_UI_PUMP_INTERVAL:
                    pump_ui(ui_pump)
                    ensure_not_cancelled(should_cancel)
                    last_pump = now
        shutil.copystat(source, destination)
        pump_ui(ui_pump)
        ensure_not_cancelled(should_cancel)
    except Exception:
        remove_path_if_exists(destination)
        raise


def copy_directory_tree(
    source: Path,
    destination: Path,
    ui_pump: Optional[Callable[[], None]] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> None:
    ensure_not_cancelled(should_cancel)
    shutil.copytree(
        source,
        destination,
        symlinks=True,
        copy_function=lambda src, dst: copy_file_with_metadata(
            Path(src),
            Path(dst),
            ui_pump=ui_pump,
            should_cancel=should_cancel,
        ),
    )
    ensure_not_cancelled(should_cancel)


def copy_file_atomically(
    source: Path,
    destination: Path,
    ui_pump: Optional[Callable[[], None]] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> None:
    ensure_not_cancelled(should_cancel)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = temporary_copy_path(destination)
    try:
        copy_file_with_metadata(source, temp_path, ui_pump=ui_pump, should_cancel=should_cancel)
        try:
            os.replace(temp_path, destination)
        except PermissionError:
            if destination.exists():
                os.chmod(destination, stat.S_IWRITE)
            os.replace(temp_path, destination)
        ensure_not_cancelled(should_cancel)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


def execute_action(
    action: SyncAction,
    trash: TrashStager,
    stats: SyncStats,
    ui_pump: Optional[Callable[[], None]] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> None:
    ensure_not_cancelled(should_cancel)
    logging.info("%s: %s", action.description, action.rel)

    if action.kind == "mkdir":
        if action.dst_path.exists() and not action.dst_path.is_dir():
            should_remove_manually = trash.trash_path(
                action.dst_path,
                action.dst_root,
                action.dst_side,
                ui_pump=ui_pump,
                should_cancel=should_cancel,
            )
            if should_remove_manually:
                remove_path(action.dst_path)
            stats.updated_files += 1
        action.dst_path.mkdir(parents=True, exist_ok=True)
        stats.created_dirs += 1
        if action.conflict:
            stats.conflict_paths += 1
        return

    if action.kind == "delete":
        if action.dst_path.exists():
            should_remove_manually = trash.trash_path(
                action.dst_path,
                action.dst_root,
                action.dst_side,
                ui_pump=ui_pump,
                should_cancel=should_cancel,
            )
            if should_remove_manually:
                remove_path(action.dst_path)
            stats.deleted_paths += 1
        return

    if action.kind == "copy":
        if action.src_path is None:
            raise RuntimeError("Внутрішня помилка: copy-дія без src_path.")

        destination_exists = action.dst_path.exists()
        destination_is_dir = action.dst_path.is_dir() and not action.dst_path.is_symlink()
        should_remove_manually = False

        if destination_exists:
            should_remove_manually = trash.trash_path(
                action.dst_path,
                action.dst_root,
                action.dst_side,
                ui_pump=ui_pump,
                should_cancel=should_cancel,
            )

        if destination_is_dir and should_remove_manually:
            remove_path(action.dst_path)
        copy_file_atomically(
            action.src_path,
            action.dst_path,
            ui_pump=ui_pump,
            should_cancel=should_cancel,
        )

        if destination_exists or action.replace_existing:
            stats.updated_files += 1
        else:
            stats.copied_files += 1
        if action.conflict:
            stats.conflict_paths += 1
        return

    raise ValueError(f"Невідомий тип дії: {action.kind}")


def build_stage_header(stage_idx: int, total_stages: int, title: str) -> str:
    return f"Етап {stage_idx}/{total_stages}: {title}"


def validate_sync_roots(source: Path, target: Path) -> None:
    if not source.exists() or not source.is_dir():
        raise FileNotFoundError(f"SOURCE не існує або не є папкою: {source}")
    if not target.exists() or not target.is_dir():
        raise FileNotFoundError(f"TARGET не існує або не є папкою: {target}")

    source_resolved = source.resolve()
    target_resolved = target.resolve()
    if source_resolved == target_resolved:
        raise ValueError("SOURCE і TARGET не можуть бути однією й тією ж папкою.")
    if is_relative_to(source_resolved, target_resolved) or is_relative_to(target_resolved, source_resolved):
        raise ValueError("SOURCE і TARGET не можуть бути вкладеними одна в одну.")


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


class SyncCancelledError(RuntimeError):
    pass


def ensure_not_cancelled(should_cancel: Optional[Callable[[], bool]]) -> None:
    if should_cancel is not None and should_cancel():
        raise SyncCancelledError("Синхронізацію скасовано користувачем.")


def removable_drive_root(path: Path) -> Optional[Path]:
    if sys.platform != "win32":
        return None
    drive = path.drive
    if not drive:
        try:
            drive = path.resolve().drive
        except Exception:
            drive = ""
    if len(drive) == 2 and drive[1] == ":":
        return Path(f"{drive}\\")
    return None


def is_path_on_removable_drive(path: Path) -> bool:
    root = removable_drive_root(path)
    if root is None:
        return False
    try:
        return ctypes.windll.kernel32.GetDriveTypeW(str(root)) == 2
    except Exception:
        return False


def remove_path_if_exists(path: Path) -> None:
    if not path.exists():
        return
    remove_path(path)


def send_path_to_trash(path: Path, should_cancel: Optional[Callable[[], bool]] = None) -> None:
    ensure_not_cancelled(should_cancel)
    try:
        send2trash(str(path))
    except PermissionError:
        try:
            os.chmod(path, stat.S_IWRITE)
        except Exception:
            pass
        send2trash(str(path))
    ensure_not_cancelled(should_cancel)


def process_all(
    source: Path,
    target: Path,
    mode: str = MODE_ONE_WAY,
    send_to_trash_enabled: bool = True,
    progress: Any | None = None,
) -> SyncStats:
    source = source.expanduser().resolve()
    target = target.expanduser().resolve()
    validate_sync_roots(source, target)

    timestamp = make_run_timestamp()
    log_file_name = make_run_log_filename(timestamp)
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s", force=True)

    trash = TrashStager(send_to_trash_enabled, timestamp)
    stats = SyncStats()
    sync_header = "Синхронізація папок"
    if mode == MODE_ONE_WAY:
        mode_label = "One-way"
    elif mode == MODE_TWO_WAY:
        mode_label = "Two-way"
    else:
        raise ValueError(f"Невідомий режим синхронізації: {mode}")

    try:
        push_progress(
            progress,
            sync_header,
            "Послідовно перевіряємо папки та одразу застосовуємо зміни.",
            f"SOURCE: {source} | TARGET: {target} | Режим: {mode_label}",
        )
        ensure_not_cancelled(progress.is_cancelled if progress else None)

        if mode == MODE_ONE_WAY:
            sync_one_way_directory(source, target, Path(), trash, stats, sync_header, progress)
        else:
            sync_two_way_directory(source, target, Path(), trash, stats, sync_header, progress)
    except SyncCancelledError as exc:
        stats.cancelled = True
        logging.info("%s", exc)
    finally:
        stats.staged_paths = trash.staged_count
        try:
            trash.write_log(
                log_file_name,
                build_run_log_text(
                    stats=stats,
                    source=source,
                    target=target,
                    mode=mode,
                    started_at=timestamp,
                    send_to_trash_enabled=send_to_trash_enabled,
                ),
            )
        except Exception as exc:
            stats.errors.append(f"Log: {exc}")
            logging.exception("Не вдалося записати log-файл запуску: %s", exc)

        push_progress(progress, "Фіналізація кошика", "Завершуємо роботу з log-файлом і backup-версіями...", "")
        try:
            trash.finalize()
        except Exception as exc:
            stats.errors.append(f"Кошик: {exc}")
            logging.exception("Не вдалося завершити відправлення у кошик: %s", exc)

        stats.staged_paths = trash.staged_count
        stats.trash_destination = trash.trash_destination
        stats.log_file_name = trash.log_file_name

        push_progress(progress, "Фіналізація кошика", "Фіналізацію завершено", stats.trash_destination)

    return stats


class ProgressWindow:
    def __init__(self, owner: Any):
        if tk is None or ttk is None:
            raise RuntimeError("Tkinter недоступний.")

        self.cancel_requested = False
        self.close_requested = False
        self.dialog = tk.Toplevel(owner)
        self.dialog.withdraw()
        install_frozen_executable_icon(self.dialog)
        self.dialog.title(f"СИНХРОНІЗАЦІЯ: DataSync")
        self.dialog.resizable(False, False)
        self.dialog.attributes("-topmost", True)
        self.dialog.protocol("WM_DELETE_WINDOW", self.request_close)
        colors = configure_launch_styles(self.dialog)
        self.dialog.configure(bg=colors["window"])
        install_dark_title_bar(self.dialog)

        self.header_var = tk.StringVar(value="Підготовка...")
        self.detail_var = tk.StringVar(value="Послідовно перевіряємо папки та синхронізуємо зміни.")
        self.file_var = tk.StringVar(value="Будь ласка, зачекайте")
        self.badge_var = tk.StringVar(value="ОБРОБКА")
        self.activity_entries: List[str] = []
        self._last_activity_signature = ""

        container = ttk.Frame(self.dialog, style="LaunchRoot.TFrame")
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=colors["header"], padx=22, pady=10)
        header.pack(fill="x")
        header.grid_columnconfigure(0, weight=1)
        tk.Label(
            header,
            text=APP_NAME,
            bg=colors["header"],
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Синхронізація виконується. Будь ласка, дочекайтеся завершення.",
            bg=colors["header"],
            fg="#D7FBF5",
            font=("Segoe UI", 9),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(3, 4))
        badge_shell = tk.Frame(header, bg=colors["accent_dark"], padx=1, pady=1)
        self.badge_label = tk.Label(
            badge_shell,
            textvariable=self.badge_var,
            bg="#053D39",
            fg="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
            padx=10,
            pady=5,
        )
        self.badge_label.pack()
        badge_shell.place(relx=1.0, x=5, y=5, anchor="ne")

        body = ttk.Frame(container, style="LaunchRoot.TFrame", padding=(18, 16, 18, 18))
        body.pack(fill="both", expand=True)

        panel_shell = tk.Frame(body, bg=colors["border"], padx=1, pady=1)
        panel_shell.pack(fill="both", expand=True)
        panel = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 12))
        panel.pack(fill="both", expand=True)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(3, weight=1)

        self.progress = ttk.Progressbar(
            panel,
            orient="horizontal",
            mode="indeterminate",
            style="Launch.Horizontal.TProgressbar",
        )
        self.progress.grid(row=0, column=0, sticky="we", pady=(4, 8))
        ttk.Label(
            panel,
            text="Останні дії",
            style="ProgressSection.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 6))

        log_shell = tk.Frame(panel, bg=colors["border"], padx=1, pady=1)
        log_shell.grid(row=2, column=0, sticky="nsew")
        log_panel = tk.Frame(log_shell, bg=colors["panel"])
        log_panel.pack(fill="both", expand=True)
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(0, weight=1)

        self.activity_text = tk.Text(
            log_panel,
            height=9,
            wrap="word",
            bd=0,
            highlightthickness=0,
            relief="flat",
            bg=colors["panel"],
            fg=colors["text"],
            font=("Segoe UI", 9),
            padx=12,
            pady=10,
            state="disabled",
            cursor="arrow",
        )
        activity_font = tkfont.Font(font=self.activity_text.cget("font"))
        activity_indent = activity_font.measure("")
        self.activity_text.tag_configure(
            "activity_entry",
            lmargin1=activity_indent,
            lmargin2=0,
        )
        activity_scrollbar = ttk.Scrollbar(
            log_panel,
            orient="vertical",
            command=self.activity_text.yview,
            style="Launch.Vertical.TScrollbar",
        )
        self.activity_text.configure(yscrollcommand=activity_scrollbar.set)
        self.activity_text.grid(row=0, column=0, sticky="nsew")
        activity_scrollbar.grid(row=0, column=1, sticky="ns")
        self.progress.start(12)
        self._append_activity(self.detail_var.get(), self.file_var.get())

        self.dialog.update_idletasks()
        self.dialog.deiconify()
        self.dialog.lift()
        self.refresh()

    def update(self, header=None, detail=None, current=None, total=None, file_name=None):
        try:
            if self.close_requested or not _widget_exists(self.dialog):
                return
            if header is not None:
                self.header_var.set(header)
                self._update_badge(header)
            if detail is not None:
                self.detail_var.set(detail)
            if file_name is not None:
                self.file_var.set(file_name)
            if detail is not None or file_name is not None:
                self._append_activity(self.detail_var.get(), self.file_var.get())
            self.refresh()
        except Exception:
            return

    def _update_badge(self, header: str) -> None:
        header_lower = header.casefold()
        if "готово" in header_lower:
            self.badge_var.set("ГОТОВО")
        elif "помилк" in header_lower:
            self.badge_var.set("УВАГА")
        elif "скас" in header_lower:
            self.badge_var.set("СТОП")
        else:
            self.badge_var.set("ОБРОБКА")

    def _append_activity(self, detail: str, file_name: str) -> None:
        if not _widget_exists(self.activity_text):
            return

        detail_text = (detail or "").strip()
        file_text = (file_name or "").strip()
        if not detail_text and not file_text:
            return

        signature = f"{detail_text}\n{file_text}"
        if signature == self._last_activity_signature:
            return
        self._last_activity_signature = signature

        entry = self._format_activity_entry(detail_text, file_text)
        self.activity_entries.append(entry)
        self.activity_entries = self.activity_entries[-18:]

        self.activity_text.configure(state="normal")
        self.activity_text.delete("1.0", tk.END)
        self.activity_text.insert("1.0", "\n".join(self.activity_entries), ("activity_entry",))
        self.activity_text.configure(state="disabled")
        self.activity_text.see(tk.END)

    @staticmethod
    def _normalize_activity_title(detail: str) -> str:
        normalized_map = (
            ("Перевірка об'єкта", "Перевірка"),
            ("Сканування папки", "Сканування"),
            ("Копіювання", "Копіювання"),
            ("Оновлення", "Оновлення"),
            ("Створення папки", "Створення папки"),
            ("Видалення", "Видалення"),
            ("Заміна", "Заміна"),
            ("Конфлікт", "Конфлікт"),
            ("Завершуємо роботу з log-файлом", "Фіналізація кошика"),
            ("Фіналізацію завершено", "Фіналізація кошика"),
        )
        for prefix, label in normalized_map:
            if detail.startswith(prefix):
                return label
        return detail

    @staticmethod
    def _normalize_activity_path(detail: str, file_name: str) -> str:
        value = file_name.strip()
        if not value:
            return ""
        if " -> " in value and (
            detail.startswith("Копіювання")
            or detail.startswith("Оновлення")
            or detail.startswith("Конфлікт")
        ):
            value = value.split(" -> ", 1)[1].strip()
        return format_progress_path_text(value)

    def _format_activity_entry(self, detail: str, file_name: str) -> str:
        title = (self._normalize_activity_title(detail or file_name)).replace("\r", " ").replace("\n", " ").strip()
        path_value = self._normalize_activity_path(detail, file_name)
        if path_value:
            return f"{title}: {path_value}"
        return title

    def refresh(self):
        try:
            if self.close_requested or not _widget_exists(self.dialog):
                return
            self.dialog.update_idletasks()
            self.dialog.update()
        except Exception:
            pass

    def close(self):
        self.close_requested = True
        try:
            if _widget_exists(self.dialog):
                self.progress.stop()
                self.dialog.destroy()
        except Exception:
            pass

    def is_cancelled(self) -> bool:
        return self.cancel_requested or self.close_requested

    def request_close(self):
        if self.cancel_requested:
            return
        self.cancel_requested = True
        self.close()

    def show_success_then_close(
        self,
        detail: str,
        file_name: str = "",
        delay_ms: int = 1500,
        header: str = "Готово",
    ):
        if self.close_requested or not _widget_exists(self.dialog):
            return
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=100)
        self.progress["value"] = 100
        self.update(header=header, detail=detail, file_name=file_name)
        deadline = time.monotonic() + delay_ms / 1000.0
        while time.monotonic() < deadline:
            if not _widget_exists(self.dialog):
                return
            self.refresh()
            time.sleep(0.05)
        self.close()


class ProgressUpdateProxy:
    def __init__(
        self,
        updates: "queue.Queue[Dict[str, Any]]",
        should_cancel: Callable[[], bool] | None = None,
    ):
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

    def refresh(self) -> None:
        return

    def is_cancelled(self) -> bool:
        if self.should_cancel is None:
            return False
        try:
            return bool(self.should_cancel())
        except Exception:
            return False


def apply_pending_progress_updates(
    progress: ProgressWindow,
    updates: "queue.Queue[Dict[str, Any]]",
) -> None:
    while True:
        try:
            payload = updates.get_nowait()
        except queue.Empty:
            return
        progress.update(**payload)


def run_process_all_with_progress(
    source: Path,
    target: Path,
    mode: str,
    send_to_trash_enabled: bool,
    progress: ProgressWindow,
) -> SyncStats:
    updates: "queue.Queue[Dict[str, Any]]" = queue.Queue()
    result: Dict[str, Any] = {"stats": None, "error": None}
    progress_proxy = ProgressUpdateProxy(updates, should_cancel=progress.is_cancelled)

    def worker() -> None:
        try:
            result["stats"] = process_all(
                source,
                target,
                mode=mode,
                send_to_trash_enabled=send_to_trash_enabled,
                progress=progress_proxy,
            )
        except Exception as exc:
            result["error"] = exc

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    while thread.is_alive():
        apply_pending_progress_updates(progress, updates)
        progress.refresh()
        time.sleep(0.03)

    thread.join()
    apply_pending_progress_updates(progress, updates)
    progress.refresh()

    if result["error"] is not None:
        raise result["error"]
    if result["stats"] is None:
        raise RuntimeError("Синхронізація завершилася без результату.")
    return result["stats"]


ToggleSwitchBase = ttk.Frame if ttk is not None else object


def _resolve_widget_background(widget: Any, fallback: str = "#FFFFFF") -> str:
    if ttk is not None:
        try:
            style_name = widget.cget("style")
            if style_name:
                background = ttk.Style(widget).lookup(style_name, "background")
                if background:
                    return str(background)
        except Exception:
            pass

    parent = getattr(widget, "master", None)
    if parent is not None and ttk is not None:
        try:
            style_name = parent.cget("style")
            if style_name:
                background = ttk.Style(widget).lookup(style_name, "background")
                if background:
                    return str(background)
        except Exception:
            pass
        try:
            background = parent.cget("background")
            if background:
                return str(background)
        except Exception:
            pass

    if ttk is not None:
        try:
            background = ttk.Style(widget).lookup("TFrame", "background")
            if background:
                return str(background)
        except Exception:
            pass
    return fallback


class ToggleSwitch(ToggleSwitchBase):
    def __init__(self, owner: Any, variable: Any):
        super().__init__(owner)
        self.variable = variable
        self.state_var = tk.StringVar()

        try:
            parent_style = owner.cget("style")
            if parent_style:
                self.configure(style=parent_style)
        except Exception:
            pass
        canvas_bg = _resolve_widget_background(self)
        state_label_style = "ToggleState.TLabel"
        try:
            ttk.Style(self).configure(
                state_label_style,
                background="#FFFFFF",
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

        self.state_label = ttk.Label(self, textvariable=self.state_var, style=state_label_style)
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


class SyncOptionsSegmentControl(ToggleSwitchBase):
    segment_width = 100
    height = 40
    labels = ("One-way", "Two-way", "Trash", "No Trash")

    def __init__(self, owner: Any, mode_variable: Any, trash_variable: Any):
        super().__init__(owner)
        self.mode_variable = mode_variable
        self.trash_variable = trash_variable

        canvas_bg = _resolve_widget_background(self)
        width = self.segment_width * len(self.labels)

        self.canvas = tk.Canvas(
            self,
            width=width,
            height=self.height,
            bd=0,
            highlightthickness=0,
            bg=canvas_bg,
            cursor="hand2",
        )
        self.canvas.grid(row=0, column=0, sticky="w")
        self.canvas.bind("<Button-1>", self._on_click)

        self.mode_variable.trace_add("write", self._redraw)
        self.trash_variable.trace_add("write", self._redraw)
        self._redraw()

    def _on_click(self, event) -> None:
        segment = max(0, min(len(self.labels) - 1, event.x // self.segment_width))
        if segment == 0:
            self.mode_variable.set(MODE_ONE_WAY)
        elif segment == 1:
            self.mode_variable.set(MODE_TWO_WAY)
        elif segment == 2:
            self.trash_variable.set(True)
        else:
            self.trash_variable.set(False)

    def _redraw(self, *_args) -> None:
        selected_segments = {
            0 if self.mode_variable.get() == MODE_ONE_WAY else 1,
            2 if bool(self.trash_variable.get()) else 3,
        }
        self.canvas.delete("all")

        border = "#AEB7C2"
        selected = "#2E8B57"
        idle = "#FFFFFF"
        text_idle = "#1F2933"
        width = self.segment_width * len(self.labels)
        bottom = self.height - 2

        self.canvas.create_rectangle(2, 2, width - 2, bottom, fill=idle, outline=border, width=1)

        for idx, label in enumerate(self.labels):
            left = 3 + idx * self.segment_width
            right = 2 + (idx + 1) * self.segment_width
            center = left + (right - left) / 2
            is_selected = idx in selected_segments

            if is_selected:
                self.canvas.create_rectangle(left, 3, right, bottom - 1, fill=selected, outline=selected)

            if idx > 0:
                x = 2 + idx * self.segment_width
                line_color = selected if idx in selected_segments or idx - 1 in selected_segments else border
                self.canvas.create_line(x, 3, x, bottom - 1, fill=line_color)

            self.canvas.create_text(
                center,
                self.height / 2,
                text=label,
                fill="#FFFFFF" if is_selected else text_idle,
                font=("Segoe UI", 9, "bold" if is_selected else "normal"),
            )


def _widget_exists(widget: Any | None) -> bool:
    if tk is None or widget is None:
        return False
    try:
        return bool(widget.winfo_exists())
    except (tk.TclError, RuntimeError):
        return False


def _destroy_widget(widget: Any | None) -> None:
    if not _widget_exists(widget):
        return
    try:
        widget.destroy()
    except Exception:
        pass


def center_window(window: Any, width: int | None = None, height: int | None = None) -> None:
    if not _widget_exists(window):
        return
    window.update_idletasks()
    requested_width = window.winfo_reqwidth()
    requested_height = window.winfo_reqheight()
    resolved_width = max(width or requested_width, requested_width)
    resolved_height = max(height or requested_height, requested_height)
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = max(0, (screen_width - resolved_width) // 2)
    y = max(0, (screen_height - resolved_height) // 2)
    window.geometry(f"{resolved_width}x{resolved_height}+{x}+{y}")


def _prepare_dialog_parent(root: Any) -> None:
    if not _widget_exists(root):
        return
    try:
        try:
            current_grab = root.grab_current()
            if current_grab is not None and current_grab is not root:
                current_grab.grab_release()
        except Exception:
            pass
        if hasattr(root, "deiconify") and not root.winfo_ismapped():
            root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.update()
    except Exception:
        pass


def _show_topmost_message(kind: str, title: str, text: str, parent: Any | None = None) -> None:
    if tk is None or messagebox is None:
        stream = sys.stderr if kind == "error" else sys.stdout
        print(f"{title}: {text}", file=stream)
        return

    if _widget_exists(parent):
        _prepare_dialog_parent(parent)
        if kind == "info":
            messagebox.showinfo(title, text, parent=parent)
        elif kind == "warning":
            messagebox.showwarning(title, text, parent=parent)
        else:
            messagebox.showerror(title, text, parent=parent)
        return

    root = tk.Tk()
    root.withdraw()
    install_frozen_executable_icon(root)
    install_dark_title_bar(root)
    root.attributes("-topmost", True)
    try:
        if kind == "info":
            messagebox.showinfo(title, text, parent=root)
        elif kind == "warning":
            messagebox.showwarning(title, text, parent=root)
        else:
            messagebox.showerror(title, text, parent=root)
    finally:
        _destroy_widget(root)


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


def looks_like_display_path(value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    if value.startswith("\\\\"):
        return True
    if "\\" in value or "/" in value:
        return True
    return len(value) >= 2 and value[1] == ":" and value[0].isalpha()


def format_progress_path_text(value: str) -> str:
    value = value.strip()
    if not value:
        return ""

    if " | " in value:
        return " | ".join(format_progress_path_text(part) for part in value.split(" | "))

    if " -> " in value:
        return " -> ".join(format_progress_path_text(part) for part in value.split(" -> "))

    for prefix in ("SOURCE: ", "TARGET: ", "SOURCE ", "TARGET "):
        if value.startswith(prefix):
            return f"{prefix}{format_progress_path_text(value[len(prefix):])}"

    normalized = value.replace("\r", " ").replace("\n", " ").strip()
    if looks_like_display_path(normalized):
        return format_path_for_entry_display(normalized)
    return normalized


def fit_path_for_entry_display(path: str, entry: Any) -> str:
    display_path = format_path_for_entry_display(path)
    if not display_path:
        return ""

    entry.update_idletasks()
    font_name = entry.cget("font") or "TkDefaultFont"
    try:
        font = tkfont.nametofont(font_name)
    except tk.TclError:
        try:
            font = tkfont.nametofont("TkDefaultFont")
        except tk.TclError:
            return display_path

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


def refresh_path_entry_display(source_var: Any, display_var: Any, entry: Any) -> None:
    if not _widget_exists(entry):
        return
    display_var.set(fit_path_for_entry_display(source_var.get(), entry))


def choose_folder(root: Any, source_var: Any, display_var: Any, entry: Any) -> None:
    _prepare_dialog_parent(root)
    path = filedialog.askdirectory(parent=root)
    if path:
        source_var.set(path)
        refresh_path_entry_display(source_var, display_var, entry)
    _prepare_dialog_parent(root)


def configure_launch_styles(root: Any) -> Dict[str, str]:
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
        style.configure("LaunchTiny.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 8))
        style.configure("LaunchStatus.TLabel", background=colors["window"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("ProgressHeader.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 12, "bold"))
        style.configure("ProgressSection.TLabel", background=colors["panel"], foreground=colors["accent_dark"], font=("Segoe UI", 9, "bold"))
        style.configure("ProgressBody.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
        style.configure("ProgressFile.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
        style.configure("ProgressCount.TLabel", background=colors["panel"], foreground=colors["accent_dark"], font=("Segoe UI", 10, "bold"))
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
            "Launch.Horizontal.TScrollbar",
            gripcount=0,
            background="#E8EEF5",
            darkcolor="#E8EEF5",
            lightcolor="#E8EEF5",
            troughcolor=colors["panel"],
            bordercolor=colors["panel"],
            arrowcolor=colors["muted"],
            relief="flat",
            arrowsize=12,
        )
        style.configure("Launch.TRadiobutton", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
        style.map("Launch.TRadiobutton", background=[("active", colors["panel"])])
        style.configure(
            "Launch.Horizontal.TProgressbar",
            background=colors["header_badge"],
            troughcolor="#E8EEF5",
            bordercolor=colors["border"],
            lightcolor=colors["header_badge"],
            darkcolor=colors["header_badge"],
            thickness=12,
        )
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
            bordercolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
            lightcolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
            darkcolor=[("pressed", colors["accent_dark"]), ("active", colors["accent_dark"])],
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
            padding=(18, 10),
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


def run_gui() -> Optional[Tuple[Path, Path, str, bool]]:
    if tk is None or ttk is None or filedialog is None:
        raise RuntimeError("Tkinter недоступний: неможливо показати вікно налаштувань.")

    selected: Dict[str, object] = {
        "source": "",
        "target": "",
        "mode": MODE_ONE_WAY,
        "send_to_trash": False,
        "confirmed": False,
    }

    def on_start() -> None:
        selected["source"] = source_var.get().strip()
        selected["target"] = target_var.get().strip()
        selected["mode"] = mode_var.get()
        selected["send_to_trash"] = bool(trash_var.get())

        if not selected["source"] or not selected["target"]:
            _show_topmost_message("error", "Помилка", "Потрібно вказати SOURCE та TARGET.", parent=root)
            return
        if send2trash is None:
            _show_topmost_message(
                "error",
                "Помилка",
                "Пакет send2trash потрібен для обов'язкового відправлення log-файлу у кошик.",
                parent=root,
            )
            return

        try:
            validate_sync_roots(
                Path(str(selected["source"])).expanduser().resolve(),
                Path(str(selected["target"])).expanduser().resolve(),
            )
        except Exception as exc:
            _show_topmost_message("error", "Помилка", str(exc), parent=root)
            return

        selected["confirmed"] = True
        root.destroy()

    def on_cancel() -> None:
        selected["confirmed"] = False
        root.destroy()

    root = tk.Tk()
    root.withdraw()
    install_frozen_executable_icon(root)
    root.title(f"НАЛАШТУВАННЯ: DataSync")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    root.protocol("WM_DELETE_WINDOW", on_cancel)
    colors = configure_launch_styles(root)
    root.configure(bg=colors["window"])
    install_dark_title_bar(root)
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    source_var = tk.StringVar(value="")
    target_var = tk.StringVar(value="")
    source_display_var = tk.StringVar(value="")
    target_display_var = tk.StringVar(value="")
    mode_var = tk.StringVar(value=MODE_ONE_WAY)
    trash_var = tk.BooleanVar(value=False)
    autodetect_status_var = tk.StringVar(
        value='Натисніть "Автодобір", щоб спробувати знайти схожу пару між USB та ПК...'
    )
    summary_var = tk.StringVar(value="")
    autodetect_button: Any = None

    def add_path_row(
        parent: Any,
        row: int,
        title: str,
        hint: str,
        path_var: Any,
        display_var: Any,
    ) -> Any:
        label_stack = ttk.Frame(parent, style="LaunchPanel.TFrame")
        label_stack.grid(row=row, column=0, padx=(0, 14), pady=7, sticky="w")
        ttk.Label(label_stack, text=title, style="LaunchField.TLabel").pack(anchor="w")
        ttk.Label(label_stack, text=hint, style="LaunchTiny.TLabel").pack(anchor="w")

        entry = ttk.Entry(parent, textvariable=display_var, width=50, state="readonly", style="Launch.TEntry")
        entry.grid(row=row, column=1, padx=(0, 10), pady=7, sticky="we")
        ttk.Button(
            parent,
            text="ОБРАТИ",
            width=13,
            style="LaunchBrowse.TButton",
            command=lambda: choose_folder(root, path_var, display_var, entry),
        ).grid(row=row, column=2, pady=7, sticky="nsew")
        return entry

    def update_launch_state(*_args) -> None:
        source_ready = bool(source_var.get().strip())
        target_ready = bool(target_var.get().strip())
        ready = source_ready and target_ready

        if not ready:
            summary_var.set("Оберіть SOURCE та TARGET, щоб запустити синхронізацію...")
            return

        mode_label = "SOURCE -> TARGET" if mode_var.get() == MODE_ONE_WAY else "SOURCE <-> TARGET"
        trash_label = "backup у кошик увімкнено" if trash_var.get() else "backup у кошик вимкнено"
        summary_var.set(f"Готово до запуску: {mode_label}; {trash_label}; log завжди у кошик.")

    def refresh_source_display(*_args) -> None:
        refresh_path_entry_display(source_var, source_display_var, source_entry)
        update_launch_state()

    def refresh_target_display(*_args) -> None:
        refresh_path_entry_display(target_var, target_display_var, target_entry)
        update_launch_state()

    def on_autodetect() -> None:
        if autodetect_button is not None:
            autodetect_button.state(["disabled"])
        autodetect_status_var.set("Шукаємо схожу пару папок для SOURCE та TARGET...")
        root.update_idletasks()
        try:
            suggestion = find_autodetect_suggestion()
        except Exception as exc:
            autodetect_status_var.set("Автодобір не завершився. Перевірте папки вручну.")
            _show_topmost_message("error", "Помилка", f"Автодобір шляхів завершився помилкою:\n{exc}", parent=root)
            return
        finally:
            if autodetect_button is not None:
                autodetect_button.state(["!disabled"])

        if suggestion is None:
            autodetect_status_var.set("Автодобір не знайшов готову пару. Оберіть обидві папки вручну...")
            return

        source_var.set(str(suggestion.source))
        target_var.set(str(suggestion.target))
        source_name = suggestion.source.name or str(suggestion.source)
        target_name = suggestion.target.name or str(suggestion.target)
        autodetect_status_var.set(f"Автодобір заповнив схожу пару: {source_name} -> {target_name}.")

    root.bind("<Escape>", lambda _event: on_cancel())
    root.bind("<Return>", lambda _event: on_start())

    container = ttk.Frame(root, style="LaunchRoot.TFrame")
    container.grid(row=0, column=0, sticky="nsew")
    container.grid_columnconfigure(0, weight=1)

    header = tk.Frame(container, bg=colors["header"], padx=22, pady=12)
    header.grid(row=0, column=0, sticky="we")
    header.grid_columnconfigure(0, weight=1)
    tk.Label(
        header,
        text=APP_NAME,
        bg=colors["header"],
        fg="#FFFFFF",
        font=("Segoe UI Semibold", 16),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        header,
        text="Підготовка синхронізації між обома директоріями...",
        bg=colors["header"],
        fg="#D7FBF5",
        font=("Segoe UI", 9),
        anchor="w",
    ).grid(row=1, column=0, sticky="w", pady=(4, 6))
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

    body = ttk.Frame(container, style="LaunchRoot.TFrame", padding=(18, 16, 18, 18))
    body.grid(row=1, column=0, sticky="nsew")
    body.grid_columnconfigure(0, weight=1)

    panel_shell = tk.Frame(body, bg=colors["border"], padx=1, pady=1)
    panel_shell.grid(row=0, column=0, sticky="we")
    settings = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 12))
    settings.pack(fill="both", expand=True)
    settings.grid_columnconfigure(1, weight=1)
    settings.grid_columnconfigure(2, minsize=104)

    ttk.Label(settings, text="Папки:", style="LaunchSection.TLabel").grid(
        row=0, column=0, columnspan=2, pady=(0, 0), sticky="w"
    )
    ttk.Label(settings, textvariable=autodetect_status_var, style="LaunchMuted.TLabel", wraplength=520).grid(
        row=1, column=0, columnspan=2, padx=(0, 2), pady=(4, 9), sticky="w"
    )
    autodetect_button = ttk.Button(
        settings,
        text="АВТОДОБІР",
        width=13,
        style="LaunchSecondary.TButton",
        command=on_autodetect,
    )
    autodetect_button.grid(row=0, column=2, rowspan=2, pady=(4, 8), sticky="se")

    source_entry = add_path_row(settings, 2, "SOURCE:", "основна папка", source_var, source_display_var)
    target_entry = add_path_row(settings, 3, "TARGET:", "папка призначення", target_var, target_display_var)

    ttk.Separator(settings, orient="horizontal").grid(
        row=4, column=0, columnspan=3, sticky="we", pady=(10, 12)
    )
    ttk.Label(settings, text="Режим роботи", style="LaunchSection.TLabel").grid(
        row=5, column=0, columnspan=3, sticky="w"
    )

    mode_frame = ttk.Frame(settings, style="LaunchPanel.TFrame")
    mode_frame.grid(row=6, column=0, columnspan=3, sticky="we", pady=(8, 4))
    mode_frame.grid_columnconfigure(0, weight=1)
    mode_frame.grid_columnconfigure(1, weight=1)

    one_way_frame = ttk.Frame(mode_frame, style="LaunchPanel.TFrame")
    one_way_frame.grid(row=0, column=0, padx=(0, 18), sticky="nwe")
    ttk.Radiobutton(
        one_way_frame,
        text="One-way: SOURCE -> TARGET",
        variable=mode_var,
        value=MODE_ONE_WAY,
        style="Launch.TRadiobutton",
    ).grid(row=0, column=0, sticky="w")
    ttk.Label(
        one_way_frame,
        text="TARGET приводиться до стану SOURCE. Найкраще для резервної копії.",
        style="LaunchMuted.TLabel",
        wraplength=275,
    ).grid(row=1, column=0, padx=(22, 0), pady=(3, 4), sticky="w")

    two_way_frame = ttk.Frame(mode_frame, style="LaunchPanel.TFrame")
    two_way_frame.grid(row=0, column=1, sticky="nwe")
    ttk.Radiobutton(
        two_way_frame,
        text="Two-way: SOURCE <-> TARGET",
        variable=mode_var,
        value=MODE_TWO_WAY,
        style="Launch.TRadiobutton",
    ).grid(row=0, column=0, sticky="w")
    ttk.Label(
        two_way_frame,
        text="Обидві папки вирівнюються між собою. Конфлікти вирішуються новішим файлом.",
        style="LaunchMuted.TLabel",
        wraplength=275,
    ).grid(row=1, column=0, padx=(22, 0), pady=(3, 4), sticky="w")

    ttk.Separator(settings, orient="horizontal").grid(
        row=7, column=0, columnspan=3, sticky="we", pady=(12, 12)
    )
    trash_frame = ttk.Frame(settings, style="LaunchPanel.TFrame")
    trash_frame.grid(row=8, column=0, columnspan=3, sticky="we")
    trash_frame.grid_columnconfigure(0, weight=1)
    ttk.Label(trash_frame, text="Backup перед видаленням", style="LaunchField.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    trash_hint = (
        "Toggle керує тільки backup; log завжди переноситься на Desktop і йде в кошик окремим файлом."
        if send2trash is not None
        else "send2trash не знайдено: запуск недоступний, бо log завжди відправляється у кошик."
    )
    ttk.Label(trash_frame, text=trash_hint, style="LaunchMuted.TLabel", wraplength=550).grid(
        row=1, column=0, pady=(4, 4), sticky="w"
    )
    ToggleSwitch(trash_frame, trash_var).grid(row=0, column=1, rowspan=2, padx=(18, 0), sticky="e")

    footer = ttk.Frame(body, style="LaunchRoot.TFrame")
    footer.grid(row=1, column=0, pady=(10, 0), sticky="we")
    footer.grid_columnconfigure(0, weight=1)
    ttk.Label(footer, textvariable=summary_var, style="LaunchStatus.TLabel", wraplength=360).grid(
        row=0, column=0, padx=(0, 12), sticky="we"
    )

    actions = ttk.Frame(footer, style="LaunchRoot.TFrame")
    actions.grid(row=0, column=1, sticky="e")
    start_button = ttk.Button(
        actions,
        text="ЗАПУСТИТИ СИНХРОНІЗАЦІЮ",
        style="LaunchPrimary.TButton",
        command=on_start,
    )
    start_button.grid(row=0, column=0, sticky="e")

    source_var.trace_add("write", refresh_source_display)
    target_var.trace_add("write", refresh_target_display)
    mode_var.trace_add("write", update_launch_state)
    trash_var.trace_add("write", update_launch_state)
    source_entry.bind("<Configure>", lambda _event: refresh_path_entry_display(source_var, source_display_var, source_entry))
    target_entry.bind("<Configure>", lambda _event: refresh_path_entry_display(target_var, target_display_var, target_entry))

    refresh_path_entry_display(source_var, source_display_var, source_entry)
    refresh_path_entry_display(target_var, target_display_var, target_entry)
    update_launch_state()

    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.grab_set()
    root.focus_force()
    root.mainloop()

    if not bool(selected["confirmed"]):
        return None

    return (
        Path(str(selected["source"])).expanduser().resolve(),
        Path(str(selected["target"])).expanduser().resolve(),
        str(selected["mode"]),
        bool(selected["send_to_trash"]),
    )


def build_done_detail(stats: SyncStats) -> str:
    if stats.cancelled:
        return "Синхронізацію скасовано користувачем."
    if stats.errors:
        return f"Завершено з помилками: {len(stats.errors)}"
    if stats.changed_count == 0:
        return "Синхронізація завершена: змін не потрібно."
    return (
        "Синхронізація завершена: "
        f"скопійовано {stats.copied_files}, "
        f"оновлено {stats.updated_files}, "
        f"видалено {stats.deleted_paths}, "
        f"папок створено {stats.created_dirs}."
    )


def build_result_text(stats: SyncStats) -> str:
    lines = [
        build_done_detail(stats),
        f"Пропущено: файлів {stats.skipped_files}, папок {stats.skipped_dirs}.",
    ]
    if stats.staged_paths:
        lines.append(f"Backup-об'єктів передано до кошика: {stats.staged_paths}.")
    if stats.trash_destination:
        lines.append(f"USB backup-папка у кошику: {stats.trash_destination}.")
    if stats.log_file_name:
        lines.append(f"Log-файл у кошику: {stats.log_file_name}.")
    if stats.conflict_paths:
        lines.append(f"Конфліктів вирішено: {stats.conflict_paths}.")
    if stats.skipped_log:
        lines.append("")
        lines.append("Пропущені недоступні об'єкти:")
        lines.extend(stats.skipped_log[:5])
    if stats.errors:
        lines.append("")
        lines.append("Конфлікти та помилки:")
        lines.extend(stats.errors[:5])
    return "\n".join(lines)


def build_summary_block_lines(stats: SyncStats) -> List[str]:
    return [
        "Підсумок:",
        f"- Змінено об'єктів: {stats.changed_count}",
        f"- Створено папок: {stats.created_dirs}",
        f"- Скопійовано файлів: {stats.copied_files}",
        f"- Оновлено файлів: {stats.updated_files}",
        f"- Видалено об'єктів: {stats.deleted_paths}",
        f"- Пропущено: файлів {stats.skipped_files}, папок {stats.skipped_dirs}",
        f"- Конфліктів вирішено: {stats.conflict_paths}",
        f"- Помилок: {len(stats.errors)}",
        f"- Backup-об'єктів передано до кошика: {stats.staged_paths}",
    ]


def build_final_summary_text(stats: SyncStats) -> str:
    lines = [
        build_done_detail(stats),
        "",
        *build_summary_block_lines(stats),
    ]

    if stats.log_file_name:
        lines.append(f"- Log-файл: {stats.log_file_name}")
    if stats.trash_destination:
        lines.append(f"- USB backup-папка: {stats.trash_destination}")

    if stats.errors:
        lines.append("")
        lines.append("Конфлікти та помилки:")
        lines.extend(f"- {entry}" for entry in stats.errors[:5])
    if stats.skipped_log:
        lines.append("")
        lines.append("Пропущені недоступні об'єкти:")
        lines.extend(f"- {entry}" for entry in stats.skipped_log[:5])

    return "\n".join(lines)


def show_final_summary_window(parent: Any | None, stats: SyncStats) -> None:
    summary_text = build_final_summary_text(stats)
    if tk is None or ttk is None:
        print(f"Підсумок:\n{summary_text}")
        return

    owns_root = False
    if not _widget_exists(parent):
        parent = tk.Tk()
        parent.withdraw()
        install_frozen_executable_icon(parent)
        install_dark_title_bar(parent)
        owns_root = True

    dialog = tk.Toplevel(parent)
    dialog.withdraw()
    install_frozen_executable_icon(dialog)
    dialog.title(f"{APP_NAME} - Підсумок")
    dialog.resizable(False, False)
    dialog.attributes("-topmost", True)
    colors = configure_launch_styles(dialog)
    dialog.configure(bg=colors["window"])
    install_dark_title_bar(dialog)

    def close_dialog() -> None:
        _destroy_widget(dialog)

    dialog.protocol("WM_DELETE_WINDOW", close_dialog)
    dialog.bind("<Escape>", lambda _event: close_dialog())
    dialog.bind("<Return>", lambda _event: close_dialog())

    container = ttk.Frame(dialog, style="LaunchRoot.TFrame")
    container.pack(fill="both", expand=True)

    header = tk.Frame(container, bg=colors["header"], padx=22, pady=12)
    header.pack(fill="x")
    header.grid_columnconfigure(0, weight=1)
    tk.Label(
        header,
        text=APP_NAME,
        bg=colors["header"],
        fg="#FFFFFF",
        font=("Segoe UI Semibold", 15),
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    tk.Label(
        header,
        text="Підсумок виконаної синхронізації",
        bg=colors["header"],
        fg="#D7FBF5",
        font=("Segoe UI", 9),
        anchor="w",
    ).grid(row=1, column=0, sticky="w", pady=(3, 4))
    badge_shell = tk.Frame(header, bg=colors["accent_dark"], padx=1, pady=1)
    badge = tk.Label(
        badge_shell,
        text="ПІДСУМОК",
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
    panel = ttk.Frame(panel_shell, style="LaunchPanel.TFrame", padding=(16, 14, 16, 12))
    panel.pack(fill="both", expand=True)
    panel.grid_columnconfigure(0, weight=1)
    panel.grid_rowconfigure(1, weight=1)

    ttk.Label(panel, text="Підсумок:", style="LaunchSection.TLabel").grid(row=0, column=0, sticky="w")

    text_shell = tk.Frame(panel, bg=colors["border"], padx=1, pady=1)
    text_shell.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
    text_shell.grid_columnconfigure(0, weight=1)
    text_shell.grid_rowconfigure(0, weight=1)

    text = tk.Text(
        text_shell,
        width=76,
        height=12,
        wrap="word",
        bd=0,
        highlightthickness=0,
        relief="flat",
        bg=colors["panel"],
        fg=colors["text"],
        font=("Segoe UI", 9),
        padx=12,
        pady=10,
        cursor="arrow",
    )
    scrollbar = ttk.Scrollbar(
        text_shell,
        orient="vertical",
        command=text.yview,
        style="Launch.Vertical.TScrollbar",
    )
    text.configure(yscrollcommand=scrollbar.set)
    text.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    text.insert("1.0", summary_text)
    text.configure(state="disabled")

    footer = ttk.Frame(body, style="LaunchRoot.TFrame")
    footer.pack(fill="x", pady=(10, 0))
    ttk.Button(
        footer,
        text="ЗАКРИТИ",
        width=14,
        style="LaunchPrimary.TButton",
        command=close_dialog,
    ).pack(side="right")

    center_window(dialog, width=680, height=430)
    dialog.deiconify()
    dialog.lift()
    dialog.focus_force()
    dialog.grab_set()
    dialog.wait_window()

    if owns_root:
        _destroy_widget(parent)


def build_run_log_text(
    stats: SyncStats,
    source: Path,
    target: Path,
    mode: str,
    started_at: str,
    send_to_trash_enabled: bool,
) -> str:
    lines = [
        APP_NAME,
        f'Дата та час запуску: {started_at}',
        f'Статус: {build_done_detail(stats)}',
        f'SOURCE: "{source}"',
        f'TARGET: "{target}"',
        f'Режим: {MODE_LABELS.get(mode, mode)}',
        f'Backup у кошик: {"увімкнено" if send_to_trash_enabled else "вимкнено"}',
        '',
        *build_summary_block_lines(stats),
    ]

    if stats.log_file_name:
        lines.append(f"- Log-файл: {stats.log_file_name}")
    if stats.trash_destination:
        lines.append(f"- USB backup-папка: {stats.trash_destination}")

    if stats.action_log:
        lines.append("")
        lines.append("Фактично виконані зміни:")
        lines.extend(f"- {entry}" for entry in stats.action_log)

    if stats.errors:
        lines.append("")
        lines.append("Конфлікти та помилки:")
        lines.extend(f"- {entry}" for entry in stats.errors[:20])
    if stats.skipped_log:
        lines.append("")
        lines.append("Пропущені недоступні об'єкти:")
        lines.extend(f"- {entry}" for entry in stats.skipped_log[:20])

    return "\n".join(lines) + "\n"


def main() -> None:
    ui_root: Any | None = None
    progress: Optional[ProgressWindow] = None
    try:
        selection = run_gui()
        if selection is None:
            return

        source, target, mode, send_to_trash_enabled = selection
        ui_root = tk.Tk()
        ui_root.withdraw()
        install_frozen_executable_icon(ui_root)
        install_dark_title_bar(ui_root)
        ui_root.attributes("-topmost", True)
        progress = ProgressWindow(ui_root)

        stats = run_process_all_with_progress(
            source,
            target,
            mode,
            send_to_trash_enabled,
            progress,
        )

        if stats.cancelled:
            progress = None
            print(build_result_text(stats))
            return

        detail = f"Підсумок: {build_done_detail(stats)}"
        result_header = "Готово" if not stats.errors else "Завершено з помилками"
        result_file = f"Змін: {stats.changed_count}"
        if stats.errors:
            result_file += f" | Помилок: {len(stats.errors)}"
        progress.show_success_then_close(detail, file_name=result_file, header=result_header)
        progress = None
        show_final_summary_window(ui_root, stats)
        print(build_result_text(stats))
    except Exception as exc:
        logging.exception("Критична помилка: %s", exc)
        try:
            message_parent = ui_root
            if progress is not None and _widget_exists(progress.dialog):
                message_parent = progress.dialog
            _show_topmost_message("error", "Помилка", str(exc), parent=message_parent)
        except Exception:
            print(f"Помилка: {exc}")
    finally:
        if progress is not None:
            progress.close()
        if ui_root is not None:
            _destroy_widget(ui_root)


if __name__ == "__main__":
    main()
