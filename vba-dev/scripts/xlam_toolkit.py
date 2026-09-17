# -*- coding: utf-8 -*-
"""
xlam_toolkit.py - Excel Add-in Development Toolkit

Provides utilities for:
- Extracting/repackaging .xlam files
- Modifying Ribbon XML (customUI) — .xlam / .xlsm
- Injecting VBA code via win32com (.xlam, .xlsm, .xlsx)
  - For .xlsx files: auto-converts to .xlsm (original file unchanged)
- Creating UserForms and controls (.xlam, .xlsm, .xlsx)
- Source export/import round-trip (export_vba_source / import_vba_source)
- Build verification: static structure checks + real VBE compile (build_check)
- Running macros / throwaway test macros with error capture (run_macro / run_test)
- UserForm geometry linting (lint_form)
- Automatic backups before every write (backup_file / restore_backup)
- Environment check & one-click VBA trust enable (check_environment / enable_vba_trust)

Safety design:
- All COM automation runs in an ISOLATED Excel instance (DispatchEx).
  The user's running Excel is never touched, hidden, or quit.
- attach=True opts into reusing the user's running instance (read-style ops);
  attached instances are never Quit and their app settings are restored.
- Every write operation backs the file up to <file>.bak/<timestamp>_ first.
- A dialog watchdog reads + dismisses modal dialogs (compile errors, MsgBox)
  so automation cannot hang; a stuck isolated instance is force-killed after
  a timeout and reported.

Supported file formats:
  - .xlam: Excel Add-in (Ribbon customization supported)
  - .xlsm: Macro-enabled workbook
  - .xlsx: Standard workbook (auto-converted to .xlsm when adding VBA)
"""

import atexit
import contextlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as _xml_escape

import win32com.client

# GBK consoles cannot encode some unicode symbols — never crash on printing
try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

# Constants
XLAM_TEMP_DIR = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "xlam_toolkit_temp")
vbCrLf = "\r\n"  # VBA line break constant

# Expected closer text per block type (for readable error messages)
_BLOCK_CLOSER_TEXT = {
    "If": "End If", "For": "Next", "Do": "Loop", "While": "Wend",
    "Select": "End Select", "With": "End With", "#If": "#End If",
}


def _proc_of_line_result(raw):
    """ProcOfLine returns (name, prockind) tuple under pywin32 — normalize."""
    if isinstance(raw, tuple):
        return raw[0]
    return raw


def _proc_decl_re(proc_name: str):
    """Compiled regex matching a procedure DECLARATION line of `proc_name`
    (line-anchored, MULTILINE). Tolerates Public/Private/Friend/Static
    modifiers and Property Get/Let/Set. A plain substring scan both misses
    'Public Sub Foo' and false-positives on comments mentioning the name."""
    return re.compile(
        rf'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
        rf'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+'
        rf'{re.escape(proc_name)}\b',
        re.IGNORECASE | re.MULTILINE)


def _xml_attr(value) -> str:
    """Escape a value for use inside a double-quoted XML attribute.
    Handles & < > and " — a raw & or " in label/screentip would produce
    an invalid customUI.xml and Excel would refuse the WHOLE ribbon."""
    return _xml_escape(str(value), {'"': "&quot;"})


def _proc_type_of_decl(decl_line: str) -> str:
    """Classify a procedure from its declaration line: Sub/Function/Property.

    Must tolerate leading Public/Private/Friend/Static modifiers — the
    CodeModule declaration line includes them, so a naive
    startswith('Function') misfiles every 'Public Function ...' as Sub.
    Unknown shapes conservatively return 'Sub' (previous behavior)."""
    m = re.match(r"(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
                 r"(Sub|Function|Property)\b", decl_line.strip(), re.IGNORECASE)
    if not m:
        return "Sub"
    kw = m.group(1).lower()
    return {"sub": "Sub", "function": "Function"}.get(kw, "Property")


def _read_vba_text(src: str) -> str:
    """Read a VBA source file: UTF-8 first, GBK fallback (legacy files).

    When the GBK fallback has to REPLACE characters outside GBK (emoji etc.),
    print an explicit warning — a silent '?' substitution corrupts VBA string
    literals with no error anywhere in the export→edit→import round-trip."""
    with open(src, "rb") as f:
        raw = f.read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("gbk", errors="replace")
        n_replaced = text.count("�")
        if n_replaced:
            print(f"Warning: {os.path.basename(src)} 含 {n_replaced} 个 GBK 外字符"
                  f"（已替换为 '?'）——VBA 字符串字面量可能损坏，请将源文件改存 UTF-8")
        return text


def _proc_span(code_mod, proc_name: str):
    """
    (start_line, line_count) of a procedure, robust against the documented
    ProcCountLines quirk (it counts from ProcStartLine, including preceding
    blank/comment lines, so a naive body_line+count can overshoot EOF).
    start covers preceding blanks/comments so deletion leaves no orphans.
    """
    start_line = code_mod.ProcStartLine(proc_name, 0)
    body_line = code_mod.ProcBodyLine(proc_name, 0)
    proc_lines = code_mod.ProcCountLines(proc_name, 0)
    total = code_mod.CountOfLines
    start = max(1, min(start_line, body_line))
    end = min(start_line + proc_lines - 1, total)
    return start, max(1, end - start + 1)


def _proc_body_span(code_mod, proc_name: str):
    """(body_line, body_line_count): the declaration line through End Sub,
    excluding preceding blank/comment lines (kept on replace)."""
    start_line = code_mod.ProcStartLine(proc_name, 0)
    body_line = code_mod.ProcBodyLine(proc_name, 0)
    proc_lines = code_mod.ProcCountLines(proc_name, 0)
    total = code_mod.CountOfLines
    body_end = min(start_line + proc_lines - 1, total)
    return body_line, max(1, body_end - body_line + 1)

# Excel save formats (win32com format codes)
XLSX_FORMAT = 51   # xlOpenXMLWorkbook (.xlsx)
XLSM_FORMAT = 52   # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XLAM_FORMAT = 55   # xlOpenXMLAddIn (.xlam) — 54 is the legacy .xla format

# RibbonX namespace (2009/07 = final version, Excel 2010+)
CUSTOMUI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
CUSTOMUI_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

# VBE component type constants
VBEXT_CT_STDMODULE = 1
VBEXT_CT_CLASSMODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOCUMENT = 100

# Automation security constants
MSO_SECURITY_LOW = 1            # macros enabled
MSO_SECURITY_FORCE_DISABLE = 3  # macros disabled on open

# Backups per file to keep
MAX_BACKUPS = 10


# ============================================================================
# Section 1: Excel COM Session Management (isolated instance)
# ============================================================================

_EXCEL_LOCK = threading.Lock()
_cached_excel = {"excel": None, "pid": None}


def _excel_pid(excel):
    """Return the OS process id of an Excel.Application COM object (best-effort)."""
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        return pid
    except Exception:
        return None


def _acquire_excel(attach: bool = False):
    """
    Get an Excel COM application.

    Default: an ISOLATED instance created via DispatchEx and cached for the
    lifetime of this process (fast repeated calls). The user's running Excel
    is never touched.

    attach=True: reuse the user's running Excel instance (GetActiveObject).
    Attached instances are never Quit; if none is running, fall back to an
    isolated instance.

    Returns (excel, attached: bool).
    """
    if attach:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            return excel, True
        except Exception:
            print("Note: no running Excel found, falling back to an isolated instance")
    with _EXCEL_LOCK:
        if _cached_excel["excel"] is None:
            excel = win32com.client.DispatchEx("Excel.Application")
            try:
                excel.Visible = False
            except Exception:
                pass
            _cached_excel["excel"] = excel
            _cached_excel["pid"] = _excel_pid(excel)
        return _cached_excel["excel"], False


def _invalidate_cached_excel() -> None:
    """Forget the cached isolated instance (e.g. after it was force-killed)."""
    with _EXCEL_LOCK:
        _cached_excel["excel"] = None
        _cached_excel["pid"] = None


def _is_attached_instance(excel) -> bool:
    """True when `excel` is the user's running instance rather than our
    cached isolated one (attach requested AND a running Excel was found —
    _acquire_excel falls back to an isolated instance otherwise)."""
    with _EXCEL_LOCK:
        return excel is not _cached_excel["excel"]


def _process_alive(pid: int) -> bool:
    try:
        # bytes 模式判断：tasklist 输出本地化文本（中文系统为 GBK），而
        # PYTHONUTF8=1 的进程会把 text=True 输出按 utf-8 解码——读线程崩溃、
        # stdout 变 None、本函数恒返回 False（残留 Excel 永远不会被强杀）。
        # 只判断 ASCII 的 "EXCEL.EXE"，不解码就没有编码问题。
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            capture_output=True, timeout=15).stdout
        return b"EXCEL.EXE" in out.upper()
    except Exception:
        return False


def shutdown_excel(force_wait: float = 5.0) -> None:
    """
    Quit the cached isolated Excel instance (never touches a user instance).

    On machines with startup add-ins, Excel.Quit() can return without error
    while the process lingers indefinitely. After Quit, the pid-tracked
    process is verified for `force_wait` seconds and force-killed if still
    alive — safe, because it is always OUR DispatchEx-spawned instance.
    """
    with _EXCEL_LOCK:
        excel = _cached_excel["excel"]
        pid = _cached_excel["pid"]
        _cached_excel["excel"] = None
        _cached_excel["pid"] = None
    if excel is None:
        return
    try:
        for wb in list(excel.Workbooks):
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
    except Exception:
        pass
    if pid:
        deadline = time.time() + force_wait
        while time.time() < deadline:
            if not _process_alive(pid):
                return
            time.sleep(0.5)
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=15)
        except Exception:
            pass


atexit.register(shutdown_excel)


def _com_error_text(e):
    """Extract a human-readable message from a pywintypes.com_error.
    Returns None when the exception carries no description (typical for VBA
    runtime errors ended via dialog button — the real text is in the dialog)."""
    ei = getattr(e, "excepinfo", None)
    if ei:
        for idx in (2, 1):
            try:
                if ei[idx]:
                    return str(ei[idx])
            except Exception:
                pass
    return None


# ============================================================================
# Section 2: Dialog Watchdog (prevents COM hangs on modal dialogs)
# ============================================================================

def _read_dialog_text(hwnd) -> str:
    """Read the static text controls of a Win32 dialog without dismissing it."""
    import win32gui

    texts = []

    def _read_children(child, _):
        try:
            if win32gui.GetClassName(child) == "Static":
                text = win32gui.GetWindowText(child)
                if text:
                    texts.append(text)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _read_children, None)
    except Exception:
        pass
    return "\n".join(texts)


# Safe dismiss buttons, in priority order. '结束/End' first: on VBA runtime
# error dialogs it terminates the macro cleanly. NEVER '继续/Continue' (the
# error re-raises, looping dialogs) or '调试/Debug' (hangs the VBE).
_DISMISS_BUTTONS = ("结束", "end", "确定", "ok", "是(&", "是", "yes")


def _dismiss_dialog(hwnd) -> bool:
    """
    Dismiss a dialog window. VBA error dialogs are hard-modal: they have no
    close button and IGNORE WM_CLOSE — the only reliable way out is clicking
    a safe button via BM_CLICK.
    """
    import win32con
    import win32gui

    buttons = []

    def _collect(child, _):
        try:
            if win32gui.GetClassName(child) == "Button":
                buttons.append((win32gui.GetWindowText(child).casefold(), child))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _collect, None)
    except Exception:
        pass

    for caption, bh in buttons:
        for want in _DISMISS_BUTTONS:
            if caption.startswith(want):
                try:
                    win32gui.PostMessage(bh, win32con.BM_CLICK, 0, 0)
                    return True
                except Exception:
                    pass
    # Fallback for ordinary dialogs that do honor WM_CLOSE
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass
    return False


class _DialogWatchdog(threading.Thread):
    """
    Poll an Excel process for modal dialogs (#32770), read their text BEFORE
    dismissing them, then close them (safe-button click first, WM_CLOSE as
    fallback) so the blocked COM call can continue.

    kill_after: if set and the dialog loop is still running after this many
    seconds, hard-kill the (isolated, tool-owned) Excel process to break a
    stuck COM call. Never use against a user-attached instance.
    """

    def __init__(self, pid, interval=0.4, kill_after=None, storm_limit=None):
        super().__init__(daemon=True)
        self.pid = pid
        self.interval = interval
        self.kill_after = kill_after
        # Kill the Excel process after this many dialog dismissals when set —
        # guards against dialog storms (e.g. a compile error re-raising on
        # every Run retry would loop forever)
        self.storm_limit = storm_limit
        self.stop_event = threading.Event()
        self.dialog_texts = []   # text of every dialog seen, in order
        self.dismissed = 0
        self.killed = False
        self.storm = False
        self._t0 = None

    def run(self):
        import win32gui
        import win32process

        self._t0 = time.time()
        while not self.stop_event.is_set():
            try:
                dialogs = []

                def _enum(hwnd, _):
                    try:
                        if not win32gui.IsWindowVisible(hwnd):
                            return True
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        if pid == self.pid and win32gui.GetClassName(hwnd) == "#32770":
                            dialogs.append(hwnd)
                    except Exception:
                        pass
                    return True

                win32gui.EnumWindows(_enum, None)
                for hwnd in dialogs:
                    text = _read_dialog_text(hwnd)
                    if text:
                        self.dialog_texts.append(text)
                    if _dismiss_dialog(hwnd):
                        self.dismissed += 1
            except Exception:
                pass

            elapsed = time.time() - self._t0
            storm_hit = (self.storm_limit is not None and self.dismissed >= self.storm_limit)
            timeout_hit = (self.kill_after is not None and self.pid
                           and elapsed > self.kill_after)
            if storm_hit or timeout_hit:
                # Stuck or storming: force-kill OUR isolated Excel to unblock
                # the COM call.
                if self.pid:
                    try:
                        subprocess.run(
                            ["taskkill", "/PID", str(self.pid), "/F"],
                            capture_output=True, timeout=15,
                        )
                        self.killed = True
                        self.storm = storm_hit and not timeout_hit
                    except Exception:
                        pass
                return
            self.stop_event.wait(self.interval)

    def stop(self, grace=1.0):
        """Signal stop, wait a grace period for late dialogs, then join."""
        time.sleep(grace)
        self.stop_event.set()
        self.join(timeout=3)


def _run_watchdog(excel, timeout: int, attached: bool,
                  interval: float = 0.5, storm_limit: int = 8) -> _DialogWatchdog:
    """
    Watchdog aimed at the Excel instance ACTUALLY in use — not the cached
    isolated one. In attach mode the cached pid is stale/None (the user's
    instance is never cached), which historically left dialogs unhandled and
    kill logic aimed at the wrong process.

    Attached (user) instances are NEVER force-killed: kill_after and
    storm_limit are disabled. Consequence: a stuck macro in attach mode
    cannot be auto-broken (the blocking COM call only returns once its
    dialog is dismissed) — dialog read/dismiss still works.
    """
    pid = _excel_pid(excel)
    if pid is None and not attached:
        with _EXCEL_LOCK:
            pid = _cached_excel["pid"]
    if attached:
        return _DialogWatchdog(pid, interval=interval,
                               kill_after=None, storm_limit=None)
    return _DialogWatchdog(pid, interval=interval,
                           kill_after=max(15, timeout), storm_limit=storm_limit)


# ============================================================================
# Section 3: Backup & Restore
# ============================================================================

def _backup_dir(file_path: str) -> str:
    return os.path.abspath(file_path) + ".bak"


def backup_file(file_path: str, reason: str = "write"):
    """
    Back up a file to <file>.bak/<timestamp>_<name>. Keeps newest MAX_BACKUPS.

    Returns backup path, or None if the file doesn't exist.
    """
    path = os.path.abspath(file_path)
    if not os.path.exists(path):
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    bdir = _backup_dir(path)
    os.makedirs(bdir, exist_ok=True)
    dest = os.path.join(bdir, f"{ts}_{os.path.basename(path)}")
    shutil.copy2(path, dest)
    # Prune old backups
    try:
        backups = sorted(os.listdir(bdir))
        for old in backups[:-MAX_BACKUPS]:
            os.remove(os.path.join(bdir, old))
    except Exception:
        pass
    print(f"[backup] {reason}: {dest}")
    return dest


def list_backups(file_path: str) -> list:
    """List available backups for a file, newest first."""
    bdir = _backup_dir(file_path)
    if not os.path.isdir(bdir):
        return []
    return sorted(os.listdir(bdir), reverse=True)


def restore_backup(file_path: str, backup: str = None) -> str:
    """
    Restore a file from backup.

    backup: backup filename inside <file>.bak/ (as listed by list_backups),
            or None for the newest.

    The current file state is itself backed up first, so a restore is
    reversible. Returns the backup path that was restored.
    """
    path = os.path.abspath(file_path)
    backups = list_backups(path)
    if not backups:
        raise FileNotFoundError(f"No backups found for {path}")
    chosen = backup if backup is not None else backups[0]
    src = os.path.join(_backup_dir(path), chosen)
    if not os.path.exists(src):
        raise FileNotFoundError(f"Backup not found: {src}")
    if os.path.exists(path):
        backup_file(path, reason="pre-restore")
    shutil.copy2(src, path)
    print(f"[restore] {path} <- {src}")
    return src


# ============================================================================
# Section 4: Environment Check & VBA Trust
# ============================================================================

def _office_versions_with_excel() -> list:
    """Office version keys (e.g. '16.0') under HKCU that have an Excel subkey."""
    import winreg
    versions = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office")
        i = 0
        while True:
            try:
                ver = winreg.EnumKey(root, i)
            except OSError:
                break
            i += 1
            if re.match(r"^\d+\.\d+$", ver):
                try:
                    winreg.OpenKey(root, f"{ver}\\Excel")
                    versions.append(ver)
                except OSError:
                    pass
    except OSError:
        pass
    return versions


def check_environment(verbose: bool = True) -> dict:
    """
    Check the local environment for VBA automation readiness.

    Returns dict:
      {python, pywin32, excel_installed, excel_version,
       vba_trust_enabled, office_versions, ok}
    """
    import winreg

    result = {
        "python": ".".join(str(v) for v in sys.version_info[:3]),
        "pywin32": True,
        "excel_installed": False,
        "excel_version": None,
        "office_versions": [],
        "vba_trust_enabled": False,
        "ok": False,
    }

    # pywin32 (importable == installed, we already imported win32com)
    # Excel installed?
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as k:
            progid = winreg.QueryValueEx(k, "")[0]  # e.g. Excel.Application.16
            result["excel_installed"] = True
            result["excel_version"] = progid.split(".")[-1]
    except OSError:
        result["pywin32_note"] = "Excel not found in registry (HKCR\\Excel.Application)"

    # VBA project object model trust per Office version
    versions = _office_versions_with_excel()
    result["office_versions"] = versions
    trust_states = {}
    for ver in versions:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                rf"Software\Microsoft\Office\{ver}\Excel\Security") as k:
                trust_states[ver] = winreg.QueryValueEx(k, "AccessVBOM")[0] == 1
        except OSError:
            trust_states[ver] = False
    result["_trust_states"] = trust_states
    result["vba_trust_enabled"] = any(trust_states.values())

    result["ok"] = result["pywin32"] and result["excel_installed"] and result["vba_trust_enabled"]

    if verbose:
        print("=== 环境检查 ===")
        print(f"Python        : {result['python']}")
        print(f"pywin32       : OK")
        print(f"Excel         : {'已安装 (Office ' + result['excel_version'] + ')' if result['excel_installed'] else '未检测到'}")
        print(f"Office 版本键 : {', '.join(versions) or '无'}")
        print(f"VBA 工程访问  : {'已信任 (AccessVBOM=1)' if result['vba_trust_enabled'] else '未信任 — 运行 enable_vba_trust() 或在 Excel 信任中心开启'}")
        print(f"总体          : {'[OK] 就绪' if result['ok'] else '[FAIL] 未就绪'}")
    return result


def enable_vba_trust() -> bool:
    """
    Enable 'Trust access to the VBA project object model' for every detected
    Office/Excel version by writing HKCU\\...\\Excel\\Security\\AccessVBOM = 1.

    Note: this relaxes an Office security guard; revert in
    Excel 信任中心 → 宏设置, or set AccessVBOM back to 0.
    Takes effect for NEW Excel instances (restart running Excel).
    """
    import winreg

    versions = _office_versions_with_excel()
    if not versions:
        print("未检测到带 Excel 的 Office 版本，无法设置")
        return False
    changed = []
    for ver in versions:
        key_path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            changed.append(ver)
        except OSError as e:
            print(f"设置失败 ({ver}): {e}")
    if changed:
        print(f"已启用 VBA 工程访问信任 (AccessVBOM=1): {', '.join(changed)}")
        print("注意：已重启的 Excel 才会生效；可在信任中心手动关闭以还原")
    return bool(changed)


# ============================================================================
# Section 5: File Format Functions
# ============================================================================

def get_file_format(file_path: str) -> str:
    """
    Detect Excel file format from file extension.

    Args:
        file_path: Path to Excel file

    Returns:
        'xlam' / 'xlsm' / 'xlsx' / 'unknown'
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xlam":
        return "xlam"
    elif ext == ".xlsm":
        return "xlsm"
    elif ext == ".xlsx":
        return "xlsx"
    else:
        return "unknown"


def convert_to_xlsm(file_path: str, output_path: str = None, attach: bool = False) -> str:
    """
    Convert xlsx to xlsm format.

    Args:
        file_path: Source xlsx file path
        output_path: Output path. If None, replaces extension with .xlsm
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        Path to the converted xlsm file
    """
    file_path = os.path.abspath(file_path)

    if get_file_format(file_path) not in ("xlsx",):
        raise ValueError(f"convert_to_xlsm only supports .xlsx files, got: {file_path}")

    if output_path is None:
        output_path = os.path.splitext(file_path)[0] + ".xlsm"
    else:
        output_path = os.path.abspath(output_path)
        if os.path.exists(output_path):
            # 显式指定的输出已存在：SaveAs(DisplayAlerts=False) 会静默覆盖，
            # 先备份（默认路径的覆盖备份由 _ensure_xlsm 负责，避免双备份）
            backup_file(output_path, reason="convert-overwrite")

    excel, attached = _acquire_excel(attach)
    _apply_app_settings(excel, attached, mode="write")
    wb = None
    try:
        wb = excel.Workbooks.Open(file_path, UpdateLinks=0)
        wb.SaveAs(output_path, FileFormat=XLSM_FORMAT)
        print(f"Converted: {file_path} -> {output_path}")
        return output_path
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if attached:
            _restore_app_settings(excel)


def _ensure_xlsm(file_path: str) -> str:
    """
    Ensure file is in xlsm/xlam format for VBA write operations.

    Args:
        file_path: Path to Excel file (.xlsx, .xlsm, or .xlam)

    Returns:
        Path to xlsm/xlam file (may be newly created for xlsx inputs)
        If xlsx was converted, returns new xlsm path (original file unchanged)
    """
    file_path = os.path.abspath(file_path)
    fmt = get_file_format(file_path)

    if fmt == "xlsx":
        target = os.path.splitext(file_path)[0] + ".xlsm"
        if os.path.exists(target):
            if os.path.getmtime(target) >= os.path.getmtime(file_path):
                # A previous conversion exists and the xlsx has not changed
                # since — REUSE it. Re-converting would silently overwrite
                # the xlsm (SaveAs + DisplayAlerts=False) and drop every VBA
                # change made in between.
                print(f"Reusing existing conversion: {target}（xlsx 未变更，"
                      f"避免覆盖丢失已写入的 VBA）")
                return target
            # The xlsx is newer than the last conversion (user edit?) —
            # reconvert, but back up the existing xlsm first: it may hold
            # VBA that would otherwise be lost.
            backup_file(target, reason="reconvert-xlsm")
            print(f"Warning: xlsx 比已转换的 {os.path.basename(target)} 新，"
                  f"将重新转换覆盖（旧 xlsm 已备份）")
        return convert_to_xlsm(file_path)
    elif fmt in ("xlsm", "xlam"):
        return file_path
    else:
        raise ValueError(f"Unsupported file format: {file_path}")


def _require_macro_file(file_path: str) -> str:
    """For read operations: file must already be .xlsm/.xlam (xlsx has no VBA)."""
    path = os.path.abspath(file_path)
    fmt = get_file_format(path)
    if fmt == "xlsx":
        raise ValueError(
            f".xlsx 不含 VBA 工程: {path}（写操作会自动转换 xlsm；读取请直接指向 xlsm/xlam）")
    if fmt not in ("xlsm", "xlam"):
        raise ValueError(f"Unsupported file format: {path}")
    return path


# ============================================================================
# Section 6: Extract and Repackage
# ============================================================================

def create_addin(xlam_path: str) -> str:
    """
    Create an empty Excel add-in (.xlam) from scratch — the '0' of a 0-to-1
    add-in build. The VBA and Ribbon layers are then added with the other
    toolkit functions.

    Args:
        xlam_path: Output .xlam path (overwrites if exists)

    Returns:
        Absolute path of the created add-in.
    """
    xlam_path = os.path.abspath(xlam_path)
    excel, attached = _acquire_excel(False)
    _apply_app_settings(excel, attached, mode="write")
    wb = None
    try:
        wb = excel.Workbooks.Add()
        if os.path.exists(xlam_path):
            backup_file(xlam_path, reason="create_addin-overwrite")
            os.remove(xlam_path)
        wb.SaveAs(xlam_path, FileFormat=XLAM_FORMAT)
        print(f"Created add-in: {xlam_path}")
        return xlam_path
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass


def unpack_xlam(xlam_path: str, output_dir: str) -> None:
    """
    Extract .xlam file to a directory.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx)
        output_dir: Directory to extract to
    """
    xlam_path = os.path.abspath(xlam_path)
    output_dir = os.path.abspath(output_dir)

    if os.path.isdir(output_dir):
        # 护栏：只清空"像本工具解包产物"的目录（Office OPC 包必含
        # [Content_Types].xml）或空目录——防止误传父目录等导致 rmtree
        # 整目录销毁，与 export_vba_source 的 manifest.json 校验同一思路
        # （空目录放行：mkdtemp 预建目录 + render_ribbon_preview 内部
        # 解包会走到这里）
        is_unpack_product = os.path.exists(
            os.path.join(output_dir, "[Content_Types].xml"))
        if not (is_unpack_product or not os.listdir(output_dir)):
            raise ValueError(
                f"目录已存在且不是本工具的解包产物（缺 [Content_Types].xml），"
                f"拒绝清空: {output_dir}")
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    with zipfile.ZipFile(xlam_path, "r") as z:
        z.extractall(output_dir)
    print(f"Extracted: {xlam_path} -> {output_dir}")


def pack_xlam(source_dir: str, output_xlam: str, vba_source: str = None) -> None:
    """
    Repackage directory to a workbook file.
    Ensures [Content_Types].xml is first in archive.

    Args:
        source_dir: Directory to package
        output_xlam: Output file path (.xlam/.xlsm)
        vba_source: Optional path to the CURRENT .xlam/.xlsm file. Its
            xl/vbaProject.bin replaces the stale copy in source_dir before
            packing. REQUIRED whenever VBA code was modified through the
            COM API (add_vba_module etc.) after unpacking — otherwise those
            changes are silently reverted to the unpack-time snapshot.
    """
    source_dir = os.path.abspath(source_dir)
    output_xlam = os.path.abspath(output_xlam)

    # Refresh vbaProject.bin from the live file when given
    if vba_source:
        vba_source = os.path.abspath(vba_source)
        if vba_source == output_xlam:
            raise ValueError(
                "output 与 vba_source 不能指向同一文件——打包会先删除输出文件，"
                "导致 VBA 源丢失。请输出到新路径，再重命名。")
        bin_entry = "xl/vbaProject.bin"
        with zipfile.ZipFile(vba_source) as zsrc:
            if bin_entry in zsrc.namelist():
                data = zsrc.read(bin_entry)
                dst = os.path.join(source_dir, "xl", "vbaProject.bin")
                with open(dst, "wb") as f:
                    f.write(data)
                print(f"[vba] refreshed vbaProject.bin from {vba_source}")
            else:
                print(f"[vba] warning: {vba_source} has no vbaProject.bin — nothing to refresh")

    # Remove existing output file (backed up first — removal is destructive)
    if os.path.exists(output_xlam):
        backup_file(output_xlam, reason="pack-overwrite")
        os.remove(output_xlam)

    with zipfile.ZipFile(output_xlam, "w", zipfile.ZIP_DEFLATED) as z:
        # Add [Content_Types].xml first (required by Office)
        content_types = os.path.join(source_dir, "[Content_Types].xml")
        if os.path.exists(content_types):
            z.write(content_types, "[Content_Types].xml")
        else:
            print(f"[pack] warning: 缺 {content_types} —— 产出的包 Excel 打不开。"
                  f"该文件由 unpack_xlam 解出或 init_custom_ui 维护，请确认解包目录完整")

        # Add all other files
        for root, dirs, files in os.walk(source_dir):
            for f in files:
                full_path = os.path.join(root, f)
                # Never pack the output into itself: callers may point the
                # output inside source_dir (e.g. pack_xlam("pkg/", "pkg/x.xlam")),
                # which would otherwise embed a half-written copy of the zip.
                if os.path.normcase(full_path) == os.path.normcase(output_xlam):
                    continue
                arcname = os.path.relpath(full_path, source_dir)
                if arcname != "[Content_Types].xml":  # Already added
                    z.write(full_path, arcname)

    print(f"Packaged: {source_dir} -> {output_xlam}")


# ============================================================================
# Section 7: Ribbon XML Functions
# ============================================================================

_DEFAULT_CUSTOMUI_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<customUI xmlns="{CUSTOMUI_NS}">\n'
    '  <ribbon>\n'
    '    <tabs>\n'
    '      <tab id="tabCustom" label="自定义">\n'
    '      </tab>\n'
    '    </tabs>\n'
    '  </ribbon>\n'
    '</customUI>\n'
)


def _read_xml_part(path: str) -> str:
    """Read an XML part (BOM-tolerant)."""
    with open(path, "r", encoding="utf-8-sig") as f:
        return f.read()


def _ensure_content_types(xlam_dir: str, ext: str, content_type: str) -> None:
    """Ensure [Content_Types].xml declares a Default for the extension."""
    ct_path = os.path.join(xlam_dir, "[Content_Types].xml")
    ct = _read_xml_part(ct_path)
    if f'extension="{ext}"' in ct.lower():
        return
    entry = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
    if "</Types>" not in ct:
        raise ValueError(f"Malformed [Content_Types].xml in {xlam_dir}")
    ct = ct.replace("</Types>", entry + "</Types>")
    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(ct)
    print(f"[content-types] declared .{ext} -> {content_type}")


def init_custom_ui(xlam_dir: str, xml: str = None) -> str:
    """
    Initialize the full customUI infrastructure in an unpacked workbook
    directory (idempotent — safe to call on an already-initialized package):

      - customUI/customUI.xml            (given xml, or a starter template
                                          with one empty 'tabCustom' tab)
      - customUI/_rels/customUI.xml.rels (empty icon-relationship list)
      - [Content_Types].xml              .xml / .rels / .png declarations
      - _rels/.rels                      root relationship to customUI.xml

    Call this FIRST when customizing the Ribbon of a workbook that has never
    had one — a fresh unpack has no customUI directory and every other
    ribbon function would fail without this initialization.

    Returns the customUI.xml content.
    """
    xlam_dir = os.path.abspath(xlam_dir)
    cui_dir = os.path.join(xlam_dir, "customUI")
    rels_dir = os.path.join(cui_dir, "_rels")
    os.makedirs(rels_dir, exist_ok=True)
    os.makedirs(os.path.join(cui_dir, "images"), exist_ok=True)

    # 1. customUI.xml
    cui_path = os.path.join(cui_dir, "customUI.xml")
    if not os.path.exists(cui_path):
        with open(cui_path, "w", encoding="utf-8") as f:
            f.write(xml if xml is not None else _DEFAULT_CUSTOMUI_XML)
        print(f"[customUI] initialized: {cui_path}")

    # 2. empty icon rels
    rels_path = os.path.join(rels_dir, "customUI.xml.rels")
    if not os.path.exists(rels_path):
        with open(rels_path, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '</Relationships>')
        print(f"[customUI] initialized: {rels_path}")

    # 3. content type declarations
    _ensure_content_types(xlam_dir, "xml", "application/xml")
    _ensure_content_types(xlam_dir, "rels",
                          "application/vnd.openxmlformats-package.relationships+xml")
    _ensure_content_types(xlam_dir, "png", "image/png")

    # 4. root relationship (append if missing)
    root_rels_path = os.path.join(xlam_dir, "_rels", ".rels")
    root_rels = _read_xml_part(root_rels_path)
    if "customUI/customUI.xml" not in root_rels:
        new_rel = (f'<Relationship Id="rIdCustomUI" Type="{CUSTOMUI_REL_TYPE}" '
                   f'Target="customUI/customUI.xml"/>')
        if "</Relationships>" not in root_rels:
            raise ValueError(f"Malformed _rels/.rels in {xlam_dir}")
        root_rels = root_rels.replace("</Relationships>", new_rel + "</Relationships>")
        with open(root_rels_path, "w", encoding="utf-8") as f:
            f.write(root_rels)
        print("[customUI] root relationship added to _rels/.rels")

    return _read_xml_part(cui_path)


def get_ribbon_xml(xlam_dir: str) -> str:
    """
    Read customUI.xml content.

    Args:
        xlam_dir: Path to extracted workbook directory

    Returns:
        Content of customUI.xml
    """
    ribbon_path = os.path.join(xlam_dir, "customUI", "customUI.xml")
    if not os.path.exists(ribbon_path):
        raise ValueError(
            "customUI.xml not found — the workbook has no Ribbon customization yet. "
            "Call init_custom_ui(xlam_dir) first to create it from scratch.")
    return _read_xml_part(ribbon_path)


def set_ribbon_xml(xlam_dir: str, xml_content: str) -> None:
    """
    Write customUI.xml content. Initializes the full customUI infrastructure
    automatically when the package has never had one (so writing a Ribbon
    from scratch is just one call).
    """
    ribbon_path = os.path.join(xlam_dir, "customUI", "customUI.xml")
    if not os.path.exists(ribbon_path):
        init_custom_ui(xlam_dir, xml=xml_content)
        return
    with open(ribbon_path, "w", encoding="utf-8") as f:
        f.write(xml_content)
    print(f"Updated: {ribbon_path}")


def add_button_to_ribbon(
    xlam_dir: str,
    group_id: str,
    button_xml: str,
    after_button_id: str = None
) -> None:
    """
    Add a button to an existing group in customUI.xml.

    Handles both normal (<group id="x">...</group>) and self-closing
    (<group id="x"/>) groups — the self-closing form is what a freshly
    initialized or hand-written minimal Ribbon uses.

    Args:
        xlam_dir: Path to extracted workbook directory
        group_id: ID of the group to add button to (matches id= / idQ=)
        button_xml: Button XML string (without leading/trailing whitespace)
        after_button_id: Optional: ID of button to insert after
    """
    ribbon_xml = get_ribbon_xml(xlam_dir)

    # Normal form: <group id="x" ...> ... </group>
    # ([^/]?> excludes self-closing tags; \s boundary avoids matching
    #  e.g. someid="x")
    group_pattern = (rf'(<group[^>]*\s(?:id|idQ)="{re.escape(group_id)}"'
                     rf'[^>]*[^/]?>)(.*?)(</group>)')
    match = re.search(group_pattern, ribbon_xml, re.DOTALL)

    if match:
        group_start = match.group(1)
        group_content = match.group(2)
        group_end = match.group(3)

        # 查重：同 id 按钮重复追加（如 AI 对结果不确定时重试）会产出
        # Excel 拒载的 customUI——先看组内是否已有该 id
        new_id_m = re.search(r'\bid="([^"]+)"', button_xml)
        if new_id_m and re.search(rf'\bid="{re.escape(new_id_m.group(1))}"',
                                  group_content):
            print(f"Warning: button id '{new_id_m.group(1)}' already in group "
                  f"'{group_id}' — skipped (duplicate ids break the whole ribbon)")
            return

        # Format button XML
        indented_button = "\n          " + button_xml.strip()

        if after_button_id:
            # Insert after specific button
            btn_pattern = rf'(<button[^>]*id="{re.escape(after_button_id)}"[^>]*/?>)'
            btn_match = re.search(btn_pattern, group_content)
            if btn_match:
                pos = btn_match.end()
                new_content = group_content[:pos] + indented_button + group_content[pos:]
            else:
                new_content = group_content + indented_button
        else:
            # Append at end of group
            new_content = group_content + indented_button

        new_ribbon = ribbon_xml[:match.start()] + group_start + new_content + group_end + ribbon_xml[match.end():]
        set_ribbon_xml(xlam_dir, new_ribbon)
        print(f"Added button to group '{group_id}'")
        return

    # Self-closing form: <group id="x" .../> — expand it
    sc_pattern = rf'<group[^>]*\s(?:id|idQ)="{re.escape(group_id)}"[^>]*/>'
    sc_match = re.search(sc_pattern, ribbon_xml)
    if sc_match:
        tag = sc_match.group(0)
        open_tag = tag[:-2].rstrip() + ">"
        replacement = (open_tag + "\n          " + button_xml.strip()
                       + "\n        </group>")
        set_ribbon_xml(xlam_dir, ribbon_xml.replace(tag, replacement, 1))
        print(f"Added button to group '{group_id}' (expanded self-closing group)")
        return

    print(f"Warning: Group '{group_id}' not found")


def add_group_to_ribbon(xlam_dir: str, group_xml: str, after_group_id: str = None,
                        tab_id: str = None) -> None:
    """
    Add a group to a tab in customUI.xml.

    Args:
        xlam_dir: Path to extracted workbook directory
        group_xml: Group XML string
        after_group_id: Optional: ID of group to insert after
        tab_id: Optional: target tab ID. Matches id=, idMso= or idQ=, so
                built-in tabs customized via idMso (e.g. 'TabHome') work too.
                When omitted, the FIRST tab is used and a note is printed
                if multiple tabs exist.
    """
    ribbon_xml = get_ribbon_xml(xlam_dir)

    if tab_id:
        tab_pattern = (rf'(<tab[^>]*\s(?:id|idMso|idQ)="{re.escape(tab_id)}"'
                       rf'[^>]*[^/]?>)(.*?)(</tab>)')
    else:
        tab_pattern = r'(<tab[^>]*[^/]?>)(.*?)(</tab>)'

    matches = list(re.finditer(tab_pattern, ribbon_xml, re.DOTALL))
    if not matches:
        hint = (f"（指定 tab_id='{tab_id}' 未匹配到" if tab_id
                else "（没有可插入的 <tab>，可用 init_custom_ui() 创建初始模板）")
        print(f"Warning: Tab not found {hint}")
        return
    if not tab_id and len(matches) > 1:
        print(f"Note: 发现 {len(matches)} 个 tab，未指定 tab_id，默认使用第一个；"
              f"如需其他 tab 请传 tab_id 参数")

    match = matches[0]
    tab_content = match.group(2)
    indented_group = "\n        " + group_xml.strip()

    if after_group_id:
        grp_pattern = rf'(<group[^>]*\s(?:id|idQ)="{re.escape(after_group_id)}"[^>]*[^/]?>.*?</group>)'
        grp_match = re.search(grp_pattern, tab_content, re.DOTALL)
        if grp_match:
            pos = grp_match.end()
            new_tab_content = tab_content[:pos] + indented_group + tab_content[pos:]
        else:
            new_tab_content = tab_content + indented_group
    else:
        new_tab_content = tab_content + indented_group

    new_ribbon = ribbon_xml[:match.start()] + match.group(1) + new_tab_content + match.group(3) + ribbon_xml[match.end():]

    set_ribbon_xml(xlam_dir, new_ribbon)
    target = tab_id or "(first tab)"
    print(f"Added group to tab {target}")


def generate_button_xml(
    id: str,
    label: str,
    on_action: str,
    size: str = "large",
    image_mso: str = None,
    image: str = None,
    screentip: str = "",
    supertip: str = ""
) -> str:
    """
    Generate button XML string.

    Args:
        id: Button ID
        label: Display label
        on_action: VBA callback procedure name
        size: 'large' or 'normal'
        image_mso: Built-in icon name (e.g., 'Copy', 'Paste')
        image: Custom icon ID (defined in customUI.xml.rels)
        screentip: Short tooltip
        supertip: Detailed tooltip

    Returns:
        Button XML string
    """
    attrs = [
        f'id="{_xml_attr(id)}"',
        f'label="{_xml_attr(label)}"',
        f'size="{size}"',
        f'onAction="{_xml_attr(on_action)}"',
    ]

    if image_mso:
        attrs.append(f'imageMso="{_xml_attr(image_mso)}"')
    elif image:
        attrs.append(f'image="{_xml_attr(image)}"')

    if screentip:
        attrs.append(f'screentip="{_xml_attr(screentip)}"')
    if supertip:
        attrs.append(f'supertip="{_xml_attr(supertip)}"')

    return f'<button {" ".join(attrs)}/>'


def generate_group_xml(
    id: str,
    label: str,
    buttons: list
) -> str:
    """
    Generate group XML with buttons.

    Args:
        id: Group ID
        label: Display label
        buttons: List of button XML strings

    Returns:
        Group XML string
    """
    buttons_xml = "\n          ".join(buttons)
    return f'''<group id="{_xml_attr(id)}" label="{_xml_attr(label)}">
          {buttons_xml}
        </group>'''


# ============================================================================
# Section 8: Icon Registration Functions
# ============================================================================

def get_icon_rels(xlam_dir: str) -> str:
    """
    Read customUI.xml.rels content.

    Args:
        xlam_dir: Path to extracted .xlam directory

    Returns:
        Content of customUI.xml.rels
    """
    rels_path = os.path.join(xlam_dir, "customUI", "_rels", "customUI.xml.rels")
    with open(rels_path, "r", encoding="utf-8") as f:
        return f.read()


def set_icon_rels(xlam_dir: str, rels_content: str) -> None:
    """
    Write customUI.xml.rels content.

    Args:
        xlam_dir: Path to extracted .xlam directory
        rels_content: New rels content
    """
    rels_path = os.path.join(xlam_dir, "customUI", "_rels", "customUI.xml.rels")
    with open(rels_path, "w", encoding="utf-8") as f:
        f.write(rels_content)
    print(f"Updated: {rels_path}")


def register_icon(xlam_dir: str, icon_id: str, icon_path: str) -> None:
    """
    Register a custom icon: copies the PNG into customUI/images/, adds the
    relationship to customUI.xml.rels (created if missing), and ensures
    [Content_Types].xml declares the png extension — all three are required
    for the icon to actually show up after packing.

    Args:
        xlam_dir: Path to extracted workbook directory
        icon_id: Icon ID (used in customUI.xml as image="icon_id")
        icon_path: Path to PNG icon file (16x16 / 32x32, transparent bg)
    """
    icon_path = os.path.abspath(icon_path)
    images_dir = os.path.join(xlam_dir, "customUI", "images")

    # Create images directory if needed
    os.makedirs(images_dir, exist_ok=True)

    # Copy icon file
    dest_path = os.path.join(images_dir, f"{icon_id}.png")
    shutil.copy2(icon_path, dest_path)
    print(f"Copied icon: {icon_path} -> {dest_path}")

    # Ensure png is declared in [Content_Types].xml (required to display)
    _ensure_content_types(xlam_dir, "png", "image/png")

    # Read rels (auto-create if the package never had one)
    rels_path = os.path.join(xlam_dir, "customUI", "_rels", "customUI.xml.rels")
    if os.path.exists(rels_path):
        rels_content = _read_xml_part(rels_path)
    else:
        os.makedirs(os.path.dirname(rels_path), exist_ok=True)
        rels_content = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        '</Relationships>')
        print(f"[customUI] initialized: {rels_path}")

    # Check if already registered
    if f'Id="{icon_id}"' in rels_content:
        print(f"Icon '{icon_id}' already registered")
        return

    # Add new relationship
    new_rel = f'<Relationship Id="{icon_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/{icon_id}.png"/>'

    # Insert before closing tag
    rels_content = rels_content.replace("</Relationships>", new_rel + "</Relationships>")

    set_icon_rels(xlam_dir, rels_content)
    print(f"Registered icon: {icon_id}")


def unregister_icon(xlam_dir: str, icon_id: str) -> None:
    """
    Unregister a custom icon.

    Args:
        xlam_dir: Path to extracted workbook directory
        icon_id: Icon ID to remove
    """
    rels_content = get_icon_rels(xlam_dir)

    # Remove relationship
    rel_pattern = rf'<Relationship[^>]*Id="{icon_id}"[^>]*/>'
    rels_content = re.sub(rel_pattern, "", rels_content)
    set_icon_rels(xlam_dir, rels_content)

    # Remove icon file
    icon_path = os.path.join(xlam_dir, "customUI", "images", f"{icon_id}.png")
    if os.path.exists(icon_path):
        os.remove(icon_path)
        print(f"Removed icon file: {icon_path}")


# ============================================================================
# Section 8.5: Icon Generation Pipeline
# (prompt -> image model OR offline drawing -> size normalization -> embed)
# ============================================================================

# imageMso names batch-verified IN-PROCESS against Office 16 (2026-09).
# Do NOT trust unverified lists — names like Find/Print/Replace look
# plausible but do not exist, and the button then shows no icon at all.
COMMON_IMAGEMSO = {
    "编辑":   ["Copy", "Cut", "Paste", "PasteValues", "Undo", "Redo", "Clear",
               "ClearContents", "Delete", "FormatPainter", "Spelling"],
    "格式":   ["Font", "FontDialog", "Bold", "Underline", "WrapText", "MergeCells",
               "AlignLeft", "TextBoxInsert"],
    "数据":   ["Sort", "Filter", "AutoSum", "Refresh", "RefreshAll",
               "TableInsertDialog", "InsertTable", "InsertChart", "ChartInsert"],
    "其他":   ["Save", "HyperlinkInsert", "Camera", "Calculator",
               "CalendarInsert", "VisualBasic", "MacroPlay", "AddInManager"],
}

# Config (icon API + user preferences) lives OUTSIDE the skill directory
# (~/.vba-dev/) so publishing the skill (e.g. to GitHub) can never leak
# credentials. The location is agent-agnostic: an earlier version used the
# Claude-Code-specific ~/.claude/vba-dev/, which _migrate_legacy_config()
# moves over automatically; a legacy config inside scripts/ is also migrated.
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".vba-dev")
ICON_CONFIG_FILE = os.path.join(CONFIG_DIR, "icon_config.json")
_LEGACY_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".claude", "vba-dev")
_LEGACY_ICON_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "icon_config.json")
# Default image model for OpenAI-compatible gateways; override anytime via
# set_icon_config(model=...)
DEFAULT_ICON_MODEL = "gpt-image-1"

# Cross-session user preferences. Same directory as the icon config (OUTSIDE
# the skill dir), so publishing the skill never carries the user's choices —
# and so a preference expressed once ("极简风格"、"用内置图标") survives into
# the next conversation instead of being re-asked every time.
PREF_FILE = os.path.join(CONFIG_DIR, "preferences.json")

# Conventional keys (free-form dict — these are what SKILL.md tells the AI
# to read/write; anything else is allowed and simply ignored by the toolkit):
#   icon_type      imageMso | draw | ai | user_png   （图标类型偏好）
#   icon_style     自由文本，如「极简单色线条」「跟 Office 自带一致」
#   theme_color    "#1F4E92"  （图标/按钮底色）
#   naming_style   verb | noun | prefixed           （label 命名风格）
#   default_group  常用分组 id，如 "grpTools"
#   default_tab    常用 tab id，如 "tabCustom"


def _migrate_legacy_config() -> None:
    """
    One-time migration of config files into CONFIG_DIR (~/.vba-dev/):
      - ~/.claude/vba-dev/{icon_config.json, preferences.json}  (agent-specific
        location used by earlier versions — keep preferences working when the
        skill runs under a different agent, e.g. ZCode)
      - scripts/icon_config.json  (very old in-skill location)
    Destinations are resolved from CONFIG_DIR at call time so tests can
    sandbox them by patching the module globals. Never overwrites: a file
    already in CONFIG_DIR wins. Best-effort, silent beyond a warning.
    """
    sources = [(_LEGACY_ICON_CONFIG_FILE, os.path.join(CONFIG_DIR, "icon_config.json"))]
    for name in ("icon_config.json", "preferences.json"):
        sources.append((os.path.join(_LEGACY_CONFIG_DIR, name),
                        os.path.join(CONFIG_DIR, name)))
    for old, new in sources:
        if old != new and os.path.exists(old) and not os.path.exists(new):
            try:
                os.makedirs(CONFIG_DIR, exist_ok=True)
                shutil.move(old, new)
                print(f"[config] 已迁移 {old} -> {new}（配置保存在 skill 目录外，发布 skill 不会泄露）")
            except Exception as e:
                print(f"Warning: config migration failed: {e}")


def get_preferences() -> dict:
    """
    Read persisted user preferences from ~/.vba-dev/preferences.json.

    Returns {} when nothing has been saved yet (never raises).
    """
    _migrate_legacy_config()
    try:
        with open(PREF_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def set_preference(key: str, value) -> dict:
    """
    Remember one user preference across sessions (merge-write).

    Args:
        key: Preference name (see the conventional keys documented above)
        value: Any JSON-serializable value

    Returns:
        The full preference dict after the update.
    """
    prefs = get_preferences()
    if value is None:
        prefs.pop(key, None)
    else:
        prefs[key] = value
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(PREF_FILE, "w", encoding="utf-8") as f:
        json.dump(prefs, f, ensure_ascii=False, indent=2)
    print(f"[prefs] {key} = {value!r} -> {PREF_FILE}")
    return prefs


def clear_preferences() -> None:
    """Forget all saved preferences (removes the file)."""
    try:
        os.remove(PREF_FILE)
        print(f"[prefs] cleared: {PREF_FILE}")
    except FileNotFoundError:
        pass


def get_icon_config(mask_key: bool = True) -> dict:
    """
    Read the icon-generation API config from ~/.vba-dev/icon_config.json
    (kept OUTSIDE the skill dir so publishing the skill never leaks it).
    The api_key resolves from env var ICON_API_KEY first, then the file.
    Legacy configs (in-skill or ~/.claude/vba-dev/) are migrated automatically.
    mask_key=True replaces the key with a masked form for display.
    """
    _migrate_legacy_config()

    cfg = {}
    if os.path.exists(ICON_CONFIG_FILE):
        try:
            with open(ICON_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as e:
            print(f"Warning: could not read {ICON_CONFIG_FILE}: {e}")
    cfg.setdefault("model", DEFAULT_ICON_MODEL)
    if os.environ.get("ICON_API_KEY"):
        cfg["api_key"] = os.environ["ICON_API_KEY"]
        cfg["api_key_source"] = "env"
    if mask_key and cfg.get("api_key"):
        key = str(cfg["api_key"])
        cfg["api_key"] = key[:8] + "..." if len(key) > 8 else "***"
    return cfg


def set_icon_config(provider: str = None, base_url: str = None,
                    model: str = None, api_key: str = None) -> dict:
    """
    Configure the image-generation API for icon creation (merged update).
    OpenAI-compatible /v1/images/generations endpoint — works with OpenAI
    gateways and most domestic providers (通义/智谱/自建网关).

    Only base_url + api_key are required — model defaults to
    DEFAULT_ICON_MODEL (gpt-image-1). Config is stored OUTSIDE the skill
    directory (~/.vba-dev/icon_config.json), so publishing the skill
    (e.g. to GitHub) never leaks credentials.
    Prefer the ICON_API_KEY env var over writing the key into the file.
    """
    cfg = get_icon_config(mask_key=False)
    for k, v in (("provider", provider), ("base_url", base_url),
                 ("model", model), ("api_key", api_key)):
        if v is not None:
            cfg[k] = v
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(ICON_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    print(f"[icon-config] saved -> {ICON_CONFIG_FILE}（skill 目录外，可安全发布）")
    return get_icon_config()  # masked echo


def generate_icon(prompt: str, out_png: str, size: int = 1024) -> str:
    """
    Generate an icon image via the configured image model (OpenAI-compatible
    /images/generations). Handles both b64_json and url responses. The raw
    output is usually large — run prepare_icon() before embedding.

    Guarded: missing config fails FAST with an actionable message (no
    network call); runtime errors (unreachable / bad key / wrong endpoint)
    are translated into short diagnostics. Pre-check with check_icon_api().
    Offline fallback without any API: draw_icon().
    """
    cfg = get_icon_config(mask_key=False)

    # --- local zero-cost guard: list everything missing, with the fix ---
    missing = [k for k in ("base_url", "model", "api_key") if not cfg.get(k)]
    if missing:
        raise ValueError(
            f"[图标API] 未配置（缺少: {', '.join(missing)}），未发起网络请求。修复：\n"
            f"  1) set_icon_config(base_url=\"https://网关/v1\", model=\"模型名\", api_key=\"sk-...\")\n"
            f"     （api_key 也可放环境变量 ICON_API_KEY，不落盘）\n"
            f"  2) check_icon_api() 验证连通后再重试\n"
            f"  无 API 场景请直接用 draw_icon() 离线绘制，不要重试本接口")

    if not str(cfg["base_url"]).startswith(("http://", "https://")):
        raise ValueError(
            f"[图标API] base_url 格式错误: {cfg['base_url']!r}（需 http(s):// 开头，"
            f"OpenAI 兼容地址通常形如 https://host/v1）")

    import requests

    base = cfg["base_url"].rstrip("/")
    # Endpoint resolution: explicit cfg["endpoint"] wins; a base_url that
    # already ends with a generation path is used as-is; otherwise try the
    # OpenAI-compatible default first, then the MiniMax-style path (404
    # fallback is free — no generation is billed for a missing endpoint).
    if cfg.get("endpoint"):
        endpoints = [base + cfg["endpoint"]]
    elif re.search(r"/generation", base):
        endpoints = [base]
    else:
        endpoints = [base + "/images/generations",      # OpenAI-compatible
                     base + "/image_generation"]        # MiniMax-style

    headers = {"Authorization": f"Bearer {cfg['api_key']}"}
    # Payload variants: rich first, leaner on rejection (providers differ in
    # which optional params they accept — e.g. MiniMax's response_format
    # only accepts 'url', so the b64_json variant gets dropped automatically)
    def payloads():
        common = {"model": cfg["model"], "prompt": prompt, "n": 1}
        yield {**common, "size": f"{size}x{size}", "response_format": "b64_json"}
        yield {**common, "size": f"{size}x{size}"}
        yield common
        yield {"model": cfg["model"], "prompt": prompt}

    resp = body = None
    last_note = ""
    param_error_keys = ("response_format", "'size'", '"size"', "'n'", '"n"')
    try:
        for ep in endpoints:
            for payload in payloads():
                resp = requests.post(ep, headers=headers, json=payload, timeout=180)
                if resp.status_code == 400:
                    continue          # optional param rejected — leaner variant
                if resp.status_code == 404:
                    last_note = f"端点不存在: {ep}"
                    resp = None
                    break             # try next endpoint
                if resp.status_code != 200:
                    break             # other error — report below
                # HTTP 200 — MiniMax-style services wrap errors inside the
                # body (base_resp.status_code != 0); treat those as failures
                body = resp.json()
                base_resp = body.get("base_resp") or {}
                err_code = base_resp.get("status_code")
                if err_code not in (None, 0):
                    msg = str(base_resp.get("status_msg", ""))
                    if any(k in msg for k in param_error_keys):
                        continue      # param rejected inside 200 — leaner variant
                    raise ValueError(
                        f"[图标API] 服务返回错误: {msg or body}"
                        f"（code {err_code}，端点 {ep}）")
                break                 # success
            if body is not None:
                break
        if body is None:
            code = resp.status_code if resp is not None else 404
            if code == 404:
                raise ValueError(
                    f"[图标API] 所有已知图像端点均不可用（{'; '.join(endpoints)}）— "
                    f"该服务可能未开通图像生成，或需在配置中指定 endpoint")
            if code in (401, 403):
                raise ValueError(
                    "[图标API] api_key 无效或过期 — set_icon_config(api_key=...) 更新，"
                    "或修正环境变量 ICON_API_KEY")
            raise ValueError(
                f"[图标API] HTTP {code}: "
                f"{(resp.text if resp is not None else last_note)[:200]}")
    except requests.exceptions.ConnectionError:
        raise ValueError(
            f"[图标API] 无法连接 {endpoints[0]} — 检查 base_url 拼写/是否缺 /v1/网络或代理。"
            f"可先 check_icon_api() 诊断")
    except requests.exceptions.Timeout:
        raise ValueError(
            "[图标API] 请求超时（180s）— 服务无响应。稍后重试或 check_icon_api() 诊断")
    except ValueError as e:
        if str(e).startswith("[图标API]"):
            raise
        # json decode failures etc.
        raise ValueError(f"[图标API] 响应非 JSON 或格式异常: {str(e)[:120]}")

    # Response parsing: OpenAI (data[0].b64_json / data[0].url),
    # MiniMax (data.image_urls[...])
    raw = None
    try:
        entry = body["data"][0] if isinstance(body.get("data"), list) else {}
        if entry.get("b64_json"):
            import base64
            raw = base64.b64decode(entry["b64_json"])
        elif entry.get("url"):
            raw = requests.get(entry["url"], timeout=180).content
        else:
            urls = body["data"].get("image_urls") if isinstance(body.get("data"), dict) else None
            if not urls:
                urls = [u for u in ([body.get("image_url")] if body.get("image_url") else [])]
            if urls:
                raw = requests.get(urls[0], timeout=180).content
    except requests.exceptions.RequestException as e:
        raise ValueError(f"[图标API] 下载生成图片失败: {e}")
    if raw is None:
        raise ValueError(
            f"[图标API] 响应里未找到图片（键: data[0].b64_json/url 或 data.image_urls）: "
            f"{str(body)[:150]}")

    out_png = os.path.abspath(out_png)
    with open(out_png, "wb") as f:
        f.write(raw)
    print(f"[icon-api] generated ({size}px) -> {out_png}")
    return out_png


def check_icon_api(quiet: bool = False) -> dict:
    """
    One-call preflight for the icon-generation API (call BEFORE
    generate_icon; cheap fail-fast, avoids blind requests):

      1. local config completeness (base_url / model / api_key)
      2. base_url format (http(s)://)
      3. endpoint reachability via GET {base_url}/models (lightweight):
         unreachable / auth_failed(401,403) / not_compatible(404) / ok

    Returns {ok, missing, reachability, detail, hint}. Prints a <=2-line
    summary unless quiet=True.
    """
    cfg = get_icon_config(mask_key=False)
    result = {"ok": False, "missing": [], "reachability": None,
              "detail": None, "hint": None}

    missing = [k for k in ("base_url", "model", "api_key") if not cfg.get(k)]
    if missing:
        result["missing"] = missing
        result["detail"] = f"缺少配置: {', '.join(missing)}"
        result["hint"] = ('set_icon_config(base_url="https://网关/v1", model="模型名", '
                          'api_key="sk-...")；key 也可用环境变量 ICON_API_KEY')
    elif not str(cfg["base_url"]).startswith(("http://", "https://")):
        result["detail"] = f"base_url 格式错误: {cfg['base_url']!r}（需 http(s):// 开头）"
        result["hint"] = "OpenAI 兼容地址通常形如 https://host/v1"
    else:
        import requests
        try:
            r = requests.get(
                cfg["base_url"].rstrip("/") + "/models",
                headers={"Authorization": f"Bearer {cfg['api_key']}"},
                timeout=15)
            if r.status_code in (401, 403):
                result["reachability"] = "auth_failed"
                result["detail"] = "端点可达，但 api_key 无效或过期"
                result["hint"] = "set_icon_config(api_key=...) 更新或修正 ICON_API_KEY"
            elif r.status_code == 404:
                result["reachability"] = "not_compatible"
                result["detail"] = f"{cfg['base_url']} 响应 404 — 可能不是 OpenAI 兼容网关"
                result["hint"] = ("确认 base_url 应含 /v1；个别网关不提供 /models，"
                                  "若地址确认无误可直接试 generate_icon()")
            else:
                # Cross-check: a 200 from /models proves gateway + key, but
                # NOT that the configured model is available on this key
                model_ids = []
                try:
                    model_ids = [m.get("id") for m in r.json().get("data", [])
                                 if m.get("id")]
                except Exception:
                    pass
                result["reachability"] = "ok"
                if model_ids and cfg["model"] not in model_ids:
                    shown = ", ".join(model_ids[:6]) + ("..." if len(model_ids) > 6 else "")
                    # Not a failure: several providers (e.g. MiniMax) keep
                    # image models off /models while they work fine — the
                    # real test is generate_icon, which auto-adapts endpoints
                    result["ok"] = True
                    result["detail"] = (f"端点可达、key 有效；/models 未列出 {cfg['model']}"
                                        f"（列出: {shown}）— 部分服务的图像模型不列入"
                                        f"/models，generate_icon 会自动适配端点")
                    result["hint"] = "若生成失败再到平台确认图像权限"
                else:
                    result["ok"] = True
                    result["detail"] = (f"配置完整，端点可达（HTTP {r.status_code}，"
                                        f"model={cfg['model']}）")
        except requests.exceptions.ConnectionError:
            result["reachability"] = "unreachable"
            result["detail"] = f"无法连接 {cfg['base_url']}"
            result["hint"] = "检查 URL 拼写/代理/网络连通性"
        except requests.exceptions.Timeout:
            result["reachability"] = "timeout"
            result["detail"] = "端点 15s 无响应"
            result["hint"] = "稍后重试或检查服务状态"

    if not quiet:
        mark = "OK" if result["ok"] else "FAIL"
        print(f"[icon-api] [{mark}] {result['detail']}"
              + (f" | {result['hint']}" if result["hint"] else ""))
    return result


def _load_cjk_font(px: int):
    """Load a CJK-capable system font (msyh/simhei), fallback to default."""
    from PIL import ImageFont
    for name in ("msyh.ttc", "msyhbd.ttc", "simhei.ttf", "seguisb.ttf"):
        path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, px)
            except Exception:
                pass
    return ImageFont.load_default()


def draw_icon(text: str, out_png: str, bg=(31, 78, 146), fg=(255, 255, 255),
              canvas: int = 64) -> str:
    """
    Offline fallback: draw a rounded-square icon with 1-2 characters
    (Chinese or Latin). Drawn at high resolution so prepare_icon() can
    downscale smoothly to 32px.
    """
    from PIL import Image, ImageDraw

    img = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, canvas - 1, canvas - 1],
                        radius=int(canvas * 0.22), fill=tuple(bg) + (255,))
    txt = str(text)[:2]
    px = int(canvas * (0.56 if len(txt) == 1 else 0.38))
    d.text((canvas / 2, canvas / 2), txt, font=_load_cjk_font(px),
           fill=tuple(fg) + (255,), anchor="mm")
    out_png = os.path.abspath(out_png)
    img.save(out_png)
    print(f"[icon-draw] drawn '{txt}' -> {out_png}")
    return out_png


def prepare_icon(src: str, out_png: str, size: int = 32, rounded: bool = True,
                 radius_ratio: float = 0.22, white_to_alpha: bool = False) -> str:
    """
    Normalize an icon image for Ribbon embedding:

      - center-crop + LANCZOS downscale to exactly size x size (models return
        1024px images; embedding those directly breaks the button)
      - optional rounded-corner alpha mask (supersampled for smooth edges)
      - optional near-white -> transparent (fallback for solid-white model
        backgrounds; ask the model for a transparent background instead)

    Args:
        src: source image (any size/format PIL can open)
        out_png: output PNG path
        size: target size in px — 32 for large buttons, 16 for normal
        rounded: apply rounded-corner alpha
        white_to_alpha: convert near-white pixels to transparent
    """
    from PIL import Image, ImageChops, ImageDraw, ImageOps

    img = Image.open(src).convert("RGBA")
    img = ImageOps.fit(img, (size, size), Image.LANCZOS)

    if white_to_alpha:
        px = img.load()
        for y in range(size):
            for x in range(size):
                r, g, b, a = px[x, y]
                if r >= 245 and g >= 245 and b >= 245:
                    px[x, y] = (r, g, b, 0)

    if rounded:
        ss = size * 4  # supersample the mask for smooth corners
        mask = Image.new("L", (ss, ss), 0)
        ImageDraw.Draw(mask).rounded_rectangle(
            [0, 0, ss - 1, ss - 1], radius=int(ss * radius_ratio), fill=255)
        mask = mask.resize((size, size), Image.LANCZOS)
        # combine with any existing alpha (e.g. white_to_alpha result)
        img.putalpha(ImageChops.multiply(img.getchannel("A"), mask))

    out_png = os.path.abspath(out_png)
    img.save(out_png)
    print(f"[icon-prepare] {size}x{size}{' rounded' if rounded else ''}"
          f"{' white->alpha' if white_to_alpha else ''} -> {out_png}")
    return out_png


def add_icon_button(xlam_dir: str, group_id: str, button_id: str, label: str,
                    on_action: str, icon_png: str, screentip: str = "",
                    supertip: str = "", size: str = "large", rounded: bool = True,
                    white_to_alpha: bool = False,
                    after_button_id: str = None) -> str:
    """
    One-call custom-icon button: prepare the image to 32x32 (rounded),
    register it under `button_id`, and add the button referencing it.

    Args:
        icon_png: source image (any size; e.g. straight from generate_icon
                  or draw_icon — it gets normalized here)
    Returns the button XML added.
    """
    prepared = os.path.join(os.environ.get("TEMP", "."),
                            f"_prepared_{button_id}.png")
    prepare_icon(icon_png, prepared, size=32, rounded=rounded,
                 white_to_alpha=white_to_alpha)
    register_icon(xlam_dir, button_id, prepared)
    btn = generate_button_xml(id=button_id, label=label, on_action=on_action,
                              size=size, image=button_id,
                              screentip=screentip, supertip=supertip)
    add_button_to_ribbon(xlam_dir, group_id, btn, after_button_id=after_button_id)
    return btn


def validate_imagemso(names, xlam_path: str = None) -> dict:
    """
    Verify imageMso names REALLY exist in this Office installation.

    Runs the check IN-PROCESS via an injected probe macro: pywin32's
    out-of-process CommandBars.GetImageMso fails for every name (binding
    quirk), so an external check is useless here. Use COMMON_IMAGEMSO for
    pre-verified names.

    Args:
        names: single name or list of names
        xlam_path: any macro-enabled workbook; a throwaway one is created
                   when omitted

    Returns:
        {name: bool} and prints the valid/invalid lists.
    """
    if isinstance(names, str):
        names = [names]

    own_tmp = False
    if xlam_path is None:
        own_tmp = True
        xlam_path = os.path.join(os.environ.get("TEMP", "."),
                                 "xlam_toolkit_mso_probe.xlsm")
        excel, attached = _acquire_excel(False)
        _apply_app_settings(excel, attached, "write")
        wb = excel.Workbooks.Add()
        if os.path.exists(xlam_path):
            os.remove(xlam_path)
        wb.SaveAs(xlam_path, FileFormat=XLSM_FORMAT)
        wb.Close(SaveChanges=False)

    lines = ["Public Function RunTest() As String",
             "    Dim s As String",
             "    On Error Resume Next"]
    for n in names:
        lines.append(f'    Err.Clear')
        lines.append(f'    Application.CommandBars.GetImageMso "{n}", 32, 32')
        lines.append(f'    s = s & "{n}=" & IIf(Err.Number = 0, "1", "0") & ","')
    lines += ["    RunTest = s", "End Function"]

    r = run_test(xlam_path, "\n".join(lines), timeout=120)
    if own_tmp:
        try:
            os.remove(xlam_path)
        except OSError:
            pass

    if not r.get("ok"):
        raise RuntimeError(f"imageMso 探针执行失败: {r.get('error')}")

    result = {k: (v == "1")
              for k, v in (pair.split("=") for pair in r["result"].rstrip(",").split(","))}
    valid = sorted(k for k, v in result.items() if v)
    invalid = sorted(k for k, v in result.items() if not v)
    print(f"[imagemso] 有效 ({len(valid)}): {', '.join(valid) or '无'}")
    print(f"[imagemso] 无效 ({len(invalid)}): {', '.join(invalid) or '无'}")
    return result


# ============================================================================
# Section 8.6: Ribbon Preview Rendering (customUI.xml -> PNG mockup)
# ============================================================================
#
# A Ribbon is invisible until the add-in is loaded in Excel — neither the user
# nor the AI can eyeball "is the button in the right group / is the label too
# long". This renders a faithful mock of tab / group / control layout as a PNG
# so mistakes are caught BEFORE installing anything. Pure Pillow drawing over
# the parsed customUI.xml — it never starts Excel.

_PV_TABSTRIP = (237, 235, 233)
_PV_BODY = (250, 249, 248)
_PV_TAB_ACTIVE = (255, 255, 255)
_PV_ACCENT = (31, 78, 146)
_PV_TEXT = (50, 49, 48)
_PV_DIM = (96, 94, 92)
_PV_LINE = (225, 223, 221)
_PV_ICON_BOX = (222, 226, 230)

# Layout metrics (px), loosely matched to Office's own proportions
_PV_PAD_X, _PV_PAD_Y = 16, 10
_PV_TAB_H, _PV_TAB_PAD = 30, 16
_PV_GROUP_LABEL_H, _PV_GROUP_SEP_W = 20, 14
_PV_LARGE_W, _PV_LARGE_H, _PV_LARGE_ICON = 70, 74, 32
_PV_SMALL_COL_W, _PV_SMALL_ROW_H, _PV_SMALL_ICON = 132, 22, 16


def _local_tag(tag: str) -> str:
    """'{ns}button' -> 'button' (customUI 2006/2009 命名空间通用)."""
    return tag.rsplit("}", 1)[-1]


def _pv_kind(elem) -> str:
    """把 group 的子节点分类为 'large' / 'small' / 'separator'。"""
    tag = _local_tag(elem.tag)
    if tag == "separator":
        return "separator"
    if (elem.get("size") or "").lower() == "large":
        return "large"
    if tag in ("menu", "splitButton", "gallery", "dynamicMenu"):
        return "large"
    return "small"


def _pv_text_w(draw, text, font) -> float:
    try:
        return draw.textlength(text, font=font)
    except Exception:
        try:
            return font.getbbox(text)[2]
        except Exception:
            return len(text) * 8


def _pv_wrap(draw, text, font, max_w, max_lines=2) -> list:
    """按宽度贪心断行（中文逐字符断，不依赖空格）；超出 max_lines 时末行加省略号。"""
    text = (text or "").strip()
    if not text:
        return []
    lines, cur = [], ""
    for ch in text:
        if not cur or _pv_text_w(draw, cur + ch, font) <= max_w:
            cur += ch
            continue
        lines.append(cur)
        if len(lines) == max_lines:
            lines[-1] = lines[-1][:-1] + "…"
            return lines
        cur = ch
    if cur and len(lines) < max_lines:
        lines.append(cur)
    return lines


def _pv_load_icons(base_dir: str) -> dict:
    """customUI.xml.rels 里的 Id -> PNG 绝对路径。"""
    rels_path = os.path.join(base_dir, "customUI", "_rels", "customUI.xml.rels")
    out = {}
    if not os.path.exists(rels_path):
        return out
    try:
        root = ET.fromstring(_read_xml_part(rels_path))
    except Exception:
        return out
    for rel in root.iter():
        if _local_tag(rel.tag) != "Relationship":
            continue
        rid, target = rel.get("Id"), rel.get("Target")
        if rid and target:
            out[rid] = os.path.join(base_dir, "customUI", target.replace("/", os.sep))
    return out


def _resolve_ribbon_source(source: str):
    """Return (xml_text, base_dir, tmp_dir_to_clean).

    `source` 可以是解包目录，也可以是 .xlam/.xlsm——后者解包到临时目录，
    tmp_dir_to_clean 非 None，调用方负责删除。
    """
    src = os.path.abspath(source)
    if os.path.isdir(src):
        return get_ribbon_xml(src), src, None
    if not os.path.exists(src):
        raise FileNotFoundError(f"路径不存在: {src}")
    tmp = tempfile.mkdtemp(prefix="ribbon_preview_")
    try:
        unpack_xlam(src, tmp)
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    return get_ribbon_xml(tmp), tmp, tmp


def _pv_draw_icon(img, draw, elem, icons, cx, cy, size):
    """画控件图标：有注册的 PNG 就用它，否则（imageMso/getImage）画占位方块。"""
    from PIL import Image
    ref = elem.get("image")
    path = icons.get(ref) if ref else None
    if path and os.path.exists(path):
        try:
            ic = Image.open(path).convert("RGBA")
            ic = ic.resize((size, size), Image.LANCZOS)
            img.alpha_composite(ic, (int(cx - size / 2), int(cy - size / 2)))
            return
        except Exception:
            pass
    draw.rounded_rectangle(
        [cx - size / 2, cy - size / 2, cx + size / 2, cy + size / 2],
        radius=max(3, size // 8), fill=_PV_ICON_BOX)


def render_ribbon_preview(source: str, out_png: str = "ribbon_preview.png",
                          width: int = 1000, tab_index: int = 0) -> str:
    """
    把 customUI.xml 渲染成一张模拟功能区 PNG（纯绘图，不启动 Excel）。

    用途：功能区在装进 Excel 之前完全不可见——用它先看一眼按钮有没有加进
    正确的分组、label 会不会太长、图标有没有生效。

    Args:
        source: 解包目录，或 .xlam/.xlsm 文件（后者会解包到临时目录）
        out_png: 输出 PNG 路径
        width: 画布宽度（分组会均分拉伸填满；内容更宽时自动加宽）
        tab_index: 预览第几个标签页（默认第一个）

    Returns:
        输出 PNG 的绝对路径。
    """
    xml_text, base_dir, tmp_dir = _resolve_ribbon_source(source)
    try:
        return _pv_render(xml_text, base_dir, out_png, width, tab_index)
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _pv_render(xml_text, base_dir, out_png, width, tab_index) -> str:
    from PIL import Image, ImageDraw

    root = ET.fromstring(xml_text)
    ribbon = next((e for e in root.iter() if _local_tag(e.tag) == "ribbon"), None)
    if ribbon is None:
        raise ValueError("customUI.xml 里没有 <ribbon> 节点")

    tabs_el = next((c for c in ribbon if _local_tag(c.tag) == "tabs"), None)
    tabs = [c for c in (list(tabs_el) if tabs_el is not None else [])
            if _local_tag(c.tag) == "tab"]
    if not tabs:
        tabs = [e for e in ribbon.iter() if _local_tag(e.tag) == "tab"]
    if not tabs:
        raise ValueError("customUI.xml 里没有 <tab>——功能区至少要有一个标签页")

    tab = tabs[min(tab_index, len(tabs) - 1)]
    groups = [c for c in tab if _local_tag(c.tag) == "group"]
    icons = _pv_load_icons(base_dir)

    f_tab = _load_cjk_font(13)
    f_group = _load_cjk_font(12)
    f_btn = _load_cjk_font(12)
    f_small = _load_cjk_font(12)

    # --- 预计算每个分组的自然宽度 / 内容高度 ---
    plan = []
    control_count = 0
    for g in groups:
        larges, smalls = [], []
        for c in list(g):
            k = _pv_kind(c)
            if k == "separator":
                larges.append(c)
            elif k == "large":
                larges.append(c)
            else:
                smalls.append(c)
        control_count += len(larges) + len(smalls)
        larges_w = sum(_PV_GROUP_SEP_W if _local_tag(c.tag) == "separator"
                       else _PV_LARGE_W for c in larges)
        small_cols = (len(smalls) + 2) // 3
        gw = larges_w + small_cols * _PV_SMALL_COL_W + 2 * _PV_PAD_X
        rows = min(max(len(smalls), 1), 3)
        gh = max(_PV_LARGE_H, rows * _PV_SMALL_ROW_H) + _PV_GROUP_LABEL_H
        plan.append({"el": g, "larges": larges, "smalls": smalls,
                     "w": max(gw, _PV_LARGE_W + 2 * _PV_PAD_X), "h": gh})

    natural = sum(p["w"] for p in plan) or width
    W = max(width, natural)
    if plan and natural < W:
        extra = (W - natural) // len(plan)
        for p in plan:
            p["w"] += extra
    body_h = (max((p["h"] for p in plan), default=_PV_LARGE_H + _PV_GROUP_LABEL_H)
              + 2 * _PV_PAD_Y)
    H = _PV_TAB_H + body_h

    img = Image.new("RGBA", (W, H), _PV_BODY + (255,))
    d = ImageDraw.Draw(img)

    # --- 标签条 ---
    d.rectangle([0, 0, W, _PV_TAB_H], fill=_PV_TABSTRIP + (255,))
    d.line([0, _PV_TAB_H, W, _PV_TAB_H], fill=_PV_LINE + (255,), width=1)
    tx = 8
    for t in tabs:
        lbl = t.get("label") or t.get("idMso") or t.get("id") or "?"
        tw = _pv_text_w(d, lbl, f_tab) + _PV_TAB_PAD * 2
        active = (t is tab)
        if active:
            d.rectangle([tx, 2, tx + tw, _PV_TAB_H], fill=_PV_TAB_ACTIVE + (255,))
            d.rectangle([tx, 0, tx + tw, 3], fill=_PV_ACCENT + (255,))
        d.text((tx + tw / 2, _PV_TAB_H / 2 + 1), lbl, font=f_tab,
               fill=(_PV_TEXT if active else _PV_DIM) + (255,), anchor="mm")
        tx += tw + 2

    # --- 分组 ---
    gx = 0
    for i, p in enumerate(plan):
        gw = p["w"]
        if i:
            sx = gx
            d.line([sx, _PV_TAB_H + _PV_PAD_Y + 4, sx, _PV_TAB_H + body_h - _PV_PAD_Y - 4],
                   fill=_PV_LINE + (255,), width=1)
        top = _PV_TAB_H + _PV_PAD_Y
        cx = gx + _PV_PAD_X
        row_cy = top + _PV_LARGE_H / 2 - 8

        # 大控件
        for c in p["larges"]:
            if _local_tag(c.tag) == "separator":
                d.line([cx + _PV_GROUP_SEP_W / 2, row_cy - 18,
                        cx + _PV_GROUP_SEP_W / 2, row_cy + 18],
                       fill=_PV_LINE + (255,), width=1)
                cx += _PV_GROUP_SEP_W
                continue
            _pv_draw_icon(img, d, c, icons, cx + _PV_LARGE_W / 2,
                          top + 8 + _PV_LARGE_ICON / 2, _PV_LARGE_ICON)
            lbl = c.get("label") or c.get("id") or ""
            extra = "▾" if _local_tag(c.tag) in ("menu", "splitButton", "dropDown") else ""
            lines = _pv_wrap(d, lbl, f_btn, _PV_LARGE_W - 6)
            ly = top + 8 + _PV_LARGE_ICON + 6
            for ln in lines:
                d.text((cx + _PV_LARGE_W / 2, ly), ln + (extra if ln == lines[-1] else ""),
                       font=f_btn, fill=_PV_TEXT + (255,), anchor="ma")
                ly += 14
            cx += _PV_LARGE_W

        # 小控件（竖排，每列 3 个）
        col_i = 0
        for j, c in enumerate(p["smalls"]):
            if j and j % 3 == 0:
                col_i += 1
            cxx = cx + col_i * _PV_SMALL_COL_W
            cyy = top + (j % 3) * _PV_SMALL_ROW_H + _PV_SMALL_ROW_H / 2
            tag = _local_tag(c.tag)
            if tag == "separator":
                d.line([cxx, cyy, cxx + _PV_SMALL_COL_W - 10, cyy],
                       fill=_PV_LINE + (255,), width=1)
                continue
            _pv_draw_icon(img, d, c, icons, cxx + 10, cyy, _PV_SMALL_ICON)
            lbl = c.get("label") or c.get("id") or ""
            if tag in ("dropDown", "comboBox"):
                lbl += " ▾"
            d.text((cxx + 22, cyy), lbl[:18], font=f_small,
                   fill=_PV_TEXT + (255,), anchor="lm")

        # 分组名（底部居中）
        glabel = p["el"].get("label") or p["el"].get("id") or ""
        d.text((gx + gw / 2, _PV_TAB_H + body_h - _PV_PAD_Y - _PV_GROUP_LABEL_H / 2),
               glabel, font=f_group, fill=_PV_DIM + (255,), anchor="mm")
        gx += gw

    out_png = os.path.abspath(out_png)
    img.convert("RGB").save(out_png)
    print(f"[preview] {len(tabs)} 个标签页 / 当前页 {len(groups)} 个分组 / "
          f"{control_count} 个控件 -> {out_png}  ({W}x{H})")
    return out_png


# ============================================================================
# Section 9: Workbook Session (open/save/close with safety)
# ============================================================================

_APP_SETTINGS = ("DisplayAlerts", "EnableEvents", "ScreenUpdating", "AutomationSecurity")


def _apply_app_settings(excel, attached: bool, mode: str, security: str = None) -> dict:
    """Apply per-mode app settings; returns the snapshot to restore (attached only).

    Modes:
      read  - read-only open, macros force-disabled
      edit  - normal open, macros force-disabled (in-memory changes, no save)
      run   - normal open, macros ENABLED, events ON
      write - normal open, macros force-disabled, save on success

    security='low' overrides the per-mode macro setting. Needed by build_check:
    AutomationSecurity is applied at OPEN time and cannot be lowered later —
    a workbook opened under ForceDisable makes VBE compile commands and macro
    runs silently ineffective ("需要重新打开此工作簿" to enable macros).
    """
    snapshot = {}
    if attached:
        # Snapshot user's settings so we can restore them
        for prop in _APP_SETTINGS:
            try:
                snapshot[prop] = getattr(excel, prop)
            except Exception:
                pass
    try:
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        if mode == "run" or security == "low":
            excel.AutomationSecurity = MSO_SECURITY_LOW
            excel.EnableEvents = (mode == "run")
        else:
            # ForceDisable blocks Workbook_Open etc. when we only edit code
            excel.AutomationSecurity = MSO_SECURITY_FORCE_DISABLE
            excel.EnableEvents = False
    except Exception as e:
        print(f"Warning: could not apply app settings: {e}")
    return snapshot


def _restore_app_settings(excel, snapshot: dict) -> None:
    for prop, value in snapshot.items():
        try:
            setattr(excel, prop, value)
        except Exception:
            pass


def _has_mark_of_the_web(path: str) -> bool:
    """True when the file carries an NTFS Zone.Identifier ADS — i.e. it was
    downloaded from the internet / saved from an email attachment. Pure
    filesystem probe (no COM); non-NTFS filesystems simply have no ADS."""
    try:
        with open(path + ":Zone.Identifier", "rb"):
            return True
    except OSError:
        return False


def _refuse_untrusted(path: str, assume_trusted: bool = False) -> None:
    """Gate before opening a file with macros ENABLED (run_macro / run_test /
    build_check's real compile). The isolated Excel instance is NOT a sandbox:
    Workbook_Open / Auto_Open run with the user's full privileges (Shell,
    FileSystemObject, WinAPI all reachable). A MOTW-flagged file must not go
    through this path without an explicit trust decision."""
    if assume_trusted or not _has_mark_of_the_web(path):
        return
    raise ValueError(
        f"文件带网络下载标记（Zone.Identifier），拒绝以宏启用方式打开: {path}\n"
        "打开瞬间 Workbook_Open/Auto_Open 将以你的用户权限执行（隔离实例不是沙箱）。\n"
        "处理方式：① export_vba_source 导出源码人工审阅（重点 Declare/Shell/"
        "CreateObject）确认无害后，传 assume_trusted=True；② 确认可信后由用户"
        "手动解除锁定再试。")


@contextlib.contextmanager
def _open_workbook(file_path: str, mode: str = "write", attach: bool = False,
                   backup_reason: str = "write", save: bool = None,
                   security: str = None):
    """
    Open a workbook in a managed COM session.

    Modes:
      'write' - ensure xlsm/xlam, back up first, save on success
      'edit'  - open normally, modify in memory, never save
      'read'  - open read-only
      'run'   - open with macros ENABLED, events ON (never saves by itself)

    save: overrides the save-on-success default (write=True, others=False).
    Used by run_macro(save=True) to persist macro effects while keeping
    macros enabled ('run' settings) — a 'write'-mode open would force-disable
    macros and every Run would fail with '宏被禁用'.

    Yields (excel, workbook, resolved_path). On exception the workbook is
    closed without saving, so a failed operation leaves the file untouched.
    """
    if save is None:
        save = (mode == "write")
    path = os.path.abspath(file_path)
    if mode == "write":
        path = _ensure_xlsm(path)
    else:
        path = _require_macro_file(path)
    if (mode == "write" or save) and os.path.exists(path):
        backup_file(path, reason=backup_reason)

    excel, attached = _acquire_excel(attach)
    snapshot = _apply_app_settings(excel, attached, mode, security=security)

    if not attached:
        # Close a stale copy left open in OUR cached instance from an earlier
        # failed call (never closes the user's workbooks — that's attach only).
        try:
            for w in list(excel.Workbooks):
                full = w.FullName
                if full and os.path.normcase(os.path.abspath(full)) == os.path.normcase(path):
                    w.Close(SaveChanges=False)
        except Exception:
            pass

    wb = None
    try:
        wb = excel.Workbooks.Open(path, ReadOnly=(mode == "read"), UpdateLinks=0)
        yield excel, wb, path
        if mode == "write" or save:
            try:
                wb.Save()
            except Exception as e:
                # 看门狗强杀后 Excel 已死，收尾的 Save 必抛 com_error——不能让
                # 已妥善构造的超时结果在退出路径上变成未处理异常
                print(f"Warning: save-on-close skipped ({e})")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if attached:
            _restore_app_settings(excel, snapshot)


# ============================================================================
# Section 10: VBA Code Read/Write Functions
# ============================================================================

def _set_component_code(proj, name: str, comp_type: int, code: str):
    """
    Set the full code of component `name` (type comp_type).

    If a same-name, same-type component exists, its code body is cleared and
    rewritten IN PLACE — module-level attributes (VB_Name, Option Private
    Module, ...) are preserved. If types differ, the component is removed and
    recreated.

    Returns (component, action) with action in:
      'replaced-in-place' / 'recreated' / 'created'
    """
    existing = None
    try:
        existing = proj.VBComponents(name)
    except Exception:
        pass

    if existing is not None:
        if existing.Type == comp_type:
            cm = existing.CodeModule
            if cm.CountOfLines:
                cm.DeleteLines(1, cm.CountOfLines)
            if code.strip():
                cm.AddFromString(code)
            return existing, "replaced-in-place"
        proj.VBComponents.Remove(existing)

    comp = proj.VBComponents.Add(comp_type)
    comp.Name = name
    if code.strip():
        comp.CodeModule.AddFromString(code)
    return comp, "created"


def add_vba_module(xlam_path: str, module_name: str, code: str, attach: bool = False) -> None:
    """
    Add a VBA standard module, or replace the code of an existing one
    (module attributes are preserved on replace).

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name for the module
        code: VBA code to add
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        _, action = _set_component_code(wb.VBProject, module_name, VBEXT_CT_STDMODULE, code)
        print(f"{action}: module {module_name} ({os.path.basename(path)})")


def add_vba_class(xlam_path: str, class_name: str, code: str, attach: bool = False) -> None:
    """
    Add a VBA class module, or replace the code of an existing one
    (class attributes are preserved on replace).

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        class_name: Name for the class module
        code: VBA class code
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        _, action = _set_component_code(wb.VBProject, class_name, VBEXT_CT_CLASSMODULE, code)
        print(f"{action}: class {class_name} ({os.path.basename(path)})")


def add_vba_callback(
    xlam_path: str,
    module_name: str,
    callback_name: str,
    code: str,
    attach: bool = False
) -> None:
    """
    Append a callback procedure to an existing VBA module (created if missing).
    No-op if a procedure with the same name already exists.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of existing module
        callback_name: Name of the new callback
        code: VBA callback code
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        proj = wb.VBProject
        try:
            module = proj.VBComponents(module_name)
        except Exception:
            module = proj.VBComponents.Add(VBEXT_CT_STDMODULE)
            module.Name = module_name

        cm = module.CodeModule
        lines = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        if _proc_decl_re(callback_name).search(lines):
            print(f"Callback '{callback_name}' already exists — skipped")
            return

        n = cm.CountOfLines
        cm.InsertLines(n + 1, code)
        print(f"Added callback: {callback_name} to module {module_name}")


def read_vba_module(xlam_path: str, module_name: str, attach: bool = False) -> str:
    """
    Read VBA code from a module.

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        module_name: Name of module to read
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        VBA code content ("" on error)
    """
    with _open_workbook(xlam_path, mode="read", attach=attach) as (excel, wb, path):
        module = wb.VBProject.VBComponents(module_name)
        cm = module.CodeModule
        return cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""


def list_procedures(xlam_path: str, module_name: str, attach: bool = False) -> dict:
    """
    Read all procedures from a VBA module.

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        module_name: Name of module to read
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        Dict mapping procedure names to their code and metadata:
        {
            "procedures": {
                "ProcName": {
                    "type": "Sub|Function|Property",
                    "body_line": 5,
                    "start_line": 5,
                    "end_line": 10,
                    "line_count": 6,
                    "code": "Sub ProcName()\n..."
                },
                ...
            },
            "declarations": "Option Explicit\n...",
            "module_name": "modAI"
        }
    """
    result = {"procedures": {}, "declarations": "", "module_name": module_name}
    with _open_workbook(xlam_path, mode="read", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule

        total_lines = code_mod.CountOfLines
        decl_count = code_mod.CountOfDeclarationLines

        if decl_count > 0:
            result["declarations"] = code_mod.Lines(1, decl_count)

        seen = set()
        for line_num in range(decl_count + 1, total_lines + 1):
            try:
                proc_name = _proc_of_line_result(code_mod.ProcOfLine(line_num, 0))
            except Exception:
                continue
            if proc_name and proc_name not in seen:
                seen.add(proc_name)
                try:
                    body_line = code_mod.ProcBodyLine(proc_name, 0)
                    _, span = _proc_body_span(code_mod, proc_name)
                    proc_lines = span
                    proc_code = code_mod.Lines(body_line, proc_lines)

                    first_line = proc_code.split('\n')[0]
                    proc_type = _proc_type_of_decl(first_line)

                    result["procedures"][proc_name] = {
                        "type": proc_type,
                        "body_line": body_line,
                        "start_line": body_line,
                        "end_line": body_line + proc_lines - 1,
                        "line_count": proc_lines,
                        "code": proc_code
                    }
                except Exception as e:
                    print(f"Warning: Could not read procedure '{proc_name}': {e}")
        return result


def delete_vba_module(xlam_path: str, module_name: str, attach: bool = False) -> None:
    """
    Delete a VBA module.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module to delete
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        module = wb.VBProject.VBComponents(module_name)
        wb.VBProject.VBComponents.Remove(module)
        print(f"Deleted VBA module: {module_name}")


def list_vba_modules(xlam_path: str, attach: bool = False) -> dict:
    """
    List all VBA components in the workbook.

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        Dict with component types and names:
        {'modules': [...], 'classes': [...], 'forms': [...], 'documents': [...]}
    """
    result = {
        "modules": [],       # Standard modules
        "classes": [],       # Class modules
        "forms": [],         # UserForms
        "documents": [],     # Sheet / ThisWorkbook document modules
    }
    with _open_workbook(xlam_path, mode="read", attach=attach) as (excel, wb, path):
        type_names = {
            VBEXT_CT_STDMODULE: "modules",
            VBEXT_CT_CLASSMODULE: "classes",
            VBEXT_CT_MSFORM: "forms",
            VBEXT_CT_DOCUMENT: "documents",
        }
        for comp in wb.VBProject.VBComponents:
            type_name = type_names.get(comp.Type)
            if type_name:
                result[type_name].append(comp.Name)
        return result


def replace_vba_module(xlam_path: str, module_name: str, new_code: str, attach: bool = False) -> None:
    """
    Replace entire module content (clear and rewrite in place — module
    attributes preserved).

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module to replace
        new_code: New VBA code content
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        try:
            module = wb.VBProject.VBComponents(module_name)
        except Exception:
            print(f"Module '{module_name}' not found — nothing replaced")
            return
        _, action = _set_component_code(wb.VBProject, module_name, module.Type, new_code)
        print(f"{action}: module {module_name}")


def update_vba_callback(
    xlam_path: str,
    module_name: str,
    callback_name: str,
    new_code: str,
    attach: bool = False
) -> bool:
    """
    Update or add a single callback procedure.

    Uses ProcBodyLine/ProcCountLines to find the procedure and replace it.
    If procedure doesn't exist, appends it.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module
        callback_name: Procedure name to update
        new_code: New VBA procedure code (must start with Sub or Function)
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        True if updated, False if appended
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule

        all_lines = code_mod.Lines(1, code_mod.CountOfLines) if code_mod.CountOfLines else ""
        # Modifier-tolerant declaration match: '^(Sub|Function)' missed
        # 'Public Sub Foo' and appended a DUPLICATE procedure (compile error)
        found = _proc_decl_re(callback_name).search(all_lines) is not None

        if found:
            try:
                start_line, proc_lines = _proc_body_span(code_mod, callback_name)
                code_mod.DeleteLines(start_line, proc_lines)
                code_mod.InsertLines(start_line, new_code)
                print(f"Updated callback: {callback_name}")
                return True
            except Exception as e:
                print(f"Warning: Could not locate procedure precisely, appending: {e}")

        n = code_mod.CountOfLines
        if n > 0:
            code_mod.InsertLines(n + 1, vbCrLf + new_code)
        else:
            code_mod.AddFromString(new_code)
        print(f"Appended callback: {callback_name}")
        return False


def replace_vba_lines(
    xlam_path: str,
    module_name: str,
    start_line: int,
    line_count: int,
    new_code: str,
    attach: bool = False
) -> None:
    """
    Replace a range of lines in a module.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module
        start_line: Starting line number (1-based)
        line_count: Number of lines to replace
        new_code: New code to insert (can be multiline)
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule
        code_mod.DeleteLines(start_line, line_count)
        code_mod.InsertLines(start_line, new_code)
        print(f"Replaced lines {start_line}-{start_line + line_count - 1} in {module_name}")


def get_procedure_info(xlam_path: str, module_name: str, procedure_name: str,
                       attach: bool = False) -> dict:
    """
    Get information about a VBA procedure.

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        module_name: Name of module
        procedure_name: Procedure name
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        Dict with start_line, end_line, line_count, or empty dict if not found
    """
    with _open_workbook(xlam_path, mode="read", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule
        try:
            body_line, span = _proc_body_span(code_mod, procedure_name)
            return {
                "start_line": body_line,
                "end_line": body_line + span - 1,
                "line_count": span
            }
        except Exception:
            return {}


def insert_vba_code(
    xlam_path: str,
    module_name: str,
    insert_after: str,
    new_code: str,
    attach: bool = False
) -> None:
    """
    Insert code after a specific procedure.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module
        insert_after: Procedure name to insert after
        new_code: Code to insert
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule
        try:
            start_line = code_mod.ProcStartLine(insert_after, 0)
            proc_lines = code_mod.ProcCountLines(insert_after, 0)
            insert_line = min(start_line + proc_lines, code_mod.CountOfLines + 1)
        except Exception:
            insert_line = code_mod.CountOfLines + 1
        code_mod.InsertLines(insert_line, vbCrLf + new_code)
        print(f"Inserted code after {insert_after}")


def delete_vba_callback(xlam_path: str, module_name: str, callback_name: str,
                        attach: bool = False) -> bool:
    """
    Delete a callback procedure.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        module_name: Name of module
        callback_name: Procedure name to delete
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        True if deleted, False if not found
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(module_name).CodeModule
        try:
            start_line, proc_lines = _proc_span(code_mod, callback_name)
        except Exception:
            print(f"Callback '{callback_name}' not found")
            return False
        code_mod.DeleteLines(start_line, proc_lines)
        print(f"Deleted callback: {callback_name}")
        return True


# ============================================================================
# Section 11: Source Export / Import (AI-friendly round-trip)
# ============================================================================

def export_vba_source(xlam_path: str, out_dir: str = None, attach: bool = False) -> dict:
    """
    Export ALL VBA components of a workbook to plain source files, ready for
    direct AI/editor editing:

        标准模块   -> <Name>.bas
        类模块     -> <Name>.cls
        用户窗体   -> <Name>.frm (+ <Name>.frx)
        文档模块   -> <Name>.cls   (Sheet1 / ThisWorkbook, code only)

    Also writes manifest.json recording component names/types/files so that
    import_vba_source can faithfully rebuild (and correctly treat document
    modules as in-place code updates rather than new components).

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        out_dir: Output directory. Default: <file-without-extension>_vba/
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        The manifest dict (also written to manifest.json).
    """
    path = _require_macro_file(xlam_path)
    if out_dir is None:
        out_dir = os.path.splitext(path)[0] + "_vba"
    out_dir = os.path.abspath(out_dir)

    # Safety: only wipe a directory that we generated before (marker file)
    if os.path.isdir(out_dir):
        if not os.path.exists(os.path.join(out_dir, "manifest.json")):
            raise ValueError(
                f"目录已存在且不是本工具导出的源码目录（缺 manifest.json），拒绝清空: {out_dir}")
        shutil.rmtree(out_dir)
    os.makedirs(out_dir)

    ext_by_type = {
        VBEXT_CT_STDMODULE: (".bas", "std"),
        VBEXT_CT_CLASSMODULE: (".cls", "class"),
        VBEXT_CT_MSFORM: (".frm", "form"),
        VBEXT_CT_DOCUMENT: (".cls", "document"),
    }

    components = []
    with _open_workbook(path, mode="read", attach=attach) as (excel, wb, _):
        proj = wb.VBProject
        for comp in proj.VBComponents:
            mapped = ext_by_type.get(comp.Type)
            if not mapped:
                continue  # ActiveX designers etc. — skip
            ext, category = mapped
            filename = f"{comp.Name}{ext}"
            dest = os.path.join(out_dir, filename)
            comp.Export(dest)
            if ext in (".bas", ".cls"):
                # VBIDE Export writes the system codepage (GBK on zh Windows);
                # rewrite as UTF-8 so AI/editors can round-trip Chinese safely.
                with open(dest, "rb") as f:
                    raw = f.read()
                try:
                    text = raw.decode("utf-8")
                except UnicodeDecodeError:
                    text = raw.decode("gbk", errors="replace")
                with open(dest, "w", encoding="utf-8", newline="") as f:
                    f.write(text)
            components.append({
                "name": comp.Name,
                "type": comp.Type,
                "category": category,
                "file": filename,
            })

    manifest = {
        "source_file": path,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "components": components,
    }
    with open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"Exported {len(components)} components -> {out_dir}")
    for c in components:
        print(f"  [{c['category']:>8}] {c['file']}")
    return manifest


def _strip_vba_attributes(text: str) -> str:
    """
    Remove the exported-file boilerplate from a module file, leaving only
    the VBA code: the 'VERSION 1.0 CLASS / BEGIN..END' header block and all
    'Attribute VB_*' lines. (The VERSION/BEGIN block injected as code is a
    compile error: '缺少: 语句结束'.)
    """
    out = []
    in_begin_block = False
    for line in text.split("\n"):
        stripped = line.strip()
        if re.match(r"^\s*Attribute\s+VB_", line, re.IGNORECASE):
            continue
        upper = stripped.upper()
        if upper == "VERSION 1.0 CLASS":
            in_begin_block = True   # skip header until its END
            continue
        if in_begin_block:
            if upper == "END":
                in_begin_block = False
            continue
        out.append(line)
    return "\n".join(out).lstrip("\r\n")


def import_vba_source(xlam_path: str, source_dir: str, sync: bool = False,
                      attach: bool = False) -> dict:
    """
    Import source files (as produced by export_vba_source) back into the
    workbook. The file is backed up automatically before any change.

    - .bas / .cls(class) / .frm: existing same-name component is removed and
      re-imported from file (attributes travel inside the files).
    - document modules (Sheet1/ThisWorkbook): code is updated IN PLACE
      (document components cannot be removed/re-imported).
    - sync=True: components present in the workbook but ABSENT from the source
      directory are deleted (standard/class/form only — never documents).

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        source_dir: Directory containing exported sources + manifest.json
        sync: Delete components missing from the source directory
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        {"imported": [...], "updated": [...], "deleted": [...], "skipped": [...]}
    """
    source_dir = os.path.abspath(source_dir)
    manifest_path = os.path.join(source_dir, "manifest.json")

    if os.path.exists(manifest_path):
        with open(manifest_path, "r", encoding="utf-8") as f:
            manifest = json.load(f)
        comps = manifest.get("components", [])

    summary = {"imported": [], "updated": [], "deleted": [], "skipped": []}

    with _open_workbook(xlam_path, mode="write", attach=attach,
                        backup_reason="import") as (excel, wb, path):
        proj = wb.VBProject

        # Existing component names/types (to classify manifest-less .cls)
        existing_types = {}
        for comp in proj.VBComponents:
            existing_types[comp.Name] = comp.Type

        if not os.path.exists(manifest_path):
            # No manifest: infer from files. A .cls whose name matches an
            # existing document component is a document update; anything
            # else is a fresh class.
            comps = []
            for fname in sorted(os.listdir(source_dir)):
                name, ext = os.path.splitext(fname)
                if ext == ".bas":
                    comps.append({"name": name, "type": VBEXT_CT_STDMODULE,
                                  "category": "std", "file": fname})
                elif ext == ".cls":
                    is_doc = existing_types.get(name) == VBEXT_CT_DOCUMENT
                    comps.append({"name": name, "type": VBEXT_CT_DOCUMENT if is_doc else VBEXT_CT_CLASSMODULE,
                                  "category": "document" if is_doc else "class", "file": fname})
                elif ext == ".frm":
                    comps.append({"name": name, "type": VBEXT_CT_MSFORM,
                                  "category": "form", "file": fname})

        for info in comps:
            src = os.path.join(source_dir, info["file"])
            if not os.path.exists(src):
                summary["skipped"].append(f"{info['file']} (文件不存在)")
                continue

            category = info.get("category")
            name = info["name"]
            # manifest-less .cls that matches a non-document component → class
            if (category == "document" and name in existing_types
                    and existing_types[name] != VBEXT_CT_DOCUMENT):
                category = "class"

            if category == "document":
                try:
                    comp = proj.VBComponents(name)
                except Exception:
                    summary["skipped"].append(f"{name} (工作簿中无此文档模块)")
                    continue
                cm = comp.CodeModule
                text = _read_vba_text(src)
                code = _strip_vba_attributes(text)
                if cm.CountOfLines:
                    cm.DeleteLines(1, cm.CountOfLines)
                if code.strip():
                    cm.AddFromString(code)
                summary["updated"].append(f"{name} (document, in place)")
                continue

            # std / class / form: remove existing then Import from file.
            # VBIDE Import reads the system codepage (GBK), but our on-disk
            # .bas/.cls are UTF-8 — re-encode to a temp GBK file first.
            import_src = src
            if os.path.splitext(src)[1] in (".bas", ".cls"):
                text = _read_vba_text(src)
                base, ext = os.path.splitext(os.path.basename(src))
                import_src = os.path.join(os.path.dirname(src),
                                          f"{base}.__import_tmp__{ext}")
                with open(import_src, "w", encoding="gbk", errors="replace",
                          newline="") as f:
                    f.write(text)
            try:
                existing = proj.VBComponents(name)
                proj.VBComponents.Remove(existing)
            except Exception:
                pass
            comp = proj.VBComponents.Import(import_src)
            if import_src != src:
                try:
                    os.remove(import_src)
                except OSError:
                    pass
            try:
                if comp.Name != name:
                    comp.Name = name
            except Exception:
                pass
            summary["imported"].append(f"{name} ({category})")

        if sync:
            keep = {c["name"] for c in comps}
            for comp in list(proj.VBComponents):
                # Capture the name BEFORE Remove — the COM object is
                # invalidated once removed from the collection
                try:
                    name = comp.Name
                    ctype = comp.Type
                except Exception:
                    continue
                if ctype in (VBEXT_CT_STDMODULE, VBEXT_CT_CLASSMODULE,
                             VBEXT_CT_MSFORM) and name not in keep:
                    try:
                        proj.VBComponents.Remove(comp)
                        summary["deleted"].append(name)
                    except Exception as e:
                        summary["skipped"].append(f"{name} (删除失败: {_com_error_text(e) or e})")

    print(f"Import from {source_dir}: "
          f"{len(summary['imported'])} imported, {len(summary['updated'])} updated, "
          f"{len(summary['deleted'])} deleted, {len(summary['skipped'])} skipped")
    return summary


# ============================================================================
# Section 12: Static VBA Analysis (pure text — no dialogs, no risk)
# ============================================================================

_RE_PROC_START = re.compile(
    r"^(?:PUBLIC\s+|PRIVATE\s+|FRIEND\s+)?(?:STATIC\s+)?"
    r"(?:SUB\s|FUNCTION\s|PROPERTY\s+(?:GET|LET|SET)\s)", re.IGNORECASE)
_RE_PROC_END = re.compile(r"^END\s+(?:SUB|FUNCTION|PROPERTY)\b", re.IGNORECASE)
_RE_PROC_NAME = re.compile(
    r"^(?:PUBLIC\s+|PRIVATE\s+|FRIEND\s+)?(?:STATIC\s+)?"
    r"(?:SUB|FUNCTION|PROPERTY\s+(?:GET|LET|SET))\s+(\w+)", re.IGNORECASE)
_RE_MODULE_LEVEL = re.compile(
    r"^(?:OPTION\s|PRIVATE\s|PUBLIC\s|GLOBAL\s|DIM\s|CONST\s|DECLARE\s|"
    r"IMPLEMENTS\s|TYPE\s|ENUM\s|#|ATTRIBUTE\s)", re.IGNORECASE)
_RE_TYPE_ENUM_START = re.compile(r"^(?:PUBLIC\s+|PRIVATE\s+)?(?:TYPE|ENUM)\s+\w", re.IGNORECASE)
_RE_TYPE_ENUM_END = re.compile(r"^END\s+(?:TYPE|ENUM)\b", re.IGNORECASE)

# Block openers (multiline only; single-line If/For/Do handled by exclusions)
_RE_IF_OPEN = re.compile(r"^IF\s+.+\sTHEN\s*(?:'.*)?$", re.IGNORECASE)
_RE_FOR_OPEN = re.compile(r"^FOR\s(?:EACH\s)?\S", re.IGNORECASE)
_RE_DO_OPEN = re.compile(r"^DO(?:$|\s+(?:WHILE|UNTIL)\s)", re.IGNORECASE)
_RE_WHILE_OPEN = re.compile(r"^WHILE\s.+", re.IGNORECASE)
_RE_SELECT_OPEN = re.compile(r"^SELECT\s+CASE\s", re.IGNORECASE)
_RE_WITH_OPEN = re.compile(r"^WITH\s.+", re.IGNORECASE)
# Block closers
_RE_END_IF = re.compile(r"^END\s+IF\b", re.IGNORECASE)
_RE_NEXT = re.compile(r"^NEXT\b(.*)", re.IGNORECASE)
_RE_LOOP = re.compile(r"^LOOP\b", re.IGNORECASE)
_RE_WEND = re.compile(r"^WEND\b", re.IGNORECASE)
_RE_END_SELECT = re.compile(r"^END\s+SELECT\b", re.IGNORECASE)
_RE_END_WITH = re.compile(r"^END\s+WITH\b", re.IGNORECASE)
_RE_ELSE = re.compile(r"^(ELSE\b|ELSEIF\s)", re.IGNORECASE)
# Single-line forms that close themselves on the same line
_RE_ONELINE_FOR = re.compile(r":\s*NEXT\b", re.IGNORECASE)
_RE_ONELINE_DO = re.compile(r":\s*LOOP\b", re.IGNORECASE)

# --- Early-binding detection (distribution safety) -------------------------
# External type libraries that are NOT default references of an Excel VBA
# project. Early-binding one of these compiles fine on THIS machine but dies
# with "用户定义类型未定义" on every recipient who has not ticked the same
# reference by hand — fatal for a distributed .xlam. Each entry below is
# reachable via CreateObject() (late binding), which needs no reference.
# Host libraries (Excel / Office / MSForms / stdole / VBA) are deliberately
# absent: they are default references and carry no distribution risk.
_EARLY_BIND_LIBS = (
    "Scripting", "ADODB", "ADOX", "MSXML2", "MSXML", "WScript", "Shell",
    "VBScript_RegExp_55", "VBScript_RegExp", "Outlook", "Word", "PowerPoint",
    "CDO", "WinHttp", "WinHTTP", "InternetExplorer", "WIA", "SAPI",
)
_RE_EARLY_BIND = re.compile(
    r"\b(?:AS|NEW)\s+(" + "|".join(_EARLY_BIND_LIBS) + r")\s*\.\s*(\w+)",
    re.IGNORECASE)
# Opt-out: a  ' @allow-early-binding  comment anywhere in the module skips it.
_RE_BIND_EXEMPT = re.compile(r"@allow-early-binding", re.IGNORECASE)
# ProgIDs that do NOT equal "<Library>.<Type>" — the classic traps.
_PROGID_ALIAS = {
    "vbscript_regexp_55.regexp": "VBScript.RegExp",
    "vbscript_regexp.regexp": "VBScript.RegExp",
    "msxml2.xmlhttp60": "MSXML2.XMLHTTP",
    "msxml2.serverxmlhttp60": "MSXML2.ServerXMLHTTP",
    "msxml2.domdocument60": "MSXML2.DOMDocument",
    "msxml2.freethreadeddomdocument60": "MSXML2.FreeThreadedDOMDocument",
}


def _join_continuations(lines: list) -> list:
    """Merge VBA line-continuation lines ('... _') into single logical lines."""
    joined = []
    buf = ""
    for line in lines:
        if buf:
            # Drop the ' _' left by the previous fragment, then join with one
            # space. The underscore is consumed HERE and only here — removing
            # it at detection time as well would eat a real character off the
            # previous fragment on the next pass ("Then" -> "The").
            buf = buf.rstrip()[:-1].rstrip() + " " + line.strip()
        else:
            buf = line
        if buf.rstrip().endswith(" _"):
            continue
        joined.append(buf)
        buf = ""
    if buf:
        joined.append(buf)
    return joined


def _check_structure_in_module(module_name: str, lines: list, errors: list) -> None:
    """Flag executable code sitting outside any Sub/Function/Property, and
    unclosed Type/Enum blocks (which silently absorb everything below them)."""
    in_proc = False
    in_block = False
    block_line = 0

    for i, raw in enumerate(_join_continuations(lines), 1):
        stripped = raw.strip()
        if not stripped or stripped.startswith("'"):
            continue

        if not in_proc and _RE_TYPE_ENUM_START.match(stripped):
            in_block = True
            block_line = i
            continue
        if in_block:
            if _RE_TYPE_ENUM_END.match(stripped):
                in_block = False
            continue

        if _RE_PROC_START.match(stripped):
            in_proc = True
            continue
        if _RE_PROC_END.match(stripped):
            in_proc = False
            continue

        if in_proc:
            continue

        if not _RE_MODULE_LEVEL.match(stripped):
            errors.append({
                "module": module_name, "line": i,
                "error": f"过程外的可执行代码: {stripped[:80]}",
            })
            return  # one error per module is enough

    if in_block:
        errors.append({
            "module": module_name, "line": block_line,
            "error": "Type/Enum 块未关闭（缺少 End Type / End Enum），其后的代码都会被吞进块内",
        })


def _check_blocks_in_module(module_name: str, lines: list, errors: list) -> None:
    """Stack-check block structure: If/End If, For/Next, Do/Loop, While/Wend,
    Select/End Select, With/End With, and #If...#End If directives."""
    stack = []  # [(block_type, line_no)]
    in_proc = False

    for i, raw in enumerate(_join_continuations(lines), 1):
        stripped = raw.strip()
        if not stripped or stripped.startswith("'"):
            continue

        # Conditional compilation directives
        if stripped.startswith("#"):
            if re.match(r"#IF\s+", stripped, re.IGNORECASE):
                stack.append(("#If", i))
            elif re.match(r"#END\s+IF", stripped, re.IGNORECASE):
                if stack and stack[-1][0] == "#If":
                    stack.pop()
            continue

        if _RE_PROC_START.match(stripped):
            in_proc = True
            stack = []
            continue

        if _RE_PROC_END.match(stripped):
            if stack:
                blk_type, blk_line = stack[0]
                closer = _BLOCK_CLOSER_TEXT.get(blk_type, f"End {blk_type}")
                errors.append({
                    "module": module_name, "line": blk_line,
                    "error": f"{blk_type} 块缺少对应的 {closer}（在第 {i} 行 End Sub/Function 处仍未关闭）",
                })
            in_proc = False
            stack = []
            continue

        if not in_proc:
            continue

        # --- openers ---
        if _RE_IF_OPEN.match(stripped):
            stack.append(("If", i))
            continue
        if _RE_ELSE.match(stripped):
            continue
        if _RE_FOR_OPEN.match(stripped):
            if not _RE_ONELINE_FOR.search(stripped):
                stack.append(("For", i))
            continue
        if _RE_DO_OPEN.match(stripped):
            if not _RE_ONELINE_DO.search(stripped):
                stack.append(("Do", i))
            continue
        if _RE_WHILE_OPEN.match(stripped):
            stack.append(("While", i))
            continue
        if _RE_SELECT_OPEN.match(stripped):
            stack.append(("Select", i))
            continue
        if _RE_WITH_OPEN.match(stripped):
            stack.append(("With", i))
            continue

        # --- closers ---
        if _RE_END_IF.match(stripped):
            if stack and stack[-1][0] in ("If", "#If"):
                stack.pop()
            elif not stack:
                errors.append({
                    "module": module_name, "line": i,
                    "error": "End If 没有对应的 If",
                })
            else:
                blk_type, blk_line = stack[-1]
                errors.append({
                    "module": module_name, "line": i,
                    "error": f"此处应为 End {blk_type}（块开始于第 {blk_line} 行），却出现了 End If",
                })
            continue
        m = _RE_NEXT.match(stripped)
        if m:
            pops = 1 + m.group(1).count(",")  # Next i, j closes two loops
            for _ in range(pops):
                if stack and stack[-1][0] == "For":
                    stack.pop()
            continue
        if _RE_LOOP.match(stripped):
            if stack and stack[-1][0] == "Do":
                stack.pop()
            continue
        if _RE_WEND.match(stripped):
            if stack and stack[-1][0] == "While":
                stack.pop()
            continue
        if _RE_END_SELECT.match(stripped):
            if stack and stack[-1][0] == "Select":
                stack.pop()
            continue
        if _RE_END_WITH.match(stripped):
            if stack and stack[-1][0] == "With":
                stack.pop()
            continue

    if in_proc:
        errors.append({
            "module": module_name, "line": len(lines),
            "error": "模块在过程内部结束（缺少 End Sub / End Function / End Property）",
        })


def _check_early_binding_in_module(module_name: str, lines: list, errors: list) -> None:
    """Flag early binding to external type libraries (a distribution hazard).

    A .xlam that early-binds, say, Scripting.Dictionary compiles only on
    machines where that reference happens to be ticked; everywhere else it
    fails with "用户定义类型未定义". Late binding (CreateObject) reaches the
    same objects with no reference at all, so early binding to these libraries
    is reported as an ERROR (blocking), not a warning.

    One error per distinct library type per module. Opt out with a
    ' @allow-early-binding comment — required for WithEvents sinks, which VBA
    cannot late-bind at all.
    """
    if _RE_BIND_EXEMPT.search("\n".join(lines)):
        return

    reported = set()
    for i, raw in enumerate(_join_continuations(lines), 1):
        stripped = raw.strip()
        if not stripped or stripped.startswith("'"):
            continue
        m = _RE_EARLY_BIND.search(stripped)
        if not m:
            continue

        full = f"{m.group(1)}.{m.group(2)}"
        if full.lower() in reported:
            continue
        reported.add(full.lower())

        if re.search(r"\bWITH\s*EVENTS\b", stripped, re.IGNORECASE):
            hint = ("WithEvents 无法后期绑定（VBA 要求编译期类型）——改为不挂事件的"
                    "设计；确需保留时在模块顶部加 ' @allow-early-binding 并注明理由")
        else:
            progid = _PROGID_ALIAS.get(full.lower(), full)
            hint = f'改后期绑定：Dim x As Object: Set x = CreateObject("{progid}")'
        errors.append({
            "module": module_name, "line": i, "kind": "early_binding",
            "error": (f"前期绑定外部类型库 {full}：分发 .xlam 后，未勾选该引用的机器会"
                      f"编译失败（用户定义类型未定义）。{hint}"),
        })


def _static_check_project(proj, strict: bool = False,
                          check_binding: bool = True) -> dict:
    """Run structural checks over all code-bearing components of a project."""
    errors, warnings = [], []
    modules_checked = 0
    public_procs = {}  # name -> module (for duplicate detection)

    for comp in proj.VBComponents:
        if comp.Type not in (VBEXT_CT_STDMODULE, VBEXT_CT_CLASSMODULE,
                             VBEXT_CT_MSFORM, VBEXT_CT_DOCUMENT):
            continue
        try:
            cm = comp.CodeModule
            n = cm.CountOfLines
            if not n:
                continue
            code = cm.Lines(1, n)
        except Exception:
            continue

        modules_checked += 1
        lines = code.replace("\r\n", "\n").split("\n")
        _check_structure_in_module(comp.Name, lines, errors)
        _check_blocks_in_module(comp.Name, lines, errors)
        if check_binding:
            _check_early_binding_in_module(comp.Name, lines, errors)

        if strict:
            if "option explicit" not in code.lower():
                warnings.append(f"{comp.Name}: 缺少 Option Explicit")
            for raw in lines:
                m = _RE_PROC_NAME.match(raw.strip())
                if m and not re.match(r"^PRIVATE\s", raw.strip(), re.IGNORECASE):
                    name = m.group(1)
                    if name in public_procs and public_procs[name] != comp.Name:
                        warnings.append(
                            f"公共过程重名: {name} 同时定义于 "
                            f"{public_procs[name]} 和 {comp.Name}")
                    else:
                        public_procs[name] = comp.Name

    return {"errors": errors, "warnings": warnings, "modules_checked": modules_checked}


# ============================================================================
# Section 13: Build Check (static + real VBE compile)
# ============================================================================

def _pane_matches_project(pane, proj) -> bool:
    """Best-effort: does a VBE CodePane belong to the given VBProject?"""
    try:
        pane_proj = pane.CodeModule.Parent.Collection.Parent
        try:
            a = os.path.normcase(os.path.abspath(pane_proj.FileName))
            b = os.path.normcase(os.path.abspath(proj.FileName))
            return a == b
        except Exception:
            return pane_proj.Name == proj.Name
    except Exception:
        return False


def _ensure_code_pane(excel, wb) -> str:
    """
    Make a code pane of THIS workbook's project the active one.

    VBE 'Debug > Compile' acts on the ACTIVE project and is only reliably
    enabled when one of its code panes has focus. Best-effort; returns a
    diagnostic string.
    """
    vbe = _get_vbe(excel)
    proj = wb.VBProject
    try:
        pane = vbe.ActiveCodePane
        if pane is not None and _pane_matches_project(pane, proj):
            return f"active pane already: {pane.CodeModule.Parent.Name}"
    except Exception:
        pass
    try:
        for pane in vbe.CodePanes:
            if _pane_matches_project(pane, proj):
                pane.Show()
                return f"shown pane: {pane.CodeModule.Parent.Name}"
    except Exception:
        pass
    try:
        for comp in proj.VBComponents:
            if comp.Type == VBEXT_CT_STDMODULE and comp.CodeModule.CountOfLines > 0:
                try:
                    comp.CodeModule.CodePane.Show()
                    return f"created pane: {comp.Name}"
                except Exception:
                    try:
                        comp.Activate()
                        return f"activated: {comp.Name}"
                    except Exception:
                        pass
    except Exception:
        pass
    return "no pane activated (best-effort failed)"


def _active_error_location(excel) -> dict:
    """After a compile error, the VBE selection sits on the offending line."""
    sel = _current_selection(excel)
    if sel:
        return {"component": sel[0], "line": sel[1]}
    return None


def _probe_compile(excel, wb, timeout: int = 30) -> dict:
    """
    Definitive compile verification via an IN-PROCESS trampoline.

    Out-of-process CommandBars Execute suppresses compile-error dialogs
    (silently no-ops or navigates), so the trampoline runs the VBE
    'Debug > Compile' command (ID 578) INSIDE the Excel process via
    Application.Run — exactly like a user clicking the menu. Compile errors
    then surface as real modal dialogs, which the watchdog reads and
    dismisses; VBE.ActiveCodePane afterwards gives the error location.

    The project is dirtied first (an already-compiled state makes the
    compile command a no-op). Trampoline is removed afterwards; the
    workbook is never saved. Requires the workbook opened with
    AutomationSecurity = Low.

    Trampoline code:
        Public Sub Probe()
            Application.VBE.CommandBars.FindControl(ID:=578).Execute
        End Sub
    """
    result = {"ok": None, "error_text": None, "dialogs": [],
              "error_location": None, "timed_out": False}
    proj = wb.VBProject
    probe_name = f"zzCmpProbe{int(time.time())}"
    comp = None
    try:
        comp = proj.VBComponents.Add(VBEXT_CT_STDMODULE)
        comp.Name = probe_name
        comp.CodeModule.AddFromString(
            "Public Sub Probe()\n"
            "    Application.VBE.CommandBars.FindControl(ID:=578).Execute\n"
            "End Sub")
    except Exception as e:
        result["error_text"] = f"无法注入编译探针: {_com_error_text(e) or e}"
        return result

    # Already-compiled projects make the compile command a silent no-op
    _dirty_project(wb)

    attached = _is_attached_instance(excel)
    if attached:
        print("⚠ attach 模式：编译探针卡死时不会强杀用户 Excel（看门狗击杀已禁用）")
    wd = _run_watchdog(excel, timeout=timeout, attached=attached,
                       interval=0.4, storm_limit=6)
    wd.start()
    err = None
    ran = False
    try:
        excel.Run(f"'{wb.Name}'!Probe")
        ran = True
    except Exception as e:
        err = _com_error_text(e) or str(e)
    finally:
        wd.stop(grace=1.5)

    try:
        proj.VBComponents.Remove(comp)
    except Exception:
        pass

    result["dialogs"] = wd.dialog_texts
    if wd.killed:
        _invalidate_cached_excel()
        result["timed_out"] = True
        result["ok"] = False
        result["error_text"] = "编译探针执行卡死，已强制结束本次隔离 Excel 进程"
        return result
    if wd.dialog_texts:
        result["ok"] = False
        result["error_text"] = wd.dialog_texts[-1]
        result["error_location"] = _active_error_location(excel)
    elif ran:
        result["ok"] = True
    else:
        # Run itself failed (macros disabled? trust off?) — inconclusive
        result["error_text"] = err
    return result


def _compile_project(excel, wb, timeout: int = 60) -> dict:
    """
    REAL compile verification of the whole project (see _probe_compile).
    """
    result = {
        "attempted": True, "ok": None,
        "trigger": None, "error_text": None,
        "error_location": None, "timed_out": False, "code_pane": None,
    }

    try:
        vbe = _get_vbe(excel)
    except Exception as e:
        result["error_text"] = f"无法访问 VBE: {_com_error_text(e) or e}"
        return result

    # Macros must be enabled (open-time setting) for the probe Run to work
    try:
        excel.AutomationSecurity = MSO_SECURITY_LOW
    except Exception:
        pass

    result["code_pane"] = _ensure_code_pane(excel, wb)

    probe = _probe_compile(excel, wb, timeout=min(timeout, 30))
    result["probe"] = probe
    result["trigger"] = "in-process-compile" if probe.get("ok") is not None else "probe-not-run"
    result["timed_out"] = probe.get("timed_out", False)
    if probe.get("ok") is False:
        result["ok"] = False
        result["error_text"] = probe.get("error_text")
        result["error_location"] = probe.get("error_location")
        if probe.get("dialogs"):
            result["dialogs"] = probe["dialogs"]
    elif probe.get("ok") is True:
        result["ok"] = True
    else:
        result["ok"] = None
        result["error_text"] = probe.get("error_text") or result["error_text"]
    return result


def _dirty_project(wb) -> bool:
    """
    Mark the VBA project 'dirty' by round-tripping one code line, so the VBE
    'Debug > Compile' menu item becomes enabled. Without this, a project
    loaded in an already-compiled state has the item disabled and Execute()
    silently no-ops — reporting a false pass. In-memory only (build_check
    never saves).
    """
    try:
        proj = wb.VBProject
        for comp in proj.VBComponents:
            if comp.Type != VBEXT_CT_STDMODULE:
                continue
            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                first = cm.Lines(1, 1)
                cm.DeleteLines(1, 1)
                cm.InsertLines(1, first)
                return True
    except Exception:
        pass
    return False


def _get_vbe(excel, retries: int = 3, delay: float = 1.0):
    """
    Access excel.VBE with retries. On a freshly spawned instance (or one
    still settling after another Excel quit), the first .VBE access can
    transiently fail with 0x800A03EC — it succeeds on retry.
    """
    last = None
    for i in range(retries):
        try:
            return excel.VBE
        except Exception as e:
            last = e
            time.sleep(delay)
    raise last


def _current_selection(excel):
    """Normalized (component, line) of the VBE active code pane selection,
    or None. After a (silent) compile error, the VBE jumps the selection to
    the offending line — comparing before/after detects it."""
    try:
        pane = excel.VBE.ActiveCodePane
        if pane is not None:
            sel = pane.GetSelection()  # (line1, col1, line2, col2)
            return (pane.CodeModule.Parent.Name, int(sel[0]))
    except Exception:
        pass
    return None


def _hide_vbe_window(vbe) -> None:
    """Best-effort hide of the VBE main window."""
    try:
        vbe.MainWindow.Visible = False
    except Exception:
        pass


def build_check(xlam_path: str, compile: bool = True, strict: bool = False,
                timeout: int = 60, attach: bool = False,
                check_binding: bool = True,
                assume_trusted: bool = False) -> dict:
    """
    检查工作簿的 VBA 工程，形成验证闭环：

    1. 静态检查（纯文本分析，无风险）：
       块结构配对（If/For/Do/While/Select/With、#If）、过程外代码、
       未关闭的 Type/Enum、截断的过程、前期绑定外部类型库。
       strict=True 时附加：缺少 Option Explicit、公共过程重名。
    2. 真实编译（compile=True）：触发 VBE 'Debug > Compile'，
       捕获编译错误文本与出错位置。

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        compile: Run the real VBE compile (default True)
        strict: Extra strict warnings (default False)
        timeout: Seconds before a stuck compile is force-killed
        attach: Use the user's running Excel instance (default: isolated)
        check_binding: Report early binding to external type libraries as an
            error (default True). These references are not standard in a fresh
            VBA project, so a distributed .xlam fails to compile on machines
            that have not ticked them. Opt a module out with a
            ' @allow-early-binding comment (needed for WithEvents sinks).
        assume_trusted: Skip the Mark-of-the-Web gate for compile=True (file
            flagged as downloaded from the internet). Review the source with
            export_vba_source before passing this.

    Returns:
        {
          "file": ...,
          "ok": bool,                      # static errors == 0 AND compile ok
          "static": {"errors": [...], "warnings": [...], "modules_checked": n},
          "compile": {...} or None,
        }
        绑定类错误额外带 "kind": "early_binding" 字段，便于区分。
    """
    path = _require_macro_file(xlam_path)
    if compile:
        # 宏启用的打开只发生在真实编译路径；compile=False 走 ForceDisable，
        # 静态层读 VBProject 不需要宏——这就是"不可信文件只允许静态检查"
        # 红线的技术落点
        _refuse_untrusted(path, assume_trusted)
    outcome = {"file": path, "ok": False,
               "static": {"errors": [], "warnings": [], "modules_checked": 0},
               "compile": None}

    # security='low': AutomationSecurity applies at OPEN time — a workbook
    # opened with macros disabled makes the VBE compile command and the
    # probe run silently ineffective (false pass). ONLY the real compile
    # needs macros enabled; the static layer reads the VBProject fine under
    # ForceDisable, so compile=False opens untrusted files safely.
    with _open_workbook(path, mode="edit", attach=attach,
                        security=("low" if compile else None)) as (excel, wb, _):
        outcome["static"] = _static_check_project(
            wb.VBProject, strict=strict, check_binding=check_binding)
        if compile:
            outcome["compile"] = _compile_project(excel, wb, timeout=timeout)

    static_ok = not outcome["static"]["errors"]
    comp = outcome["compile"]
    # A real compile error (False) blocks; an inconclusive trigger (None)
    # does not — the static layer already gated, and a false failure would
    # send the AI chasing phantom problems.
    compile_ok = comp is None or comp.get("ok") is not False
    outcome["ok"] = static_ok and compile_ok

    # Human-readable summary
    print(f"=== build_check: {os.path.basename(path)} ===")
    s = outcome["static"]
    print(f"静态检查: {s['modules_checked']} 个模块, "
          f"{len(s['errors'])} 错误, {len(s['warnings'])} 警告")
    for e in s["errors"]:
        tag = "[绑定]" if e.get("kind") == "early_binding" else "[错误]"
        print(f"  {tag} {e['module']}:{e['line']} {e['error']}")
    for w in s["warnings"]:
        print(f"  [警告] {w}")
    if comp is not None:
        if comp.get("ok") is True:
            print("编译: [OK] 通过")
        elif comp.get("timed_out"):
            print(f"编译: [TIMEOUT] — {comp.get('error_text')}")
        elif comp.get("error_text"):
            loc = comp.get("error_location")
            loc_s = f" @ {loc['component']}:{loc['line']}" if loc else ""
            print(f"编译: [FAIL] {comp['error_text']}{loc_s}")
        else:
            print(f"编译: [WARN] 无法确定（{comp.get('trigger')}）")
    print(f"总体: {'[OK] 通过' if outcome['ok'] else '[FAIL] 未通过'}")
    return outcome


# ============================================================================
# Section 14: Run Macros / Throwaway Tests
# ============================================================================

def run_macro(xlam_path: str, macro_name: str, args: list = None,
              timeout: int = 120, save: bool = False, attach: bool = False,
              assume_trusted: bool = False) -> dict:
    """
    Run a macro in the workbook and capture its result / errors.

    Modal dialogs the macro raises (MsgBox etc.) are read and auto-dismissed
    by a watchdog, so a stray MsgBox cannot hang the call; if the macro is
    still stuck after `timeout` seconds, the isolated Excel is force-killed.

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        macro_name: Macro/procedure name (Public Sub/Function)
        args: Optional positional arguments
        timeout: Kill stuck Excel after this many seconds
        save: Save the workbook after the macro ran (default False)
        attach: Use the user's running Excel instance (default: isolated)
        assume_trusted: Skip the Mark-of-the-Web gate (file downloaded from
            the internet) — review the source first, see SKILL.md red-line

    Returns:
        {"ok": bool, "result": Any, "error": str|None,
         "dialogs": [...], "timed_out": bool}
    """
    _refuse_untrusted(_require_macro_file(xlam_path), assume_trusted)
    out = {"ok": False, "result": None, "error": None, "dialogs": [], "timed_out": False}
    mode = "run"  # macros must be ENABLED for Run — never 'write' (would disable them)

    with _open_workbook(xlam_path, mode=mode, attach=attach, save=save,
                        backup_reason="run+save") as (excel, wb, path):
        attached = _is_attached_instance(excel)
        if attached:
            print("⚠ attach 模式：运行卡死时不会强杀用户 Excel（看门狗击杀已禁用，"
                  "挂死需人工结束）；弹窗仍会被自动关闭（含用户自己触发的对话框）")
        wd = _run_watchdog(excel, timeout=timeout, attached=attached)
        wd.start()
        try:
            out["result"] = excel.Run(f"'{wb.Name}'!{macro_name}", *(args or []))
            out["ok"] = True
        except Exception as e:
            out["error"] = _com_error_text(e) or str(e)
        finally:
            wd.stop(grace=1.0)

        out["dialogs"] = wd.dialog_texts
        if not out["ok"] and out["dialogs"]:
            # VBA runtime errors carry no COM description — the dismissed
            # dialog text IS the real error message
            out["error"] = out["dialogs"][-1]
        if wd.killed:
            _invalidate_cached_excel()
            out["timed_out"] = not wd.storm
            out["ok"] = False
            if wd.storm:
                out["error"] = ((out["dialogs"][-1] if out["dialogs"] else "弹窗风暴")
                                + " — 弹窗反复出现触发风暴保护，已结束本次隔离 Excel 进程")
            else:
                out["error"] = (out["error"] or
                                f"宏执行超时（>{timeout}s），已强制结束本次隔离 Excel 进程")

    status = "[OK]" if out["ok"] else "[FAIL]"
    extra = f" result={out['result']!r}" if out["ok"] and out["result"] is not None else ""
    print(f"{status} run {macro_name}{extra}")
    if out["error"]:
        print(f"  error: {out['error']}")
    if out["dialogs"]:
        print(f"  自动关闭的弹窗: {out['dialogs']}")
    return out


def run_test(xlam_path: str, code: str, proc_name: str = "RunTest", args: list = None,
             timeout: int = 120, keep: bool = False, attach: bool = False,
             assume_trusted: bool = False) -> dict:
    """
    Inject a THROWAWAY test module, run it, then remove it again.
    The workbook file itself is never modified (unless keep=True).

    `code` must define a Public Sub/Function named `proc_name` ('RunTest' by
    default). If it is a Function, its return value is captured — the easiest
    way to get output back:

        Public Function RunTest() As String
            RunTest = "sum=" & (2 + 2)
        End Function

    Compile errors, runtime errors and unexpected MsgBoxes are all captured
    (dialogs auto-dismissed; stuck Excel force-killed after `timeout`).

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        code: VBA code of the test module
        proc_name: Procedure inside `code` to run
        args: Optional positional arguments
        timeout: Kill stuck Excel after this many seconds
        keep: Keep the test module and save (debugging aid; default False)
        attach: Use the user's running Excel instance (default: isolated)
        assume_trusted: Skip the Mark-of-the-Web gate (file downloaded from
            the internet) — review the source first, see SKILL.md red-line

    Returns:
        {"ok": bool, "result": Any, "error": str|None,
         "dialogs": [...], "timed_out": bool, "module": name}
    """
    out = {"ok": False, "result": None, "error": None, "dialogs": [],
           "timed_out": False, "module": None}
    module_name = f"zzTmpTest{int(time.time())}"
    out["module"] = module_name

    _refuse_untrusted(_require_macro_file(xlam_path), assume_trusted)

    # 恒用 'run' 模式：keep=True 若走 write，_apply_app_settings 会在 OPEN 前
    # 强禁宏，注入的测试宏 Run 必报"宏被禁用"（open 后重设无效——
    # AutomationSecurity 只在打开时生效）。保存需求由 save=keep 表达，
    # 与 run_macro(save=True) 同一套路。
    with _open_workbook(xlam_path, mode="run", attach=attach, save=keep,
                        backup_reason="run_test+keep") as (excel, wb, path):
        proj = wb.VBProject
        # Inject throwaway module (macros are enabled in 'run' mode)
        comp = proj.VBComponents.Add(VBEXT_CT_STDMODULE)
        comp.Name = module_name
        comp.CodeModule.AddFromString(code)

        attached = _is_attached_instance(excel)
        if attached:
            print("⚠ attach 模式：运行卡死时不会强杀用户 Excel（看门狗击杀已禁用，"
                  "挂死需人工结束）；弹窗仍会被自动关闭（含用户自己触发的对话框）")
        wd = _run_watchdog(excel, timeout=timeout, attached=attached)
        wd.start()
        try:
            out["result"] = excel.Run(f"'{wb.Name}'!{proc_name}", *(args or []))
            out["ok"] = True
        except Exception as e:
            out["error"] = _com_error_text(e) or str(e)
        finally:
            wd.stop(grace=1.0)

        out["dialogs"] = wd.dialog_texts
        if not out["ok"] and out["dialogs"]:
            # VBA runtime errors carry no COM description — the dismissed
            # dialog text IS the real error message
            out["error"] = out["dialogs"][-1]
        if wd.killed:
            _invalidate_cached_excel()
            out["timed_out"] = not wd.storm
            out["ok"] = False
            if wd.storm:
                out["error"] = ((out["dialogs"][-1] if out["dialogs"] else "弹窗风暴")
                                + " — 弹窗反复出现触发风暴保护，已结束本次隔离 Excel 进程")
            else:
                out["error"] = (out["error"] or
                                f"测试执行超时（>{timeout}s），已强制结束本次隔离 Excel 进程")

        if not keep:
            try:
                proj.VBComponents.Remove(comp)
            except Exception:
                pass

    status = "[OK]" if out["ok"] else "[FAIL]"
    extra = f" result={out['result']!r}" if out["ok"] and out["result"] is not None else ""
    print(f"{status} run_test ({module_name}.{proc_name}){extra}")
    if out["error"]:
        print(f"  error: {out['error']}")
    if out["dialogs"]:
        print(f"  自动关闭的弹窗: {out['dialogs']}")
    return out


# ============================================================================
# Section 14.5: Installed Add-in Load/Unload (upgrade path for locked files)
# ============================================================================

def is_file_locked(file_path: str) -> bool:
    """
    True when another process holds a write lock on the file — typically a
    .xlam loaded in the user's running Excel. Pure filesystem probe (no COM,
    no Excel started).

    Args:
        file_path: Path to check

    Returns:
        bool
    """
    path = os.path.abspath(file_path)
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def _match_addin(excel, addin: str):
    """Find an entry in this instance's AddIns collection by name
    ('MyTools'), file name ('MyTools.xlam') or full path. None if absent."""
    base = os.path.basename(str(addin))
    want_file = base.lower()
    want_name = os.path.splitext(base)[0].lower()
    for ai in excel.AddIns:
        try:
            nm = (ai.Name or "").lower()
            if nm == want_file or os.path.splitext(nm)[0].lower() == want_name:
                return ai
            if (ai.Title or "").lower() == want_name:
                return ai
            try:
                full = os.path.basename(ai.FullName or "").lower()
                if full == want_file:
                    return ai
            except Exception:
                pass
        except Exception:
            continue
    return None


def unload_addin(addin: str, attach: bool = True) -> dict:
    """
    Uninstall (uncheck) an add-in in the user's running Excel, releasing the
    file lock so the toolkit can modify / repack the .xlam. First step of the
    installed-add-in upgrade loop: unload_addin -> edit/pack -> reload_addin.

    The AddIns collection is PER INSTANCE — the lock is held by the instance
    that has the add-in loaded, normally the USER's Excel. Hence attach=True
    by default (falls back to an isolated instance when no user Excel runs,
    in which case there is no lock to release anyway). Attached instances are
    never Quit.

    Args:
        addin: Add-in name ('MyTools'), file name or full path
        attach: operate on the user's running Excel (default True)

    Returns:
        {"found", "unloaded", "was_installed", "note"}
    """
    out = {"found": False, "unloaded": False, "was_installed": False, "note": ""}
    excel, attached = _acquire_excel(attach)
    if not attached:
        out["note"] = "未连接到用户 Excel（未在运行？）——文件本就无锁；若仍锁定请让用户关闭后重试"
        print(f"[addin] {out['note']}")
        # 提前返回：隔离实例的 AddIns 集合同样读取（用户）注册表，在它上面
        # 执行 Installed=False 会随实例退出被持久化——用户下次启动 Excel
        # 加载项直接消失。无用户实例时根本没有锁要释放。
        return out
    try:
        ai = _match_addin(excel, addin)
        if ai is None:
            out["note"] = f"加载项列表中未找到 '{addin}'（该实例未加载它？）"
            print(f"[addin] {out['note']}")
            return out
        out["found"] = True
        out["was_installed"] = bool(ai.Installed)
        if ai.Installed:
            ai.Installed = False  # unloads in that instance, lock released
            out["unloaded"] = True
            print(f"[addin] 已卸载 '{ai.Name}'（文件锁已释放；注册保留，"
                  f"修改后用 reload_addin 恢复）")
        else:
            out["note"] = f"'{ai.Name}' 在列表中但本就未加载"
            print(f"[addin] {out['note']}")
    except Exception as e:
        out["note"] = _com_error_text(e) or str(e)
        print(f"[addin] unload 失败: {out['note']}")
    return out


def reload_addin(addin: str, path: str = None, attach: bool = True) -> dict:
    """
    (Re)load an add-in — the last step of the installed-add-in upgrade loop;
    effective immediately in the user's running Excel (no restart needed).

    Args:
        addin: Add-in name ('MyTools'), file name or full path
        path: full path to the .xlam — only needed when the add-in is not yet
              in the AddIns list (first install): registers it via
              AddIns.Add(path) first
        attach: load into the user's running Excel (default True; an isolated
                instance only registers for the NEXT user Excel start)

    Returns:
        {"found", "loaded", "workbook_loaded", "note"}
        workbook_loaded: True/False/None(None=无法复核)——Installed=True 只代表标志置位，
        同会话「卸载→重载」后 Office 16 可能并不真正装载工程（需重启 Excel）。
    """
    out = {"found": False, "loaded": False, "note": ""}
    excel, attached = _acquire_excel(attach)
    try:
        ai = _match_addin(excel, addin)
        if ai is None and path:
            # AddIns.Add fails with '类 AddIns 的 Add 方法无效' while the
            # instance has NO open workbook (verified on Office 16) — hold a
            # throwaway workbook open around the call. On an attached
            # (user's) instance ScreenUpdating is briefly disabled so the
            # flash stays invisible; the holder is closed right after.
            holder = None
            su = None
            try:
                if excel.Workbooks.Count == 0:
                    if attached:
                        try:
                            su = excel.ScreenUpdating
                            excel.ScreenUpdating = False
                        except Exception:
                            pass
                    holder = excel.Workbooks.Add()
                ai = excel.AddIns.Add(os.path.abspath(path))
                print(f"[addin] 已注册加载项: {ai.FullName}")
            finally:
                if holder is not None:
                    try:
                        holder.Close(SaveChanges=False)
                    except Exception:
                        pass
                if su is not None:
                    try:
                        excel.ScreenUpdating = su
                    except Exception:
                        pass
        if ai is None:
            out["note"] = (f"加载项列表中无 '{addin}' 且未提供 path —— "
                           f"reload_addin('{addin}', path='...\\xxx.xlam') 完成注册")
            print(f"[addin] {out['note']}")
            return out
        out["found"] = True
        ai.Installed = True
        out["loaded"] = True
        # Installed=True 只代表标志置位：同会话「卸载→重载」后 Office 16 可能并不真正
        # 装载工程（实测系统提示"重启 Excel 前不会出现"）——用 Workbooks 复核真实状态。
        import time
        try:
            _target = (ai.Name or "").lower()
            wb_loaded = False
            for _ in range(3):
                wb_loaded = any((w.Name or "").lower() == _target for w in excel.Workbooks)
                if wb_loaded:
                    break
                time.sleep(0.5)
        except Exception:
            wb_loaded = None
        out["workbook_loaded"] = wb_loaded
        where = ("用户 Excel（立即生效）" if attached
                 else "当前用户配置（用户下次启动 Excel 时生效）")
        if wb_loaded is False:
            if attached:
                where += ("；⚠️ 工程未实际装载（本 Office 构建 COM 安装不即时装载，实测与是否"
                          "首次安装/是否卸载过无关）——请让用户重启 Excel 或在加载项对话框手动确认")
            else:
                where += "（隔离实例仅完成注册，属预期）"
        print(f"[addin] 已加载 '{ai.Name}' -> {where}")
    except Exception as e:
        out["note"] = _com_error_text(e) or str(e)
        print(f"[addin] reload 失败: {out['note']}")
    return out


# ============================================================================
# Section 15: UserForm and Control Functions
# ============================================================================

def create_userform(xlam_path: str, form_name: str, attach: bool = False) -> None:
    """
    Create a new UserForm (recreates an existing one).

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        form_name: Name for the new UserForm
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        proj = wb.VBProject
        try:
            existing = proj.VBComponents(form_name)
            proj.VBComponents.Remove(existing)
        except Exception:
            pass

        new_form = proj.VBComponents.Add(VBEXT_CT_MSFORM)
        new_form.Name = form_name
        new_form.Properties("Caption").Value = form_name
        new_form.Properties("Width").Value = 300
        new_form.Properties("Height").Value = 200
        print(f"Created UserForm: {form_name}")


def _resolve_container_controls(container_ctrl, page=None):
    """Resolve the controls collection of a container control.

    Frame 等常规容器走 .Controls；MultiPage 在 IDispatch 层不暴露 Controls
    （实测 IMultiPage 上 GetIDsOfNames("Controls") 返回未知名称），必须经
    Pages(page).Controls 定位到具体页。page 为页索引（0 起）或页名/页标题。
    """
    try:
        pages = container_ctrl.Pages
    except AttributeError:
        if page is not None:
            raise ValueError("page 参数仅对 MultiPage 容器有效")
        return container_ctrl.Controls
    if page is None:
        page = 0
    if isinstance(page, int):
        if page < 0 or page >= pages.Count:
            raise ValueError(f"MultiPage 页索引越界: {page}（共 {pages.Count} 页）")
        return pages.Item(page).Controls
    for i in range(pages.Count):
        pg = pages.Item(i)
        if pg.Name == page or pg.Caption == page:
            return pg.Controls
    raise ValueError(f"MultiPage 页 '{page}' 不存在")


def add_control(
    xlam_path: str,
    form_name: str,
    control_type: str,
    name: str,
    properties: dict = None,
    event_handler: str = None,
    container: str = None,
    page=None,
    attach: bool = False
) -> None:
    """
    Add a control to a UserForm.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        form_name: Name of the UserForm
        control_type: Control ProgID (e.g., 'Forms.CommandButton.1')
        name: Name for the control
        properties: Dict of properties to set
        event_handler: Deprecated/no-op for MSForms (kept for backward compatibility).
                       MSForms designer controls have no OnClick/OnChange property (that
                       is an Access/VB6 convention) — events bind automatically by naming
                       convention; write handlers via add_form_event_handler().
        container: Optional container name (frame, or a MultiPage control) to add
                   the control inside
        page: For MultiPage containers only — page index (0-based) or page name/caption.
              Ignored for other containers. Default: page 0.
        attach: Use the user's running Excel instance (default: isolated)
    """
    properties = properties or {}

    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        form = wb.VBProject.VBComponents(form_name)
        designer = form.Designer

        if container:
            container_ctrl = designer.Controls(container)
            controls = _resolve_container_controls(container_ctrl, page)
        else:
            controls = designer.Controls

        try:
            existing = controls(name)
            print(f"Control '{name}' already exists, removing")
            controls.Remove(existing)
        except Exception:
            pass

        ctrl = controls.Add(control_type)
        ctrl.Name = name

        # MSForms controls expose their properties DIRECTLY on the control
        # object (ctrl.Caption = ...) — there is no usable Properties
        # collection on them (that exists on VBComponent, not its controls).
        for prop_name, prop_value in properties.items():
            try:
                if prop_name == "List" and isinstance(prop_value, list):
                    for item in prop_value:
                        ctrl.AddItem(item)
                else:
                    setattr(ctrl, prop_name, prop_value)
            except Exception:
                # Fallback: some hosts do expose a Properties collection
                try:
                    ctrl.Properties(prop_name).Value = prop_value
                except Exception as e:
                    print(f"Warning: Could not set property '{prop_name}': {e}")

        # MSForms 设计器控件没有 OnClick/OnChange 属性（那是 Access/VB6 约定）——事件由
        # VBE 按"控件名_事件名"自动绑定，处理器用 add_form_event_handler() 添加即可。
        # event_handler 参数仅为兼容旧调用保留，不再尝试设置（此前必然失败并打 warning）。

        print(f"Added control '{name}' ({control_type}) to {form_name}")


def add_form_event_handler(
    xlam_path: str,
    form_name: str,
    event_name: str,
    code: str,
    attach: bool = False
) -> None:
    """
    Add event handler code to a UserForm.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        form_name: Name of the UserForm
        event_name: Event name (e.g., 'UserForm_Initialize', 'btnOK_Click')
        code: VBA event handler code
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        code_mod = wb.VBProject.VBComponents(form_name).CodeModule

        lines = code_mod.Lines(1, code_mod.CountOfLines) if code_mod.CountOfLines else ""
        # 词边界正则判重（与 add_vba_callback 一致）：裸子串 "Sub btnOK_Click"
        # 会误命中 btnOK_Click2，导致事件处理器被静默跳过、事件永不触发
        if _proc_decl_re(event_name).search(lines):
            print(f"Event handler '{event_name}' already exists")
            return

        n = code_mod.CountOfLines
        if n > 0:
            code_mod.InsertLines(n + 1, vbCrLf)
            n += 1
        code_mod.InsertLines(n + 1, code)
        print(f"Added event handler: {event_name} to {form_name}")


def set_form_properties(xlam_path: str, form_name: str, properties: dict,
                        attach: bool = False) -> None:
    """
    Set UserForm properties.

    Args:
        xlam_path: Path to Excel file (.xlam/.xlsm/.xlsx). .xlsx will be auto-converted to .xlsm.
        form_name: Name of the UserForm
        properties: Dict of properties
        attach: Use the user's running Excel instance (default: isolated)
    """
    with _open_workbook(xlam_path, mode="write", attach=attach) as (excel, wb, path):
        form = wb.VBProject.VBComponents(form_name)
        # 实例化 Designer 后组件 Properties 才可写：否则重开的工程会报"发生意外"（实测）
        try:
            form.Designer
        except Exception:
            pass
        for prop_name, prop_value in properties.items():
            try:
                form.Properties(prop_name).Value = prop_value
            except Exception as e:
                print(f"Warning: Could not set property '{prop_name}': {e}")
        print(f"Set properties for {form_name}")


def lint_form(xlam_path: str, form_name: str, attach: bool = False) -> dict:
    """
    Check UserForm control geometry. Flags:
      - 负坐标 / 零尺寸控件
      - 控件超出窗体（或容器）边界
      - 同级控件互相重叠（重叠面积 > 较小控件的 20%）

    遍历窗体直系控件及容器（Frame / MultiPage 页）内的子控件；子控件
    坐标按容器边界检查。不可见控件跳过。

    Args:
        xlam_path: Path to Excel file (.xlsm/.xlam)
        form_name: Name of the UserForm
        attach: Use the user's running Excel instance (default: isolated)

    Returns:
        {"form": form_name, "findings": [{"level","control","issue"}], "ok": bool}
    """
    findings = []

    def _check_level(controls, bounds, level_label, container_name=None):
        entries = []
        try:
            count = controls.Count
        except Exception:
            return
        for idx in range(count):
            try:
                c = controls.Item(idx)
                name = c.Name
                left, top = float(c.Left), float(c.Top)
                width, height = float(c.Width), float(c.Height)
                try:
                    visible = bool(c.Visible)
                except Exception:
                    visible = True
                try:
                    parent_name = c.Parent.Name
                except Exception:
                    parent_name = None
                entries.append((name, left, top, width, height, visible, c, parent_name))
            except Exception:
                continue

        # designer.Controls 是扁平集合（含 Frame/MultiPage 页内子控件，坐标为容器内相对
        # 坐标）：跨容器两两比对会把页内子控件与页本身/其他页子控件误判"重叠"（实测全为
        # 误报）。因此只保留"父容器 == 当前容器"的条目做检查，嵌套子控件交给下方递归按
        # 其真实容器检查；Parent 语义异常（过滤后为空）时回退为不过滤，保持旧行为不漏检。
        scoped = entries
        if container_name is not None:
            filtered = [e for e in entries if e[7] == container_name]
            if filtered:
                scoped = filtered

        # Per-control checks
        for name, left, top, width, height, visible, _, _ in scoped:
            if not visible:
                continue
            if left < 0 or top < 0:
                findings.append({"level": "error", "control": name,
                                 "issue": f"负坐标 Left={left:.0f} Top={top:.0f} ({level_label})"})
            if width <= 0 or height <= 0:
                findings.append({"level": "error", "control": name,
                                 "issue": f"零尺寸 W={width:.0f} H={height:.0f} ({level_label})"})
            if bounds and (left + width > bounds[0] + 4 or top + height > bounds[1] + 4):
                findings.append({"level": "warning", "control": name,
                                 "issue": (f"超出边界: 右缘 {left + width:.0f}/下缘 {top + height:.0f} "
                                           f"vs 容器 {bounds[0]:.0f}x{bounds[1]:.0f} ({level_label})")})

        # Pairwise overlap among visible siblings (same real parent only)
        vis = [e for e in scoped if e[5]]
        for i in range(len(vis)):
            for j in range(i + 1, len(vis)):
                n1, l1, t1, w1, h1, _, _, _ = vis[i]
                n2, l2, t2, w2, h2, _, _, _ = vis[j]
                ox = max(0, min(l1 + w1, l2 + w2) - max(l1, l2))
                oy = max(0, min(t1 + h1, t2 + h2) - max(t1, t2))
                overlap = ox * oy
                smaller = min(w1 * h1, w2 * h2)
                if smaller > 0 and overlap > 0.2 * smaller:
                    findings.append({"level": "warning", "control": f"{n1} × {n2}",
                                     "issue": f"控件重叠 {overlap:.0f}pt² ({level_label})"})

        # Recurse into containers (Frame / MultiPage pages)
        for name, l, t, w, h, visible, c, _p in scoped:
            try:
                inner = c.Controls
                if inner.Count > 0:
                    _check_level(inner, (w, h), f"{level_label}/{name}", container_name=name)
            except Exception:
                pass
            try:
                pages = c.Pages
                for p in pages:
                    try:
                        _check_level(p.Controls, (w, h), f"{level_label}/{name}/{p.Name}",
                                     container_name=p.Name)
                    except Exception:
                        pass
            except Exception:
                pass

    with _open_workbook(xlam_path, mode="edit", attach=attach) as (excel, wb, path):
        form = wb.VBProject.VBComponents(form_name)
        if form.Type != VBEXT_CT_MSFORM:
            raise ValueError(f"'{form_name}' 不是 UserForm (type={form.Type})")
        designer = form.Designer
        try:
            inside_w = float(designer.InsideWidth)
            inside_h = float(designer.InsideHeight)
        except Exception:
            inside_w = float(designer.Width) - 6
            inside_h = float(designer.Height) - 30
        _check_level(designer.Controls, (inside_w, inside_h), "form",
                     container_name=form.Name)

    result = {"form": form_name, "findings": findings,
              "ok": not any(f["level"] == "error" for f in findings)}
    print(f"=== lint_form: {form_name} — {len(findings)} 项发现 ===")
    for f in findings:
        print(f"  [{f['level']}] {f['control']}: {f['issue']}")
    if not findings:
        print("  [OK] 无几何问题")
    return result


# ============================================================================
# Section 16: VBA Code Generation Helpers
# ============================================================================

def generate_vba_callback(
    name: str,
    control_param: bool = True,
    doc: str = ""
) -> str:
    """
    Generate VBA callback procedure template.

    Args:
        name: Callback procedure name
        control_param: Include IRibbonControl parameter
        doc: Documentation comment

    Returns:
        VBA callback code string
    """
    lines = []

    if doc:
        lines.append(f"' {doc}")

    if control_param:
        lines.append(f"Sub {name}(control As IRibbonControl)")
    else:
        lines.append(f"Sub {name}()")

    lines.append(f"    ' TODO: Implement")
    lines.append("End Sub")

    return "\n".join(lines)


def generate_form_init_handler(form_name: str, form_properties: dict = None) -> str:
    """
    Generate UserForm_Initialize handler.

    The event procedure name is FIXED as 'UserForm_Initialize' — the
    UserForm's Initialize event belongs to the form class itself and never
    carries the form's name ('{form_name}_Initialize' would compile but
    never fire). `form_name` is kept for call compatibility / documentation.

    Args:
        form_name: UserForm name (not used in the generated code — the
            handler is self-bound via the fixed event name; controls are
            addressed through `Me`)
        form_properties: Dict of control initializations

    Returns:
        VBA event handler code
    """
    lines = ["Private Sub UserForm_Initialize()"]

    if form_properties:
        for ctrl_name, props in form_properties.items():
            if "AddItem" in props:
                for item in props["AddItem"]:
                    lines.append(f'    Me.{ctrl_name}.AddItem "{item}"')

    lines.append("End Sub")
    return "\n".join(lines)


def generate_full_feature_module(
    feature_name: str,
    buttons: list,
    doc: str = ""
) -> str:
    """
    Generate a complete VBA module with multiple button callbacks.

    Args:
        feature_name: Feature name (used for module name)
        buttons: List of dicts with keys: id, label, callback, imageMso, description
        doc: Module documentation

    Returns:
        Complete VBA module code
    """
    lines = ["Option Explicit"]

    if doc:
        lines.append(f"' {doc}")

    for btn in buttons:
        if doc_text := btn.get("description"):
            lines.append(f"")
            lines.append(f"' {doc_text}")

        callback = btn.get("callback", f"{btn['id']}_Click")
        lines.append(f"Sub {callback}(control As IRibbonControl)")
        lines.append(f"    ' {btn.get('label', btn['id'])}")
        lines.append("    ' TODO: Implement")
        lines.append("End Sub")

    return "\n".join(lines)


# ============================================================================
# Main (for testing)
# ============================================================================

if __name__ == "__main__":
    print("xlam_toolkit.py - Excel Add-in Development Toolkit")
    print("")
    print("Supported formats: .xlam, .xlsm, .xlsx")
    print("  - .xlsx files will be auto-converted to .xlsm when adding VBA")
    print("  - Ribbon customization only for .xlam files")
    print("  - All COM automation runs in an isolated Excel instance")
    print("    (user's running Excel is never touched; attach=True to reuse it)")
    print("  - Every write operation backs up to <file>.bak/ first")
    print("")
    print("Environment:")
    print("  - check_environment()                    检查 pywin32/Excel/VBA 信任")
    print("  - enable_vba_trust()                     一键开启 VBA 工程访问信任")
    print("")
    print("Core workflow (recommended):")
    print("  - export_vba_source(file, dir)           导出全部 VBA 为 .bas/.cls/.frm")
    print("  - (edit the exported source files directly)")
    print("  - import_vba_source(file, dir, sync)     导回工作簿（写前自动备份）")
    print("  - build_check(file)                      静态检查 + 真实 VBE 编译")
    print("  - run_test(file, code)                   注入临时测试宏并运行（文件不变）")
    print("  - run_macro(file, name)                  运行宏并捕获错误")
    print("")
    print("Backup & safety:")
    print("  - backup_file / list_backups / restore_backup")
    print("  - shutdown_excel()                       关闭本工具缓存的隔离实例")
    print("")
    print("VBA edit functions (all formats):")
    print("  - add_vba_module / add_vba_class / add_vba_callback")
    print("  - replace_vba_module / update_vba_callback / replace_vba_lines")
    print("  - insert_vba_code / delete_vba_callback / delete_vba_module")
    print("  - read_vba_module / list_procedures / list_vba_modules / get_procedure_info")
    print("")
    print("UserForm functions:")
    print("  - create_userform / add_control / add_form_event_handler / set_form_properties")
    print("  - lint_form(file, form)                  控件几何检查")
    print("")
    print("Add-in lifecycle (installed add-in upgrade):")
    print("  - is_file_locked(path)                  检测文件是否被用户 Excel 锁定")
    print("  - unload_addin(name) / reload_addin     卸载释放锁 / 改完重载（默认 attach 用户实例）")
    print("")
    print("Ribbon functions (.xlam / .xlsm):")
    print("  - unpack_xlam / pack_xlam / get_ribbon_xml / set_ribbon_xml")
    print("  - add_button_to_ribbon / add_group_to_ribbon / register_icon")
