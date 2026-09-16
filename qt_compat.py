# -*- coding: utf-8 -*-
"""
Qt binding compatibility layer (ASCII output only).

Why this exists
---------------
PySide6 is built on Qt 6, and Qt 6 requires **Windows 10 1809 or later**.
Qt officially dropped Windows 7 / 8.x support in Qt 6.

On older systems such as Windows Server 2008 / 2008 R2 / Windows 7, the
Qt 6 DLLs fail to load:

    ImportError: DLL load failed while importing QtCore:
                 The specified procedure could not be found.

That message means the DLL file is present, but it imports Win10-only
entry points that do not exist on the old OS.

The fix is to use a Qt 5 based binding instead:

    PySide6 (Qt 6)  ->  Windows 10 / 11
    PySide2 (Qt 5)  ->  Windows 7 SP1 / Server 2008 R2   <-- old servers
    PyQt5   (Qt 5)  ->  fallback

All three share the short enum names used by this project
(``Qt.AlignCenter``, ``QMessageBox.Yes``, ``QHeaderView.ResizeToContents``),
so no call-site changes are needed.

Usage
-----
Replace::

    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import QApplication, ...

with::

    from qt_compat import Qt, QFont, QApplication, ...

Run ``python qt_compat.py`` to see which binding your system picked.
"""

from __future__ import annotations

import os
import platform
import sys

# ---------------------------------------------------------------- detection

#: Binding search order. Set env var QT_BINDING=PySide2 to force one.
_ORDER = ["PySide6", "PySide2", "PySide5", "PyQt5"]

BINDING: str = ""
QT_VERSION: str = ""

#: Why each binding failed, for the diagnostic report
_ERRORS: dict = {}


def _try_import(name: str):
    """Return (QtCore, QtGui, QtWidgets) for the given binding, or None.

    The reason is recorded in _ERRORS[name] so it can be shown later.
    """
    try:
        if name in ("PySide6", "PySide2", "PySide5"):
            qc = __import__(f"{name}.QtCore", fromlist=["QtCore"])
            qg = __import__(f"{name}.QtGui", fromlist=["QtGui"])
            qw = __import__(f"{name}.QtWidgets", fromlist=["QtWidgets"])
        elif name == "PyQt5":
            qc = __import__("PyQt5.QtCore", fromlist=["QtCore"])
            qg = __import__("PyQt5.QtGui", fromlist=["QtGui"])
            qw = __import__("PyQt5.QtWidgets", fromlist=["QtWidgets"])
        else:
            _ERRORS[name] = "unknown binding name"
            return None
        return qc, qg, qw
    except ModuleNotFoundError as e:
        _ERRORS[name] = "not installed (%s)" % e
        return None
    except ImportError as e:
        # Most common on old Windows: Qt 6 DLL fails to load
        _ERRORS[name] = "import failed: %s" % e
        return None
    except Exception as e:
        _ERRORS[name] = "%s: %s" % (type(e).__name__, e)
        return None


def _load():
    global BINDING, QT_VERSION

    order = list(_ORDER)
    forced = os.environ.get("QT_BINDING", "").strip()
    if forced:
        order.insert(0, forced)

    last_err = None
    for name in order:
        got = _try_import(name)
        if got:
            BINDING = name
            # QT_VERSION_STR does not exist in PySide6; qVersion() works everywhere
            try:
                QT_VERSION = got[0].qVersion()
            except Exception:
                QT_VERSION = getattr(got[0], "QT_VERSION_STR", "?")
            return got
    return None


# ---------------------------------------------------------------- diagnostics

#: Set to the path of a log file once something has been written to it
LAST_LOG: str = ""


def _app_dir() -> str:
    """Directory for log files: next to the EXE when frozen, else source dir."""
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
    except Exception:
        pass
    return os.path.dirname(os.path.abspath(__file__))


def _emit(msg: str) -> bool:
    """
    Write to stderr, but never crash.

    Under PyInstaller with console=False, sys.stderr is None - calling
    .write() on it raises AttributeError, which is exactly the confusing
    "'NoneType' object has no attribute 'write'" error.
    """
    try:
        out = sys.stderr
        if out is not None and hasattr(out, "write"):
            out.write(msg)
            out.flush()
            return True
    except Exception:
        pass
    return False


def _write_log(msg: str) -> str:
    """Append to error.log in the app directory. Returns the path, or ''."""
    global LAST_LOG
    path = os.path.join(_app_dir(), "error.log")
    try:
        from datetime import datetime
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 66 + "\n")
            f.write("时间  ：%s\n" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            f.write("模式  ：%s\n" % ("打包EXE" if getattr(sys, "frozen", False) else "源码运行"))
            f.write("Python：%s\n" % sys.version.split()[0])
            f.write("系统  ：%s\n\n" % platform.platform())
            f.write(msg)
        LAST_LOG = path
        return path
    except Exception:
        return ""


def _message_box(title: str, msg: str) -> bool:
    """
    Show a native Windows message box via ctypes.

    Deliberately does NOT use Qt: this runs precisely when Qt is broken.
    """
    try:
        if platform.system() != "Windows":
            return False
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, title, 0x10)  # MB_ICONERROR
        return True
    except Exception:
        return False


def windows_nt():
    """Windows NT (major, minor), or None on other platforms / on failure."""
    try:
        if platform.system() == "Windows":
            v = sys.getwindowsversion()
            return (v[0], v[1])
    except Exception:
        pass
    return None


def _analyze(errors: dict) -> str:
    """Turn the collected import errors into an actionable explanation."""
    lines = []
    nt = windows_nt()

    is_dll = any("dll" in str(e).lower() or "找不到指定的程序" in str(e)
                 or "procedure could not be found" in str(e).lower()
                 for e in errors.values())
    missing = [n for n, e in errors.items() if "not installed" in str(e)]

    if is_dll and nt and nt < (6, 1):
        # NT 6.0 = Vista / Server 2008 (non-R2): PySide2 needs Win7 SP1,
        # and Python 3.9 (last Win7-capable release) will not install here.
        lines.append("诊断：本系统 Windows NT %d.%d 过于陈旧，无法运行本工具。" % nt)
        lines.append("")
        lines.append("    Qt 6  需要 Windows 10 1809 或更高")
        lines.append("    Qt 5  需要 Windows 7 SP1 或更高")
        lines.append("    Python 3.8（最后一个支持 Win7 的版本）也要求 Win7 SP1+")
        lines.append("")
        lines.append("建议：升级到 Windows Server 2008 R2 或更高版本，")
        lines.append("      或换一台机器运行本工具。")
    elif is_dll and nt and nt < (10, 0):
        lines.append("诊断：Qt 6 无法在 Windows NT %d.%d 上运行。" % nt)
        lines.append("")
        lines.append("Qt 6 要求 Windows 10 1809 或更高版本，本系统是老版本 Windows，")
        lines.append("Qt 6 的动态库链接了 Win10 才有的系统接口，加载时会失败。")
        lines.append("")
        lines.append("解决办法：改用 PySide2（基于 Qt 5）：")
        lines.append("    %s -m pip install PySide2" % sys.executable)
        lines.append("")
        lines.append("注意：Python 请用 3.8 — 它是最后一个支持 Win7 / Server 2008 R2")
        lines.append("      的版本（3.9 起安装程序直接拒绝在 Win7 上安装）。")
        lines.append("      PySide2 的 wheel 也只覆盖到 Python 3.9。")
        lines.append("      且必须在目标机器上重新打包，不能拷贝别的机器上打的包。")
    elif is_dll:
        lines.append("诊断：Qt 动态库加载失败。")
        lines.append("")
        lines.append("常见原因：")
        lines.append("  1) 杀毒软件拦截或删除了 Qt 的 DLL，加白名单后重新打包")
        lines.append("  2) 缺少 Visual C++ 运行库，安装 vcredist")
        lines.append("  3) 打包时没有收集到 Qt 插件，重新打包")
    elif missing and len(missing) == len(errors):
        lines.append("诊断：没有安装任何 Qt 绑定。")
        lines.append("")
        lines.append("安装方式：")
        lines.append("    Windows 10 / 11             ->  pip install PySide6")
        lines.append("    Windows 7 / Server 2008 R2  ->  pip install PySide2")
        lines.append("")
        lines.append("或运行：python install_deps.py")
    else:
        lines.append("诊断：Qt 绑定加载失败，原因见下方各条明细。")
    return "\n".join(lines)


def _report_failure():
    """Build the report, log it, show it, then let the ImportError fly."""
    lines = ["无法加载 Qt 图形库，程序无法启动。", ""]
    lines.append("逐个尝试的结果：")
    for name in _ORDER:
        reason = _ERRORS.get(name, "未尝试")
        lines.append("  %-9s %s" % (name, reason))
    lines.append("")
    lines.append(_analyze(_ERRORS))
    lines.append("")
    lines.append("本消息同时已写入 error.log。")
    msg = "\n".join(lines)

    framed = "\n" + "=" * 66 + "\n  No Qt binding available\n" + "=" * 66 + "\n\n" + msg + "\n"

    log = _write_log(msg)
    _emit(framed)

    shown = _message_box("IP Tool - Cannot Start", msg)
    if not shown:
        # Last resort: no Qt, no Windows box (e.g. Linux server)
        pass


_loaded = _load()

if _loaded is None:
    _report_failure()
    raise ImportError("No Qt binding found (tried: %s)" % ", ".join(_ORDER))

QtCore, QtGui, QtWidgets = _loaded

# ---------------------------------------------------------------- exports

Qt = QtCore.Qt
QDate = QtCore.QDate

QFont = QtGui.QFont
QColor = QtGui.QColor
QBrush = QtGui.QBrush

QApplication = QtWidgets.QApplication
QMainWindow = QtWidgets.QMainWindow
QWidget = QtWidgets.QWidget
QVBoxLayout = QtWidgets.QVBoxLayout
QHBoxLayout = QtWidgets.QHBoxLayout
QGridLayout = QtWidgets.QGridLayout
QTabWidget = QtWidgets.QTabWidget
QTableWidget = QtWidgets.QTableWidget
QTableWidgetItem = QtWidgets.QTableWidgetItem
QHeaderView = QtWidgets.QHeaderView
QAbstractItemView = QtWidgets.QAbstractItemView
QPushButton = QtWidgets.QPushButton
QLabel = QtWidgets.QLabel
QLineEdit = QtWidgets.QLineEdit
QSpinBox = QtWidgets.QSpinBox
QComboBox = QtWidgets.QComboBox
QGroupBox = QtWidgets.QGroupBox
QSplitter = QtWidgets.QSplitter
QFileDialog = QtWidgets.QFileDialog
QMessageBox = QtWidgets.QMessageBox
QListWidget = QtWidgets.QListWidget
QListWidgetItem = QtWidgets.QListWidgetItem
QCheckBox = QtWidgets.QCheckBox
QStatusBar = QtWidgets.QStatusBar
QInputDialog = QtWidgets.QInputDialog
QDateEdit = QtWidgets.QDateEdit


# ---------------------------------------------------------------- helpers

def qt_version() -> str:
    """Qt version string, e.g. '6.11.2' or '5.15.2'."""
    return QT_VERSION


def binding_name() -> str:
    """Which binding is in use: PySide6 / PySide2 / PyQt5."""
    return BINDING


def is_qt6() -> bool:
    return QT_VERSION.startswith("6")


def exec_app(app) -> int:
    """
    Start the event loop.

    PySide6 exposes ``exec()``; PySide2/PyQt5 expose ``exec_()``
    (PySide2 also aliases ``exec()``, but not all builds do).
    """
    fn = getattr(app, "exec", None)
    if callable(fn):
        try:
            return fn()
        except TypeError:
            pass
    return app.exec_()


def setup_high_dpi() -> bool:
    """
    Best-effort high-DPI setup. Qt 6 only; silently skipped on Qt 5.

    Returns True if the policy was applied.
    """
    try:
        if not is_qt6():
            return False
        policy = getattr(Qt, "HighDpiScaleFactorRoundingPolicy", None)
        setter = getattr(QApplication, "setHighDpiScaleFactorRoundingPolicy", None)
        if policy is None or setter is None:
            return False
        setter(policy.PassThrough)
        return True
    except Exception:
        return False


def windows_info() -> str:
    """Human-readable Windows version, or '' on other platforms."""
    try:
        if platform.system() != "Windows":
            return ""
        return platform.platform()
    except Exception:
        return ""


def recommend() -> str:
    """
    Recommend a binding for the current Windows version.
    Returns one of _ORDER, or '' when not on Windows.
    """
    if platform.system() != "Windows":
        return "PySide6"
    try:
        major, minor = sys.getwindowsversion()[:2]
    except Exception:
        return "PySide6"
    # NT 6.0 = Vista / Server 2008
    # NT 6.1 = Win7 / Server 2008 R2
    # NT 6.2 = Win8 / Server 2012
    # NT 6.3 = Win8.1 / Server 2012 R2
    # NT 10  = Win10 / Server 2016+
    if (major, minor) >= (10, 0):
        return "PySide6"      # Win10+ works with Qt6
    return "PySide2"          # Win7/2008 R2 and older need Qt5


# ---------------------------------------------------------------- self-test

if __name__ == "__main__":
    print("=" * 66)
    print("  Qt binding report")
    print("=" * 66)
    print("  Python   : %s (%s)" % (platform.python_version(),
                                    "64-bit" if sys.maxsize > 2 ** 32 else "32-bit"))
    print("  Platform : %s" % platform.platform())
    print("  Binding  : %s" % BINDING)
    print("  Qt       : %s   (%s)" % (QT_VERSION, "Qt6" if is_qt6() else "Qt5"))
    print("  Recommended for this OS: %s" % recommend())
    print("=" * 66)

    if platform.system() == "Windows" and is_qt6():
        try:
            major, minor = sys.getwindowsversion()[:2]
            if (major, minor) < (10, 0):
                print()
                print("  WARNING: Qt 6 requires Windows 10 1809 or later.")
                print("           This machine is NT %d.%d." % (major, minor))
                print("           If the app fails to start, install PySide2:")
                print("               pip install PySide2")
                print("=" * 66)
        except Exception:
            pass

    actually = binding_name()
    best = recommend()
    if actually != best:
        print()
        print("  NOTE: using %s, but %s is recommended here." % (actually, best))
