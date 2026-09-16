# -*- coding: utf-8 -*-
"""
Smart dependency installer (ASCII output only).

Why this exists:
    python-docx needs lxml. On Windows, pip sometimes cannot find a
    matching prebuilt lxml wheel and falls back to compiling from source,
    which fails with:

        error: Microsoft Visual C++ 14.0 or greater is required

    That is NOT a problem with this tool. The usual causes are:

      1. OLD PIP      - cannot read modern wheel tags -> picks the source
                        tarball instead of the wheel.  FIX: upgrade pip.
      2. NEW PYTHON   - e.g. 3.15, pip resolves an lxml version that has
                        no wheel for it yet.           FIX: pin a newer lxml.
      3. NO WHEEL     - genuinely none available.      FIX: use Python 3.11/3.12.

    This script detects which case you are in and applies the right fix.

    It also installs PyInstaller, which is needed to build an EXE.
    Pass --no-build if you only want to run the app, not package it.

    It also picks the right Qt binding for your Windows version:
      - Windows 10 / 11            ->  PySide6  (Qt 6)
      - Windows 7 / Server 2008 R2 ->  PySide2  (Qt 5)
    Qt 6 requires Windows 10 1809 or later and will NOT load on older
    systems (ImportError: DLL load failed ... procedure could not be found).
    See qt_compat.py for details.

Run:
    python install_deps.py
"""

import os
import platform
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REQ = os.path.join(HERE, "requirements.txt")

MIRROR = "https://mirrors.cloud.tencent.com/pypi/simple"
MIRROR_HOST = "mirrors.cloud.tencent.com"

OK = "  [ OK ]"
NG = "  [FAIL]"
INFO = "  [ .. ]"

# lxml versions known to ship Windows wheels, newest first.
# Tried in order when the plain install fails.
LXML_FALLBACKS = [
    "6.1.3", "6.1.2", "6.1.1", "6.1.0",
    "6.0.4", "6.0.3", "6.0.2", "6.0.1",
]

# Python versions with the best third-party wheel coverage
RECOMMENDED_PY = ("3.11", "3.12")

# Non-Qt packages (the Qt binding is chosen separately, see pick_qt_binding)
BASE_PACKAGES = ["openpyxl", "python-docx"]

# PyInstaller is only needed for building an EXE, so it is installed
# separately. Newest-first fallback chain: if the latest release drops
# support for this Python/Windows, an older one usually still works.
BUILD_PACKAGE = "pyinstaller"
BUILD_FALLBACKS = ["6.3.0", "5.13.2"]
# 5.13.2 is the last release supporting Windows 7 / Python 3.7


def line(ch="="):
    print(ch * 66)


def run(args, show=False, timeout=900, cwd=None):
    """Run a command, return (returncode, combined_output)."""
    try:
        p = subprocess.run(
            args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            universal_newlines=True, errors="replace", timeout=timeout,
            cwd=cwd)
        out = p.stdout or ""
        if show:
            for ln in out.splitlines():
                print("        | " + ln)
        return p.returncode, out
    except subprocess.TimeoutExpired:
        return 124, "TIMEOUT"
    except Exception as e:
        return 1, str(e)


def pip(*args, show=False, cwd=None):
    return run([sys.executable, "-m", "pip", *args], show=show, cwd=cwd)


def pip_version():
    try:
        code, out = pip("--version")
        if code == 0:
            parts = out.split()
            if len(parts) >= 2:
                return parts[1]
    except Exception:
        pass
    return "0.0"


def v_tuple(s):
    try:
        return tuple(int(x) for x in s.split(".")[:3] if x.isdigit())
    except Exception:
        return (0,)


# ---------------------------------------------------------------- report

def windows_nt():
    """Return (major, minor) of the Windows NT version, or None."""
    try:
        if platform.system() != "Windows":
            return None
        v = sys.getwindowsversion()
        return (v[0], v[1])
    except Exception:
        return None


def pick_qt_binding():
    """
    Choose PySide6 or PySide2 based on the Windows version.

    Qt 6 needs Windows 10 1809+. Anything older (Win7, Server 2008 R2,
    Server 2008, Vista) must use Qt 5, i.e. PySide2.
    """
    nt = windows_nt()
    if nt is None:
        return "PySide6", "non-Windows (default)"
    major, minor = nt
    if (major, minor) >= (10, 0):
        return "PySide6", "Windows NT %d.%d (Win10+ / Server 2016+)" % nt
    return "PySide2", "Windows NT %d.%d (older than Win10)" % nt


def python_hint_for(pkg_major, pkg_minor, binding):
    """Warn when the Python version is a poor match for the binding."""
    msgs = []
    if binding == "PySide2":
        # PySide2 ships wheels for Python 3.5 - 3.9 only
        if (pkg_major, pkg_minor) >= (3, 10):
            msgs.append(
                "PySide2 has no wheels for Python %d.%d; use Python 3.8."
                % (pkg_major, pkg_minor))
        # Python 3.9+ installers refuse to run on Windows 7 (see PEP 11)
        nt = windows_nt()
        if nt and nt < (10, 0) and (pkg_major, pkg_minor) >= (3, 9):
            msgs.append(
                "Python %d.%d cannot run on this old Windows; use Python 3.8 "
                "(last release supporting Windows 7 / Server 2008 R2)."
                % (pkg_major, pkg_minor))
    return msgs


def report_env():
    major, minor = sys.version_info[:2]
    arch = "64-bit" if sys.maxsize > 2 ** 32 else "32-bit"
    print("  Python      : %d.%d.%d  (%s)" % (
        sys.version_info[0], major and minor or 0, sys.version_info[2], arch))
    print("  Executable  : %s" % sys.executable)
    print("  Platform    : %s" % platform.platform())
    print("  pip         : %s" % pip_version())
    print("  Mirror      : %s" % MIRROR_HOST)
    binding, why = pick_qt_binding()
    print("  Qt binding  : %s   (%s)" % (binding, why))
    return major, minor, arch


def diagnose_python(major, minor, arch):
    """Warn about Python versions that are known to cause wheel trouble."""
    warns = []
    if (major, minor) < (3, 8):
        warns.append("Python %d.%d is too old (need 3.8+)." % (major, minor))
    if (major, minor) >= (3, 15):
        warns.append(
            "Python %d.%d is very new; some packages may not have wheels yet."
            % (major, minor))
    if arch == "32-bit":
        warns.append(
            "32-bit Python: fewer prebuilt wheels are published for win32.")
    return warns


# ---------------------------------------------------------------- steps

def step_upgrade_pip():
    print("\n[1/4] Upgrading pip (this alone fixes most wheel problems)...")
    code, out = pip("install", "--upgrade", "pip", "-i", MIRROR)
    if code == 0:
        print("%s pip is now %s" % (OK, pip_version()))
        return True
    print("%s pip upgrade failed (continuing anyway)" % NG)
    for ln in out.splitlines()[-5:]:
        print("        | " + ln)
    return False


def step_try(req_args, label, show=False):
    """Try a pip install. Return (ok, output)."""
    print("\n%s" % label)
    code, out = pip(*req_args, show=show)
    if code == 0:
        print("%s installed" % OK)
        return True, out
    print("%s failed" % NG)
    low = out.lower()
    if "microsoft visual c++" in low or "error: command" in low:
        print("        -> pip tried to COMPILE from source (no matching wheel)")
    for ln in out.splitlines()[-8:]:
        print("        | " + ln)
    return False, out


def step_lxml_fallback():
    """Install lxml alone, pinning versions that are known to have wheels."""
    print("\n[3/4] Trying pinned lxml versions (wheel-only, no compiling)...")
    for v in LXML_FALLBACKS:
        print("%s lxml==%s ..." % (INFO, v))
        code, out = pip(
            "install", "--only-binary=:all:", "lxml==%s" % v, "-i", MIRROR)
        if code == 0:
            print("%s lxml==%s installed from wheel" % (OK, v))
            return True
        low = out.lower()
        if "no matching distribution" in low:
            print("        no wheel for this Python, trying next...")
        else:
            print("        failed")
    return False


def step_pyinstaller():
    """
    Install PyInstaller (needed only for building an EXE).

    Tries the newest release first, then walks the fallback chain.
    Returns True on success.
    """
    print("\n[3/4] Installing PyInstaller (needed to build the EXE)...")
    code, out = pip("install", "--only-binary=:all:", BUILD_PACKAGE, "-i", MIRROR)
    if code == 0:
        print("%s %s installed" % (OK, BUILD_PACKAGE))
        return True

    low = out.lower()
    if "no matching distribution" in low or "requires a different python" in low:
        print("%s latest release does not support this Python" % NG)
    else:
        print("%s latest release failed, trying older versions" % NG)

    for v in BUILD_FALLBACKS:
        print("%s pyinstaller==%s ..." % (INFO, v))
        code, out = pip("install", "--only-binary=:all:",
                        "%s==%s" % (BUILD_PACKAGE, v), "-i", MIRROR)
        if code == 0:
            print("%s pyinstaller==%s installed" % (OK, v))
            return True
        print("        failed")

    print("%s could not install PyInstaller" % NG)
    return False


def step_verify(need_build=True):
    print("\n[4/4] Verifying installation...")
    # (display label, importable module name)
    mods = [("openpyxl", "openpyxl"), ("python-docx", "docx"), ("lxml", "lxml")]
    bad = []
    for label, mod in mods:
        code, out = run([sys.executable, "-c", "import %s" % mod], cwd=HERE)
        if code == 0:
            print("%s %s" % (OK, label))
        else:
            print("%s %s" % (NG, label))
            bad.append(label)

    # Qt: go through qt_compat so any binding counts
    code, out = run(
        [sys.executable, "-c",
         "import qt_compat as q; print(q.BINDING, q.QT_VERSION)"],
        cwd=HERE)
    if code == 0:
        print("%s %-14s %s" % (OK, "Qt binding", out.strip()))
    else:
        print("%s %-14s no working Qt binding" % (NG, "Qt binding"))
        for ln in out.strip().splitlines()[-3:]:
            print("        | " + ln)
        bad.append("Qt")
    if need_build:
        code, out = run([sys.executable, "-c", "import PyInstaller"], cwd=HERE)
        if code == 0:
            try:
                import PyInstaller
                print("%s %-14s %s" % (OK, "PyInstaller",
                                       PyInstaller.__version__))
            except Exception:
                print("%s %-14s installed" % (OK, "PyInstaller"))
        else:
            print("%s %-14s not available - cannot build EXE" % (NG, "PyInstaller"))
            bad.append("PyInstaller")

    # entry point
    code, out = run([sys.executable, "-c",
                     "import app_gui; assert callable(app_gui.main)"],
                    cwd=HERE)
    if code == 0:
        print("%s app_gui entry point" % OK)
    else:
        print("%s app_gui entry point" % NG)
        bad.append("app_gui")
    return bad


# ---------------------------------------------------------------- main

def main():
    line()
    print("  IP Address Allocation Tool - Dependency Installer")
    line()
    print("\n[0/4] Environment")
    major, minor, arch = report_env()

    warns = diagnose_python(major, minor, arch)
    for w in warns:
        print("  [WARN] %s" % w)

    if not os.path.exists(REQ):
        print("\n%s requirements.txt not found: %s" % (NG, REQ))
        return 1

    # --- 1. upgrade pip first ---
    step_upgrade_pip()

    # --- 2. wheel-only install (fails fast instead of compiling) ---
    binding, why = pick_qt_binding()
    for m in python_hint_for(major, minor, binding):
        print("  [WARN] %s" % m)

    print("\n[2/4] Installing %s + base packages (wheel-only)..." % binding)
    ok, out = step_try(
        ["install", "--only-binary=:all:", binding,
         *BASE_PACKAGES, "-i", MIRROR],
        "      %s / openpyxl / python-docx" % binding)

    # --- 3. fallback: pin lxml, then retry the rest ---
    if not ok:
        low = out.lower()
        needs_lxml = ("lxml" in low or "visual c++" in low or
                      "no matching distribution" in low)
        if needs_lxml:
            if step_lxml_fallback():
                print("\n[3/4] Retrying remaining packages...")
                ok, _ = step_try(
                    ["install", "--only-binary=:all:",
                     binding, *BASE_PACKAGES, "-i", MIRROR],
                    "      installing %s / openpyxl / python-docx ..." % binding)

    # --- 3b. PyInstaller (skip with --no-build) ---
    need_build = "--no-build" not in sys.argv
    if need_build:
        step_pyinstaller()

    # --- 4. verify ---
    bad = step_verify(need_build=need_build)

    line()
    if not bad:
        print("  RESULT: All dependencies installed successfully.")
        print()
        print("  Start the app with:")
        print("      %s app_gui.py" % sys.executable)
        print("  or double-click START.vbs")
        line()
        return 0

    print("  RESULT: %d package(s) still missing: %s" % (len(bad), ", ".join(bad)))
    print()
    print("  Why this happens:")
    if "PyInstaller" in bad:
        print("    PyInstaller is only needed to build an EXE. You can still")
        print("    run the app with:  %s app_gui.py" % sys.executable)
        print("    To build later:    %s -m pip install pyinstaller" % sys.executable)
        print()
    if "Qt" in bad:
        nt = windows_nt()
        if nt and nt < (10, 0):
            print("    This is Windows NT %d.%d. Qt 6 cannot run here -" % nt)
            print("    it requires Windows 10 1809 or later. Use PySide2 instead:")
            print("        %s -m pip install PySide2" % sys.executable)
            print("    Note: PySide2 needs Python 3.9 or 3.10 on old systems.")
        else:
            print("    No working Qt binding. Try:")
            print("        %s -m pip install PySide6" % sys.executable)
        print()
    if "lxml" in bad:
        print("    pip could not find a prebuilt lxml wheel for")
        print("    Python %d.%d (%s), so it tried to build lxml from" % (major, minor, arch))
        print("    C source, which needs a C++ compiler that is not installed.")
        print("    lxml is required by python-docx, so python-docx fails too.")
    else:
        print("    lxml installed fine, but a dependent package still fails")
        print("    to import. This is usually a partial or corrupted install.")
        print("    Try: %s -m pip install --force-reinstall -r requirements.txt"
              % sys.executable)
    print()
    print("  Choose ONE of these fixes:")
    print()
    print("  A) Use Python %s or %s  (RECOMMENDED)" % RECOMMENDED_PY)
    print("     These have the best coverage of prebuilt wheels.")
    print("     Download: https://www.python.org/downloads/")
    print("     Then delete the .venv folder and run this script again.")
    print()
    print("  B) Install the C++ compiler, then rerun this script")
    print("     https://visualstudio.microsoft.com/visual-cpp-build-tools/")
    print("     Check 'Desktop development with C++' during install.")
    print()
    print("  C) Let pip compile anyway (slow, needs option B first)")
    print("     %s -m pip install -r requirements.txt" % sys.executable)
    line()
    return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n  Cancelled.")
        sys.exit(130)
