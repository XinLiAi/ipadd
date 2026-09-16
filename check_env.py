# -*- coding: utf-8 -*-
"""
Environment diagnostics (ASCII output only).

Run directly:  python check_env.py
Or double-click CHECK.bat

Checks Python version, dependencies, local modules, entry point,
template files and the history database, then reports what is missing.
"""

import glob
import os
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
OK = "  [ OK ]"
NG = "  [FAIL]"
fails = []


def line(ch="="):
    print(ch * 62)


def chk(label, fn):
    try:
        detail = fn()
        print("%s %-14s %s" % (OK, label, detail))
        return True
    except Exception as e:
        print("%s %-14s %s" % (NG, label, e))
        fails.append((label, str(e)))
        return False


line()
print("  IP Address Allocation Tool - Environment Check")
line()

# 1. Python
print("\n[1/6] Python")
v = sys.version.split()[0]
print("%s %-14s %s" % (OK, "version", v))
major, minor = sys.version_info[:2]
if (major, minor) < (3, 8):
    print("%s %-14s need 3.8+, got %s" % (NG, "too old", v))
    fails.append(("python-version", v))
else:
    print("%s %-14s 3.8+ required, satisfied" % (OK, "compatible",))
print("%s %-14s %s" % (OK, "executable", sys.executable))
print("%s %-14s %s" % (OK, "mode", "frozen EXE" if getattr(sys, "frozen", False) else "source"))

# 2. Dependencies
print("\n[2/6] Dependencies")
def _qt():
    import qt_compat as q
    if q.recommend() != q.binding_name():
        return "%s %s  (this OS prefers %s)" % (
            q.binding_name(), q.QT_VERSION, q.recommend())
    return "%s %s" % (q.binding_name(), q.QT_VERSION)
def _openpyxl():
    import openpyxl
    return openpyxl.__version__
def _docx():
    import docx
    return "python-docx"
def _lxml():
    import lxml
    return "required by python-docx"

chk("Qt binding", _qt)
chk("openpyxl", _openpyxl)
chk("python-docx", _docx)
chk("lxml", _lxml)

# 3. Local modules
print("\n[3/6] Local modules")
sys.path.insert(0, HERE)
for m in ("ip_core", "ip_word", "ip_export", "ip_ledger"):
    def _mk(name):
        def f():
            mod = __import__(name)
            return "loaded"
        return f
    chk(m, _mk(m))

# 4. Entry point
print("\n[4/6] Entry point")
def _gui():
    import app_gui
    if not callable(getattr(app_gui, "main", None)):
        raise RuntimeError("main() is missing - app cannot start")
    return "main() present, v%s" % getattr(app_gui, "APP_VER", "?")
chk("app_gui", _gui)

# 5. Template
print("\n[5/6] Built-in template")
def _tpl():
    files = [f for f in glob.glob(os.path.join(HERE, "templates", "*.docx"))
             if not os.path.basename(f).startswith("~$")]
    if not files:
        raise RuntimeError("no .docx found in templates/")
    return "%d file(s): %s" % (len(files), ", ".join(os.path.basename(f) for f in files))
chk("templates", _tpl)

# 6. Database
print("\n[6/6] History database")
db = os.path.join(HERE, "history.db")
if os.path.exists(db):
    print("%s %-14s %s (%.1f KB)" % (OK, "history.db", "exists", os.path.getsize(db) / 1024))
    try:
        import ip_ledger
        h = ip_ledger.History(db)
        st = h.stats()
        print("%s %-14s %s" % (OK, "readable", st))
        h.close()
    except Exception as e:
        print("%s %-14s %s" % (NG, "readable", e))
        fails.append(("database", str(e)))
else:
    print("%s %-14s not created yet (auto-created on first run)" % (OK, "history.db"))

# Summary
line()
if fails:
    print("  RESULT: %d problem(s) found" % len(fails))
    for k, v in fails:
        print("    - %s: %s" % (k, v))
    print()
    if any(k in ("PySide6", "openpyxl", "python-docx") for k, _ in fails):
        print("  Fix: run this command")
        print("       %s -m pip install -r requirements.txt" % sys.executable)
        print("       (add: -i https://mirrors.cloud.tencent.com/pypi/simple for speed)")
else:
    print("  RESULT: All checks passed.")
    print()
    print("  The environment looks fine. To start the app run:")
    print("       %s app_gui.py" % sys.executable)
    print("  or double-click START.bat / START.vbs")
line()

if os.environ.get("CHECK_PAUSE") == "1":
    try:
        input("\nPress Enter to continue...")
    except Exception:
        pass

sys.exit(1 if fails else 0)
