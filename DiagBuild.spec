# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置

block_cipher = None

# ---------------------------------------------------------------- Qt binding
# PySide6 (Qt 6) needs Windows 10 1809+; on older Windows use PySide2 (Qt 5).
# Pick whichever is actually installed so the bundle matches the runtime.
def _qt_binding():
    for name in ("PySide6", "PySide2", "PyQt5"):
        try:
            __import__(name)
            return name
        except ImportError:
            continue
    raise SystemExit("No Qt binding installed. Run: python install_deps.py")


QT_BINDING = _qt_binding()
print("Bundling Qt binding: %s" % QT_BINDING)

a = Analysis(
    ['app_gui.py'],
    pathex=['.'],
    binaries=[],
    datas=[('templates', 'templates')],
    hiddenimports=(
        # local modules
        ['ip_core', 'ip_word', 'ip_export', 'ip_ledger']
        # Qt modules (name depends on the binding)
        + ['%s.QtCore' % QT_BINDING,
           '%s.QtGui' % QT_BINDING,
           '%s.QtWidgets' % QT_BINDING]
        # other third-party
        + ['openpyxl', 'docx', 'docx.oxml', 'ipaddress']
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=(
        ['tkinter', 'matplotlib', 'numpy', 'pandas', 'PIL']
        # drop the Qt binding we are NOT bundling, plus heavy unused Qt modules
        + ['%s' % b for b in ('PySide6', 'PySide2', 'PyQt5') if b != QT_BINDING]
        + ['%s.Qt%s' % (QT_BINDING, m) for m in (
            'WebEngineCore', 'WebEngineWidgets', 'WebEngineQuick', 'WebChannel',
            'WebSockets', 'Network', 'Qml', 'Quick', 'Quick3D', 'QuickWidgets',
            '3DCore', 'Multimedia', 'Charts', 'DataVisualization', 'Bluetooth',
            'Nfc', 'Positioning', 'SerialPort', 'Svg', 'Test', 'Designer',
            'Help', 'OpenGL', 'OpenGLWidgets', 'PrintSupport', 'Sql',
            'StateMachine', 'UiTools', 'Xml', 'Concurrent', 'NetworkAuth',
            'RemoteObjects', 'Scxml', 'Sensors', 'SpatialAudio',
            'TextToSpeech', 'Pdf', 'PdfWidgets')]
    ),
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='IPAddressTool_' + QT_BINDING + '_Diag',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,                  # Diag build: keep console to show errors
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
    version='version_info.txt',
)
