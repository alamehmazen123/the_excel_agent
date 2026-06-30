# PyInstaller spec -> ONEDIR build (folder) for robust distribution.
# A folder build has no giant appended archive, so it avoids the
# "Could not load PyInstaller's embedded PKG archive" corruption error and is
# far friendlier to antivirus. The folder is then wrapped by Inno Setup
# (installer.iss) into ExcelIntelligenceAgent-Setup.exe.
#
# Build with:  pyinstaller build.spec --noconfirm
import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None
project_dir = os.path.abspath('.')
ICON = 'ui/resources/app.ico'

# keyring loads its Windows backend dynamically; pull all submodules in.
# pywin32 (win32com/pythoncom) powers the Excel COM finalizer.
# openpyxl (pure Python) is detected by PyInstaller's static analysis, but its
# chart/drawing/serialization submodules are loaded DYNAMICALLY at runtime and
# would be MISSING without an explicit collect_submodules call -- this was the
# root cause of the v1.15.4 bug where only the Pivot Analysis sheet appeared
# (the openpyxl write phase failed with an ImportError during wb.save() because
# openpyxl.drawing, openpyxl.chart.axis/series/plotarea, etc. were not bundled).
hidden = (collect_submodules('keyring')
          + collect_submodules('win32com')
          + collect_submodules('openpyxl')
          + collect_submodules('et_xmlfile')
          # Bundle the ENTIRE engine package so every analyzer / new module
          # (derived_metrics, periods, context, insights, semantic, …) is present
          # in the frozen exe exactly like the local project — otherwise a module
          # PyInstaller's static analysis happens to miss fails at runtime and that
          # sheet silently drops out ("not loading all the generated sheets").
          + collect_submodules('core')
          + [
    'win32ctypes',
    'win32ctypes.pywin32',
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    'win32api',
    'win32con',
])

# PySide6 ships dozens of large Qt modules; this app uses only QtWidgets/QtGui/
# QtCore (+ QtSvg for the checkbox tick). Excluding the rest cuts the bundle from
# ~190 MB to a fraction of that (WebEngine alone is >120 MB).
_QT_EXCLUDES = [
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets',
    'PySide6.QtWebEngineQuick', 'PySide6.QtWebChannel', 'PySide6.QtWebSockets',
    'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuick3D',
    'PySide6.QtQuickWidgets', 'PySide6.QtQuickControls2',
    'PySide6.Qt3DCore', 'PySide6.Qt3DRender', 'PySide6.Qt3DInput',
    'PySide6.Qt3DAnimation', 'PySide6.Qt3DExtras', 'PySide6.Qt3DLogic',
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets',
    'PySide6.QtCharts', 'PySide6.QtDataVisualization',
    'PySide6.QtPdf', 'PySide6.QtPdfWidgets', 'PySide6.QtSql', 'PySide6.QtDBus',
    'PySide6.QtBluetooth', 'PySide6.QtNfc', 'PySide6.QtPositioning',
    'PySide6.QtLocation', 'PySide6.QtSensors', 'PySide6.QtSerialPort',
    'PySide6.QtTest', 'PySide6.QtDesigner', 'PySide6.QtUiTools',
    'PySide6.QtHelp', 'PySide6.QtRemoteObjects', 'PySide6.QtScxml',
    'PySide6.QtStateMachine', 'PySide6.QtTextToSpeech', 'PySide6.QtWebView',
    'PySide6.QtOpenGL', 'PySide6.QtOpenGLWidgets',
]

# Bundle the gitignored secrets module (holds the default Groq key) if present.
# It is imported conditionally in config.py, so PyInstaller needs the hint.
if os.path.exists(os.path.join(project_dir, 'local_secrets.py')):
    hidden.append('local_secrets')

a = Analysis(
    ['main.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        ('ui/resources/style.qss', 'ui/resources'),
        ('ui/resources/app.ico', 'ui/resources'),
        ('ui/resources/check.svg', 'ui/resources'),
        # The reference library (the engine 'brain') -- ship the JSON data so
        # colleagues get the same decoded Smart Tables / summaries.
        ('core/library/data/headers.json', 'core/library/data'),
        ('core/library/data/codes.json', 'core/library/data'),
        ('core/library/data/meta.json', 'core/library/data'),
    ],
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'pytest', 'numpy', 'pandas',
              'scipy', 'IPython', 'PIL', 'pythonwin'] + _QT_EXCLUDES,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Drop the unused Qt MODULE DLLs + the qml/ tree that PySide6's hook bundles as
# BINARIES regardless of the python-module excludes above (this is what keeps the
# bundle large). We KEEP opengl32sw / d3dcompiler (software GL — needed to render
# over RustDesk / RDP / VMs) and Qt6Network/translations (safe).
_DROP_BINS = (
    'qt6qml', 'qt6quick', 'qt6quick3d', 'qt63d', 'qt6designer', 'qt6charts',
    'qt6datavisualization', 'qt6multimedia', 'qt6pdf', 'qt6sql', 'qt6test',
    'qt6webengine', 'qt6webchannel', 'qt6websockets', 'qt6sensors',
    'qt6location', 'qt6positioning', 'qt6serialport', 'qt6bluetooth', 'qt6nfc',
    'qt6remoteobjects', 'qt6scxml', 'qt6statemachine', 'qt6texttospeech',
    'qt6virtualkeyboard', 'qt6help', 'qt6designercomponents',
    r'\qml\\', '/qml/', 'pyside6\\examples', 'pythonwin',
)


def _keep(entry):
    path = str(entry[0]).lower().replace('/', '\\')
    return not any(tok.replace('/', '\\') in path for tok in _DROP_BINS)


a.binaries = [b for b in a.binaries if _keep(b)]
a.datas = [d for d in a.datas if _keep(d)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# exclude_binaries=True  ->  ONEDIR (binaries live next to the exe, not inside).
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ExcelIntelligenceAgent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX can trigger AV false positives; keep off
    console=False,             # windowed app -- no terminal
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON if os.path.exists(ICON) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ExcelIntelligenceAgent',
)
