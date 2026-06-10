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
hidden = collect_submodules('keyring') + collect_submodules('win32com') + [
    'win32ctypes',
    'win32ctypes.pywin32',
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    'win32api',
    'win32con',
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
    ],
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

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
