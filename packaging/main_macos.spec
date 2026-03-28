# -*- mode: python ; coding: utf-8 -*-
import os

ROOT = os.path.normpath(os.path.join(SPECPATH, '..'))

a = Analysis(
    [os.path.join(ROOT, 'main.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'memorylab.ico'), '.'),
        (os.path.join(ROOT, 'procurar_objeto.py'), '.'),
        (os.path.join(ROOT, 'procurar_distvel.py'), '.'),
        (os.path.join(ROOT, 'updater.py'), '.'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.utils.dataframe',
        'openpyxl.worksheet',
        'openpyxl.worksheet.worksheet',
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.sip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.distutils',
        'tkinter',
        'test',
        'unittest',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NeuroTrace',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ROOT, 'memorylab.ico'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='NeuroTrace',
)

app = BUNDLE(
    coll,
    name='NeuroTrace.app',
    icon=os.path.join(ROOT, 'memorylab.ico'),
    bundle_identifier='com.rodrigoorvate.neurotrace',
    info_plist={
        'CFBundleDisplayName': 'NeuroTrace',
        'CFBundleShortVersionString': '2.0.0',
        'CFBundleVersion': '2.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.15',
        'NSRequiresAquaSystemAppearance': False,
    },
)
