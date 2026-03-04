# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for the menu bar agent (menubar.py).
Produces a self-contained .app bundle — no Python runtime required.
"""

import os
from PyInstaller.building.api import PYZ, EXE, COLLECT
from PyInstaller.building.build_main import Analysis
from PyInstaller.building.osx import BUNDLE

HERE = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    [os.path.join(HERE, 'menubar.py')],
    pathex=[HERE],
    binaries=[],
    datas=[
        # Bundle the proto_pb2 bindings
        (os.path.join(HERE, 'proto_pb2'), 'proto_pb2'),
        # Bundle the main converter module
        (os.path.join(HERE, 'pro_to_pptx.py'), '.'),
        # Bundle the menu bar icons
        (os.path.join(HERE, 'ProPresenter Converter.app', 'Contents', 'Resources', 'menubar_idle.png'), '.'),
        (os.path.join(HERE, 'ProPresenter Converter.app', 'Contents', 'Resources', 'menubar_active.png'), '.'),
        (os.path.join(HERE, 'ProPresenter Converter.app', 'Contents', 'Resources', 'AppIcon.icns'), '.'),
    ],
    hiddenimports=[
        'proto_pb2',
        'rumps',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
        'watchdog.observers.fsevents',
        'striprtf',
        'striprtf.striprtf',
        'pptx',
        'lxml',
        'lxml.etree',
        'google.protobuf',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'PIL', 'Pillow', 'IPython'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ProPresenter Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch='arm64',
    codesign_identity=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ProPresenter Converter',
)

app = BUNDLE(
    coll,
    name='ProPresenter Converter.app',
    icon=os.path.join(HERE, 'ProPresenter Converter.app', 'Contents', 'Resources', 'AppIcon.icns'),
    bundle_identifier='com.localapp.propresenter-converter',
    info_plist={
        'CFBundleName': 'ProPresenter Converter',
        'CFBundleDisplayName': 'ProPresenter Converter',
        'CFBundleShortVersionString': '1.1.0',
        'CFBundleVersion': '1.1.0',
        'NSHighResolutionCapable': True,
        'LSUIElement': True,          # hide dock icon — menu bar only
        'NSPrincipalClass': 'NSApplication',
        'LSMinimumSystemVersion': '12.0',
        'NSHumanReadableCopyright': '© 2025',
        'CFBundleDocumentTypes': [{
            'CFBundleTypeName': 'ProPresenter 7 Presentation',
            'CFBundleTypeExtensions': ['pro'],
            'CFBundleTypeRole': 'Viewer',
        }],
    },
)
