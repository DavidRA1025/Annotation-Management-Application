# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

# Specify the icon path relative to the location of the spec file
icon_path = 'AnnotAPP.ico'

a = Analysis(
    ['AnnotAPP.py'],
    pathex=[],
    binaries=[],
    datas=[
        (icon_path, '.'),
        ('file_version_info.txt', '.'),
        ('AnnotAPP.manifest', '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AnnotAPP',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
    version='file_version_info.txt',
    manifest='AnnotAPP.manifest'
)