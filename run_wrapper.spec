# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['run_wrapper.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.py', '.'),
        ('config.toml','.')
        ],
    hiddenimports=[
        'streamlit.runtime.scriptrunner.magic_funcs',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.workbook',
        'openpyxl.writer.excel',
        'openpyxl.reader.excel',
        'et_xmlfile',
        'jdcal'
        ],
    hookspath=['./hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='run_wrapper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
