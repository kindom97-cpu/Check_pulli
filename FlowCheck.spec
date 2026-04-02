# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['flowcheck_app.py'],
    pathex=[],
    binaries=[],
    datas=[('flowcheck_engine.py', '.')],
    hiddenimports=['duckdb', 'pandas', 'pandas._libs.tslibs.base', 'pandas._libs.tslibs.nattype', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.timestamps', 'pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.timezones', 'pandas._libs.tslibs.parsing', 'pandas._libs.tslibs.offsets', 'pandas._libs.tslibs.period', 'pandas._libs.tslibs.vectorized', 'pandas._libs.hashtable', 'pandas._libs.lib', 'pandas._libs.missing', 'pandas._libs.writers', 'pandas._libs.ops', 'pandas._libs.interval', 'pandas._libs.indexing', 'pandas._libs.join', 'pandas._libs.reduction', 'pandas._libs.groupby', 'pandas._libs.window.aggregations', 'pandas._libs.window.indexers', 'pandas.io.formats.excel', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'openpyxl.chart', 'openpyxl.drawing', 'openpyxl.worksheet', 'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.scrolledtext'],
    hookspath=[],
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
    name='FlowCheck',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
