# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates')]
          + collect_data_files('rapidocr_onnxruntime')
          + collect_data_files('pdfminer')      # pdfminer cmap 编码表数据
          + collect_data_files('pdfplumber'),
    hiddenimports=['flask', 'docx', 'openpyxl', 'ai_providers', 'discovery']
             + collect_submodules('rapidocr_onnxruntime')
             + collect_submodules('pdfplumber')
             + collect_submodules('pdfminer'),   # pdfminer 含动态加载，需全量收集
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
    [],
    exclude_binaries=True,
    name='server',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='install\\app.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='server',
)
