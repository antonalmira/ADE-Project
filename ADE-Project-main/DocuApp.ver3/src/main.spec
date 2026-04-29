# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
    ('..\\templates\\DER-template.docx', 'templates'),
    ('..\\performancedata_testnames.json', '.'),
    ('..\\waveform_testnames.json', '.'),
    ('..\\resources\\*', 'resources'),
    ('..\\DocuApp_ver6.ui', '.')
    ('tardis.png', '.'),      
    ('..\\..\\templates', 'templates'), 
    ('..\\..\\resource\\BOM_PIXL.xlsx', 'resource')

    ],
    hiddenimports=[
    'openpyxl',  # For excel_handlers.py, excel_utils.py, chart_extractor.py
    'docx',      # For document_generator.py, document_handler.py, word_utils.py
    'PyQt5.QtWidgets',
    'PyQt5.QtGui',
    'PyQt5.QtCore',
    'PIL'        # For image_utils.py (Pillow)
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['__pycache__', 'myenv'],
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
    name='TARDIS',
    icon='tardis.ico',
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
