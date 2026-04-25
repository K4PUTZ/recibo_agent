# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_data_files

def easyocr_model_data():
    # Coleta os arquivos de modelo do EasyOCR
    import easyocr
    model_dir = os.path.join(os.path.dirname(easyocr.__file__), 'model')
    datas = []
    if os.path.exists(model_dir):
        for root, dirs, files in os.walk(model_dir):
            for f in files:
                full_path = os.path.join(root, f)
                rel_path = os.path.relpath(full_path, model_dir)
                datas.append((full_path, os.path.join('easyocr/model', rel_path)))
    return datas

block_cipher = None

a = Analysis([
    '../recibo_agent_gui.py',
],
    pathex=['.'],
    binaries=[],
    datas=collect_data_files('easyocr') + collect_data_files('PIL') + collect_data_files('fitz') + collect_data_files('docx') + easyocr_model_data() + [('../token.json', '.'), ('../user_credentials.json', '.'), ('../config.py', '.'), ('../requirements.txt', '.'), ('../run.py', '.'), ('../processor.py', '.'), ('../graph_client.py', '.'), ('../auth.py', '.'), ('../DOCUMENTACAO.md', '.'), ('../AppleScript', '.'), ('../ocr_data.txt', '.'), ('../credentials.json', '.')],
    hiddenimports=['easyocr', 'PIL', 'fitz', 'docx', 'tkinter', 'colorama', 'requests'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Recibo Agent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Recibo Agent'
)
