# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['new-app.py'],
             pathex=['C:\\Users\\SS078074\\PycharmProjects\\Schedule Comments\\venv\\Scripts'],
             binaries=[],
             datas=[],
             hiddenimports=['FileDialog'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='new-app',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               upx_exclude=[],
               name='new-app')
