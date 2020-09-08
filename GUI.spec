# -*- mode: python ; coding: utf-8 -*-

block_cipher = None
from PyInstaller.utils.hooks import collect_submodules

hiddenimports_pycel = collect_submodules('pycel')
hiddenimports_pycelLib = collect_submodules('pycel.lib')
all_hidden_imports = hiddenimports_pycel + hiddenimports_pycelLib

a = Analysis(['GUI.py'],
             pathex=['C:\\Users\\paul.jones\\Documents\\GitHub\\ModiffyEfficiency\\ModiffyEfficiency'],
             binaries=[],
             datas=[],
             hiddenimports=all_hidden_imports,
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

a.datas += [('images.png','C:\\Users\\paul.jones\\Documents\\GitHub\\ModiffyEfficiency\\ModiffyEfficiency\\images.png', "DATA"), ('package.json','C:\\Users\\paul.jones\\Documents\\GitHub\\ModiffyEfficiency\\ModiffyEfficiency\\package.json', "DATA")]
			 
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='GUI',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )