# -*- mode: python -*-
import distutils
if distutils.distutils_path.endswith('__init__.py'):
    distutils.distutils_path = os.path.dirname(distutils.distutils_path)

import json
with open("build_meta_info.json", "r") as f:
    build_version = json.load(f)["build_version"] + 1

filename = f'Load-Daily-Stats v{build_version}'

print(json.dumps({"build_version" : build_version}), file=open("build_meta_info.json", "w"))

block_cipher = None

a = Analysis(['load.py'],
             pathex=['C:\\Users\\ramnautk\\OneDrive - Bristol-Myers Squibb (O365-D)\\Projects\\vitalize-outlook-download'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name=filename,
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True , icon='icon\daily-stats.ico')
