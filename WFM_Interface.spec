# -*- mode: python -*-
a = Analysis(['WFM_Interface.py'],
             pathex=['D:\\Clients\\Docomo\\WFM'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='WFM_Interface.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False )
