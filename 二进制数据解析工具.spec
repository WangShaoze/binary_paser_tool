# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['test-02.py'],
             pathex=['D:\\Devleop\\code\\pythonCode\\DemoTest\\TEST'],
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

# 在spec文件中找到并修改这一行
exe = EXE(pyz,
         a.scripts,
         a.binaries,
         a.zipfiles,
         a.datas,
         [],
         name='二进制数据解析工具',
         debug=False,
         bootloader_ignore_signals=False,
         strip=False,
         upx=True,
         console=False,  # 这里设置为False表示无控制台
         disable_windowed_traceback=False,
         target_arch=None,
         codesign_identity=None,
         entitlements_file=None,
         icon=None)  # 可以添加icon参数指定图标文件


coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='二进制数据解析工具')
