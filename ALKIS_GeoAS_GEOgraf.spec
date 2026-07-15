# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
od,ob,oh=collect_all('openpyxl'); ld,lb,lh=collect_all('lxml')
a=Analysis(['ALKIS_GeoAS_GEOgraf.py'],pathex=[],binaries=ob+lb,datas=od+ld,hiddenimports=oh+lh,hookspath=[],hooksconfig={},runtime_hooks=[],excludes=[],noarchive=False,optimize=0)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,a.binaries,a.datas,[],name='ALKIS_GeoAS_GEOgraf_V1.2.0',debug=False,bootloader_ignore_signals=False,strip=False,upx=True,console=False,disable_windowed_traceback=False,argv_emulation=False,target_arch=None,codesign_identity=None,entitlements_file=None)
