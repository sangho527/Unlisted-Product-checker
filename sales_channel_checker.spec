# PyInstaller 설정 — 판매채널 상품 점검 Windows exe 빌드용 (GitHub Actions: .github/workflows/build-exe.yml)
#   pyinstaller sales_channel_checker.spec
import os
block_cipher = None
ROOT = os.path.abspath('.')

a = Analysis(
    ['app/gui.py'],
    pathex=[ROOT, os.path.join(ROOT, 'app'), os.path.join(ROOT, 'src')],
    datas=[('src', 'src'), ('search', 'search')],
    hiddenimports=['config', 'runner', 'search_check', 'ui_search', 'cm29_names', 'cm29_match', 'cm29_search', 'pickle', 'shutil', 'websocket', 'openpyxl', 'openpyxl.workbook.properties', 'pandas', 'numpy', 'difflib', 'urllib.parse'],
    excludes=['matplotlib', 'scipy', 'IPython', 'pytest', 'PyQt5', 'PySide2'],
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
    name='SalesChannelChecker',
    console=False,
    upx=False,
)
