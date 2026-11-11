# 在 macOS 上执行。使用 onedir，避免 onefile 启动时解压到临时目录。
import os
from pathlib import Path
import re

from PyInstaller.utils.hooks import collect_submodules

# 版本号只在 wordshield/__init__.py 维护。
VERSION = re.search(r'^__version__ = "([^"]+)"$', Path('wordshield/__init__.py').read_text(), re.M).group(1)
# 留空为当前架构；universal2 要求全部原生扩展都含 x86_64 与 arm64。
TARGET_ARCH = os.environ.get('WORDSHIELD_TARGET_ARCH') or None

analysis = Analysis(
    ['scripts/app_entry.py'], pathex=['.'],
    binaries=[], datas=[], hiddenimports=collect_submodules('docx'),
    hookspath=[], runtime_hooks=[], excludes=['pytest', 'unittest', 'requests', 'httpx'],
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(pyz, analysis.scripts, [], exclude_binaries=True,
          name='WordShield', debug=False, bootloader_ignore_signals=False,
          strip=False, upx=False, console=False, argv_emulation=False,
          target_arch=TARGET_ARCH, codesign_identity=None, entitlements_file=None)
collection = COLLECT(exe, analysis.binaries, analysis.datas,
                     strip=False, upx=False, name='WordShield')
app = BUNDLE(collection, name='WordShield.app', icon=None,
             bundle_identifier='local.wordshield.offline',
             info_plist={'CFBundleDisplayName': 'Word 离线脱敏与回填',
                         'CFBundleShortVersionString': VERSION,
                         'CFBundleVersion': VERSION,
                         'LSMinimumSystemVersion': '11.0',
                         'NSHighResolutionCapable': True,
                         'NSPrincipalClass': 'NSApplication'})
