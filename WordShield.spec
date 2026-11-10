# 在 macOS 上执行。使用 onedir，避免 onefile 启动时解压到临时目录。
from PyInstaller.utils.hooks import collect_submodules

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
          target_arch=None, codesign_identity=None, entitlements_file=None)
collection = COLLECT(exe, analysis.binaries, analysis.datas,
                     strip=False, upx=False, name='WordShield')
app = BUNDLE(collection, name='WordShield.app', icon=None,
             bundle_identifier='local.wordshield.offline',
             info_plist={'CFBundleDisplayName': 'Word 离线脱敏与回填',
                         'CFBundleShortVersionString': '0.2.0',
                         'NSHighResolutionCapable': True,
                         'NSPrincipalClass': 'NSApplication'})
