"""构建前原生依赖检查与打包入口自检。"""
import os
from pathlib import Path
import subprocess
import sys

import pytest

from scripts.check_native_deps import external_dependencies, external_rpaths, main
from wordshield.paths import PROJECT_ROOT

SOURCE_BUILT = """\
/venv/site-packages/cryptography/hazmat/bindings/_rust.abi3.so:
\t@rpath/cryptography.hazmat.bindings._rust.abi3.so (compatibility version 0.0.0, current version 0.0.0)
\t/usr/local/opt/openssl@3/lib/libssl.3.dylib (compatibility version 3.0.0, current version 3.0.0)
\t/usr/local/opt/openssl@3/lib/libcrypto.3.dylib (compatibility version 3.0.0, current version 3.0.0)
\t/usr/lib/libiconv.2.dylib (compatibility version 7.0.0, current version 7.0.0)
\t/usr/lib/libSystem.B.dylib (compatibility version 1.0.0, current version 1351.0.0)
"""

OFFICIAL_UNIVERSAL = """\
/venv/site-packages/lxml/etree.cpython-312-darwin.so (architecture x86_64):
\t/usr/lib/libz.1.dylib (compatibility version 1.0.0, current version 1.2.12)
\t/System/Library/Frameworks/Security.framework/Versions/A/Security (compatibility version 1.0.0, current version 1.0.0)
\t@loader_path/../lib/libhelper.dylib (compatibility version 1.0.0, current version 1.0.0)
/venv/site-packages/lxml/etree.cpython-312-darwin.so (architecture arm64):
\t/usr/lib/libSystem.B.dylib (compatibility version 1.0.0, current version 1351.0.0)
"""


def test_source_built_extension_linking_homebrew_openssl_is_reported():
    assert external_dependencies(SOURCE_BUILT) == [
        "/usr/local/opt/openssl@3/lib/libssl.3.dylib",
        "/usr/local/opt/openssl@3/lib/libcrypto.3.dylib",
    ]


def test_system_and_relative_dependencies_are_allowed():
    assert external_dependencies(OFFICIAL_UNIVERSAL) == []


def test_absolute_rpath_outside_system_is_reported():
    commands = """\
Load command 12
          cmd LC_RPATH
      cmdsize 40
         path /opt/homebrew/opt/openssl@3/lib (offset 12)
Load command 13
          cmd LC_RPATH
      cmdsize 32
         path @loader_path/../lib (offset 12)
"""
    assert external_rpaths(commands) == ["/opt/homebrew/opt/openssl@3/lib"]


@pytest.mark.skipif(sys.platform != "darwin", reason="otool 仅在 macOS 上可用")
def test_current_environment_passes_native_dependency_check(capsys):
    assert main() == 0, capsys.readouterr().err


def test_entry_self_test_runs_without_opening_the_interface():
    environment = dict(os.environ, WORDSHIELD_SELF_TEST="1", PYTHONPATH=str(PROJECT_ROOT),
                       PYTHONDONTWRITEBYTECODE="1")
    result = subprocess.run([sys.executable, "-B", str(PROJECT_ROOT / "scripts" / "app_entry.py")],
                            env=environment, capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "WordShield self-test OK"


def test_build_script_runs_both_checks_in_order():
    script = (PROJECT_ROOT / "scripts" / "build_macos.sh").read_text()
    order = [script.index(marker) for marker in (
        "scripts/check_native_deps.py", "-m pytest", "-m PyInstaller", "WORDSHIELD_SELF_TEST=1")]
    assert order == sorted(order)
