"""构建前检查：已安装的原生扩展不能依赖 Homebrew 等外部动态库。

PyInstaller 会把同名动态库合并为一份。例如在 Intel Mac 上从源码编译的 cryptography
会链接 Homebrew 的 libssl.3.dylib，而 python.org Python 自带一份同名但更旧的 libssl.3.dylib，
打包后应用启动即因缺少符号而崩溃。官方二进制 wheel 已静态链接 OpenSSL，不受影响。
"""
from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import sysconfig

SYSTEM_PREFIXES = ("/usr/lib/", "/System/Library/")
RELATIVE_PREFIXES = ("@rpath/", "@loader_path/", "@executable_path/")


def external_dependencies(otool_output: str) -> list[str]:
    """从 `otool -L` 输出中找出既非系统库、也非相对路径的依赖。"""
    found = []
    for line in otool_output.splitlines()[1:]:
        if not line.startswith(("\t", " ")):
            continue  # 胖二进制中每个架构的标题行
        dependency = line.strip().split(" (", 1)[0]
        if dependency and not dependency.startswith(SYSTEM_PREFIXES + RELATIVE_PREFIXES):
            found.append(dependency)
    return found


def external_rpaths(otool_output: str) -> list[str]:
    """从 `otool -l` 输出中找出指向系统目录之外的绝对 LC_RPATH。"""
    found, in_rpath = [], False
    for line in otool_output.splitlines():
        words = line.split()
        if words[:2] == ["cmd", "LC_RPATH"]:
            in_rpath = True
        elif in_rpath and words[:1] == ["path"] and len(words) > 1:
            path = words[1]
            if path.startswith("/") and not path.startswith(SYSTEM_PREFIXES):
                found.append(path)
            in_rpath = False
    return found


def native_files(roots: list[Path]):
    for root in roots:
        for pattern in ("*.so", "*.dylib"):
            yield from sorted(root.rglob(pattern))


def main() -> int:
    roots = sorted({Path(sysconfig.get_paths()[key]) for key in ("purelib", "platlib")})
    problems = []
    for binary in native_files(roots):
        linked = subprocess.run(["/usr/bin/otool", "-L", str(binary)],
                                capture_output=True, text=True, check=True).stdout
        commands = subprocess.run(["/usr/bin/otool", "-l", str(binary)],
                                  capture_output=True, text=True, check=True).stdout
        for dependency in external_dependencies(linked) + external_rpaths(commands):
            relative = next((binary.relative_to(r) for r in roots if binary.is_relative_to(r)), binary)
            problems.append(f"  {relative} -> {dependency}")
    if problems:
        print("以下原生扩展依赖了外部动态库，打包后的应用可能无法启动：", file=sys.stderr)
        print("\n".join(problems), file=sys.stderr)
        print("这通常是因为 pip 从源码编译了扩展。请改装官方二进制 wheel 后重新构建，例如：\n"
              f"  {sys.executable} -m pip install --only-binary=:all: --force-reinstall "
              '"cryptography>=43,<51"', file=sys.stderr)
        return 1
    print("原生扩展依赖检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
