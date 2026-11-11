"""把当前虚拟环境中的原生扩展换成 universal2 版本，供构建同时支持 Intel 与 Apple Silicon 的应用。

优先下载官方 universal2 wheel；某个包只发布了单架构 wheel 时（例如 cffi），分别下载
x86_64 与 arm64 的官方 wheel，用 delocate-merge 合并后安装。所有 wheel 都来自 PyPI 二进制包，
不从源码编译。
"""
from __future__ import annotations

from importlib import metadata
from pathlib import Path
import subprocess
import sys
import tempfile

REQUIRED = {"x86_64", "arm64"}
# 只在构建时使用、不会打进应用的工具包；delocate 自带单架构测试数据。
BUILD_TOOLS = {"delocate"}
# pip 会把 macosx_11_0 自动扩展到更早的兼容版本标签。
PLATFORMS = {"universal2": "macosx_11_0_universal2",
             "x86_64": "macosx_11_0_x86_64", "arm64": "macosx_11_0_arm64"}


def archs(binary: Path) -> set[str]:
    output = subprocess.run(["/usr/bin/lipo", "-archs", str(binary)],
                            capture_output=True, text=True, check=True).stdout
    return set(output.split())


def thin_distributions() -> dict[str, tuple[str, list[str]]]:
    """返回含有单架构原生文件的发行包：名称 -> (版本, 文件列表)。"""
    found = {}
    for dist in metadata.distributions():
        if dist.metadata["Name"].lower() in BUILD_TOOLS:
            continue
        for file in dist.files or ():
            if file.suffix in (".so", ".dylib"):
                path = Path(dist.locate_file(file))
                if path.exists() and not REQUIRED <= archs(path):
                    found.setdefault(dist.metadata["Name"], (dist.version, []))[1].append(str(file))
    return found


def pip(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, "-B", "-m", "pip", *args], capture_output=True, text=True)


def download(name: str, version: str, platform: str, target: Path) -> Path | None:
    target.mkdir(parents=True, exist_ok=True)
    python_version = f"{sys.version_info.major}.{sys.version_info.minor}"
    result = pip("download", "--no-deps", "--only-binary=:all:", "--platform", platform,
                 "--python-version", python_version, "--dest", str(target), f"{name}=={version}")
    wheels = sorted(target.glob("*.whl"))
    return wheels[0] if result.returncode == 0 and wheels else None


def universal_wheel(name: str, version: str, work: Path) -> Path:
    wheel = download(name, version, PLATFORMS["universal2"], work / "universal2")
    if wheel:
        return wheel
    thin = [download(name, version, PLATFORMS[arch], work / arch) for arch in ("x86_64", "arm64")]
    if not all(thin):
        raise SystemExit(f"{name} {version} 没有可用的 x86_64 与 arm64 二进制 wheel，无法构建 universal2。")
    merged = work / "merged"
    tool = Path(sys.executable).parent / "delocate-merge"
    subprocess.run([str(tool), "-w", str(merged), *map(str, thin)], check=True)
    return next(merged.glob("*.whl"))


def main() -> int:
    if not REQUIRED <= archs(Path(sys.executable).resolve()):
        print("当前 Python 不是 universal2 构建。请使用 python.org 的 macOS 安装包创建虚拟环境。", file=sys.stderr)
        return 1
    with tempfile.TemporaryDirectory(prefix="wordshield-universal2-") as temporary:
        for name, (version, _) in sorted(thin_distributions().items()):
            wheel = universal_wheel(name, version, Path(temporary) / name)
            result = pip("install", "--no-deps", "--force-reinstall", "--no-index", str(wheel))
            if result.returncode:
                print(result.stdout + result.stderr, file=sys.stderr)
                return 1
            print(f"{name} {version} -> {wheel.name}")
    remaining = thin_distributions()
    if remaining:
        print("以下原生文件仍缺少 x86_64 或 arm64：", file=sys.stderr)
        for name, (_, files) in remaining.items():
            print(f"  {name}: " + ", ".join(files), file=sys.stderr)
        return 1
    print("全部原生扩展均为 universal2")
    return 0


if __name__ == "__main__":
    sys.exit(main())
