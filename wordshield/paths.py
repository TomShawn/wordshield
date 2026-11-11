"""项目只读目录与外置开发目录。业务文档不进入开发缓存。"""
from __future__ import annotations

from hashlib import sha256
import os
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def protected_directories() -> tuple[Path, ...]:
    if not getattr(sys, 'frozen', False):
        return (PROJECT_ROOT,)
    executable = Path(sys.executable).resolve()
    bundle = next((parent for parent in executable.parents if parent.suffix == '.app'), None)
    installation = bundle if bundle is not None else executable.parent
    # 从源码目录的构建结果启动时，也保护整个源码目录。
    for parent in (installation, *installation.parents):
        if (parent / 'pyproject.toml').is_file() and (parent / 'wordshield').is_dir():
            return (parent,)
    return (installation,)


def require_external(path: Path) -> Path:
    resolved = Path(path).expanduser().resolve()
    in_checkout = any((parent / 'pyproject.toml').is_file() and (parent / 'wordshield').is_dir()
                      for parent in (resolved, *resolved.parents))
    if in_checkout or any(resolved.is_relative_to(root.resolve()) for root in protected_directories()):
        raise ValueError('请选择项目目录和应用包之外的文件夹，以保持程序目录干净。')
    return resolved


def development_directory() -> Path:
    checkout = sha256(str(PROJECT_ROOT).encode('utf-8')).hexdigest()[:12]
    default = Path.home() / 'Library' / 'Caches' / 'WordShield-Development' / checkout
    return require_external(Path(os.environ.get('WORDSHIELD_DEV_HOME', str(default))))
