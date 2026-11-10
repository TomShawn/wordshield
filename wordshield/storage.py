"""排他创建输出，权限 0600。不覆盖源文件，不写明文中间文件。"""
from __future__ import annotations

import os
from pathlib import Path

from .paths import require_external


class OutputError(ValueError):
    pass


def write_outputs(outputs: dict[Path, bytes], inputs: tuple[Path, ...] = ()) -> None:
    paths = [Path(p).absolute() for p in outputs]
    if len({p.resolve() for p in paths}) != len(paths):
        raise OutputError("输出文件路径不能重复。")
    protected = {Path(p).resolve() for p in inputs}
    for path in paths:
        try:
            require_external(path)
        except ValueError as exc:
            raise OutputError(str(exc)) from None
        if path.resolve() in protected or path.exists() or path.is_symlink():
            raise OutputError("输出路径已存在或与输入文件相同，请选择新文件名。")
        if not path.parent.is_dir():
            raise OutputError("输出文件夹不存在。")
    created: list[Path] = []
    try:
        for target, data in outputs.items():
            target = Path(target).absolute()
            fd = os.open(target, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            created.append(target)
            with os.fdopen(fd, "wb") as stream:
                stream.write(data)
                stream.flush()
                os.fsync(stream.fileno())
    except BaseException:
        for path in created:
            path.unlink(missing_ok=True)
        raise
