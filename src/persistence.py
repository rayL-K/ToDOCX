"""共享的 JSON 持久化辅助。"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any


def atomic_write_json(path: Path, data: Any) -> None:
    """原子写入 JSON 文件。"""

    path.parent.mkdir(parents=True, exist_ok=True)
    file_handle = None
    temp_path = None

    try:
        file_handle = tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            delete=False,
            suffix=".tmp",
        )
        temp_path = Path(file_handle.name)
        json.dump(data, file_handle, ensure_ascii=False, indent=2)
        file_handle.flush()
        os.fsync(file_handle.fileno())
        file_handle.close()
        file_handle = None
        os.replace(temp_path, path)
    finally:
        if file_handle is not None:
            file_handle.close()
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
