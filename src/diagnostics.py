"""诊断日志与审计事件。"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from .app_paths import get_log_file_path

LOGGER_NAME = "todocx"
_configured = False


def _configure_logging() -> None:
    global _configured
    if _configured:
        return

    log_path = get_log_file_path()
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        formatter = logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    _configured = True


def get_logger(name: str | None = None) -> logging.Logger:
    """获取应用日志器。"""

    _configure_logging()
    if not name:
        return logging.getLogger(LOGGER_NAME)
    return logging.getLogger(f"{LOGGER_NAME}.{name}")


def get_log_path() -> Path:
    """获取日志文件路径。"""

    _configure_logging()
    return get_log_file_path()


def _serialize_context(context: dict[str, Any]) -> str:
    if not context:
        return "{}"
    return json.dumps(context, ensure_ascii=False, default=str, sort_keys=True)


def log_event(logger: logging.Logger, message: str, **context: Any) -> None:
    """记录普通审计事件。"""

    logger.info("%s | context=%s", message, _serialize_context(context))


def log_exception(
    logger: logging.Logger,
    message: str,
    error: Exception,
    **context: Any,
) -> None:
    """记录带堆栈的异常。"""

    logger.exception(
        "%s | error=%s | context=%s",
        message,
        getattr(error, "code", error.__class__.__name__),
        _serialize_context(context),
    )
