"""统一的应用异常类型。"""

from __future__ import annotations


class ToDOCXError(Exception):
    """带错误码和用户提示的基础异常。"""

    default_code = "TODX000"

    def __init__(
        self,
        user_message: str,
        *,
        code: str | None = None,
        hint: str | None = None,
        details: str | None = None,
    ) -> None:
        super().__init__(user_message)
        self.code = code or self.default_code
        self.user_message = user_message
        self.hint = hint
        self.details = details

    def format_user_message(self) -> str:
        if self.hint:
            return f"{self.user_message}\n\n建议：{self.hint}"
        return self.user_message

    def __str__(self) -> str:
        return self.format_user_message()


class ValidationError(ToDOCXError):
    default_code = "TODX100"


class ResourceAccessError(ToDOCXError):
    default_code = "TODX200"


class TemplateStorageError(ToDOCXError):
    default_code = "TODX300"


class AnalysisError(ToDOCXError):
    default_code = "TODX400"


class OperationCancelledError(ToDOCXError):
    default_code = "TODX900"


def user_message_for_error(error: Exception) -> str:
    """把异常转换为给用户看的消息。"""

    if isinstance(error, ToDOCXError):
        return error.format_user_message()
    return "操作失败，请重试；如果问题仍然存在，请查看日志文件。"
