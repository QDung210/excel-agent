"""Typed errors mapped to the HTTP error contract (spec 05/08)."""
from __future__ import annotations

from typing import Any


class WorkspaceError(Exception):
    status = 400
    code = "BAD_REQUEST"

    def __init__(self, message: str, **extra: Any):
        super().__init__(message)
        self.message = message
        self.extra = extra

    def body(self) -> dict[str, Any]:
        return {"code": self.code, "message": self.message, **self.extra}


class NotFound(WorkspaceError):
    status, code = 404, "NOT_FOUND"


class UnsupportedFile(WorkspaceError):
    status, code = 400, "UNSUPPORTED_FILE"


class FileTooLarge(WorkspaceError):
    status, code = 413, "FILE_TOO_LARGE"


class InvalidOperation(WorkspaceError):
    status, code = 422, "INVALID_OPERATION"


class RevisionConflict(WorkspaceError):
    status, code = 409, "REVISION_CONFLICT"


class ValidationFailed(WorkspaceError):
    status, code = 422, "VALIDATION_FAILED"


class ModelUnavailable(WorkspaceError):
    status, code = 503, "MODEL_UNAVAILABLE"
