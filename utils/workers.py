from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import QThread, Signal

from utils.error_handler import AppError, friendly_error

ProgressCb = Callable[[int, str], None]
JobFn = Callable[[ProgressCb], object]


class JobThread(QThread):
    progress = Signal(int, str)
    succeeded = Signal(object)
    failed = Signal(str, str)

    def __init__(self, job: JobFn, parent=None) -> None:
        super().__init__(parent)
        self._job = job

    def run(self) -> None:
        def callback(value: int, text: str = "正在处理……") -> None:
            self.progress.emit(max(0, min(100, int(value))), text)

        try:
            result = self._job(callback)
            self.succeeded.emit(result)
        except AppError as exc:
            self.failed.emit(exc.title, exc.hint)
        except Exception as exc:
            err = friendly_error(exc)
            self.failed.emit(err.title, err.hint)
