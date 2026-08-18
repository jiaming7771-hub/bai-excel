from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import QObject

from utils.workers import JobFn, JobThread


class TaskHost(QObject):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._thread: JobThread | None = None

    def running(self) -> bool:
        return bool(self._thread and self._thread.isRunning())

    def start(
        self,
        job: JobFn,
        on_ok: Callable,
        on_fail: Callable,
        on_progress: Callable,
    ) -> None:
        if self.running():
            return
        thread = JobThread(job, self)
        self._thread = thread
        thread.progress.connect(on_progress)
        thread.succeeded.connect(on_ok)
        thread.failed.connect(on_fail)
        thread.finished.connect(self._clear)
        thread.start()

    def cancel(self) -> None:
        if self._thread and self._thread.isRunning():
            self._thread.request_cancel()

    def _clear(self) -> None:
        self._thread = None
