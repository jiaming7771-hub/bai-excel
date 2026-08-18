from __future__ import annotations

from pathlib import Path

from pandas.errors import EmptyDataError, ParserError


class AppError(Exception):
    def __init__(self, title: str, hint: str = "") -> None:
        super().__init__(title)
        self.title = title
        self.hint = hint


def friendly_error(exc: BaseException) -> AppError:
    if isinstance(exc, AppError):
        return exc

    text = str(exc).lower()
    name = type(exc).__name__.lower()

    if isinstance(exc, PermissionError) or "permission" in text or "being used" in text:
        return AppError(
            "Excel 文件读取失败",
            "请确认文件没有被其他程序占用，并且您有权限访问这个文件。",
        )
    if isinstance(exc, FileNotFoundError):
        return AppError("找不到文件", "请重新选择 Excel 文件后再试。")
    if isinstance(exc, EmptyDataError):
        return AppError("Excel 文件是空的", "请确认表格里至少有一行表头和数据。")
    if isinstance(exc, ParserError) or "excel" in text or "workbook" in text:
        return AppError(
            "Excel 文件读取失败",
            "请确认文件没有被其他程序占用，并且文件格式正确。",
        )
    if "no columns to parse" in text:
        return AppError("表格没有可用的列", "请确认第一行是表头，并且文件不是空表。")
    if "illegal" in text or "invalid" in name:
        return AppError("文件无法处理", "请确认这是一个正常的 Excel 文件，而不是已损坏的文件。")
    return AppError(
        "处理失败",
        "请确认文件格式正确，并且没有被其他程序占用。如果问题一直出现，请换一份 Excel 再试。",
    )
