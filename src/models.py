from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class CellIssue:
    row: int
    column: int
    message: str
    kind: str

    @property
    def position(self) -> tuple[int, int]:
        return self.row, self.column


@dataclass
class WorkbookDocument:
    source_path: Path
    cache_path: Path | None = None
    document_type: str = "class"
    sheet_name: str = ""
    values: list[list[Any]] = field(default_factory=list)
    issues: list[CellIssue] = field(default_factory=list)
    header_row: int = 1
    columns: dict[str, int] = field(default_factory=dict)
    load_error: str | None = None

    @property
    def filename(self) -> str:
        return self.source_path.name

    @property
    def has_issues(self) -> bool:
        return bool(self.issues)

    @property
    def issue_positions(self) -> set[tuple[int, int]]:
        return {issue.position for issue in self.issues if issue.row > 0 and issue.column > 0}

    def cell(self, row: int, column: int) -> Any:
        if row < 1 or column < 1:
            return None
        try:
            return self.values[row - 1][column - 1]
        except IndexError:
            return None

    def set_cell(self, row: int, column: int, value: Any) -> None:
        while len(self.values) < row:
            self.values.append([])
        while len(self.values[row - 1]) < column:
            self.values[row - 1].append(None)
        self.values[row - 1][column - 1] = value


@dataclass(frozen=True)
class ExpectedTotals:
    evening: float
    noon: float
    afternoon: float
    weekend_monitor: float = 0.0

    def as_dict(self) -> dict[str, float]:
        return {
            "晚自修": self.evening,
            "高中中午": self.noon,
            "全程下午": self.afternoon,
            "周六日监考": self.weekend_monitor,
        }


@dataclass(frozen=True)
class CalculationResult:
    workbook_path: Path
    sheet_name: str
    values: list[list[Any]]
    total_amount: float
    teacher_count: int
    source_file_count: int
    supplemental_file_count: int = 0
