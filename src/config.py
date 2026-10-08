from __future__ import annotations

import json
import math
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class AppPaths:
    """项目中固定文件的位置。"""

    root: Path
    template: Path
    salary: Path
    cache_root: Path

    @classmethod
    def from_root(cls, root: Path | str) -> "AppPaths":
        root_path = Path(root).resolve()
        return cls(
            root=root_path,
            template=root_path / "docs" / "module.xlsx",
            salary=root_path / "salary.json",
            cache_root=root_path / "__cntcache__",
        )


@dataclass
class SalaryRates:
    grade_evening: dict[str, float] = field(
        default_factory=lambda: {"1": 60.0, "2": 65.0, "3": 70.0}
    )
    evening_reading: float = 7.0
    noon: float = 25.0
    afternoon: float = 7.0

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "SalaryRates":
        defaults = cls()
        grade_raw = raw.get("grade_evening", {})
        grades = {
            grade: _non_negative_number(grade_raw.get(grade, default), f"G{grade}晚")
            for grade, default in defaults.grade_evening.items()
        }
        return cls(
            grade_evening=grades,
            evening_reading=_non_negative_number(
                raw.get("evening_reading", defaults.evening_reading), "晚读"
            ),
            noon=_non_negative_number(raw.get("noon", defaults.noon), "高中中午"),
            afternoon=_non_negative_number(
                raw.get("afternoon", defaults.afternoon), "全程下午"
            ),
        )

    def as_json_dict(self) -> dict[str, Any]:
        return asdict(self)


def _non_negative_number(value: Any, label: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label}工资必须是数字") from exc
    if not math.isfinite(number) or number < 0:
        raise ValueError(f"{label}工资必须是有限的非负数字")
    return number


class SalaryConfig:
    def __init__(self, path: Path):
        self.path = path

    def load(self) -> SalaryRates:
        if not self.path.exists() or not self.path.read_text(encoding="utf-8").strip():
            rates = SalaryRates()
            self.save(rates)
            return rates
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(raw, dict):
                raise ValueError("工资配置必须是 JSON 对象")
            return SalaryRates.from_dict(raw)
        except (json.JSONDecodeError, ValueError) as exc:
            raise ValueError(f"工资配置无效：{exc}") from exc

    def save(self, rates: SalaryRates) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(
            json.dumps(rates.as_json_dict(), ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temporary, self.path)
