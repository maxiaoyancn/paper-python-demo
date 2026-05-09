from __future__ import annotations

from pathlib import Path

import openpyxl

from scripts.generate_excel_random_data import GroupSpec, RowSpec, read_specs


def _make_xlsx(tmp_path: Path, rows: list[tuple]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["指标", "N", "C均值", "D-SD", "M", "F均值", "G-SD"])
    for r in rows:
        ws.append(list(r))
    path = tmp_path / "in.xlsx"
    wb.save(path)
    return path


def test_read_specs_normal_row(tmp_path):
    path = _make_xlsx(tmp_path, [("体重", 10, 120.6263, 10.3698, 12, 110.36, 9.8635)])
    specs = read_specs(path, sheet=None)
    assert len(specs) == 1
    spec = specs[0]
    assert isinstance(spec, RowSpec)
    assert spec.metric == "体重"
    assert spec.row_index == 2
    assert spec.group1 == GroupSpec(name="第一组", n=10, mean=120.6263, sd=10.3698)
    assert spec.group2 == GroupSpec(name="第二组", n=12, mean=110.36, sd=9.8635)


def test_read_specs_skips_invalid_rows(tmp_path, caplog):
    rows = [
        ("体重", 10, 120.6263, 10.3698, 12, 110.36, 9.8635),  # ok
        ("空 N", None, 100.0, 5.0, 10, 90.0, 5.0),
        ("N=1", 1, 100.0, 5.0, 10, 90.0, 5.0),
        ("SD=0", 10, 100.0, 0.0, 10, 90.0, 5.0),
        ("非数值", 10, "x", 5.0, 10, 90.0, 5.0),
    ]
    path = _make_xlsx(tmp_path, rows)
    with caplog.at_level("WARNING"):
        specs = read_specs(path, sheet=None)
    assert [s.metric for s in specs] == ["体重"]
    assert sum(1 for r in caplog.records if r.levelname == "WARNING") >= 4


def test_read_specs_empty_sheet(tmp_path):
    path = _make_xlsx(tmp_path, [])
    assert read_specs(path, sheet=None) == []
