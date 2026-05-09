from __future__ import annotations

import shutil
from pathlib import Path

import numpy as np
import openpyxl
import pytest
from openpyxl.utils import column_index_from_string
from scipy import stats as sp_stats

from scripts.generate_excel_random_data import LEVENE_ALPHA, main

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_XLSX = ROOT / "20260509-随机数生成.xlsx"


@pytest.fixture
def out_path(tmp_path: Path) -> Path:
    src = tmp_path / "in.xlsx"
    shutil.copy(SAMPLE_XLSX, src)
    return src


def test_cli_writes_expected_layout(out_path: Path):
    rc = main(["--input", str(out_path), "--output", str(out_path), "--seed", "42"])
    assert rc == 0

    wb_in = openpyxl.load_workbook(SAMPLE_XLSX, data_only=True)
    ws_in = wb_in.active
    wb_out = openpyxl.load_workbook(out_path, data_only=True)
    ws_out = wb_out.active

    for r in range(1, ws_in.max_row + 1):
        for c in range(1, 9):
            assert (
                ws_out.cell(row=r, column=c).value == ws_in.cell(row=r, column=c).value
            ), (r, c)

    col_S = column_index_from_string("S")
    col_AF = column_index_from_string("AF")
    col_R = column_index_from_string("R")
    col_AE = column_index_from_string("AE")
    col_AG = column_index_from_string("AG")

    assert ws_out.cell(row=1, column=col_R).value == "10（N）"
    assert ws_out.cell(row=1, column=col_AE).value == "12（M）"

    for r in range(2, 5):
        assert ws_out.cell(row=r, column=col_S).value is None
        assert ws_out.cell(row=r, column=col_AF).value is None

        g1 = [
            ws_out.cell(row=r, column=column_index_from_string(L)).value
            for L in ["I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"]
        ]
        g2 = [
            ws_out.cell(row=r, column=column_index_from_string(L)).value
            for L in ["T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"]
        ]
        assert all(isinstance(v, float) for v in g1), g1
        assert all(isinstance(v, float) for v in g2), g2
        assert all(round(v, 4) == v for v in g1)
        assert all(round(v, 4) == v for v in g2)
        assert len(g1) == 10
        assert len(g2) == 12

        a1 = np.asarray(g1)
        a2 = np.asarray(g2)
        levene_p = sp_stats.levene(a1, a2, center="median").pvalue
        equal_var = levene_p >= LEVENE_ALPHA
        expected = (
            float(np.mean(a1)),
            float(np.std(a1, ddof=1)),
            float(np.mean(a2)),
            float(np.std(a2, ddof=1)),
            float(sp_stats.ttest_ind(a1, a2, equal_var=equal_var).pvalue),
        )
        actual = tuple(ws_out.cell(row=r, column=col_AG + i).value for i in range(5))
        for got, exp in zip(actual, expected, strict=True):
            assert got == pytest.approx(exp, rel=1e-9, abs=1e-9), (got, exp)


def test_cli_requires_input_argument(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 2
    err = capsys.readouterr().err
    assert "--input" in err


def test_cli_seed_reproducible(out_path: Path, tmp_path: Path):
    out2 = tmp_path / "second.xlsx"
    shutil.copy(SAMPLE_XLSX, out2)

    main(["--input", str(out_path), "--output", str(out_path), "--seed", "42"])
    main(["--input", str(out2), "--output", str(out2), "--seed", "42"])

    wb1 = openpyxl.load_workbook(out_path, data_only=True).active
    wb2 = openpyxl.load_workbook(out2, data_only=True).active
    for r in range(1, wb1.max_row + 1):
        for c in range(1, wb1.max_column + 1):
            assert wb1.cell(row=r, column=c).value == wb2.cell(row=r, column=c).value, (
                r,
                c,
            )
