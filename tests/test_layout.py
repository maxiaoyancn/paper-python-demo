from __future__ import annotations

from openpyxl.utils import column_index_from_string

from scripts.generate_excel_random_data import compute_layout


def _col(letter: str) -> int:
    return column_index_from_string(letter)


def test_layout_matches_sample_when_n10_m12():
    layout = compute_layout(n=10, m=12)
    assert list(layout.group1_cols) == list(range(_col("I"), _col("R") + 1))
    assert list(layout.group2_cols) == list(range(_col("T"), _col("AE") + 1))
    assert list(layout.stat_cols) == list(range(_col("AG"), _col("AK") + 1))


def test_layout_other_sizes():
    layout = compute_layout(n=5, m=7)
    assert list(layout.group1_cols) == [9, 10, 11, 12, 13]
    assert list(layout.group2_cols) == [15, 16, 17, 18, 19, 20, 21]
    assert list(layout.stat_cols) == [23, 24, 25, 26, 27]
