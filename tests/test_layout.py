from __future__ import annotations

from openpyxl.utils import column_index_from_string

from scripts.generate_excel_random_data import compute_layout


def _col(letter: str) -> int:
    return column_index_from_string(letter)


def test_layout_k2_decimals_at_h():
    layout = compute_layout(decimals_col=8, group_sizes=(10, 12))
    assert list(layout.group_cols[0]) == list(range(_col("J"), _col("S") + 1))
    assert list(layout.group_cols[1]) == list(range(_col("U"), _col("AF") + 1))
    stats = layout.stat_cols
    assert stats.mean_sd_pairs[0] == (_col("AH"), _col("AI"))
    assert stats.mean_sd_pairs[1] == (_col("AJ"), _col("AK"))
    assert stats.levene == _col("AL")
    assert stats.shapiro_min == _col("AM")
    assert stats.overall == _col("AN")
    assert stats.pairwise_raw == (_col("AO"),)
    assert stats.pairwise_q == (_col("AP"),)


def test_layout_k3_decimals_at_k():
    layout = compute_layout(decimals_col=11, group_sizes=(8, 10, 12))
    assert layout.group_cols[0].start == _col("M")
    assert layout.group_cols[0].stop == _col("M") + 8
    assert layout.group_cols[1].start == _col("V")
    assert layout.group_cols[1].stop == _col("V") + 10
    assert layout.group_cols[2].start == _col("AG")
    assert layout.group_cols[2].stop == _col("AG") + 12
    stats = layout.stat_cols
    assert stats.mean_sd_pairs[0][0] == _col("AT")
    assert stats.levene == 52
    assert stats.shapiro_min == 53
    assert stats.overall == 54
    assert stats.pairwise_raw == (55, 56, 57)
    assert stats.pairwise_q == (58, 59, 60)


def test_layout_k4_total_stat_cols():
    layout = compute_layout(decimals_col=14, group_sizes=(6, 8, 10, 12))
    stats = layout.stat_cols
    n_stat = 2 * 4 + 3 + 2 * 6
    flat = (
        [c for pair in stats.mean_sd_pairs for c in pair]
        + [stats.levene, stats.shapiro_min, stats.overall]
        + list(stats.pairwise_raw)
        + list(stats.pairwise_q)
    )
    assert len(flat) == n_stat
    assert flat == list(range(flat[0], flat[0] + n_stat))
