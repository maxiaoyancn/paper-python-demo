from __future__ import annotations

from itertools import combinations
from math import isclose

import numpy as np
import pytest
from scipy import stats as sp_stats

from scripts.generate_excel_random_data import (
    LEVENE_ALPHA,
    compute_levene,
    compute_overall,
    compute_pairwise,
    compute_shapiro_min,
)


def _eq_var_data(rng):
    return [
        rng.normal(10.0, 2.0, 30),
        rng.normal(11.0, 2.0, 30),
        rng.normal(12.0, 2.0, 30),
    ]


def _diff_var_data(rng):
    return [
        rng.normal(10.0, 1.0, 30),
        rng.normal(10.0, 10.0, 30),
        rng.normal(10.0, 5.0, 30),
    ]


def test_compute_levene_equal_var():
    rng = np.random.default_rng(0)
    groups = _eq_var_data(rng)
    p, equal_var = compute_levene(groups)
    expected = float(sp_stats.levene(*groups, center="median").pvalue)
    assert isclose(p, expected)
    assert equal_var == (expected > LEVENE_ALPHA)
    assert equal_var is True


def test_compute_levene_unequal_var():
    rng = np.random.default_rng(1)
    groups = _diff_var_data(rng)
    _, equal_var = compute_levene(groups)
    assert equal_var is False


def test_compute_shapiro_min_normal_groups():
    rng = np.random.default_rng(1)  # verified seed: all 3 SW p > 0.05
    groups = [rng.normal(0.0, 1.0, 50) for _ in range(3)]
    p, all_normal = compute_shapiro_min(groups)
    assert all_normal is True
    assert p > 0.05


def test_compute_shapiro_min_with_skewed_group():
    rng = np.random.default_rng(0)
    groups = [rng.normal(0.0, 1.0, 50), rng.exponential(1.0, 50)]
    p, all_normal = compute_shapiro_min(groups)
    assert all_normal is False
    assert p <= 0.05


def test_compute_shapiro_min_handles_n_lt_3(caplog):
    g_small = np.array([1.0, 2.0])
    g_ok = np.random.default_rng(0).normal(0, 1, 20)
    with caplog.at_level("WARNING"):
        p, all_normal = compute_shapiro_min([g_small, g_ok])
    assert p == 0.0
    assert all_normal is False


def test_compute_overall_anova_branch():
    rng = np.random.default_rng(0)
    groups = _eq_var_data(rng)
    p, label = compute_overall(groups, equal_var=True)
    expected = float(sp_stats.f_oneway(*groups).pvalue)
    assert isclose(p, expected)
    assert label == "ANOVA"


def test_compute_overall_kw_branch():
    rng = np.random.default_rng(1)
    groups = _diff_var_data(rng)
    p, label = compute_overall(groups, equal_var=False)
    expected = float(sp_stats.kruskal(*groups).pvalue)
    assert isclose(p, expected)
    assert label == "KW"


def test_compute_pairwise_tukey_branch():
    rng = np.random.default_rng(0)
    groups = _eq_var_data(rng)
    raw, q, label = compute_pairwise(groups, equal_var=True, all_normal=True)
    assert label == "Tukey"
    assert len(raw) == 3
    assert q == [None, None, None]
    expected_matrix = sp_stats.tukey_hsd(*groups).pvalue
    expected_p = [float(expected_matrix[i][j]) for i, j in combinations(range(3), 2)]
    for got, exp in zip(raw, expected_p):
        assert isclose(got, exp)


def test_compute_pairwise_welch_bonferroni_branch():
    rng = np.random.default_rng(1)
    groups = _diff_var_data(rng)
    raw, q, label = compute_pairwise(groups, equal_var=False, all_normal=True)
    assert label == "Welch+Bonferroni"
    n_pairs = 3
    assert len(raw) == n_pairs
    assert all(qv is not None for qv in q)
    for r, qv in zip(raw, q):
        assert qv == pytest.approx(min(1.0, r * n_pairs))


def test_compute_pairwise_welch_when_not_all_normal():
    rng = np.random.default_rng(2)
    groups = _eq_var_data(rng)
    _, q, label = compute_pairwise(groups, equal_var=True, all_normal=False)
    assert label == "Welch+Bonferroni"
    assert all(qv is not None for qv in q)
