from __future__ import annotations

import numpy as np
from scipy import stats as sp_stats

from scripts.generate_excel_random_data import LEVENE_ALPHA, compute_stats


def test_compute_stats_equal_var_branch():
    rng = np.random.default_rng(0)
    g1 = rng.normal(10.0, 2.0, 50)
    g2 = rng.normal(11.0, 2.0, 50)
    mean1, sd1, mean2, sd2, p, equal_var = compute_stats(g1, g2)

    assert mean1 == float(np.mean(g1))
    assert sd1 == float(np.std(g1, ddof=1))
    assert mean2 == float(np.mean(g2))
    assert sd2 == float(np.std(g2, ddof=1))

    levene_p = sp_stats.levene(g1, g2, center="median").pvalue
    expected_equal_var = bool(levene_p >= LEVENE_ALPHA)
    assert equal_var == expected_equal_var
    expected_p = sp_stats.ttest_ind(g1, g2, equal_var=expected_equal_var).pvalue
    assert p == float(expected_p)


def test_compute_stats_welch_branch():
    rng = np.random.default_rng(1)
    g1 = rng.normal(10.0, 1.0, 30)
    g2 = rng.normal(10.0, 10.0, 30)
    _, _, _, _, p, equal_var = compute_stats(g1, g2)

    levene_p = sp_stats.levene(g1, g2, center="median").pvalue
    assert equal_var == bool(levene_p >= LEVENE_ALPHA)
    expected_p = sp_stats.ttest_ind(g1, g2, equal_var=equal_var).pvalue
    assert p == float(expected_p)
    assert equal_var is False
