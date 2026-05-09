from __future__ import annotations

import numpy as np
import pytest

from scripts.generate_excel_random_data import (
    SD_TOLERANCE,
    generate_one_group,
    generate_with_retry,
)


def test_generate_one_group_decimals_4():
    rng = np.random.default_rng(42)
    x = generate_one_group(
        target_mean=100.0, target_sd=15.0, size=200, rng=rng, decimals=4
    )
    assert np.all(np.isclose(x * 10000, np.round(x * 10000)))
    assert abs(np.std(x, ddof=1) - 15.0) / 15.0 <= SD_TOLERANCE


def test_generate_one_group_decimals_0_yields_integers():
    rng = np.random.default_rng(0)
    x = generate_one_group(
        target_mean=10.0, target_sd=3.0, size=200, rng=rng, decimals=0
    )
    assert np.all(x == np.round(x))


def test_generate_one_group_decimals_6():
    rng = np.random.default_rng(0)
    x = generate_one_group(
        target_mean=1.0, target_sd=0.1, size=200, rng=rng, decimals=6
    )
    assert np.all(np.isclose(x * 10**6, np.round(x * 10**6)))


def test_generate_with_retry_passes_decimals():
    rng = np.random.default_rng(0)
    x = generate_with_retry(
        metric="t", target_mean=5.0, target_sd=2.0, size=10, rng=rng, decimals=2
    )
    assert np.all(np.isclose(x * 100, np.round(x * 100)))


def test_generate_with_retry_rejects_zero_sd():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(
            metric="bad", target_mean=1.0, target_sd=0.0, size=10, rng=rng, decimals=4
        )


def test_generate_with_retry_rejects_size_lt_2():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(
            metric="bad", target_mean=1.0, target_sd=1.0, size=1, rng=rng, decimals=4
        )
