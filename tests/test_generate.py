from __future__ import annotations

import numpy as np
import pytest

from scripts.generate_excel_random_data import (
    SD_TOLERANCE,
    generate_one_group,
    generate_with_retry,
)


def test_generate_one_group_matches_target_within_rounding():
    rng = np.random.default_rng(42)
    x = generate_one_group(target_mean=100.0, target_sd=15.0, size=200, rng=rng)
    assert x.shape == (200,)
    assert np.all(np.isclose(x * 10000, np.round(x * 10000)))
    assert abs(np.mean(x) - 100.0) < 1e-2
    assert abs(np.std(x, ddof=1) - 15.0) / 15.0 <= SD_TOLERANCE


def test_generate_with_retry_small_sample_within_tolerance():
    rng = np.random.default_rng(0)
    for _ in range(5):
        sub_rng = np.random.default_rng(rng.integers(0, 2**32 - 1))
        x = generate_with_retry(
            metric="t", target_mean=5.0, target_sd=2.0, size=2, rng=sub_rng
        )
        assert x.shape == (2,)
        assert abs(np.std(x, ddof=1) - 2.0) / 2.0 <= SD_TOLERANCE


def test_generate_with_retry_rejects_zero_sd():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(
            metric="bad", target_mean=1.0, target_sd=0.0, size=10, rng=rng
        )


def test_generate_with_retry_rejects_size_lt_2():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(
            metric="bad", target_mean=1.0, target_sd=1.0, size=1, rng=rng
        )
