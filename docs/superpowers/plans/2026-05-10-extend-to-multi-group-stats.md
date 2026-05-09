# Extend to Multi-Group Stats Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 `excel-random-data-generator` 从 2 组扩展为通用 K 组（K≥2）+ 行级 decimals + 完整统计流水线（Levene → Shapiro-Wilk → ANOVA/Kruskal-Wallis → Tukey HSD/Welch+Bonferroni）。

**Architecture:** 单文件 `scripts/generate_excel_random_data.py` 内分五层：read（扫描行 1 表头定位"位数列"，按 col(2+3i) 解析 K 组）→ generate（按行 decimals round）→ layout（动态计算 K 组数据列 + 2K+3+2C(K,2) 个统计列）→ stats（拆出 compute_levene/compute_shapiro_min/compute_overall/compute_pairwise 4 个新函数）→ write（按新布局写入数据 + 统计列）。CLI 与可复现性逻辑保留。

**Tech Stack:** Python 3.10+、`openpyxl`、`numpy`、`scipy>=1.11`（含 `tukey_hsd`）、`pytest`、`ruff`。

**OpenSpec 文档**：
- 提案 [openspec/changes/extend-to-multi-group-stats/proposal.md](../../../openspec/changes/extend-to-multi-group-stats/proposal.md)
- 设计 [openspec/changes/extend-to-multi-group-stats/design.md](../../../openspec/changes/extend-to-multi-group-stats/design.md)
- 规范 [openspec/changes/extend-to-multi-group-stats/specs/excel-random-data-generator/spec.md](../../../openspec/changes/extend-to-multi-group-stats/specs/excel-random-data-generator/spec.md)
- 任务 [openspec/changes/extend-to-multi-group-stats/tasks.md](../../../openspec/changes/extend-to-multi-group-stats/tasks.md)

---

## File Structure

| 路径 | 责任 |
|---|---|
| `scripts/generate_excel_random_data.py` | 整体重写：dataclass / read / generate / layout / 4 stats / write / CLI |
| `tests/fixtures/sample-k2.xlsx` | K=2 黄金 fixture（拷自当前 `20260509-随机数生成.xlsx`） |
| `tests/fixtures/build_fixtures.py` | 现场构造 K=3、K=4 fixture 的 helper（用 openpyxl 生成测试用 xlsx） |
| `tests/test_read_specs.py` | 重写：覆盖 K=2/K=3/无位数列/非法行 |
| `tests/test_generate.py` | 重写：覆盖 decimals=0/4/6 + ±10% 容差 + 重试上限 |
| `tests/test_stats.py` | 重写：4 个统计函数各自至少 2 case + 分支选择 |
| `tests/test_layout.py` | 重写：K=2/K=3/K=4 列号断言 |
| `tests/test_cli.py` | 重写：K=2/K=3 端到端 + Tukey/Welch 分支 + seed 复现 + 必填 input |
| `scripts/README.md` | 重写列布局段（K 组动态，位数列定位） |

---

## Task 1: 准备 fixtures 与目录骨架

**Files:**
- Create: `tests/fixtures/__init__.py`
- Create: `tests/fixtures/build_fixtures.py`
- Copy: `20260509-随机数生成.xlsx` → `tests/fixtures/sample-k2.xlsx`

- [ ] **Step 1.1: 拷贝 K=2 黄金 fixture**

```bash
mkdir -p tests/fixtures
cp "20260509-随机数生成.xlsx" tests/fixtures/sample-k2.xlsx
touch tests/fixtures/__init__.py
```

- [ ] **Step 1.2: 写 `tests/fixtures/build_fixtures.py`**

```python
"""Build K=3 / K=4 fixture xlsx files for v2 tests.

Layout reminder (K groups):
    A1=指标; B-D=组1 (N,μ,σ); E-G=组2; H-J=组3 (when K≥3); ... ;
    decimals_col = col(2 + 3K); decimals_col + 1 = remarks (optional);
    decimals_col + 2 onward = data area (脚本会写)
"""

from __future__ import annotations

from pathlib import Path

import openpyxl


def build_k3(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "指标",
        "G1 N", "G1 mean", "G1 SD",
        "G2 N", "G2 mean", "G2 SD",
        "G3 N", "G3 mean", "G3 SD",
        "原始数据小数点后位数",
        "备注（脚本不读）",
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c).value = h
    rows = [
        ("体重", 8, 70.0, 5.0, 10, 75.0, 6.0, 12, 80.0, 7.0, 4),
        ("身高", 8, 165.0, 8.0, 10, 170.0, 9.0, 12, 175.0, 10.0, 2),
    ]
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c).value = v
    wb.save(path)


def build_k4(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "指标",
        "G1 N", "G1 mean", "G1 SD",
        "G2 N", "G2 mean", "G2 SD",
        "G3 N", "G3 mean", "G3 SD",
        "G4 N", "G4 mean", "G4 SD",
        "原始数据小数点后位数",
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c).value = h
    rows = [("血糖", 6, 5.0, 1.0, 8, 6.0, 1.2, 10, 7.0, 1.5, 12, 8.0, 1.8, 3)]
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c).value = v
    wb.save(path)


if __name__ == "__main__":
    here = Path(__file__).parent
    build_k3(here / "sample-k3.xlsx")
    build_k4(here / "sample-k4.xlsx")
    print("ok")
```

- [ ] **Step 1.3: 跑一次生成 fixtures**

```bash
.venv/bin/python tests/fixtures/build_fixtures.py
ls tests/fixtures/
```
Expected: 出现 `sample-k2.xlsx`、`sample-k3.xlsx`、`sample-k4.xlsx`、`build_fixtures.py`、`__init__.py`

---

## Task 2: 重写读取层（Dataclass + read_specs）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（重写 `RowSpec`、新增 `_find_decimals_col`、重写 `read_specs`）
- Replace: `tests/test_read_specs.py`

- [ ] **Step 2.1: 写失败测试**

完整新内容覆盖 `tests/test_read_specs.py`：

```python
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from scripts.generate_excel_random_data import (
    GroupSpec,
    RowSpec,
    _find_decimals_col,
    read_specs,
)

FIXTURES = Path(__file__).parent / "fixtures"


def test_find_decimals_col_k2():
    wb = openpyxl.load_workbook(FIXTURES / "sample-k2.xlsx", data_only=True)
    assert _find_decimals_col(wb.active) == 8


def test_find_decimals_col_k3():
    wb = openpyxl.load_workbook(FIXTURES / "sample-k3.xlsx", data_only=True)
    assert _find_decimals_col(wb.active) == 11


def test_find_decimals_col_missing(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["指标", "N1", "μ1", "σ1", "N2", "μ2", "σ2"])
    p = tmp_path / "no_decimals.xlsx"
    wb.save(p)
    wb2 = openpyxl.load_workbook(p, data_only=True)
    with pytest.raises(LookupError):
        _find_decimals_col(wb2.active)


def test_read_specs_k2():
    specs = read_specs(FIXTURES / "sample-k2.xlsx", sheet=None)
    assert len(specs) == 3  # 体重 / 身高 / 血糖
    spec = specs[0]
    assert isinstance(spec, RowSpec)
    assert spec.metric == "体重"
    assert spec.row_index == 2
    assert spec.decimals == 4
    assert len(spec.groups) == 2
    assert spec.groups[0] == GroupSpec(name="G1", n=10, mean=120.6263, sd=10.3698)
    assert spec.groups[1] == GroupSpec(name="G2", n=12, mean=110.36, sd=9.8635)


def test_read_specs_k3():
    specs = read_specs(FIXTURES / "sample-k3.xlsx", sheet=None)
    assert len(specs) == 2
    spec = specs[0]
    assert spec.metric == "体重"
    assert spec.decimals == 4
    assert len(spec.groups) == 3
    assert spec.groups[0].n == 8
    assert spec.groups[2].mean == 80.0


def test_read_specs_skips_invalid_rows(tmp_path, caplog):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["指标", "N1", "μ1", "σ1", "N2", "μ2", "σ2", "原始数据小数点后位数"])
    ws.append(["ok", 10, 100.0, 5.0, 10, 90.0, 5.0, 4])         # ok
    ws.append(["bad-N", 1, 100.0, 5.0, 10, 90.0, 5.0, 4])       # N=1
    ws.append(["bad-SD", 10, 100.0, 0.0, 10, 90.0, 5.0, 4])     # SD=0
    ws.append(["bad-dec", 10, 100.0, 5.0, 10, 90.0, 5.0, -1])   # decimals<0
    ws.append(["miss-mean", 10, None, 5.0, 10, 90.0, 5.0, 4])   # 缺均值
    p = tmp_path / "mixed.xlsx"
    wb.save(p)

    with caplog.at_level("WARNING"):
        specs = read_specs(p, sheet=None)
    assert [s.metric for s in specs] == ["ok"]
    assert sum(1 for r in caplog.records if r.levelname == "WARNING") >= 4


def test_read_specs_no_decimals_col_raises(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["指标", "N1", "μ1", "σ1", "N2", "μ2", "σ2"])
    ws.append(["x", 10, 1.0, 1.0, 10, 1.0, 1.0])
    p = tmp_path / "no_dec.xlsx"
    wb.save(p)
    with pytest.raises(LookupError):
        read_specs(p, sheet=None)
```

- [ ] **Step 2.2: 跑测试，确认失败**

`.venv/bin/pytest tests/test_read_specs.py -v` Expected: ImportError or fail.

- [ ] **Step 2.3: 重写实现**

把 `scripts/generate_excel_random_data.py` 中 `GroupSpec / RowSpec / read_specs` 替换为如下（保留其他部分暂不动；后续 task 会改 generate / layout / stats / write）：

```python
@dataclass(slots=True, frozen=True)
class GroupSpec:
    name: str
    n: int
    mean: float
    sd: float


@dataclass(slots=True, frozen=True)
class RowSpec:
    row_index: int
    metric: str
    groups: tuple[GroupSpec, ...]
    decimals: int


def _find_decimals_col(ws: Worksheet) -> int:
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if isinstance(v, str) and "位数" in v:
            return c
    raise LookupError("未在第 1 行找到包含'位数'的表头列")


def _coerce_int(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return None


def _coerce_float(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def read_specs(path: Path, sheet: str | None) -> list[RowSpec]:
    wb = load_workbook(path, data_only=True)
    ws: Worksheet = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    decimals_col = _find_decimals_col(ws)
    if (decimals_col - 2) % 3 != 0:
        raise LookupError(
            f"位数列在第 {decimals_col} 列，与 K 组三元组布局不兼容（需满足 (col-2) % 3 == 0）"
        )
    k = (decimals_col - 2) // 3
    if k < 2:
        raise LookupError(f"组数 K={k} < 2，至少需要两组数据")

    out: list[RowSpec] = []
    for row_idx in range(2, ws.max_row + 1):
        metric = ws.cell(row=row_idx, column=1).value
        if metric is None or (isinstance(metric, str) and not metric.strip()):
            # Allow blank rows to be silently skipped
            if all(ws.cell(row=row_idx, column=c).value is None for c in range(1, decimals_col + 1)):
                continue

        problems: list[str] = []
        if not isinstance(metric, str) or not metric.strip():
            problems.append("缺少指标名")

        groups: list[GroupSpec] = []
        for i in range(k):
            base = 2 + 3 * i
            n = _coerce_int(ws.cell(row=row_idx, column=base).value)
            mean = _coerce_float(ws.cell(row=row_idx, column=base + 1).value)
            sd = _coerce_float(ws.cell(row=row_idx, column=base + 2).value)
            if n is None or n <= 1:
                problems.append(f"G{i + 1} 的 N 必须是 ≥2 的整数")
            if mean is None:
                problems.append(f"G{i + 1} 的均值缺失或非数值")
            if sd is None or sd <= 0:
                problems.append(f"G{i + 1} 的 SD 必须是 >0 的数值")
            if n is not None and n > 1 and mean is not None and sd is not None and sd > 0:
                groups.append(GroupSpec(name=f"G{i + 1}", n=n, mean=mean, sd=sd))

        decimals = _coerce_int(ws.cell(row=row_idx, column=decimals_col).value)
        if decimals is None or decimals < 0:
            problems.append("decimals 缺失或为负")

        if problems:
            logger.warning(
                "跳过第 %d 行（%r）：%s", row_idx, metric, "；".join(problems)
            )
            continue

        out.append(
            RowSpec(
                row_index=row_idx,
                metric=metric,  # type: ignore[arg-type]
                groups=tuple(groups),
                decimals=decimals,  # type: ignore[arg-type]
            )
        )
    return out
```

并把模块顶部 `STAT_HEADERS` 改为：

```python
STAT_HEADERS_PER_GROUP = ("第{}组均值", "第{}组SD值")
LEVENE_HEADER = "Levene p"
SHAPIRO_HEADER = "Shapiro-Wilk min p"
OVERALL_HEADER = "整体 p（ANOVA/KW）"
```

注意：保留旧 `STAT_HEADERS` 直到 Task 6 写完后再删（避免中间状态测试集体崩）。

- [ ] **Step 2.4: 跑测试**

`.venv/bin/pytest tests/test_read_specs.py -v` Expected: 全部 PASS。

---

## Task 3: 生成层接收行级 decimals

**Files:**
- Modify: `scripts/generate_excel_random_data.py:generate_one_group, generate_with_retry`
- Replace: `tests/test_generate.py`

- [ ] **Step 3.1: 写失败测试**

```python
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
    x = generate_one_group(target_mean=100.0, target_sd=15.0, size=200, rng=rng, decimals=4)
    assert np.all(np.isclose(x * 10000, np.round(x * 10000)))
    assert abs(np.std(x, ddof=1) - 15.0) / 15.0 <= SD_TOLERANCE


def test_generate_one_group_decimals_0_yields_integers():
    rng = np.random.default_rng(0)
    x = generate_one_group(target_mean=10.0, target_sd=3.0, size=200, rng=rng, decimals=0)
    assert np.all(x == np.round(x))


def test_generate_one_group_decimals_6():
    rng = np.random.default_rng(0)
    x = generate_one_group(target_mean=1.0, target_sd=0.1, size=200, rng=rng, decimals=6)
    assert np.all(np.isclose(x * 10**6, np.round(x * 10**6)))


def test_generate_with_retry_passes_decimals():
    rng = np.random.default_rng(0)
    x = generate_with_retry(metric="t", target_mean=5.0, target_sd=2.0, size=10, rng=rng, decimals=2)
    assert np.all(np.isclose(x * 100, np.round(x * 100)))


def test_generate_with_retry_rejects_zero_sd():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(metric="bad", target_mean=1.0, target_sd=0.0, size=10, rng=rng, decimals=4)


def test_generate_with_retry_rejects_size_lt_2():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(metric="bad", target_mean=1.0, target_sd=1.0, size=1, rng=rng, decimals=4)
```

- [ ] **Step 3.2: 跑失败**

`.venv/bin/pytest tests/test_generate.py -v` Expected: TypeError（缺 decimals 参数）

- [ ] **Step 3.3: 实现**

修改 `scripts/generate_excel_random_data.py`：

```python
def generate_one_group(
    target_mean: float,
    target_sd: float,
    size: int,
    rng: np.random.Generator,
    decimals: int,
) -> np.ndarray:
    if size < 2:
        raise ValueError(f"size 必须 ≥2，收到 {size}")
    if target_sd <= 0:
        raise ValueError(f"target_sd 必须 >0，收到 {target_sd}")
    if decimals < 0:
        raise ValueError(f"decimals 必须 ≥0，收到 {decimals}")
    z = rng.standard_normal(size)
    z = (z - z.mean()) / z.std(ddof=1)
    x = target_mean + target_sd * z
    return np.round(x, decimals)


def generate_with_retry(
    metric: str,
    target_mean: float,
    target_sd: float,
    size: int,
    rng: np.random.Generator,
    decimals: int,
) -> np.ndarray:
    last_actual_sd: float | None = None
    for attempt in range(1, MAX_RETRY + 1):
        seed = int(rng.integers(0, 2**32 - 1))
        sub_rng = np.random.default_rng(seed)
        x = generate_one_group(target_mean, target_sd, size, sub_rng, decimals)
        actual_sd = float(np.std(x, ddof=1))
        last_actual_sd = actual_sd
        rel_err = abs(actual_sd - target_sd) / target_sd
        logger.debug(
            "metric=%s attempt=%d size=%d decimals=%d target_sd=%.6f actual_sd=%.6f rel_err=%.4f",
            metric, attempt, size, decimals, target_sd, actual_sd, rel_err,
        )
        if rel_err <= SD_TOLERANCE:
            return x
    raise RuntimeError(
        f"无法在容差内还原 SD: metric={metric}, target_sd={target_sd}, "
        f"last_actual_sd={last_actual_sd}, retries={MAX_RETRY}"
    )
```

也删除模块顶部 `ROUND_DIGITS = 4`（已经不再用）。

- [ ] **Step 3.4: 跑通过**

`.venv/bin/pytest tests/test_generate.py -v` Expected: 6 个 PASS。

---

## Task 4: 列布局动态化

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（重写 `Layout`、`compute_layout`）
- Replace: `tests/test_layout.py`

- [ ] **Step 4.1: 写失败测试**

```python
from __future__ import annotations

from openpyxl.utils import column_index_from_string

from scripts.generate_excel_random_data import compute_layout


def _col(letter: str) -> int:
    return column_index_from_string(letter)


def test_layout_k2_decimals_at_h():
    layout = compute_layout(decimals_col=8, group_sizes=(10, 12))
    # data starts at H+2 = 10 (J)
    assert list(layout.group_cols[0]) == list(range(_col("J"), _col("S") + 1))   # J-S
    assert list(layout.group_cols[1]) == list(range(_col("U"), _col("AF") + 1))  # U-AF
    # stats start after data + 1 blank
    # data ends at AF=32; +2 = AH=34
    stats = layout.stat_cols
    assert stats.mean_sd_pairs[0] == (_col("AH"), _col("AI"))   # μ1, σ1
    assert stats.mean_sd_pairs[1] == (_col("AJ"), _col("AK"))   # μ2, σ2
    assert stats.levene == _col("AL")
    assert stats.shapiro_min == _col("AM")
    assert stats.overall == _col("AN")
    assert stats.pairwise_raw == (_col("AO"),)                   # 1 pair (1,2)
    assert stats.pairwise_q == (_col("AP"),)


def test_layout_k3_decimals_at_k():
    layout = compute_layout(decimals_col=11, group_sizes=(8, 10, 12))
    # data starts at K+2 = 13 (M)
    assert layout.group_cols[0].start == _col("M")
    assert layout.group_cols[0].stop == _col("M") + 8        # M..T
    # gap then G2 starts at T+2 = 22 (V)
    assert layout.group_cols[1].start == _col("V")
    assert layout.group_cols[1].stop == _col("V") + 10       # V..AE
    assert layout.group_cols[2].start == _col("AG")
    assert layout.group_cols[2].stop == _col("AG") + 12      # AG..AR
    # stats: data ends at AR (column 44); +2 = AT (col 46)
    stats = layout.stat_cols
    assert stats.mean_sd_pairs[0][0] == _col("AT")
    # 6 mean/SD cols (cols 46-51), then Levene 52, SW 53, Overall 54,
    # then 3 raw (55-57), then 3 Q (58-60)
    assert stats.levene == 52
    assert stats.shapiro_min == 53
    assert stats.overall == 54
    assert stats.pairwise_raw == (55, 56, 57)
    assert stats.pairwise_q == (58, 59, 60)


def test_layout_k4_total_stat_cols():
    layout = compute_layout(decimals_col=14, group_sizes=(6, 8, 10, 12))
    stats = layout.stat_cols
    # 2K + 3 + 2*C(K,2) = 8 + 3 + 12 = 23 stat cols
    n_stat = 2 * 4 + 3 + 2 * 6
    flat = (
        [c for pair in stats.mean_sd_pairs for c in pair]
        + [stats.levene, stats.shapiro_min, stats.overall]
        + list(stats.pairwise_raw)
        + list(stats.pairwise_q)
    )
    assert len(flat) == n_stat
    assert flat == list(range(flat[0], flat[0] + n_stat))   # contiguous
```

- [ ] **Step 4.2: 跑失败**

Expected: ImportError or AttributeError.

- [ ] **Step 4.3: 实现**

替换 `Layout` + `compute_layout`：

```python
@dataclass(slots=True, frozen=True)
class StatColIndices:
    mean_sd_pairs: tuple[tuple[int, int], ...]
    levene: int
    shapiro_min: int
    overall: int
    pairwise_raw: tuple[int, ...]
    pairwise_q: tuple[int, ...]


@dataclass(slots=True, frozen=True)
class Layout:
    decimals_col: int
    group_cols: tuple[range, ...]
    stat_cols: StatColIndices


def compute_layout(decimals_col: int, group_sizes: tuple[int, ...]) -> Layout:
    if len(group_sizes) < 2:
        raise ValueError(f"组数必须 ≥2，收到 {len(group_sizes)}")
    if any(n < 2 for n in group_sizes):
        raise ValueError(f"每组 N 必须 ≥2，收到 {group_sizes}")

    data_start = decimals_col + 2
    group_cols: list[range] = []
    cursor = data_start
    for n in group_sizes:
        group_cols.append(range(cursor, cursor + n))
        cursor += n + 1  # n cells + 1 blank between groups

    # last cursor includes a trailing +1 blank that we don't need;
    # subtract it then add 2 (1 trailing blank + 1 spacer) → +1 net
    stats_start = cursor + 1

    pair_cols: list[tuple[int, int]] = []
    c = stats_start
    for _ in group_sizes:
        pair_cols.append((c, c + 1))
        c += 2

    levene = c; c += 1
    shapiro = c; c += 1
    overall = c; c += 1

    k = len(group_sizes)
    n_pairs = k * (k - 1) // 2
    raw = tuple(range(c, c + n_pairs))
    c += n_pairs
    q = tuple(range(c, c + n_pairs))

    return Layout(
        decimals_col=decimals_col,
        group_cols=tuple(group_cols),
        stat_cols=StatColIndices(
            mean_sd_pairs=tuple(pair_cols),
            levene=levene,
            shapiro_min=shapiro,
            overall=overall,
            pairwise_raw=raw,
            pairwise_q=q,
        ),
    )
```

- [ ] **Step 4.4: 跑通过**

`.venv/bin/pytest tests/test_layout.py -v` Expected: 3 PASS.

---

## Task 5: 统计层 4 个新函数

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（删 `compute_stats`，新增 4 个函数）
- Replace: `tests/test_stats.py`

- [ ] **Step 5.1: 写失败测试**

```python
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
    return [rng.normal(10.0, 2.0, 30), rng.normal(11.0, 2.0, 30), rng.normal(12.0, 2.0, 30)]


def _diff_var_data(rng):
    return [rng.normal(10.0, 1.0, 30), rng.normal(10.0, 10.0, 30), rng.normal(10.0, 5.0, 30)]


def test_compute_levene_equal_var():
    rng = np.random.default_rng(0)
    groups = _eq_var_data(rng)
    p, equal_var = compute_levene(groups)
    expected = float(sp_stats.levene(*groups, center="median").pvalue)
    assert isclose(p, expected)
    assert equal_var == (expected >= LEVENE_ALPHA)
    assert equal_var is True  # constructed equal-variance


def test_compute_levene_unequal_var():
    rng = np.random.default_rng(1)
    groups = _diff_var_data(rng)
    _, equal_var = compute_levene(groups)
    assert equal_var is False


def test_compute_shapiro_min_normal_groups():
    rng = np.random.default_rng(0)
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
    expected_p = [
        float(expected_matrix[i][j]) for i, j in combinations(range(3), 2)
    ]
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
    # equal_var True but mark as not all_normal → must go Welch+Bonferroni
    groups = _eq_var_data(rng)
    _, q, label = compute_pairwise(groups, equal_var=True, all_normal=False)
    assert label == "Welch+Bonferroni"
    assert all(qv is not None for qv in q)
```

- [ ] **Step 5.2: 跑失败**

`.venv/bin/pytest tests/test_stats.py -v` Expected: ImportError.

- [ ] **Step 5.3: 实现**

替换 `compute_stats` 为下面 4 个函数：

```python
def compute_levene(
    groups: list[np.ndarray], alpha: float = LEVENE_ALPHA
) -> tuple[float, bool]:
    res = sp_stats.levene(*groups, center="median")
    p = float(res.pvalue)
    if np.isnan(p):
        return float("nan"), False
    return p, p > alpha


def compute_shapiro_min(
    groups: list[np.ndarray], alpha: float = LEVENE_ALPHA
) -> tuple[float, bool]:
    ps: list[float] = []
    for i, g in enumerate(groups, start=1):
        if len(g) < 3:
            logger.warning("G%d 样本量 N=%d < 3，Shapiro-Wilk 视为 0（不正态）", i, len(g))
            ps.append(0.0)
            continue
        try:
            res = sp_stats.shapiro(g)
            p = float(res.pvalue)
            if np.isnan(p):
                p = 0.0
            ps.append(p)
        except Exception as e:
            logger.warning("G%d Shapiro-Wilk 失败: %s，视为 0", i, e)
            ps.append(0.0)
    min_p = min(ps)
    return min_p, min_p > alpha


def compute_overall(
    groups: list[np.ndarray], equal_var: bool
) -> tuple[float, str]:
    if equal_var:
        res = sp_stats.f_oneway(*groups)
        return float(res.pvalue), "ANOVA"
    res = sp_stats.kruskal(*groups)
    return float(res.pvalue), "KW"


def compute_pairwise(
    groups: list[np.ndarray], equal_var: bool, all_normal: bool
) -> tuple[list[float], list[float | None], str]:
    k = len(groups)
    pairs = list(itertools.combinations(range(k), 2))
    if equal_var and all_normal:
        matrix = sp_stats.tukey_hsd(*groups).pvalue
        raw = [float(matrix[i][j]) for i, j in pairs]
        q: list[float | None] = [None] * len(pairs)
        return raw, q, "Tukey"
    raw = [
        float(sp_stats.ttest_ind(groups[i], groups[j], equal_var=False).pvalue)
        for i, j in pairs
    ]
    n_pairs = len(pairs)
    q = [min(1.0, r * n_pairs) for r in raw]
    return raw, q, "Welch+Bonferroni"
```

并在文件顶部 `import` 区加 `import itertools`。

- [ ] **Step 5.4: 跑通过**

`.venv/bin/pytest tests/test_stats.py -v` Expected: 9 PASS.

---

## Task 6: 写入层 + CLI 入口（重写 write_row + _process_rows + main）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（重写 write_row、_process_rows；main 仅微调日志）
- Replace: `tests/test_cli.py`

- [ ] **Step 6.1: 重写 write_row**

替换原 `write_row` 为：

```python
def _write_group_data_headers(
    ws: Worksheet,
    group_cols: tuple[range, ...],
    group_sizes: tuple[int, ...],
) -> None:
    """第 1 行写入每组的列序号 1..N-1，最后一格写 '<N>（Gi）'."""
    for i, (cols, n) in enumerate(zip(group_cols, group_sizes, strict=True), start=1):
        last = cols.stop - 1
        for idx, col in enumerate(cols, start=1):
            ws.cell(row=1, column=col).value = f"{n}（G{i}）" if col == last else idx


def _write_stat_headers(ws: Worksheet, stat_cols: StatColIndices, k: int) -> None:
    headers: list[tuple[int, str]] = []
    for i, (mu_col, sd_col) in enumerate(stat_cols.mean_sd_pairs, start=1):
        headers.append((mu_col, f"第{i}组均值"))
        headers.append((sd_col, f"第{i}组SD值"))
    headers.append((stat_cols.levene, LEVENE_HEADER))
    headers.append((stat_cols.shapiro_min, SHAPIRO_HEADER))
    headers.append((stat_cols.overall, OVERALL_HEADER))
    pair_labels = [
        f"G{i + 1}-G{j + 1}" for i, j in itertools.combinations(range(k), 2)
    ]
    for col, label in zip(stat_cols.pairwise_raw, pair_labels, strict=True):
        headers.append((col, f"{label} raw p"))
    for col, label in zip(stat_cols.pairwise_q, pair_labels, strict=True):
        headers.append((col, f"{label} Q-value"))
    for col, text in headers:
        if ws.cell(row=1, column=col).value in (None, ""):
            ws.cell(row=1, column=col).value = text


def _round4(v: float | None) -> float | None:
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    return round(float(v), 4)


def write_row(
    ws: Worksheet,
    row: RowSpec,
    generated: list[np.ndarray],
    layout: Layout,
    levene_p: float,
    sw_min_p: float,
    overall_p: float,
    raw_ps: list[float],
    q_values: list[float | None],
) -> None:
    group_sizes = tuple(g.n for g in row.groups)
    _write_group_data_headers(ws, layout.group_cols, group_sizes)
    _write_stat_headers(ws, layout.stat_cols, len(row.groups))

    for cols, data in zip(layout.group_cols, generated, strict=True):
        for value, col in zip(data.tolist(), cols, strict=True):
            ws.cell(row=row.row_index, column=col).value = value

    actual_means = [float(np.mean(d)) for d in generated]
    actual_sds = [float(np.std(d, ddof=1)) for d in generated]
    for (mu_col, sd_col), mu, sd in zip(
        layout.stat_cols.mean_sd_pairs, actual_means, actual_sds, strict=True
    ):
        ws.cell(row=row.row_index, column=mu_col).value = _round4(mu)
        ws.cell(row=row.row_index, column=sd_col).value = _round4(sd)

    ws.cell(row=row.row_index, column=layout.stat_cols.levene).value = _round4(levene_p)
    ws.cell(row=row.row_index, column=layout.stat_cols.shapiro_min).value = _round4(sw_min_p)
    ws.cell(row=row.row_index, column=layout.stat_cols.overall).value = _round4(overall_p)
    for col, p in zip(layout.stat_cols.pairwise_raw, raw_ps, strict=True):
        ws.cell(row=row.row_index, column=col).value = _round4(p)
    for col, q in zip(layout.stat_cols.pairwise_q, q_values, strict=True):
        ws.cell(row=row.row_index, column=col).value = _round4(q)
```

- [ ] **Step 6.2: 重写 _process_rows**

```python
def _process_rows(
    ws: Worksheet, specs: Iterable[RowSpec], rng: np.random.Generator, decimals_col: int
) -> int:
    count = 0
    for spec in specs:
        group_sizes = tuple(g.n for g in spec.groups)
        layout = compute_layout(decimals_col, group_sizes)

        generated: list[np.ndarray] = []
        for i, g in enumerate(spec.groups, start=1):
            seed = int(rng.integers(0, 2**32 - 1))
            arr = generate_with_retry(
                metric=f"{spec.metric}/G{i}",
                target_mean=g.mean,
                target_sd=g.sd,
                size=g.n,
                rng=np.random.default_rng(seed),
                decimals=spec.decimals,
            )
            generated.append(arr)

        levene_p, equal_var = compute_levene(generated)
        sw_min_p, all_normal = compute_shapiro_min(generated)
        overall_p, overall_label = compute_overall(generated, equal_var)
        raw_ps, q_values, branch_label = compute_pairwise(generated, equal_var, all_normal)

        write_row(ws, spec, generated, layout, levene_p, sw_min_p, overall_p, raw_ps, q_values)
        count += 1
        logger.info(
            "处理完成 row=%d metric=%s K=%d levene_p=%.4f all_normal=%s overall=%s(%.4f) pairwise=%s",
            spec.row_index, spec.metric, len(spec.groups),
            levene_p, all_normal, overall_label, overall_p, branch_label,
        )
    return count
```

- [ ] **Step 6.3: 微调 main 适配新 _process_rows 签名**

把原 `n = _process_rows(ws, specs, rng)` 改为：

```python
decimals_col = _find_decimals_col(ws)
n = _process_rows(ws, specs, rng, decimals_col)
```

并在 `read_specs` 调用之后立即记录：

```python
logger.info("识别到位数列在第 %d 列, K=%d", decimals_col, (decimals_col - 2) // 3)
```

- [ ] **Step 6.4: 写端到端测试**

完整覆盖 `tests/test_cli.py`：

```python
from __future__ import annotations

import shutil
from itertools import combinations
from pathlib import Path

import numpy as np
import openpyxl
import pytest
from openpyxl.utils import column_index_from_string
from scipy import stats as sp_stats

from scripts.generate_excel_random_data import LEVENE_ALPHA, main

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_K2 = FIXTURES / "sample-k2.xlsx"
SAMPLE_K3 = FIXTURES / "sample-k3.xlsx"


def _read_groups(ws, layout_cols):
    return [
        np.asarray(
            [ws.cell(row=row, column=c).value for c in cols]
        )
        for cols in layout_cols
    ]


def test_cli_requires_input_argument(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 2
    assert "--input" in capsys.readouterr().err


def test_cli_writes_expected_layout_k2(tmp_path: Path):
    src = tmp_path / "in.xlsx"
    shutil.copy(SAMPLE_K2, src)
    rc = main(["--input", str(src), "--output", str(src), "--seed", "42"])
    assert rc == 0

    wb_in = openpyxl.load_workbook(SAMPLE_K2, data_only=True)
    wb_out = openpyxl.load_workbook(src, data_only=True)
    ws_in = wb_in.active
    ws_out = wb_out.active

    # A-H (1..8) preserved
    for r in range(1, ws_in.max_row + 1):
        for c in range(1, 9):
            assert ws_out.cell(row=r, column=c).value == ws_in.cell(row=r, column=c).value, (r, c)

    # Data starts at J (col 10)
    col_J, col_S, col_T, col_U, col_AF = (column_index_from_string(L) for L in ["J", "S", "T", "U", "AF"])
    assert ws_out.cell(row=1, column=col_S).value == "10（G1）"
    assert ws_out.cell(row=1, column=col_AF).value == "12（G2）"
    for r in range(2, 5):
        assert ws_out.cell(row=r, column=col_T).value is None  # blank between groups

    # Row 2 stats: Levene (AL=38), SW (AM=39), Overall (AN=40), raw (AO=41), Q (AP=42)
    col_AL = column_index_from_string("AL")
    g1 = np.asarray([ws_out.cell(row=2, column=c).value for c in range(col_J, col_J + 10)])
    g2 = np.asarray([ws_out.cell(row=2, column=c).value for c in range(col_U, col_U + 12)])

    expected_levene = float(sp_stats.levene(g1, g2, center="median").pvalue)
    expected_sw = min(
        float(sp_stats.shapiro(g1).pvalue), float(sp_stats.shapiro(g2).pvalue)
    )
    equal_var = expected_levene > LEVENE_ALPHA
    if equal_var:
        expected_overall = float(sp_stats.f_oneway(g1, g2).pvalue)
    else:
        expected_overall = float(sp_stats.kruskal(g1, g2).pvalue)

    got_levene = ws_out.cell(row=2, column=col_AL).value
    got_sw = ws_out.cell(row=2, column=col_AL + 1).value
    got_overall = ws_out.cell(row=2, column=col_AL + 2).value
    assert got_levene == pytest.approx(round(expected_levene, 4))
    assert got_sw == pytest.approx(round(expected_sw, 4))
    assert got_overall == pytest.approx(round(expected_overall, 4))


def test_cli_writes_expected_layout_k3(tmp_path: Path):
    src = tmp_path / "in.xlsx"
    shutil.copy(SAMPLE_K3, src)
    rc = main(["--input", str(src), "--output", str(src), "--seed", "42"])
    assert rc == 0

    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb.active

    # A..K (1..11) preserved
    wb_in = openpyxl.load_workbook(SAMPLE_K3, data_only=True)
    ws_in = wb_in.active
    for r in range(1, ws_in.max_row + 1):
        for c in range(1, 12):
            assert ws.cell(row=r, column=c).value == ws_in.cell(row=r, column=c).value

    # Data starts at M (col 13). G1 N=8 → M..T; gap U; G2 N=10 → V..AE; gap AF; G3 N=12 → AG..AR
    col_T = column_index_from_string("T")
    col_AE = column_index_from_string("AE")
    col_AR = column_index_from_string("AR")
    assert ws.cell(row=1, column=col_T).value == "8（G1）"
    assert ws.cell(row=1, column=col_AE).value == "10（G2）"
    assert ws.cell(row=1, column=col_AR).value == "12（G3）"

    # Stat block starts at AT (col 46): μ1,σ1,μ2,σ2,μ3,σ3 (6) + Levene + SW + Overall + 3 raw + 3 Q = 15
    stat_start = 46
    pair_count = 3
    raw_start = stat_start + 6 + 3
    q_start = raw_start + pair_count

    # Verify row 2 raw vs computed
    g_cols_per = [(13, 8), (22, 10), (33, 12)]
    groups = []
    for (start, n) in g_cols_per:
        groups.append(np.asarray([ws.cell(row=2, column=c).value for c in range(start, start + n)]))

    pairs = list(combinations(range(3), 2))
    levene_p = float(sp_stats.levene(*groups, center="median").pvalue)
    if levene_p > LEVENE_ALPHA:
        sw_min = min(float(sp_stats.shapiro(g).pvalue) for g in groups)
        if sw_min > LEVENE_ALPHA:
            matrix = sp_stats.tukey_hsd(*groups).pvalue
            expected_raw = [float(matrix[i][j]) for i, j in pairs]
            expected_q = [None] * pair_count
        else:
            expected_raw = [
                float(sp_stats.ttest_ind(groups[i], groups[j], equal_var=False).pvalue)
                for i, j in pairs
            ]
            expected_q = [min(1.0, r * pair_count) for r in expected_raw]
    else:
        expected_raw = [
            float(sp_stats.ttest_ind(groups[i], groups[j], equal_var=False).pvalue)
            for i, j in pairs
        ]
        expected_q = [min(1.0, r * pair_count) for r in expected_raw]

    for i, exp in enumerate(expected_raw):
        got = ws.cell(row=2, column=raw_start + i).value
        assert got == pytest.approx(round(exp, 4)), (i, got, exp)
    for i, exp in enumerate(expected_q):
        got = ws.cell(row=2, column=q_start + i).value
        if exp is None:
            assert got is None
        else:
            assert got == pytest.approx(round(exp, 4))


def test_cli_seed_reproducible(tmp_path: Path):
    src1 = tmp_path / "a.xlsx"
    src2 = tmp_path / "b.xlsx"
    shutil.copy(SAMPLE_K2, src1)
    shutil.copy(SAMPLE_K2, src2)
    main(["--input", str(src1), "--output", str(src1), "--seed", "42"])
    main(["--input", str(src2), "--output", str(src2), "--seed", "42"])
    wb1 = openpyxl.load_workbook(src1, data_only=True).active
    wb2 = openpyxl.load_workbook(src2, data_only=True).active
    for r in range(1, wb1.max_row + 1):
        for c in range(1, wb1.max_column + 1):
            assert wb1.cell(row=r, column=c).value == wb2.cell(row=r, column=c).value, (r, c)


def test_cli_no_decimals_col_returns_3(tmp_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["指标", "N1", "μ1", "σ1", "N2", "μ2", "σ2"])
    ws.append(["x", 10, 1.0, 1.0, 10, 1.0, 1.0])
    p = tmp_path / "no_dec.xlsx"
    wb.save(p)
    rc = main(["--input", str(p), "--output", str(p)])
    assert rc == 3
```

- [ ] **Step 6.5: 在 `main` 里捕获 `LookupError` 返回 3**

```python
try:
    specs = read_specs(input_path, args.sheet)
except LookupError as e:
    logger.error("无法解析输入文件结构: %s", e)
    return 3
```

- [ ] **Step 6.6: 跑全量测试**

`.venv/bin/pytest -v` Expected: 全绿。

- [ ] **Step 6.7: 删除 v1 残留**

- 删除模块顶部 `STAT_HEADERS` 常量（已被 LEVENE_HEADER 等替换）
- 删除 `compute_stats` 函数体（如果还在）
- ruff 自动清理 unused imports：`.venv/bin/ruff format scripts tests && .venv/bin/ruff check scripts tests --fix`

---

## Task 7: 文档与端到端冒烟

- [ ] **Step 7.1: 重写 scripts/README.md 列布局段**

把 v1 的"列布局（动态，按 N、M 计算）"段改为：

```markdown
## 列布局（动态，按 K 与每组 N 计算，不写死）

输入区（脚本不修改）：
- A 列：指标名
- 紧跟 K 个 (N、μ、σ) 三元组：第 1 组在 B-D、第 2 组在 E-G、第 3 组在 H-J、…
- 第 K 组之后**紧跟一列"原始数据小数点后位数"**（脚本通过扫描第 1 行表头中含"位数"的列定位它，K 由位置反推）
- 位数列后**1 列空白**（可选填备注，脚本不读不写）

数据区（脚本写入）：
- 数据起始列 = 位数列 + 2
- 第 i 组占 N_i 列，组间空 1 列
- 第 i 组最后一格的第 1 行表头会被覆写为 `<N_i>（G<i>）`

统计区（脚本写入，紧跟数据区 + 1 空列）：
- 第 i 组均值、第 i 组 SD（共 2K 列）
- Levene p（中心 median；> 0.05 视为方差齐）
- Shapiro-Wilk min p（每组 SW 取最小；> 0.05 视为全部正态）
- 整体 p（齐 → ANOVA；不齐 → Kruskal-Wallis）
- 两两 raw p × C(K,2) 列（按 (1,2)(1,3)(2,3)... 顺序）
- 两两 Q-value × C(K,2) 列（齐+全正态走 Tukey 时为空；否则 Welch+Bonferroni 校正值）

例（K=2、N₁=10、N₂=12、位数列 H）：数据 J-S（G1）/ U-AF（G2）；统计 AH-AP 共 9 列。
例（K=3、N₁=8、N₂=10、N₃=12、位数列 K）：数据 M-T / V-AE / AG-AR；统计 AT 起共 15 列。
```

- [ ] **Step 7.2: 实际跑一次 K=2 样本**

```bash
.venv/bin/python scripts/generate_excel_random_data.py \
    --input "20260509-随机数生成.xlsx" \
    --output /tmp/out-v2-k2.xlsx \
    --seed 42 -v
```
Expected: 日志显示 `识别到位数列在第 8 列, K=2`，`处理完成 row=2 metric=体重 K=2 ...`，肉眼复核 `/tmp/out-v2-k2.xlsx` 的统计区 9 列。

- [ ] **Step 7.3: 实际跑一次 K=3 fixture**

```bash
.venv/bin/python scripts/generate_excel_random_data.py \
    --input tests/fixtures/sample-k3.xlsx \
    --output /tmp/out-v2-k3.xlsx \
    --seed 42 -v
```
Expected: `识别到位数列在第 11 列, K=3`，统计列 15 列。

- [ ] **Step 7.4: ruff & pytest**

```bash
.venv/bin/ruff format scripts tests
.venv/bin/ruff check scripts tests
.venv/bin/pytest -v
```
Expected: 都绿。

---

## Task 8: 收尾（OpenSpec 状态 + tasks 勾选）

- [ ] **Step 8.1: 把 OpenSpec tasks.md 全部 `- [ ]` 改成 `- [x]`**

```bash
sed -i '' 's/- \[ \]/- [x]/g' openspec/changes/extend-to-multi-group-stats/tasks.md
openspec validate "extend-to-multi-group-stats" --strict
openspec status --change "extend-to-multi-group-stats"
```
Expected: 4/4 done & valid.

- [ ] **Step 8.2: git status 检查改动范围**

```bash
git status
git diff --stat HEAD
```
Expected: 仅 `scripts/`、`tests/`、`docs/superpowers/plans/`、`openspec/changes/extend-to-multi-group-stats/`。

---

## Self-Review

**1. Spec coverage（对照 [spec.md](../../../openspec/changes/extend-to-multi-group-stats/specs/excel-random-data-generator/spec.md)）：**
- 读取目标统计量 → Task 2
- 生成两组正态分布原始数据（含 decimals=0/4/6） → Task 3
- 按动态列布局写入原始数据（K=2/K=3） → Task 4 + Task 6
- 在表头标注样本量（Gi 格式） → Task 6（_write_group_data_headers）
- 不破坏未涉及的单元格 → Task 6 测试中 A-H/A-K 列断言
- Levene 方差齐性检验 → Task 5（compute_levene）
- Shapiro-Wilk 正态性检验 → Task 5（compute_shapiro_min，含 N<3 fallback）
- 整体差异 p 值 → Task 5（compute_overall）
- 两两对比 p 值与多重比较校正 → Task 5（compute_pairwise，含 Tukey 与 Welch 双分支测试）
- 按统计列布局写入回算结果 → Task 6（_write_stat_headers + write_row）

**2. Placeholder 扫描：** 没有 TBD/TODO；每个 step 都有可粘贴代码或可执行命令。

**3. 类型/命名一致：** `_find_decimals_col` / `compute_layout(decimals_col, group_sizes)` / `Layout(decimals_col, group_cols, stat_cols: StatColIndices)` / `StatColIndices(mean_sd_pairs, levene, shapiro_min, overall, pairwise_raw, pairwise_q)` / `compute_levene/shapiro_min/overall/pairwise` / `_round4` / `_write_group_data_headers / _write_stat_headers / write_row(ws, row, generated, layout, levene_p, sw_min_p, overall_p, raw_ps, q_values)` 全部一致。

---

## Execution Handoff

计划保存。本次会话内联执行（superpowers:executing-plans）：Task 1 → Task 8 顺序，Task 2-5 严格 TDD（写失败测 → 跑失败 → 写实现 → 跑通过）；遇到错误立即切到 systematic-debugging。
