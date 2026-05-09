# Excel 随机数据生成器 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现一个 Python 脚本：从 Excel 读取每行的两组目标统计量（N、均值、SD），生成符合正态分布、保留 4 位小数的两组原始数据，回算出真实的均值/SD 与 Levene→Student/Welch 双尾 p 值，再按规定列布局写回同一份 Excel。

**Architecture:** 单文件 CLI 脚本 `scripts/generate_excel_random_data.py`，内部分四层：read（解析目标统计量）→ generate（采样+线性变换+SD 容差重试）→ stats（Levene 分支算 p）→ write（动态列布局回写 + 表头样本量标注）。openpyxl 负责 IO，numpy/scipy 负责数值与统计。

**Tech Stack:** Python 3.10+、`openpyxl`、`numpy`、`scipy`、`pytest`、`ruff`。

**关键 OpenSpec 文档（实现时随时回查）：**
- 提案：[../../../openspec/changes/add-excel-random-data-generator/proposal.md](../../../openspec/changes/add-excel-random-data-generator/proposal.md)
- 设计：[../../../openspec/changes/add-excel-random-data-generator/design.md](../../../openspec/changes/add-excel-random-data-generator/design.md)
- 规范：[../../../openspec/changes/add-excel-random-data-generator/specs/excel-random-data-generator/spec.md](../../../openspec/changes/add-excel-random-data-generator/specs/excel-random-data-generator/spec.md)
- 任务（OpenSpec 版）：[../../../openspec/changes/add-excel-random-data-generator/tasks.md](../../../openspec/changes/add-excel-random-data-generator/tasks.md)

---

## File Structure

将要创建/修改的文件：

| 路径 | 责任 |
|---|---|
| `requirements.txt` | 锁定 `openpyxl`、`numpy`、`scipy`、`pytest`、`ruff` 版本上界 |
| `scripts/__init__.py` | 让 `scripts` 成为可导入包，方便测试直接 `from scripts.generate_excel_random_data import ...` |
| `scripts/generate_excel_random_data.py` | 单文件实现：dataclass、read/generate/stats/write 四层、CLI |
| `scripts/README.md` | 命令示例、列布局说明、就地修改的备份提示 |
| `tests/__init__.py` | 测试包标记（空文件） |
| `tests/conftest.py` | `sys.path` 注入项目根，便于 `import scripts.xxx` |
| `tests/test_read_specs.py` | 单元测试：读取目标统计量 |
| `tests/test_generate.py` | 单元测试：生成 + SD 容差重试 |
| `tests/test_stats.py` | 单元测试：Levene → t / Welch 分支 |
| `tests/test_layout.py` | 单元测试：列偏移与样本表 I/R/T/AE/AG/AK 一致 |
| `tests/test_cli.py` | 端到端测试：用真实 xlsx 跑 CLI 入口 |

---

## Task 1: 项目骨架与依赖

**Files:**
- Create: `requirements.txt`
- Create: `scripts/__init__.py`
- Create: `scripts/generate_excel_random_data.py`（仅骨架）
- Create: `scripts/README.md`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`

- [ ] **Step 1.1: 写 `requirements.txt`**

```text
openpyxl>=3.1,<4
numpy>=1.26,<3
scipy>=1.11,<2
pytest>=8,<9
ruff>=0.5,<1
```

- [ ] **Step 1.2: 创建空的 `scripts/__init__.py`**

文件内容仅一个空行。

- [ ] **Step 1.3: 写 `scripts/generate_excel_random_data.py` 骨架**

```python
"""按目标统计量为 Excel 中每行指标生成两组正态分布原始数据，并回算 mean/SD/p 值。"""

from __future__ import annotations

import argparse
import logging
import secrets
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from scipy import stats as sp_stats

logger = logging.getLogger(__name__)

GROUP1_START_COL = 9  # I 列
LEVENE_ALPHA = 0.05
SD_TOLERANCE = 0.10  # ±10%
MAX_RETRY = 5
ROUND_DIGITS = 4
STAT_HEADERS = ("第一组均值", "第一组SD值", "第二组均值", "第二组SD值", "两组的p值")


def main(argv: list[str] | None = None) -> int:
    raise NotImplementedError("filled in Task 7")


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 1.4: 写 `scripts/README.md`**

```markdown
# Excel 随机数据生成脚本

## 用途

读取一份 Excel 表（第 1 行为表头，第 2 行起每行一个指标），按 A-G 列描述的两组目标统计量（样本量 N/M、均值、SD）生成两组正态分布原始数据，并回算实际 mean/SD 与 Levene→Student/Welch 双尾 p 值。

## 列布局（动态，按 N、M 计算，不写死）

- A-H 列：原表头与备注（脚本不修改）。
- I 列起：第一组数据，长度 N。最后一格的第 1 行表头会被覆写为 `<N>（N）`。
- 第一组之后空 1 列。
- 紧跟第二组数据，长度 M。最后一格的第 1 行表头会被覆写为 `<M>（M）`。
- 第二组之后空 1 列。
- 接着 5 列统计：第一组均值 / 第一组SD值 / 第二组均值 / 第二组SD值 / 两组的p值。

当 N=10、M=12 时，三块区域恰好对应 I-R / T-AE / AG-AK。

## 命令

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# 默认就地修改根目录的 20260509-随机数生成.xlsx
python scripts/generate_excel_random_data.py

# 写到别处（推荐先备份）
python scripts/generate_excel_random_data.py --output /tmp/out.xlsx --seed 42
```

## CLI 参数

- `--input PATH` 默认 `20260509-随机数生成.xlsx`
- `--output PATH` 默认与 `--input` 相同（**就地修改源文件**，请先备份）
- `--sheet NAME` 默认第一个 sheet
- `--seed INT` 不传则随机；INFO 日志会打印实际种子，便于复现
- `-v, --verbose` 把日志级别从 INFO 调到 DEBUG
```

- [ ] **Step 1.5: 创建 `tests/__init__.py`（空）和 `tests/conftest.py`**

`tests/conftest.py`:
```python
"""把项目根加入 sys.path，便于测试直接 `import scripts.generate_excel_random_data`。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
```

- [ ] **Step 1.6: 安装依赖并跑一次冒烟检查**

Run:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -c "import openpyxl, numpy, scipy; print('ok')"
```
Expected: `ok`

- [ ] **Step 1.7: Commit**

```bash
git add requirements.txt scripts/ tests/__init__.py tests/conftest.py
git commit -m "feat: scaffold excel random data generator (deps + skeleton)"
```

---

## Task 2: 数据结构 + 读取层（TDD）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（新增 dataclass 与 `read_specs`）
- Create: `tests/test_read_specs.py`

- [ ] **Step 2.1: 写 dataclass 与 `read_specs` 的失败测试**

`tests/test_read_specs.py`:
```python
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

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
        ("空 N", None, 100.0, 5.0, 10, 90.0, 5.0),            # missing N
        ("N=1", 1, 100.0, 5.0, 10, 90.0, 5.0),                # N too small
        ("SD=0", 10, 100.0, 0.0, 10, 90.0, 5.0),              # SD invalid
        ("非数值", 10, "x", 5.0, 10, 90.0, 5.0),               # non-numeric mean
    ]
    path = _make_xlsx(tmp_path, rows)
    with caplog.at_level("WARNING"):
        specs = read_specs(path, sheet=None)
    assert [s.metric for s in specs] == ["体重"]
    # Each invalid row should produce a warning record
    assert sum(1 for r in caplog.records if r.levelname == "WARNING") >= 4


def test_read_specs_empty_sheet(tmp_path):
    path = _make_xlsx(tmp_path, [])
    assert read_specs(path, sheet=None) == []
```

- [ ] **Step 2.2: 跑测试，确认失败**

Run: `pytest tests/test_read_specs.py -v`
Expected: ImportError / 失败（`GroupSpec` / `RowSpec` / `read_specs` 都还没实现）

- [ ] **Step 2.3: 实现 dataclass + `read_specs`**

在 `scripts/generate_excel_random_data.py` 顶部 `logger = ...` 之后插入：

```python
@dataclass(slots=True, frozen=True)
class GroupSpec:
    name: str
    n: int
    mean: float
    sd: float


@dataclass(slots=True, frozen=True)
class RowSpec:
    row_index: int  # 1-based, openpyxl convention
    metric: str
    group1: GroupSpec
    group2: GroupSpec


def _coerce_positive_int(value: object) -> int | None:
    if isinstance(value, bool):  # bool is subclass of int — exclude
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
    out: list[RowSpec] = []
    for row_idx in range(2, ws.max_row + 1):
        metric = ws.cell(row=row_idx, column=1).value
        n = _coerce_positive_int(ws.cell(row=row_idx, column=2).value)
        mean1 = _coerce_float(ws.cell(row=row_idx, column=3).value)
        sd1 = _coerce_float(ws.cell(row=row_idx, column=4).value)
        m = _coerce_positive_int(ws.cell(row=row_idx, column=5).value)
        mean2 = _coerce_float(ws.cell(row=row_idx, column=6).value)
        sd2 = _coerce_float(ws.cell(row=row_idx, column=7).value)

        if metric is None and n is None and mean1 is None:
            continue  # 整行空，安静跳过

        problems: list[str] = []
        if not isinstance(metric, str) or not metric.strip():
            problems.append("缺少指标名")
        for label, val in (("N", n), ("M", m)):
            if val is None or val <= 1:
                problems.append(f"{label} 必须是 ≥2 的整数")
        for label, val in (("第一组均值", mean1), ("第二组均值", mean2)):
            if val is None:
                problems.append(f"{label} 缺失或非数值")
        for label, val in (("第一组 SD", sd1), ("第二组 SD", sd2)):
            if val is None or val <= 0:
                problems.append(f"{label} 必须是 >0 的数值")
        if problems:
            logger.warning("跳过第 %d 行（%r）：%s", row_idx, metric, "；".join(problems))
            continue

        out.append(
            RowSpec(
                row_index=row_idx,
                metric=metric,
                group1=GroupSpec(name="第一组", n=n, mean=mean1, sd=sd1),
                group2=GroupSpec(name="第二组", n=m, mean=mean2, sd=sd2),
            )
        )
    return out
```

- [ ] **Step 2.4: 跑测试，确认通过**

Run: `pytest tests/test_read_specs.py -v`
Expected: 全部 PASS

- [ ] **Step 2.5: Commit**

```bash
git add scripts/generate_excel_random_data.py tests/test_read_specs.py
git commit -m "feat: implement read_specs + GroupSpec/RowSpec dataclasses"
```

---

## Task 3: 数据生成层（TDD）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（新增 `generate_one_group` 与 `generate_with_retry`）
- Create: `tests/test_generate.py`

- [ ] **Step 3.1: 写失败测试**

`tests/test_generate.py`:
```python
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
    # 4 位小数检查：x*10000 应当是整数
    assert np.all(np.isclose(x * 10000, np.round(x * 10000)))
    # round 之后允许微小漂移，但应当落在容差内
    assert abs(np.mean(x) - 100.0) < 1e-2
    assert abs(np.std(x, ddof=1) - 15.0) / 15.0 <= SD_TOLERANCE


def test_generate_with_retry_small_sample_within_tolerance():
    rng = np.random.default_rng(0)
    for _ in range(5):  # 多次随机起始 seed，都应能在容差内成功
        sub_rng = np.random.default_rng(rng.integers(0, 2**32 - 1))
        x = generate_with_retry(metric="t", target_mean=5.0, target_sd=2.0, size=2, rng=sub_rng)
        assert x.shape == (2,)
        assert abs(np.std(x, ddof=1) - 2.0) / 2.0 <= SD_TOLERANCE


def test_generate_with_retry_rejects_zero_sd():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(metric="bad", target_mean=1.0, target_sd=0.0, size=10, rng=rng)


def test_generate_with_retry_rejects_size_lt_2():
    rng = np.random.default_rng(0)
    with pytest.raises(ValueError):
        generate_with_retry(metric="bad", target_mean=1.0, target_sd=1.0, size=1, rng=rng)
```

- [ ] **Step 3.2: 跑测试，确认失败**

Run: `pytest tests/test_generate.py -v`
Expected: ImportError 或 fail

- [ ] **Step 3.3: 实现生成函数**

在 `scripts/generate_excel_random_data.py` 内 `read_specs` 之后追加：

```python
def generate_one_group(
    target_mean: float,
    target_sd: float,
    size: int,
    rng: np.random.Generator,
) -> np.ndarray:
    """采样标准正态后做线性变换，使未 round 前 mean/SD 严格匹配目标，再 round 到 4 位小数。"""
    if size < 2:
        raise ValueError(f"size 必须 ≥2，收到 {size}")
    if target_sd <= 0:
        raise ValueError(f"target_sd 必须 >0，收到 {target_sd}")
    z = rng.standard_normal(size)
    z = (z - z.mean()) / z.std(ddof=1)
    x = target_mean + target_sd * z
    return np.round(x, ROUND_DIGITS)


def generate_with_retry(
    metric: str,
    target_mean: float,
    target_sd: float,
    size: int,
    rng: np.random.Generator,
) -> np.ndarray:
    """在 SD ±10% 容差内重试，最多 MAX_RETRY 次，仍失败抛 RuntimeError。"""
    last_actual_sd: float | None = None
    for attempt in range(1, MAX_RETRY + 1):
        # 用主 rng 派生独立子 rng，保证整体可复现 + 每次 attempt 不同
        seed = int(rng.integers(0, 2**32 - 1))
        sub_rng = np.random.default_rng(seed)
        x = generate_one_group(target_mean, target_sd, size, sub_rng)
        actual_sd = float(np.std(x, ddof=1))
        last_actual_sd = actual_sd
        rel_err = abs(actual_sd - target_sd) / target_sd
        logger.debug(
            "metric=%s attempt=%d size=%d target_sd=%.6f actual_sd=%.6f rel_err=%.4f",
            metric, attempt, size, target_sd, actual_sd, rel_err,
        )
        if rel_err <= SD_TOLERANCE:
            return x
    raise RuntimeError(
        f"无法在容差内还原 SD: metric={metric}, target_sd={target_sd}, "
        f"last_actual_sd={last_actual_sd}, retries={MAX_RETRY}"
    )
```

- [ ] **Step 3.4: 跑测试**

Run: `pytest tests/test_generate.py -v`
Expected: 全部 PASS

- [ ] **Step 3.5: Commit**

```bash
git add scripts/generate_excel_random_data.py tests/test_generate.py
git commit -m "feat: add normal-distribution generators with SD tolerance retry"
```

---

## Task 4: 统计回算层（TDD）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（新增 `compute_stats`）
- Create: `tests/test_stats.py`

- [ ] **Step 4.1: 写失败测试**

`tests/test_stats.py`:
```python
from __future__ import annotations

import numpy as np
from scipy import stats as sp_stats

from scripts.generate_excel_random_data import LEVENE_ALPHA, compute_stats


def test_compute_stats_equal_var_branch():
    rng = np.random.default_rng(0)
    g1 = rng.normal(10.0, 2.0, 50)
    g2 = rng.normal(11.0, 2.0, 50)  # similar variance → expect equal_var True
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
    g2 = rng.normal(10.0, 10.0, 30)  # very different variance → expect Welch
    _, _, _, _, p, equal_var = compute_stats(g1, g2)

    levene_p = sp_stats.levene(g1, g2, center="median").pvalue
    assert equal_var == bool(levene_p >= LEVENE_ALPHA)
    expected_p = sp_stats.ttest_ind(g1, g2, equal_var=equal_var).pvalue
    assert p == float(expected_p)
    # And in this constructed scenario we actually expect equal_var=False
    assert equal_var is False
```

- [ ] **Step 4.2: 跑测试，确认失败**

Run: `pytest tests/test_stats.py -v`
Expected: ImportError

- [ ] **Step 4.3: 实现 `compute_stats`**

追加：
```python
def compute_stats(
    g1: np.ndarray, g2: np.ndarray, alpha: float = LEVENE_ALPHA
) -> tuple[float, float, float, float, float, bool]:
    """返回 (mean1, sd1, mean2, sd2, p_value, equal_var)。"""
    mean1 = float(np.mean(g1))
    sd1 = float(np.std(g1, ddof=1))
    mean2 = float(np.mean(g2))
    sd2 = float(np.std(g2, ddof=1))
    levene_p = float(sp_stats.levene(g1, g2, center="median").pvalue)
    equal_var = levene_p >= alpha
    p_value = float(sp_stats.ttest_ind(g1, g2, equal_var=equal_var).pvalue)
    logger.debug(
        "compute_stats levene_p=%.4f equal_var=%s p_value=%.4f", levene_p, equal_var, p_value
    )
    return mean1, sd1, mean2, sd2, p_value, equal_var
```

- [ ] **Step 4.4: 跑测试**

Run: `pytest tests/test_stats.py -v`
Expected: 全部 PASS

- [ ] **Step 4.5: Commit**

```bash
git add scripts/generate_excel_random_data.py tests/test_stats.py
git commit -m "feat: add compute_stats with Levene -> Student/Welch branch"
```

---

## Task 5: 列布局（TDD）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（新增 `Layout` + `compute_layout`）
- Create: `tests/test_layout.py`

- [ ] **Step 5.1: 写失败测试**

`tests/test_layout.py`:
```python
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
```

- [ ] **Step 5.2: 跑测试，确认失败**

Run: `pytest tests/test_layout.py -v`
Expected: ImportError

- [ ] **Step 5.3: 实现 `Layout` + `compute_layout`**

```python
@dataclass(slots=True, frozen=True)
class Layout:
    group1_cols: range
    group2_cols: range
    stat_cols: range


def compute_layout(n: int, m: int) -> Layout:
    if n < 2 or m < 2:
        raise ValueError(f"N、M 必须 ≥2，收到 N={n}、M={m}")
    g1_start = GROUP1_START_COL
    g1_end = g1_start + n - 1
    g2_start = g1_end + 2  # 空 1 列
    g2_end = g2_start + m - 1
    stat_start = g2_end + 2  # 空 1 列
    stat_end = stat_start + 4  # 5 个统计列
    return Layout(
        group1_cols=range(g1_start, g1_end + 1),
        group2_cols=range(g2_start, g2_end + 1),
        stat_cols=range(stat_start, stat_end + 1),
    )
```

- [ ] **Step 5.4: 跑测试**

Run: `pytest tests/test_layout.py -v`
Expected: 全部 PASS

- [ ] **Step 5.5: Commit**

```bash
git add scripts/generate_excel_random_data.py tests/test_layout.py
git commit -m "feat: add Layout + compute_layout (dynamic column offsets)"
```

---

## Task 6: 写入层（无独立单测，由 Task 7 端到端测试覆盖）

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（新增 `write_row`）

- [ ] **Step 6.1: 实现 `write_row`**

```python
def write_row(
    ws: Worksheet,
    row: RowSpec,
    g1: np.ndarray,
    g2: np.ndarray,
    stats: tuple[float, float, float, float, float, bool],
    layout: Layout,
) -> None:
    mean1, sd1, mean2, sd2, p_value, _equal_var = stats

    # 第 1 行：组 1 的列序号 1..N-1，最后一列写 "<N>（N）"
    for idx, col in enumerate(layout.group1_cols, start=1):
        if col == layout.group1_cols.stop - 1:
            ws.cell(row=1, column=col).value = f"{row.group1.n}（N）"
        else:
            ws.cell(row=1, column=col).value = idx

    # 第 1 行：组 2 的列序号 1..M-1，最后一列写 "<M>（M）"
    for idx, col in enumerate(layout.group2_cols, start=1):
        if col == layout.group2_cols.stop - 1:
            ws.cell(row=1, column=col).value = f"{row.group2.n}（M）"
        else:
            ws.cell(row=1, column=col).value = idx

    # 第 1 行：5 个统计列表头（空时补写）
    for header, col in zip(STAT_HEADERS, layout.stat_cols, strict=True):
        existing = ws.cell(row=1, column=col).value
        if existing in (None, ""):
            ws.cell(row=1, column=col).value = header

    # 数据行
    for value, col in zip(g1.tolist(), layout.group1_cols, strict=True):
        ws.cell(row=row.row_index, column=col).value = value
    for value, col in zip(g2.tolist(), layout.group2_cols, strict=True):
        ws.cell(row=row.row_index, column=col).value = value

    # 统计列
    stat_values = (mean1, sd1, mean2, sd2, p_value)
    for value, col in zip(stat_values, layout.stat_cols, strict=True):
        ws.cell(row=row.row_index, column=col).value = float(value)
```

- [ ] **Step 6.2: Commit（在 Task 7 写完端到端测试后再确认通过）**

```bash
git add scripts/generate_excel_random_data.py
git commit -m "feat: add write_row (dynamic column writes + group size markers)"
```

---

## Task 7: CLI 入口 + 端到端测试

**Files:**
- Modify: `scripts/generate_excel_random_data.py`（实现 `main`）
- Create: `tests/test_cli.py`

- [ ] **Step 7.1: 写失败的端到端测试**

`tests/test_cli.py`:
```python
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

    # A-H 列保持完全不变（只比对前几行）
    for r in range(1, ws_in.max_row + 1):
        for c in range(1, 9):  # A-H = 1..8
            assert ws_out.cell(row=r, column=c).value == ws_in.cell(row=r, column=c).value, (r, c)

    col_S = column_index_from_string("S")
    col_AF = column_index_from_string("AF")
    col_R = column_index_from_string("R")
    col_AE = column_index_from_string("AE")
    col_AG = column_index_from_string("AG")

    # 第 1 行：R1 == "10（N）"，AE1 == "12（M）"
    assert ws_out.cell(row=1, column=col_R).value == "10（N）"
    assert ws_out.cell(row=1, column=col_AE).value == "12（M）"

    # 数据行（第 2-4 行）：S 与 AF 列必须为空
    for r in range(2, 5):
        assert ws_out.cell(row=r, column=col_S).value is None
        assert ws_out.cell(row=r, column=col_AF).value is None

        # 第一组 10 个浮点数，4 位小数；第二组 12 个浮点数
        g1 = [ws_out.cell(row=r, column=column_index_from_string(L)).value
              for L in ["I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"]]
        g2 = [ws_out.cell(row=r, column=column_index_from_string(L)).value
              for L in ["T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"]]
        assert all(isinstance(v, float) for v in g1), g1
        assert all(isinstance(v, float) for v in g2), g2
        assert all(round(v, 4) == v for v in g1)
        assert all(round(v, 4) == v for v in g2)
        assert len(g1) == 10
        assert len(g2) == 12

        # AG-AK 应当与 numpy/scipy 直接计算一致
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
        actual = tuple(
            ws_out.cell(row=r, column=col_AG + i).value for i in range(5)
        )
        for got, exp in zip(actual, expected, strict=True):
            assert got == pytest.approx(exp, rel=1e-9, abs=1e-9), (got, exp)


def test_cli_seed_reproducible(out_path: Path, tmp_path: Path):
    out2 = tmp_path / "second.xlsx"
    shutil.copy(SAMPLE_XLSX, out2)

    main(["--input", str(out_path), "--output", str(out_path), "--seed", "42"])
    main(["--input", str(out2), "--output", str(out2), "--seed", "42"])

    wb1 = openpyxl.load_workbook(out_path, data_only=True).active
    wb2 = openpyxl.load_workbook(out2, data_only=True).active
    for r in range(1, wb1.max_row + 1):
        for c in range(1, wb1.max_column + 1):
            assert wb1.cell(row=r, column=c).value == wb2.cell(row=r, column=c).value, (r, c)
```

- [ ] **Step 7.2: 跑测试，确认失败**

Run: `pytest tests/test_cli.py -v`
Expected: `main` 还是 `NotImplementedError`，失败

- [ ] **Step 7.3: 实现 `main`**

替换 `scripts/generate_excel_random_data.py` 末尾的 `main()`：

```python
def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="按 Excel 表格中的目标统计量生成两组正态分布原始数据并回写。"
    )
    parser.add_argument(
        "--input", type=Path, default=Path("20260509-随机数生成.xlsx"),
        help="输入 Excel 文件路径（默认 20260509-随机数生成.xlsx）",
    )
    parser.add_argument(
        "--output", type=Path, default=None,
        help="输出 Excel 文件路径（默认与 --input 相同，即就地修改）",
    )
    parser.add_argument("--sheet", default=None, help="指定 sheet 名（默认第一个 sheet）")
    parser.add_argument("--seed", type=int, default=None, help="随机种子（不传则随机生成并打印）")
    parser.add_argument("-v", "--verbose", action="store_true", help="启用 DEBUG 级别日志")
    return parser


def _process_rows(
    ws: Worksheet, specs: Iterable[RowSpec], rng: np.random.Generator
) -> int:
    """对每行 spec 执行 generate→compute→write，返回成功处理的行数。"""
    count = 0
    for spec in specs:
        layout = compute_layout(spec.group1.n, spec.group2.n)
        # 给每行派生独立子 rng，便于追踪
        seed1 = int(rng.integers(0, 2**32 - 1))
        seed2 = int(rng.integers(0, 2**32 - 1))
        g1 = generate_with_retry(
            metric=f"{spec.metric}/G1",
            target_mean=spec.group1.mean,
            target_sd=spec.group1.sd,
            size=spec.group1.n,
            rng=np.random.default_rng(seed1),
        )
        g2 = generate_with_retry(
            metric=f"{spec.metric}/G2",
            target_mean=spec.group2.mean,
            target_sd=spec.group2.sd,
            size=spec.group2.n,
            rng=np.random.default_rng(seed2),
        )
        stats_tuple = compute_stats(g1, g2)
        write_row(ws, spec, g1, g2, stats_tuple, layout)
        count += 1
        logger.info(
            "处理完成 row=%d metric=%s p_value=%.4f equal_var=%s",
            spec.row_index, spec.metric, stats_tuple[4], stats_tuple[5],
        )
    return count


def main(argv: list[str] | None = None) -> int:
    args = _build_arg_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    seed = args.seed if args.seed is not None else secrets.randbits(32)
    logger.info("使用随机种子 seed=%d", seed)
    rng = np.random.default_rng(seed)

    input_path = args.input
    output_path = args.output or input_path
    if not input_path.exists():
        logger.error("输入文件不存在: %s", input_path)
        return 2

    specs = read_specs(input_path, args.sheet)
    if not specs:
        logger.warning("没有有效行可处理，退出")
        return 0

    wb = load_workbook(input_path)
    ws: Worksheet = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
    n = _process_rows(ws, specs, rng)
    wb.save(output_path)
    logger.info("已写入 %d 行 → %s", n, output_path)
    return 0
```

- [ ] **Step 7.4: 跑端到端测试**

Run: `pytest tests/test_cli.py -v`
Expected: 全部 PASS

- [ ] **Step 7.5: 跑全量测试**

Run: `pytest -v`
Expected: 全部 PASS

- [ ] **Step 7.6: 实际跑一次脚本，肉眼复核**

Run:
```bash
python scripts/generate_excel_random_data.py --seed 42 --output /tmp/out.xlsx
python -c "
import openpyxl
wb = openpyxl.load_workbook('/tmp/out.xlsx', data_only=True)
ws = wb.active
for r in range(1, 5):
    print([ws.cell(row=r, column=c).value for c in range(1, 38)])
"
```
Expected: 第 1 行 R1 是 `'10（N）'`、AE1 是 `'12（M）'`；数据行四位小数；S/AF 为空；AG-AK 是统计值。

- [ ] **Step 7.7: 跑 ruff**

Run:
```bash
ruff format scripts tests
ruff check scripts tests --fix
```
Expected: 无错误。

- [ ] **Step 7.8: Commit**

```bash
git add scripts/generate_excel_random_data.py tests/test_cli.py
git commit -m "feat: implement CLI main + end-to-end coverage"
```

---

## Task 8: 完成校验

- [ ] **Step 8.1: OpenSpec 状态校验**

Run: `openspec status --change add-excel-random-data-generator`
Expected: 4/4 artifacts complete

- [ ] **Step 8.2: 把 `openspec/changes/.../tasks.md` 中所有 `- [ ]` 改成 `- [x]`**

按 OpenSpec 工作流要求，apply 阶段需要把每个完成的 task 勾上。

- [ ] **Step 8.3: 跑 `git status` / `git diff --stat` 复核改动范围**

Run: `git status && git diff --stat HEAD`
Expected: 仅涉及 `requirements.txt` / `scripts/` / `tests/` / `docs/superpowers/` / `openspec/changes/...`。

- [ ] **Step 8.4: Commit final tasks.md update**

```bash
git add openspec/changes/add-excel-random-data-generator/tasks.md
git commit -m "chore(openspec): mark add-excel-random-data-generator tasks done"
```

---

## Self-Review

- **Spec coverage（对照 [spec.md](../../../openspec/changes/add-excel-random-data-generator/specs/excel-random-data-generator/spec.md)）：**
  - 读取目标统计量 → Task 2
  - 生成两组正态分布原始数据 + 容差重试 → Task 3
  - 按动态列布局写入 → Task 5（layout）+ Task 6（write_row）
  - 表头标注样本量 → Task 6（`write_row` 中的 `<N>（N）` / `<M>（M）` 写法）
  - 用生成数据回算统计量并算 p 值 → Task 4
  - 写入 AG-AK → Task 6（stat_cols 写入）
  - CLI 与可复现性 → Task 7（main + seed 测试）
  - 不破坏未涉及单元格 → Task 7 端到端测试（A-H 不变 / S / AF 为空）

- **Placeholder 扫描：** 没有 TBD/TODO，所有代码片段是完整可粘贴的。

- **类型/命名一致：** `GroupSpec`、`RowSpec`、`Layout`、`compute_layout`、`generate_one_group`、`generate_with_retry`、`compute_stats`、`write_row`、`main` 全部一致。

如发现遗漏会内联修正后继续。

---

## Execution Handoff

计划已保存到 [docs/superpowers/plans/2026-05-09-excel-random-data-generator.md](docs/superpowers/plans/2026-05-09-excel-random-data-generator.md)。

将使用 **superpowers:executing-plans / 内联执行** 推进：当前会话内按 Task 1 → Task 8 顺序、每个 Task 内严格 TDD（写失败测试 → 跑失败 → 写实现 → 跑成功 → commit），出错时切到 `superpowers:systematic-debugging`。
