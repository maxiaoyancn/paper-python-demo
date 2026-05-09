# 项目架构（paper-python-demo）

> 本文件由 OpenSpec archive 流程同步维护，记录全局架构与关键能力的现状。详细需求请见 [openspec/specs/](openspec/specs/) 下的 capability spec；变更历史见 [openspec/changes/archive/](openspec/changes/archive/)。

## 项目定位

Python 脚本工作区（非长期服务、非 DDD 项目）。核心交付物是单文件 CLI 脚本与配套测试。

## 工作流

- **OpenSpec** 驱动需求管理：`/opsx:propose` → `/opsx:apply` → `/opsx:archive` 三阶段。
- **Superpowers** 在 apply 阶段强制嵌套：`writing-plans` → `test-driven-development` → 出错时 `systematic-debugging`。
- 详见根目录 [CLAUDE.md](CLAUDE.md) §二（OpenSpec 与 Superpowers 协同规约）。

## 当前 capabilities

### excel-random-data-generator

按 Excel 表格中每行的目标统计量生成 K 组（K≥2）正态分布原始数据，并回算 mean/SD 与多组统计 p 值。

**入口**：[scripts/generate_excel_random_data.py](scripts/generate_excel_random_data.py)（CLI：`--input` 必填、`--output` / `--sheet` / `--seed` / `--verbose`）。
**测试**：[tests/](tests/)，31 个用例覆盖 read / generate / stats / layout / CLI 五层 + K=2/K=3 端到端。
**Spec**：[openspec/specs/excel-random-data-generator/spec.md](openspec/specs/excel-random-data-generator/spec.md)。

#### 接口 / 列布局约定

输入区（脚本不修改）：

```
A 列      指标名
B-D       第 1 组 (N, μ, σ)
E-G       第 2 组
H-J       第 3 组（K≥3 时）
…
col(2+3K) 原始数据小数点后位数（脚本扫描行 1 表头中含"位数"的列定位它）
col(3+3K) 备注列（可选，脚本不读不写）
```

数据区（脚本写入）：

```
data_start = decimals_col + 2
G_i 占 N_i 列，组间空 1 列
G_i 最后一格表头覆写为 "<N_i>（G<i>）"
```

统计区（数据区结束 + 1 空列起，按顺序）：

| 列 | 含义 | 备注 |
|---|---|---|
| 2K 列 | μ_i, σ_i 交替（i=1..K） | 4 位小数 |
| 1 列 | Levene p（center=median, Brown-Forsythe） | > 0.05 视为方差齐 |
| 1 列 | Shapiro-Wilk min p（每组 SW 取最小） | > 0.05 视为全部正态 |
| 1 列 | 整体 p | 齐 → ANOVA(`f_oneway`)；不齐 → KW(`kruskal`) |
| C(K,2) 列 | 两两 raw p | 顺序：(1,2)(1,3)(2,3)(1,4)... |
| C(K,2) 列 | 两两 Q-value | 齐+全正态 → Tukey HSD（留空）；否则 → Welch t + Bonferroni |

#### 关键决策（参见已归档 changes 的 design.md）

- **位数列动态识别**：扫描行 1 表头含子串"位数"的列，K = (decimals_col - 2) / 3，无需 CLI 显式传 K。
- **生成器精确还原 μ/σ**：`x = μ + σ * (z - z.mean()) / z.std(ddof=1)` 后 round 到行级 decimals；SD 在 ±10% 容差内可重试 5 次，仍失败抛 `RuntimeError`。
- **统计层 4 个独立函数**：`compute_levene`、`compute_shapiro_min`（min 聚合 + N<3 容错）、`compute_overall`、`compute_pairwise`（双分支 dispatcher）。
- **Tukey HSD vs Welch+Bonferroni 二选一**：仅当 Levene p > 0.05 **且** SW min p > 0.05 才走 Tukey；否则一律 Welch+Bonferroni（封顶 1.0）。
- **可复现性**：CLI `--seed` 固定根种子，每行/每组派生独立子种子；同 seed 同输入完全一致。

#### 依赖

- `openpyxl >= 3.1, < 4`（IO）
- `numpy >= 1.26, < 3`（生成 + round）
- `scipy >= 1.11, < 2`（含 `levene` / `shapiro` / `f_oneway` / `kruskal` / `tukey_hsd` / `ttest_ind`）
- 测试：`pytest >= 8`、`ruff >= 0.5`

## 历史

- **2026-05-09** [add-excel-random-data-generator](openspec/changes/archive/2026-05-09-add-excel-random-data-generator/)：v1 落地，仅支持 K=2，输出 5 列统计（μ1/σ1/μ2/σ2/p）。
- **2026-05-09** [extend-to-multi-group-stats](openspec/changes/archive/2026-05-09-extend-to-multi-group-stats/)：扩展为 K 组 + 行级 decimals + 完整多组统计流水线（Levene → SW → ANOVA/KW → Tukey HSD / Welch+Bonferroni）。BREAKING：列布局重定义，统计列从 5 个变为 2K + 3 + 2·C(K,2) 个。
