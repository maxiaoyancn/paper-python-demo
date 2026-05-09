## Context

- v1 已 archive，capability spec 在 [openspec/specs/excel-random-data-generator/spec.md](../../specs/excel-random-data-generator/spec.md)。本次为该 capability 的全面演进。
- 用户提供的样本 [20260509-随机数生成.xlsx](../../../20260509-随机数生成.xlsx) 已被调整为 v2 输入示例：H1 = `'原始数据小数点后位数'`、H2-H4 = 4；I1 仍是旧版备注文本（不再被脚本读取）；J 列起为数据区（v1 是 I 起，v2 后挪 1 列）。
- 用户期望流水线（按需求顺序）：
  1. K 组数据按目标 (N、μ、σ) 生成正态分布原始数据，按行级 `decimals` round。
  2. 用 Levene's test 判方差齐性 → 真实 p 值写一列。
  3. 用 Shapiro-Wilk 判每组正态性 → 取 min(p) 写一列。
  4. 整体差异：方差齐 → ANOVA；不齐 → Kruskal-Wallis。结果写一列。
  5. 两两对比：方差齐 + 全部正态 → Tukey HSD（无校正后值）；否则 → Welch t + Bonferroni（含 raw 与 Q-value）。
- 中文输出（CLAUDE.md §六）；Python 准则（[docs/lang/lang-python.md](../../../docs/lang/lang-python.md)）。
- 依赖现状：scipy 1.11+ 已含 `scipy.stats.tukey_hsd`、`shapiro`、`levene`、`f_oneway`、`kruskal`、`ttest_ind`，**无需新增依赖**。

## Goals / Non-Goals

**Goals:**
- 任意 K（K≥2）支持，K 由每行数据**自适应**确定（不强制全表统一）。
- 行级 decimals 决定原始数据的小数位数；统计回算列（μ、σ、p、Q）固定 4 位小数。
- 在原 SD ±10% 容差内重试机制保留。
- 两两对比的列顺序按组对索引字典序：(1,2)(1,3)(2,3)(1,4)(2,4)(3,4)…（即 C(K,2) 个组合按 itertools.combinations 默认顺序）。
- 写入范围只覆盖"位数列+1（空列）→ 数据区 → 统计列"；A 列至位数列保留。
- 同一 `--seed` + 同一输入 → 完全可复现。

**Non-Goals:**
- 不支持 K=1（无组间对比）。读取阶段报错跳过该行。
- 不支持非正态分布生成（用户场景只要正态）。
- 不做 sheet 内多 sheet 批量、行内 K 统一性校验。
- 不实现自定义 α（Levene/SW 都用 0.05 硬阈值）；之后如需可加 CLI。
- 不重写已 archive 的 v1 spec/tests 文件——v2 提交时**整体替换** spec.md，并把 `tests/test_*.py` 全量重写。

## Decisions

### 1. 列布局：动态识别"位数列"
- 表头扫描：行 1 第一个 cell value（去除前后空白）含子串 `位数` 的列即为**位数列**（`decimals_col`）。
- K = `(decimals_col - 2) // 3`；要求 `decimals_col == 2 + 3K` 且 K ≥ 2，否则 `logger.warning` 跳过该 sheet。
- 数据起始列 = `decimals_col + 2`（位数列后空 1 列再开始；这一空列允许填备注，脚本不读不写）。
- 拒选方案：让 CLI 显式传 `--groups K`。理由：与"K 由数据自适应"诉求冲突；当不同行 K 不同时无解。

### 2. 行级 K 与 N 的不一致问题
- 每行独立读取自己的 K 与 (N₁, …, N_K)。`compute_layout(group_sizes: tuple[int, ...]) -> Layout` 接收一个变长元组返回该行的列布局。
- 不同行的 layout 完全可能不同：当行 i 的 K=2 而行 i+1 的 K=3 时，行 1 的 σ 列、统计列与行 2 的位置不同。`write_row` 只在该行 row_index 对应的列写入，互不干扰。
- 第 1 行的样本量表头 `<count>（Gi）` 同样按"该行 row 的 layout 写到第 1 行的对应列"——意味着**最后一行被写入的内容会留在第 1 行**。这与 v1 的隐患同源；当所有行 K 与 N 都一致时无影响。在 README 里加一条："建议同一 sheet 内所有指标的 K 与各组 N 保持一致"。

### 3. 表头标注样本量：`<count>（Gi）`
- 用户已确认统一用 `Gi`（与 v1 的 N/M 不再兼容）。
- 第 1 行的统计列表头：`第i组均值`、`第i组SD值`、`Levene p`、`Shapiro-Wilk min p`、`整体 p（ANOVA/KW）`、`Gi-Gj raw p`、`Gi-Gj Q-value`。表头若为空则脚本写入；非空则保留（防覆盖用户人工备注）。

### 4. 生成器：按行级 decimals round
- 现有 `generate_one_group(target_mean, target_sd, size, rng)` 增加 `decimals: int` 参数，把 `np.round(x, ROUND_DIGITS)` 改为 `np.round(x, decimals)`。
- decimals 校验：≥0 的整数（允许 0 表示整数列）。
- ±10% 容差仍以"四舍五入后 SD 与 target 的相对误差"为准，重试上限沿用 `MAX_RETRY = 5`。

### 5. 统计层：5 个新函数 + 1 个 dispatcher
```
compute_overall(groups: list[ndarray], levene_p: float, alpha: float) -> tuple[float, str]
    # returns (overall_p, method_label) where method_label ∈ {"ANOVA", "KW"}

compute_pairwise(groups: list[ndarray], levene_p: float, sw_min_p: float, alpha: float)
    -> tuple[list[float], list[float | None], str]
    # returns (raw_ps, q_values, method_label)
    #   q_values are None when Tukey branch (留空)
```
- Levene：`scipy.stats.levene(*groups, center='median')`（即 Brown-Forsythe 变种）。**为什么用 `median` 而不是经典 Levene 的 `mean`**：① scipy / SPSS / R 的默认值，业界主流；② 即便我们生成的是严格正态数据，用户后续可能拿这份输出与论文里的**真实实测数据**做交叉对照，那时分布偏态不可控，median 在偏态/尾部不规则时更稳健；③ 对单点极端值不敏感。代价是对正态对称数据**功效略低于** `mean` 版本。如果未来需要切换，把这一处的 `center` 改成 `'mean'` 或暴露成 CLI 参数即可。
- ANOVA：`scipy.stats.f_oneway(*groups)`
- Kruskal-Wallis：`scipy.stats.kruskal(*groups)`
- Tukey HSD：`scipy.stats.tukey_hsd(*groups).pvalue` 是 K×K 矩阵；按 `combinations(range(K), 2)` 顺序提取上三角对应的 p。
- Welch t：`scipy.stats.ttest_ind(g_i, g_j, equal_var=False).pvalue`
- Bonferroni：`q_i = min(1.0, raw_p_i * C(K,2))`

### 6. 数据流入 statistics：用四舍五入后的最终数据，不是预 round 数据
- 与 v1 一致——所有统计量都基于 round 后的真实 cell 值，确保 Excel 里看到的 μ/σ/p 与 cell 数据**完全自洽**。

### 7. K=2 时的兼容
- K=2 走通用流水线：Levene → SW → ANOVA(齐) / KW(不齐) → 1 对（1,2）的 raw + Q。
- 与 v1 的 5 列输出（μ1σ1μ2σ2p）**完全不兼容**——v2 输出 9 列（含 Levene/SW/Overall/raw/Q）。
- v1 的 `compute_stats` 函数被新 `compute_overall + compute_pairwise` 替代。

### 8. 错误处理与日志
- 读取阶段非法行（K<2、N<2、SD≤0、decimals<0、缺失）→ `logger.warning` 跳过。
- 生成阶段重试 5 次仍超 ±10% → `RuntimeError` 抛出，整个进程终止（不 swallow）。
- 数值层 NaN（如 Levene 在某组方差为 0 时返回 NaN）→ 视作"不齐"分支处理（`p < 0.05`）；写入 cell 时把 NaN 转为空字符串。

## Risks / Trade-offs

- [Tukey HSD 在 K=2 时退化] → 用户期望 Tukey 是 K≥3 的"标准多重比较"。K=2 + 走 Tukey 分支 → `tukey_hsd` 仍可工作，结果与 Welch 类似但同质。可接受。
- [Shapiro-Wilk 对小样本（N<3）报错或不可靠] → SW 要求 N ≥ 3。当任意组 N<3 时 `scipy.stats.shapiro` 返回 NaN 或抛错；策略：将该组 SW p 视为 0（视作不正态），同时 logger.warning。但用户场景 N≥10，无实际风险。
- [ANOVA 假设要求方差齐 + 正态] → 当 Levene p ≤ 0.05 时本就走 KW；当 SW min p ≤ 0.05（任一组不正态）但 Levene p > 0.05 时——按用户描述仍走 ANOVA（用户 #7 只看 Levene），但**两两对比走 Welch+Bonferroni**（按 #11 的"齐且正态"才用 Tukey）。这种"整体 ANOVA / 两两 Welch+Bonferroni"组合方式与用户描述完全一致。
- [Bonferroni 在 K=2 时是恒等] → C(2,2) = 1，q = min(1, raw × 1) = raw。用户的样表 K=2，q 列实际等于 raw 列。这是 Bonferroni 在两组时的数学事实，不是 bug。
- [行级 K 不同 → 第 1 行表头互相覆盖] → 文档约束 + 实际多数场景 K 一致；不阻塞。
- [decimals 列识别启发式 "包含'位数'"] → 健壮性受表头文案影响。如果用户填了"小数位"而非"位数"，会找不到。可在 design 里要求**严格匹配子串"位数"**；不灵活但简单。
