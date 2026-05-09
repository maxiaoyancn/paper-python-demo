## MODIFIED Requirements

### Requirement: 在表头标注样本量

脚本 SHALL 把每行 K 组数据的列表头**全部**写为 `<letter><idx>` 形式：第 i 组用第 i 个英文大写字母（i=1..K → A..Z），列内 idx 从 1 起递增到 N_i。脚本 SHALL NOT 再使用 v2 的"裸数字 + 末位 `<N>（Gi）`"混合写法；样本量 N_i 可从该组最大 idx 看出。所有 K（包括 K=2）SHALL 统一使用新字母格式。

#### Scenario: K=2 时第 1 组 N=10
- **WHEN** K=2、第 1 组样本量为 10
- **THEN** 第 1 行第 1 组所占的 10 列依次写入 `"A1", "A2", "A3", …, "A10"`

#### Scenario: K=2 时第 2 组 M=12
- **WHEN** K=2、第 2 组样本量为 12
- **THEN** 第 1 行第 2 组所占的 12 列依次写入 `"B1", "B2", …, "B12"`

#### Scenario: K=3 时第 3 组样本量
- **WHEN** K=3、第 3 组样本量为 15
- **THEN** 第 1 行第 3 组所占的 15 列依次写入 `"C1", "C2", …, "C15"`

#### Scenario: K=4 时使用 D 字母
- **WHEN** K=4
- **THEN** 第 4 组列表头使用字母 D（如 `"D1", …, "D<N_4>"`）

### Requirement: 两两对比 p 值与多重比较校正

脚本 SHALL 计算 K 组之间的两两对比 p 值并写入两组连续列：先 C(K,2) 个 raw p 列（按 `itertools.combinations(range(K), 2)` 顺序，即 (1,2)、(1,3)、(2,3)、(1,4)、(2,4)、(3,4)、…），再 C(K,2) 个校正后 Q-value 列（顺序与 raw 列一一对应）。两两对比方法 SHALL 由方差齐性 + 全组正态性联合决定：当 Levene p > 0.05 **且** Shapiro-Wilk min p > 0.05（齐 + 全部正态）时使用 `scipy.stats.tukey_hsd`，否则使用 Welch t（`ttest_ind(equal_var=False)`）+ Bonferroni 校正。所有 p / Q 值 SHALL 保留 4 位小数。两两对比的列**表头** SHALL 使用字母对形式 `<letter_i>-<letter_j>`（如 `A-B raw p`、`A-B Q-value`），不再使用 `Gi-Gj` 形式。

#### Scenario: 方差齐 + 全部正态 → Tukey HSD
- **WHEN** Levene p > 0.05 且 Shapiro-Wilk min p > 0.05
- **THEN** raw p 列写入 `tukey_hsd(*groups).pvalue` 上三角部分按对索引顺序提取的值；Q-value 列**留空**（写 None / 空字符串）

#### Scenario: 方差不齐 → Welch+Bonferroni
- **WHEN** Levene p ≤ 0.05
- **THEN** raw p 列依次写入每对 (i, j) 的 `ttest_ind(g_i, g_j, equal_var=False).pvalue`；Q-value 列写入 `min(1.0, raw_p × C(K,2))`

#### Scenario: 方差齐但有组不正态 → Welch+Bonferroni
- **WHEN** Levene p > 0.05 但 Shapiro-Wilk min p ≤ 0.05
- **THEN** 走 Welch+Bonferroni 分支（不使用 Tukey），raw 与 Q 写法同上

#### Scenario: K=3 时 raw 与 Q 表头使用字母对
- **WHEN** K=3
- **THEN** raw 列表头依次为 `"A-B raw p"`、`"A-C raw p"`、`"B-C raw p"`；Q-value 列表头为 `"A-B Q-value"`、`"A-C Q-value"`、`"B-C Q-value"`

#### Scenario: K=2 时单对 Bonferroni
- **WHEN** K=2 且走 Welch+Bonferroni 分支
- **THEN** Q_(1,2) = min(1.0, raw_(1,2) × 1) = raw_(1,2)（即 Bonferroni 在两组场景下与 raw 数值相等）

### Requirement: 按统计列布局写入回算结果

脚本 SHALL 在数据区结束后空 1 列，从下一列起按以下顺序连续写入统计列（K 组场景）：μ_1、σ_1、μ_2、σ_2、…、μ_K、σ_K、Levene p、**是否方差齐**、Shapiro-Wilk min p、**是否正态**、整体 p、Pairwise raw p × C(K,2)、Pairwise Q-value × C(K,2)。所有数值 SHALL 保留 4 位小数。脚本 SHALL 在第 1 行同时写入对应表头：`<letter_i> 组均值`、`<letter_i> 组SD值`（i=1..K）、`Levene p`、`是否方差齐`、`Shapiro-Wilk min p`、`是否正态`、`整体 p（ANOVA/KW）`、`<letter_i>-<letter_j> raw p`、`<letter_i>-<letter_j> Q-value`。如果第 1 行某统计列已有非空表头，SHALL 保留不覆盖。

#### Scenario: K=2 时统计列总数
- **WHEN** K=2
- **THEN** 统计列共 11 列：A 组均值、A 组SD值、B 组均值、B 组SD值、Levene p、**是否方差齐**、Shapiro-Wilk min p、**是否正态**、整体 p、A-B raw p、A-B Q-value

#### Scenario: K=3 时统计列总数
- **WHEN** K=3
- **THEN** 统计列共 17 列：A/B/C 三组的均值与 SD（共 6 列）、Levene p、**是否方差齐**、Shapiro-Wilk min p、**是否正态**、整体 p、3 个 raw p 列、3 个 Q-value 列

#### Scenario: K=4 时统计列总数
- **WHEN** K=4
- **THEN** 统计列共 25 列：A/B/C/D 四组的均值与 SD（共 8 列）、Levene p、**是否方差齐**、Shapiro-Wilk min p、**是否正态**、整体 p、6 个 raw p 列、6 个 Q-value 列

#### Scenario: 已有表头不覆盖
- **WHEN** 第 1 行某统计列已有非空文本
- **THEN** 脚本 SHALL 保留原文本，不写入默认表头

## ADDED Requirements

### Requirement: 是否方差齐 Y/N 列

脚本 SHALL 在 Levene p 列之后**紧跟一列** "是否方差齐"，写入单字符 `Y` / `N`：当 Levene p > 0.05（方差齐）时写 `Y`；当 Levene p ≤ 0.05（方差不齐）时写 `N`。该列与 Levene p 列**并存**——前者是结论、后者是数值。判定 SHALL 复用 `compute_levene` 已经计算的 `equal_var` 布尔值，不重复执行 Levene 检验。

#### Scenario: 方差齐
- **WHEN** Levene p > 0.05
- **THEN** "是否方差齐"列写 `"Y"`

#### Scenario: 方差不齐
- **WHEN** Levene p ≤ 0.05
- **THEN** "是否方差齐"列写 `"N"`

#### Scenario: Levene 返回 NaN
- **WHEN** Levene 由于某组方差为 0 等原因返回 NaN
- **THEN** "是否方差齐"列写 `"N"`（与 `compute_levene` 把 NaN 视作不齐的策略一致）

### Requirement: 是否正态 Y/N 列

脚本 SHALL 在 Shapiro-Wilk min p 列之后**紧跟一列** "是否正态"，写入单字符 `Y` / `N`：当 SW min p > 0.05（即所有组都通过正态性检验）时写 `Y`；当 SW min p ≤ 0.05（至少一组未通过）时写 `N`。该列与 Shapiro-Wilk min p 列**并存**——前者是结论、后者是数值。判定 SHALL 复用 `compute_shapiro_min` 已经计算的 `all_normal` 布尔值，不重复执行 SW 检验。

#### Scenario: 全部组正态
- **WHEN** Shapiro-Wilk min p > 0.05
- **THEN** "是否正态"列写 `"Y"`

#### Scenario: 任一组不正态
- **WHEN** Shapiro-Wilk min p ≤ 0.05
- **THEN** "是否正态"列写 `"N"`

#### Scenario: 某组样本量过小被视为不正态
- **WHEN** 某组 N < 3 导致 SW p 被记为 0
- **THEN** "是否正态"列写 `"N"`

### Requirement: K 上限 26

脚本 SHALL 把组数 K 限制在 1..26（即字母 A..Z 范围内）。当推导出 K > 26 时，脚本 SHALL 通过 `logging.error` 报告"K 超过 26 上限"，并以非 0 退出码（3）终止该 sheet 的处理。脚本 SHALL NOT 自动启用双字母（AA、AB…）扩展。

#### Scenario: K=26（边界值）
- **WHEN** sheet 推导出 K=26
- **THEN** 脚本正常处理，第 26 组使用字母 `Z`

#### Scenario: K=27（超限）
- **WHEN** sheet 推导出 K=27
- **THEN** 脚本 SHALL `logger.error` 含"K=27 超过 26 上限"信息，并以退出码 3 终止
