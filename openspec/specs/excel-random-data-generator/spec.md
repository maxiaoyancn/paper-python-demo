# excel-random-data-generator Specification

## Purpose
TBD - created by archiving change add-excel-random-data-generator. Update Purpose after archive.
## Requirements
### Requirement: 读取目标统计量

脚本 SHALL 通过**必填**的 `--input` 参数接收 Excel 文件路径，并接受可选 `--sheet` 参数，从指定 sheet 的第 2 行起逐行读取每个指标的目标统计量。脚本 SHALL 通过扫描第 1 行表头识别"位数列"——表头包含子串 `位数` 的列即为位数列（`decimals_col`），并据此推算该 sheet 的组数 K = (decimals_col - 2) / 3。每行的列布局 SHALL 为：A 列指标名；col(2 + 3(i-1)) 起的连续 3 列为第 i 组（i=1..K）的 N、目标均值、目标 SD；col(2 + 3K) 为该行的"原始数据小数点后位数"（非负整数）。脚本 SHALL 不再硬编码 K=2，K 由数据自适应决定。

#### Scenario: 缺少 --input
- **WHEN** 用户执行 `python scripts/generate_excel_random_data.py` 不带 `--input`
- **THEN** argparse SHALL 立即报错并以退出码 2 终止，错误消息中包含 `--input`；脚本不会处理任何文件

#### Scenario: 指定输入与 sheet
- **WHEN** 用户传入 `--input some.xlsx --sheet Data`
- **THEN** 脚本读取 `some.xlsx` 中名为 `Data` 的 sheet 并按上述列约定解析每个数据行的目标统计量

#### Scenario: K=2 + H 列为位数
- **WHEN** sheet 第 1 行 H 列表头为 `'原始数据小数点后位数'`
- **THEN** 脚本识别 decimals_col=8、K=2，并按 B-D（第 1 组）、E-G（第 2 组）解析每行的 (N, μ, σ)，再读 H 列为该行 decimals

#### Scenario: K=3 时位数列后挪
- **WHEN** sheet 第 1 行 K 列（col 11）表头包含 `位数`
- **THEN** 脚本识别 decimals_col=11、K=3，并按 B-D / E-G / H-J 解析三组的 (N, μ, σ)，再读 K 列为该行 decimals

#### Scenario: 行内目标统计量缺失或非法
- **WHEN** 某一行的任一组 N、μ、σ 为空、非数值、N≤1、SD≤0，或 decimals 缺失/为负、或没有任何一组（K<2）
- **THEN** 脚本 SHALL 通过 `logging.warning` 记录该行被跳过的原因，并继续处理后续行；不阻断整体执行

#### Scenario: 表头未声明位数列
- **WHEN** sheet 第 1 行所有表头都不含子串 `位数`
- **THEN** 脚本 SHALL 通过 `logging.error` 记录"找不到位数列"，并以非 0 退出码终止该 sheet 的处理

### Requirement: 生成两组正态分布原始数据

脚本 SHALL 为每个有效行生成 K 组（K≥2）正态分布的原始数据，第 i 组样本量为 N_i、目标均值为 μ_i、目标 SD 为 σ_i。每个数值的小数位数 SHALL 等于该行的 `decimals` 字段（不再硬编码 4 位）。生成方式 SHALL 保证：未四舍五入前的样本均值严格等于目标均值、样本标准差（ddof=1）严格等于目标 SD；四舍五入后实际 SD 与目标 SD 的相对误差 SHALL 在 ±10% 之内，否则脚本 SHALL 用不同子种子重试，最多 5 次。重试仍失败时 SHALL 抛 `RuntimeError` 而非静默写入。

#### Scenario: SD 在容差内
- **WHEN** 生成并按行 decimals round 后，所有组的实际 SD 与目标 SD 相对误差均 ≤ 10%
- **THEN** 该次生成结果直接被采纳

#### Scenario: 某组 SD 超出容差但重试成功
- **WHEN** 首次生成的某组实际 SD 与目标 SD 相对误差 > 10%
- **THEN** 脚本 SHALL 仅对该组换子种子重试，最多 5 次；任何一次落入容差即采纳

#### Scenario: 重试 5 次仍失败
- **WHEN** 5 次重试后某组实际 SD 仍与目标 SD 相对误差 > 10%
- **THEN** 脚本 SHALL 抛 `RuntimeError`，错误消息中包含指标名、组序号、目标 SD、最近一次实际 SD

#### Scenario: 行内 decimals = 0
- **WHEN** 某行 decimals=0
- **THEN** 该行所有原始数据 SHALL 为整数（无小数）

#### Scenario: 行内 decimals = 6
- **WHEN** 某行 decimals=6
- **THEN** 该行所有原始数据 SHALL 保留 6 位小数

### Requirement: 按动态列布局写入原始数据

脚本 SHALL 把 K 组原始数据按以下动态列布局写入：数据起始列 = `decimals_col + 2`（位数列后空 1 列）；第 i 组（i=1..K）数据连续占 N_i 列；相邻两组之间空 1 列。脚本 SHALL NOT 硬编码任何具体列号——必须按 (decimals_col, N_1, …, N_K) 计算每组的起止列。

#### Scenario: K=2，N=10、M=12，decimals_col=8
- **WHEN** decimals_col=8、N_1=10、N_2=12
- **THEN** 数据起始列 = 10（J 列）；第 1 组占 J-S（10-19），T 列空，第 2 组占 U-AF（21-32）

#### Scenario: K=3，N_1=8、N_2=10、N_3=12，decimals_col=11
- **WHEN** decimals_col=11、N_1=8、N_2=10、N_3=12
- **THEN** 数据起始列 = 13（M 列）；第 1 组占 M-T（13-20），U 空，第 2 组占 V-AE（22-31），AF 空，第 3 组占 AG-AR（33-44）

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

### Requirement: CLI 参数与可复现性

脚本 SHALL 提供 CLI：`--input PATH`（**必填**，`required=True`，避免相对路径踩 CWD 与"默认文件"误改）、`--output PATH`（默认与 `--input` 相同，即就地修改）、`--sheet NAME`（默认第一个 sheet）、`--seed INT`（不传则随机选取并通过日志打印）、`--verbose` / `-v`（DEBUG 日志级别）。同一 `--seed` 在同一输入下 SHALL 产生完全一致的输出。

#### Scenario: 指定 seed 复现
- **WHEN** 用户两次执行脚本且都传入 `--seed 42`
- **THEN** 两次生成的所有原始数据完全一致

#### Scenario: 未传 seed
- **WHEN** 用户未传 `--seed`
- **THEN** 脚本 SHALL 取随机种子并以 INFO 级别打印实际使用的种子值，便于复现

### Requirement: 不破坏未涉及的单元格

脚本 SHALL 仅写入「数据起始列（decimals_col + 2）至最后一个统计列」范围内的单元格，以及第 1 行同范围的列号/样本量标注/统计列表头；SHALL NOT 修改 A 列至位数列、以及位数列+1（备注列）以外的任何单元格内容。

#### Scenario: 输入区列保持原样
- **WHEN** 脚本运行结束
- **THEN** A 列至位数列（含）每一行的值与运行前完全一致（以单元格 value 比较）

#### Scenario: 备注列保持原样
- **WHEN** 脚本运行结束
- **THEN** 位数列+1 的"备注列"内容（如有）与运行前完全一致

### Requirement: Levene 方差齐性检验

脚本 SHALL 对每个有效行用 `scipy.stats.levene(*groups, center='median')` 检验 K 组数据的方差齐性，把真实 p 值（保留 4 位小数）写入"Levene p"列。Levene p > 0.05 视作方差齐；Levene p ≤ 0.05 视作方差不齐。Levene p 列 SHALL 写在所有组 (μ, σ) 列之后。

#### Scenario: 方差齐
- **WHEN** Levene 检验返回 p > 0.05
- **THEN** Levene p 列写入该 p（4 位小数），且后续整体差异检验走 ANOVA 分支

#### Scenario: 方差不齐
- **WHEN** Levene 检验返回 p ≤ 0.05
- **THEN** Levene p 列写入该 p（4 位小数），且后续整体差异检验走 Kruskal-Wallis 分支

#### Scenario: K=2
- **WHEN** K=2
- **THEN** Levene 检验仍然执行（两组也可计算方差齐性），写入真实 p

### Requirement: Shapiro-Wilk 正态性检验

脚本 SHALL 对每组分别用 `scipy.stats.shapiro` 检验正态性，把所有组 SW p 值的最小值（保留 4 位小数）写入"Shapiro-Wilk min p"列。min p > 0.05 视作"全部组正态"；min p ≤ 0.05 视作"至少一组不正态"。该列 SHALL 紧跟 Levene p 列。

#### Scenario: 全部组正态
- **WHEN** 每组的 SW p 值都 > 0.05
- **THEN** 写入 min p（4 位小数），且若同时方差齐则两两对比走 Tukey HSD 分支

#### Scenario: 任一组不正态
- **WHEN** 某组的 SW p 值 ≤ 0.05
- **THEN** 写入 min p（4 位小数），两两对比 SHALL 走 Welch+Bonferroni 分支（即使方差齐）

#### Scenario: 某组样本量过小
- **WHEN** 某组 N < 3 导致 `scipy.stats.shapiro` 报错或返回 NaN
- **THEN** 该组 SW p 视为 0（不正态），脚本 SHALL `logger.warning` 该指标 + 组序号 + 实际 N

### Requirement: 整体差异 p 值（ANOVA / Kruskal-Wallis）

脚本 SHALL 计算 K 组之间的整体差异 p 值并写入"整体 p"列（紧跟 Shapiro-Wilk min p 列）。当 Levene p > 0.05（方差齐）时使用 `scipy.stats.f_oneway`（ANOVA）；当 Levene p ≤ 0.05（方差不齐）时使用 `scipy.stats.kruskal`（Kruskal-Wallis）。结果 SHALL 保留 4 位小数。

#### Scenario: 方差齐 → ANOVA
- **WHEN** Levene p > 0.05
- **THEN** "整体 p"列写入 `f_oneway(*groups).pvalue` 的 4 位小数

#### Scenario: 方差不齐 → Kruskal-Wallis
- **WHEN** Levene p ≤ 0.05
- **THEN** "整体 p"列写入 `kruskal(*groups).pvalue` 的 4 位小数

#### Scenario: ANOVA / KW 共用一列
- **WHEN** 不同行因 Levene 结果不同分别走 ANOVA 与 KW
- **THEN** 两类结果都写入同一列（"整体 p"列），不分列存放

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

