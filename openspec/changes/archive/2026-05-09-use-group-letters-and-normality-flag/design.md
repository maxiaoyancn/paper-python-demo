## Context

- v2 capability spec 在 [openspec/specs/excel-random-data-generator/spec.md](../../specs/excel-random-data-generator/spec.md)：定义了 K 组（K≥2）流水线、`Gi` 组标记、`第i组均值/SD值` 统计表头。
- 用户反馈 v2 表头不直观——原始数据列中间是裸数字（看不出哪一组），统计列两套术语（"第 1 组均值"和"G1-G2 raw p"）混用。本次微调统一为字母。
- 同时希望在已有的 Levene p / SW min p 基础上各加一列直观的 Y/N（"是否方差齐" 与 "是否正态"），让结论一目了然——避免每次回头看 0.05 阈值。

## Goals / Non-Goals

**Goals:**
- 第 i 组在所有表头中用第 i 个英文大写字母（A=1, B=2, …, Z=26）。
- 原始数据每列表头 `<letter><idx>`，idx 从 1 起；最后一格不再用 `<N>（letter）` 混合写法（样本量从最大 idx 看出）。
- 统计列表头：均值/SD/对比都用字母；保留 `Levene p`、`Shapiro-Wilk min p`、`整体 p（ANOVA/KW）` 不变。
- "是否方差齐"列写 `Y` / `N`（半角字母，1 字符），紧跟 Levene p 后。
- "是否正态"列写 `Y` / `N`，紧跟 SW min p 后。
- 写入范围、数据起始列、生成与统计逻辑、CLI、可复现性均不变。

**Non-Goals:**
- K > 26 不支持（不去做 AA/AB/AC 这种 spreadsheet 列号风格的扩展）。读取时若推导出 K>26，`logger.error` 跳过整 sheet。
- 不引入 `--enforce-normality` 这种 CLI 参数（已在用户选项中剔除）；如未来需要再单独 propose。
- 不动 v1/v2 已 archive 的 spec 历史。

## Decisions

### 1. 字母映射
- `_letter(i: int) -> str`：i=1..26 → 'A'..'Z'；i>26 抛 `ValueError("K 不支持超过 26")`，由 `read_specs` 上层捕获后转换为 `LookupError("K=... 超过 26 上限")`。
- 选择大写字母而非小写：研究表格惯用大写代号（A/B 组），更显眼。

### 2. 数据列表头的纯字母+索引格式
- 第 i 组共 N_i 列，表头依次 `<letter><1>` `<letter><2>` ... `<letter><N_i>`。
- 例：第 1 组 N=10 → A1, A2, ..., A10。
- 与 v2 比：去掉了"最后一格 <N>（letter）"混写。样本量从"最后一个表头的数字部分"读出，仍然容易识别。
- 拒选方案：保留"<N>（letter）"风格；理由是它打破列宽的均匀视觉，且需要两套规则（"中间列 vs 最后一列"），不如统一。

### 3. 统计列表头改名规则
- 当前 v2：`第 i 组均值` / `第 i 组SD值` / `G<i>-G<j> raw p` / `G<i>-G<j> Q-value`
- v3：`<letter_i> 组均值` / `<letter_i> 组SD值` / `<letter_i>-<letter_j> raw p` / `<letter_i>-<letter_j> Q-value`
- `Levene p`、`Shapiro-Wilk min p`、`整体 p（ANOVA/KW）` 三个表头保持原样（与组无关）。
- 两个新结论表头：`是否方差齐`（紧跟 Levene p 后）与 `是否正态`（紧跟 SW min p 后）。

### 4. 列布局变化（关键）
- v2 的统计列顺序：`mean_sd × K, levene, shapiro_min, overall, pairwise_raw × C(K,2), pairwise_q × C(K,2)`
- v3 改为：`mean_sd × K, levene, **levene_flag**, shapiro_min, **normality_flag**, overall, pairwise_raw × C(K,2), pairwise_q × C(K,2)`
- 即分别在 `levene` 与 `shapiro_min` 之间、`shapiro_min` 与 `overall` 之间**各插入 1 列**，所有 SW 之后的列号相对 v2 都向右挪 **+2**（pairwise_raw 与 pairwise_q 起始列也跟着 +2）。
- `StatColIndices` 新增两个字段：`levene_flag: int` 与 `normality_flag: int`。
- `compute_layout` 实现里在 levene 后 +1（levene_flag），在 shapiro 后 +1（normality_flag），其余累加逻辑不变。

### 5. 两个结论列的判定逻辑
- "是否方差齐" 列沿用 `compute_levene` 返回的 `equal_var: bool`：`Y` 当 equal_var=True、`N` 当 False。
- "是否正态" 列沿用 `compute_shapiro_min` 返回的 `all_normal: bool`：`Y` 当 all_normal=True、`N` 当 False。
- 两个结论都直接从已计算的中间结果取值，不重做检验——保持单一数据源。
- `write_row` 函数签名相应增加 `equal_var: bool` 与 `all_normal: bool` 两个参数；`_process_rows` 把这两个值连同对应 p 值一起传入。

### 6. 错误处理
- `_letter` 抛 `ValueError` 在内部边界，统一向外抛 `LookupError`（与"找不到位数列"的处理一致），main 捕获后退出码 3。
- 添加测试覆盖 K>26 的退出码。

### 7. 兼容性策略
- v2 → v3 是表头层与列偏移层的小幅破坏性变更；spec deltas 用 MODIFIED + ADDED 描述。
- 用户已反复确认 v2 输出文件不会作为输入回灌（输出区每次脚本写都是覆盖），所以不需要兼容旧表头解析。
- archive v3 后，spec.md 会再次被 +/~/- 应用（与 v2 的合并机制一致）。

## Risks / Trade-offs

- [字母 K 上限 26] → 论文场景几乎不会超 26 组（罕见 >5 组）；硬上限简单且明确。如真碰到 K=27+，未来可扩 AA/AB（双字母），代价是表头宽度。
- [列偏移变化导致使用者旧脚本拿坐标计算] → 用户没有外部脚本依赖输出列号；纯手工查看。
- [N=1..9 的纯字母表头与 N=10..99 的字母+两位数视觉宽度不一致] → Excel 列宽自动处理；可忽略。
- [`是否方差齐` / `是否正态` 列与 Levene/SW min p 列冗余] → 这是用户明确要求的"一个数一个结论"双轨呈现，不是冗余 bug——研究人员经常希望一眼看出"齐/不齐 / 正态/不正态"，无需在脑子里对照 0.05 阈值。
- [v2 测试需要重写 cell value 断言] → 测试已经覆盖 K=2/K=3/K=4 端到端，调整字符串断言即可，不影响测试结构。
