## Why

当前 v2 的表格里，原始数据列只有最后一格表头才标注组身份（`<N>（G1）`），中间 N-1 列是裸数字 `1..N-1`，看不出属于哪组。统计列的表头也用 `第 i 组` / `Gi-Gj` 这种用户不直觉的写法。研究人员希望：① 用 **A/B/C/...** 字母代表组，所有列表头都能直接看出列属于哪组；② 在 Shapiro-Wilk min p 数值之外**额外加一列 Y/N** 直接给出"所有组是否符合正态分布"的结论，免去人工对照 0.05 阈值。

## What Changes

- **BREAKING（表头）：原始数据列表头改为 `<字母><idx>` 格式**
  - 第 i 组用第 i 个英文大写字母（i=1..K → A..Z）
  - 第 i 组占 N_i 列，表头依次为 `<letter>1`、`<letter>2`、…、`<letter><N_i>`（不再用裸数字 + 末位 `<N>（Gi）` 的混合方式）
  - K 上限 26（K>26 时读取阶段报错跳过该 sheet；目前论文场景远小于 26）

- **BREAKING（统计列表头）：跟着字母统一**
  - 每组均值/SD：`第 i 组均值` / `第 i 组SD值` → `<letter> 组均值` / `<letter> 组SD值`
  - 两两对比表头：`G<i>-G<j> raw p` / `G<i>-G<j> Q-value` → `<letter_i>-<letter_j> raw p` / `<letter_i>-<letter_j> Q-value`
  - 其他三个统计表头（Levene p / Shapiro-Wilk min p / 整体 p）保持不变

- **新增两列结论列：每个数值列后面紧跟它的"结论"列（Y/N），让读者一眼可读**
  - **"是否方差齐" 列**：紧跟 `Levene p`。Y = 方差齐（Levene p > 0.05）；N = 方差不齐（≤ 0.05）
  - **"是否正态" 列**：紧跟 `Shapiro-Wilk min p`。Y = 全部组都正态（min p > 0.05）；N = 至少一组不正态（≤ 0.05）
  - 两列分别与对应数值列**并存**（一个看具体 p、一个看结论）
  - 这两列共把所有后续统计列偏移 **+2**（整体 p 与 pairwise 列号相对 v2 都向右挪 2）

- **保留**：列布局其余规则（位数列动态识别、数据起始 = 位数列+2、组间空 1 列、统计区前空 1 列）；`±10%` SD 容差 + 5 次重试；CLI 接口；可复现性。

## Capabilities

### New Capabilities
<!-- 无新 capability。 -->

### Modified Capabilities
- `excel-random-data-generator`: 调整三类需求——原始数据表头格式（v2 的"在表头标注样本量"扩展为"按字母 + 索引标注每列"）、统计列表头命名（每组与两两对比都改字母）、统计列布局（新增"是否正态"列在 Shapiro-Wilk min p 后）。

## Impact

- 代码：
  - [scripts/generate_excel_random_data.py](../../../scripts/generate_excel_random_data.py)：调整 `_write_group_data_headers` 与 `_write_stat_headers`；扩展 `StatColIndices` 增加 **`levene_flag`** 与 **`normality_flag`** 两个字段；调整 `compute_layout` 的列号计算（pairwise 起始相对 v2 向右挪 2）；调整 `write_row` 写入两个新列；新增 `_letter(i)` 辅助函数（K>26 时抛错）
- 测试：
  - [tests/test_layout.py](../../../tests/test_layout.py)：列号断言全部 +2（`overall`、pairwise）；新增 `levene_flag` 与 `normality_flag` 列号断言
  - [tests/test_cli.py](../../../tests/test_cli.py)：K=2/K=3 端到端断言新表头字符串（`S1 == "A10"`、`AT1 == "A 组均值"` 等）；新增 Levene/SW 两个分支下的 Y/N 列断言
  - [tests/test_read_specs.py](../../../tests/test_read_specs.py) / [tests/test_generate.py](../../../tests/test_generate.py) / [tests/test_stats.py](../../../tests/test_stats.py)：基本不动（不依赖表头字符串）
- 文档：[scripts/README.md](../../../scripts/README.md) 列布局段需更新示例与统计列数（K=2 从 9 列变 **11**、K=3 从 15 变 **17**、K=4 从 23 变 **25**）
- 兼容性：与 v2 的输出列布局**列偏移不兼容**（pairwise 列向右挪 1 列），且数据表头与统计列表头字符串完全不同；用户的旧 v2 输出文件无法被本版本"再处理"得到一致结果——但旧输出本就不是脚本输入，不影响实际工作流。
- 依赖：无新增。
