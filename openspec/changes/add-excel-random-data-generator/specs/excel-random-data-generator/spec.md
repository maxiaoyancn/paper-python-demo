## ADDED Requirements

### Requirement: 读取目标统计量

脚本 SHALL 通过**必填**的 `--input` 参数接收 Excel 文件路径，并接受可选 `--sheet` 参数，从指定 sheet 的第 2 行起逐行读取每个指标的两组目标统计量：A 列（指标名）、B 列（第一组样本量 N）、C 列（第一组目标均值）、D 列（第一组目标 SD）、E 列（第二组样本量 M）、F 列（第二组目标均值）、G 列（第二组目标 SD）。

#### Scenario: 缺少 --input
- **WHEN** 用户执行 `python scripts/generate_excel_random_data.py` 不带 `--input`
- **THEN** argparse SHALL 立即报错并以退出码 2 终止，错误消息中包含 `--input`；脚本不会处理任何文件

#### Scenario: 指定输入与 sheet
- **WHEN** 用户传入 `--input some.xlsx --sheet Data`
- **THEN** 脚本读取 `some.xlsx` 中名为 `Data` 的 sheet 并按上述列约定解析每个数据行的目标统计量

#### Scenario: 行内目标统计量缺失或非法
- **WHEN** 某一行的 N、M、目标均值、目标 SD 任一为空、非数值、N≤1、M≤1 或 SD≤0
- **THEN** 脚本 SHALL 通过 `logging.warning` 记录该行被跳过的原因，并继续处理后续行；不阻断整体执行

### Requirement: 生成两组正态分布原始数据

脚本 SHALL 为每个有效行分别生成两组正态分布的原始数据：第一组 N 个、第二组 M 个，每个数值 SHALL 保留 4 位小数。生成方式 SHALL 保证：未四舍五入前的样本均值严格等于目标均值、样本标准差（ddof=1）严格等于目标 SD；四舍五入后实际 SD 与目标 SD 的相对误差 SHALL 在 ±10% 之内，否则脚本 SHALL 用不同子种子重试，最多 5 次。重试仍失败时 SHALL 抛 `RuntimeError` 而非静默写入。

#### Scenario: SD 在容差内
- **WHEN** 生成并取 4 位小数后实际 SD 与目标 SD 相对误差 ≤ 10%
- **THEN** 该次生成结果直接被采纳

#### Scenario: SD 超出容差但重试成功
- **WHEN** 首次生成实际 SD 与目标 SD 相对误差 > 10%
- **THEN** 脚本 SHALL 换子种子重试，最多 5 次；任何一次落入容差即采纳

#### Scenario: 重试 5 次仍失败
- **WHEN** 5 次重试后实际 SD 仍与目标 SD 相对误差 > 10%
- **THEN** 脚本 SHALL 抛 `RuntimeError`，错误消息中包含指标名、目标 SD、最近一次实际 SD

### Requirement: 按动态列布局写入原始数据

脚本 SHALL 将第一组原始数据写入从第 9 列（即 I 列）起的连续 N 列；空 1 列后将第二组原始数据写入接下来的 M 列。脚本 SHALL NOT 把列号写死——必须按 N、M 动态计算偏移。

#### Scenario: N=10, M=12
- **WHEN** 某行 N=10、M=12
- **THEN** 第一组数据落在第 9-18 列（I-R），第 19 列（S）为空，第二组数据落在第 20-31 列（T-AE）

#### Scenario: 不同 N、M 组合
- **WHEN** 某行 N=5、M=7
- **THEN** 第一组数据落在第 9-13 列，第 14 列为空，第二组数据落在第 15-21 列

### Requirement: 在表头标注样本量

脚本 SHALL 在第 1 行同时写入两组数据的列序号（从 1 起），且 SHALL 把每组**最后一个数据列**的表头覆写为带样本量标注的形式：第一组写 `<count>（N）`、第二组写 `<count>（M）`，使用全角中文括号以与既有表格风格一致。

#### Scenario: N=10
- **WHEN** 第一组样本量 N=10
- **THEN** 第 1 行第一组所占的 10 列依次写入 `1, 2, 3, 4, 5, 6, 7, 8, 9, "10（N）"`

#### Scenario: M=12
- **WHEN** 第二组样本量 M=12
- **THEN** 第 1 行第二组所占的 12 列依次写入 `1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, "12（M）"`

### Requirement: 用生成数据回算统计量并计算组间 p 值

脚本 SHALL 使用四舍五入后的最终生成数据回算两组的样本均值与样本标准差（ddof=1），并计算组间双尾 p 值：先用 Levene 检验（`center='median'`）判定方差齐性，显著性阈值 α=0.05。若 Levene p ≥ 0.05 视作方差齐，使用 `scipy.stats.ttest_ind(equal_var=True)`；否则使用 `equal_var=False`（Welch's t-test）。

#### Scenario: 方差齐
- **WHEN** Levene 检验 p ≥ 0.05
- **THEN** 组间 p 值采用标准两样本 t 检验（双尾）的结果

#### Scenario: 方差不齐
- **WHEN** Levene 检验 p < 0.05
- **THEN** 组间 p 值采用 Welch 校正 t 检验（双尾）的结果

### Requirement: 写入回算统计量到 AG-AK 风格列

脚本 SHALL 把回算的第一组均值、第一组 SD、第二组均值、第二组 SD、组间 p 值依次写入第二组数据之后空 1 列起的连续 5 列。列号 SHALL 按 N、M 动态计算（第二组结束列 = 9 + N + M，第 5 个统计列 = 9 + N + M + 6）。当 N=10、M=12 时这 5 列恰为 AG-AK。

#### Scenario: 列号匹配既有表头
- **WHEN** N=10、M=12，且原表 AG1-AK1 已存在「第一组均值 / 第一组SD值 / 第二组均值 / 第二组SD值 / 两组的p值」表头
- **THEN** 脚本 SHALL 把回算结果写入第 2 行起的 AG-AK 单元格，且不覆盖第 1 行表头

#### Scenario: 表头缺失时补写
- **WHEN** 第 1 行对应统计列的表头为空
- **THEN** 脚本 SHALL 自动写入对应的中文表头

### Requirement: CLI 参数与可复现性

脚本 SHALL 提供 CLI：`--input PATH`（**必填**，`required=True`，避免相对路径踩 CWD 与"默认文件"误改）、`--output PATH`（默认与 `--input` 相同，即就地修改）、`--sheet NAME`（默认第一个 sheet）、`--seed INT`（不传则随机选取并通过日志打印）、`--verbose` / `-v`（DEBUG 日志级别）。同一 `--seed` 在同一输入下 SHALL 产生完全一致的输出。

#### Scenario: 指定 seed 复现
- **WHEN** 用户两次执行脚本且都传入 `--seed 42`
- **THEN** 两次生成的所有原始数据完全一致

#### Scenario: 未传 seed
- **WHEN** 用户未传 `--seed`
- **THEN** 脚本 SHALL 取随机种子并以 INFO 级别打印实际使用的种子值，便于复现

### Requirement: 不破坏未涉及的单元格

脚本 SHALL 仅写入「I 列至最后一个统计列（9 + N + M + 6）」范围内的单元格，以及第 1 行同范围的列号/样本量标注；SHALL NOT 修改 A-H 列以及范围以外的任何单元格内容。

#### Scenario: A-H 列保持原样
- **WHEN** 脚本运行结束
- **THEN** A-H 列每一行的值与运行前完全一致（以单元格 value 比较）
