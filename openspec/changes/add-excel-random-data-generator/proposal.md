## Why

研究人员（用户）手中有一份 Excel 表格 `20260509-随机数生成.xlsx`：A-G 列描述每行指标的两组目标统计量（样本量 N/M、均值、SD），但缺少符合该统计量的"原始数据"。手工凑数效率低、可重复性差。需要一个 Python 脚本，按目标统计量自动生成符合正态分布、保留 4 位小数的两组原始数据，并基于生成数据回算出真实的均值、SD 和组间 p 值（Levene 方差齐性 → t 检验 / Welch 校正），把所有结果按既定列布局写回同一份 Excel。

## What Changes

- 新增 Python 脚本 `scripts/generate_excel_random_data.py`：
  - 读取目标 Excel（路径**必须**通过 `--input` 显式传入，避免相对路径踩 CWD 的坑、避免误改某个"默认文件"）。
  - 对每个数据行（默认从第 2 行起），按 B/C/D 与 E/F/G 两组目标统计量生成正态分布原始数据。
  - 第一组从 I 列起占 N 列、空一列后第二组从 T 列起占 M 列；每组**最后一个数据单元格**的列标题（第 1 行）会被覆写为 `<count>（N）` / `<count>（M）` 形式以标注样本量。
  - 当目标 SD 不可严格还原时，允许在目标 SD ±10% 范围内取实际 SD 重生成，直至数值稳定。
  - 用生成的数据回算两组均值、SD，并计算组间 p 值：先用 Levene 检验方差齐性，齐用标准两样本 t 检验（双尾），不齐用 Welch's t-test。
  - 把回算结果写入第二组之后空一列起的 AG-AK 列（第一组均值 / 第一组 SD / 第二组均值 / 第二组 SD / p 值）。
- 新增依赖：`openpyxl`、`numpy`、`scipy`（写入 `requirements.txt`，沿用 pip + venv 方式）。
- 新增 `scripts/README.md`，说明运行方式与列布局约定（脚本目录下，避免污染根 README）。

## Capabilities

### New Capabilities
- `excel-random-data-generator`: 从 Excel 读取每行目标统计量、按正态分布生成两组原始数据并回算统计量与 p 值，将结果按规定列布局写回同一份 Excel。

### Modified Capabilities
<!-- 暂无。仓库内 openspec/specs/ 目前为空，无需修改既有 capability。 -->

## Impact

- 代码：新增 `scripts/generate_excel_random_data.py`、`scripts/README.md`。
- 依赖：新增 `openpyxl` / `numpy` / `scipy`，需写入 `requirements.txt`；建议在虚拟环境内安装。
- 数据：当 `--output` 不传或与 `--input` 相同的，脚本会**就地修改输入文件**（覆盖 I 列至 AK 列范围内的数据与首行样本量标注）；运行前建议另存备份，或始终用 `--output` 指向旁路文件。
- 文档：`scripts/README.md` 描述列布局与运行命令。
- 不影响现有代码（仓库目前没有其它 Python 模块）。
