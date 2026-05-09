## 1. 项目搭建与依赖

- [x] 1.1 在仓库根创建（或更新）`requirements.txt`，加入 `openpyxl`、`numpy`、`scipy`，并在 venv 中安装；脚本入口选 `scripts/generate_excel_random_data.py`
- [x] 1.2 新建 `scripts/__init__.py`（空）与 `scripts/generate_excel_random_data.py` 的骨架（shebang、模块 docstring、`from __future__ import annotations`、`logging` 与 `argparse` 导入）
- [x] 1.3 在 `scripts/README.md` 中说明列布局约定（I 起 → 第一组 N 列 → 空 1 列 → 第二组 M 列 → 空 1 列 → 5 个统计列）、命令示例、就地修改的备份提示

## 2. 数据结构与读取层

- [x] 2.1 定义 `@dataclass(slots=True) GroupSpec`：`name: str`、`n: int`、`mean: float`、`sd: float`；以及 `RowSpec`：`row_index: int`、`metric: str`、`group1: GroupSpec`、`group2: GroupSpec`
- [x] 2.2 写 `read_specs(path: Path, sheet: str | None) -> list[RowSpec]`：用 `openpyxl.load_workbook` 打开，从第 2 行起按 A-G 列解析；对 N≤1、M≤1、SD≤0、缺失或非数值的行 `logger.warning` 并跳过
- [x] 2.3 单元测试 `tests/test_read_specs.py`：用 `openpyxl` 现场构造一个临时 xlsx，覆盖正常行、缺失行、N=1 行三种情况

## 3. 数据生成层

- [x] 3.1 写 `generate_one_group(target_mean: float, target_sd: float, size: int, rng: np.random.Generator) -> np.ndarray`：先生成标准正态 → 减均值除标准差（ddof=1）→ 线性变换到目标 → `np.round(x, 4)`
- [x] 3.2 包装 `generate_with_retry(...)`：检查 round 后实际 SD，若与目标相对误差 > 10% 则换子种子重试，最多 5 次；都失败抛 `RuntimeError(f"无法在容差内还原 SD: metric={...}")`
- [x] 3.3 单元测试 `tests/test_generate.py`：覆盖（a）样本量较大时实际 SD 严格匹配目标；（b）N=2 时仍能在容差内（多次随机 seed）；（c）目标 SD 为 0 时直接抛错（按约定不接受 SD≤0）

## 4. 统计回算层

- [x] 4.1 写 `compute_stats(g1: np.ndarray, g2: np.ndarray, alpha: float = 0.05) -> tuple[float, float, float, float, float, bool]`，返回 `(mean1, sd1, mean2, sd2, p_value, equal_var)`；内部用 `scipy.stats.levene(..., center='median')` 与 `scipy.stats.ttest_ind(..., equal_var=...)`
- [x] 4.2 单元测试 `tests/test_stats.py`：构造一组明显方差齐与一组明显方差不齐的数据，分别断言 `equal_var` 分支与 p 值与 `scipy` 直接调用的结果一致（避免硬编码具体 p 值）

## 5. 列布局与写入层

- [x] 5.1 写 `compute_layout(n: int, m: int) -> Layout`，返回 `group1_cols: range`、`group2_cols: range`、`stat_cols: range`，全部基于 `9` 起算的偏移；同时在文档中标明 N=10、M=12 时与样本表 I/R/T/AE/AG/AK 完全一致（断言式校验）
- [x] 5.2 写 `write_row(ws, row: RowSpec, g1: np.ndarray, g2: np.ndarray, stats: tuple, layout: Layout)`：
  - 写第 1 行序号（仅当该格当前为空或与新值不同时写，避免无谓改动）
  - 把第一组最后一格表头改为 `"<n>（N）"`、第二组最后一格改为 `"<m>（M）"`
  - 写第 `row.row_index` 行的 N 个 / M 个数据
  - 写 5 个统计列（mean1、sd1、mean2、sd2、p）
  - 若统计列第 1 行表头为空，则补写 `第一组均值 / 第一组SD值 / 第二组均值 / 第二组SD值 / 两组的p值`
- [x] 5.3 单元测试 `tests/test_layout.py`：断言 `compute_layout(10, 12)` 返回的列号集对应 I-R / T-AE / AG-AK；并覆盖 N=5、M=7 的非默认场景

## 6. CLI 入口与日志

- [x] 6.1 用 `argparse` 实现 `--input / --output / --sheet / --seed / --verbose` 五个参数；`--input` **必填**（`required=True`），output 默认 = input
- [x] 6.2 实现 `main()`：配置 `logging`（`-v` → DEBUG，否则 INFO；格式化为 `%(asctime)s %(levelname)s %(message)s`）；处理种子（未传时 `secrets.randbits(32)`，记录 INFO）；按行循环 generate→compute→write；最后 `wb.save(output)`
- [x] 6.3 集成测试 `tests/test_cli.py`：拷贝项目自带的 `20260509-随机数生成.xlsx` 到 tmp_path，固定 `--seed 42` 运行脚本入口（直接 import main 后调用，不走 subprocess），随后用 openpyxl 打开输出 xlsx 校验：
  - A-H 列保持不变（与原文件逐格相等）
  - I-R 列每行恰好 10 个数；T-AE 列每行恰好 12 个数；S 与 AF 为空
  - 第 1 行 R1 == `'10（N）'`、AE1 == `'12（M）'`
  - AG-AK 与 numpy/scipy 直接计算的 mean/SD/p 一致（在浮点容差内）

## 7. 端到端验证与归档

- [x] 7.1 在仓库 venv 中实际运行 `python scripts/generate_excel_random_data.py --seed 42 --output /tmp/out.xlsx`，肉眼复核首行数据是否符合"四位小数 + 两组中间空一列 + 表头标注 + 统计列"的预期
- [x] 7.2 运行 `ruff format . && ruff check .` 与 `pytest`（若环境没有 ruff/pytest，用 `pip install` 临时装）；全部通过
- [x] 7.3 提交前以 `openspec status --change add-excel-random-data-generator` 复核所有 artifact 状态为 done，再用 `git status` / `git diff` 复核改动只涉及 `scripts/` / `requirements.txt` / `tests/` / `openspec/changes/...` 范围
