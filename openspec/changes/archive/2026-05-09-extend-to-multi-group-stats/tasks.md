## 1. 准备：依赖与样本数据

- [x] 1.1 确认 `requirements.txt` 中 `scipy>=1.11` 已含 `tukey_hsd`、`shapiro`、`f_oneway`、`kruskal`、`levene`、`ttest_ind`，本次**不新增**第三方依赖
- [x] 1.2 在 `tests/fixtures/` 新建（或现场构造）三个 sheet 样本：K=2（H 列位数）、K=3（K 列位数）、K=4（N 列位数），各含 2-3 行不同指标，用于读取/布局/E2E 测试
- [x] 1.3 备份当前 `20260509-随机数生成.xlsx` 为 `tests/fixtures/sample-k2.xlsx`，用作 E2E 黄金样本

## 2. 数据结构演进（重写读取层）

- [x] 2.1 在 `scripts/generate_excel_random_data.py` 把 `RowSpec` 改为 `RowSpec(row_index, metric, groups: tuple[GroupSpec, ...], decimals: int)`，移除 `group1` / `group2` 两字段；`GroupSpec` 不变
- [x] 2.2 写 `_find_decimals_col(ws) -> int`：扫描第 1 行表头，返回首个 value 含子串 `位数` 的列号；找不到抛 `LookupError`
- [x] 2.3 重写 `read_specs(path, sheet) -> list[RowSpec]`：先调 `_find_decimals_col` 算出 K=（decimals_col-2）/3，再循环每行按 col(2 + 3(i-1)) 起 3 列读 (N, μ, σ)，最后读 decimals；非法行 warn 跳过
- [x] 2.4 重写 `tests/test_read_specs.py`：覆盖 K=2 / K=3 / 无位数列报错 / 非法 decimals 跳过 / N=1 跳过 / SD=0 跳过

## 3. 生成层：行级 decimals

- [x] 3.1 修改 `generate_one_group(target_mean, target_sd, size, rng, decimals)`：把 `np.round(x, ROUND_DIGITS)` 改成 `np.round(x, decimals)`；删除模块级 `ROUND_DIGITS = 4`（仅保留 `STAT_DIGITS = 4` 给统计层用）
- [x] 3.2 修改 `generate_with_retry(metric, target_mean, target_sd, size, rng, decimals)`：透传 decimals；保留 ±10% 容差 + 5 次重试逻辑
- [x] 3.3 重写 `tests/test_generate.py`：覆盖 decimals=0（整数）、decimals=4（默认场景）、decimals=6；以及 ±10% 容差与重试上限

## 4. 列布局：动态多组

- [x] 4.1 把 `Layout` 重定义为 `Layout(decimals_col, group_cols: tuple[range, ...], stat_cols: StatColIndices)`，`StatColIndices` 是 dataclass（mean_sd_pairs、levene、shapiro_min、overall, pairwise_raw, pairwise_q）
- [x] 4.2 写 `compute_layout(decimals_col: int, group_sizes: tuple[int, ...]) -> Layout`：按"位数列+2 起 → 第 i 组 N_i 列 → 组间空 1 列 → 数据结束后空 1 列 → 2K 个 (μ, σ) → Levene → SW → Overall → C(K,2) raw → C(K,2) Q"计算所有列号
- [x] 4.3 重写 `tests/test_layout.py`：断言 K=2 / decimals_col=8 / N=10、M=12 时各列号；K=3 / decimals_col=11 / N_1=8、N_2=10、N_3=12 时各列号；K=4 时验证统计列总数 = 2K + 3 + 2·C(K,2)

## 5. 统计层：拆 4 个新函数

- [x] 5.1 写 `compute_levene(groups: list[ndarray], alpha: float = 0.05) -> tuple[float, bool]`：返回 `(levene_p, equal_var)`；用 `levene(*groups, center='median')`，NaN 视作 equal_var=False
- [x] 5.2 写 `compute_shapiro_min(groups: list[ndarray], alpha: float = 0.05) -> tuple[float, bool]`：每组分别 SW，组 N<3 时 p 视作 0；返回 `(min_p, all_normal)`
- [x] 5.3 写 `compute_overall(groups: list[ndarray], equal_var: bool) -> tuple[float, str]`：齐 → `f_oneway`、不齐 → `kruskal`；返回 `(p, label)`，label ∈ {"ANOVA", "KW"}
- [x] 5.4 写 `compute_pairwise(groups: list[ndarray], equal_var: bool, all_normal: bool) -> tuple[list[float], list[float | None], str]`：返回 (raw_ps, q_values, branch_label)；齐+全正态走 `tukey_hsd` 提取上三角对索引顺序的 p（Q 全 None）；否则每对 Welch t + Bonferroni
- [x] 5.5 删掉 v1 的 `compute_stats` 函数与对它的所有引用
- [x] 5.6 重写 `tests/test_stats.py`：每函数至少 2 case；针对 `compute_pairwise` 用人工构造的「齐+正态」与「不齐」两组样本，断言分支选择与 q 列恰当为 None / Bonferroni 校正值

## 6. 写入层：新统计列布局

- [x] 6.1 重写 `write_row(ws, row, generated_groups, layout, levene_p, sw_min_p, overall_p, raw_ps, q_values, branch_label)`：
  - 第 1 行写 K 组的样本量标注（统一 `<N_i>（G<i>）`）
  - 第 1 行补写统计列表头（仅当原表头为空）
  - 数据行写 K 组数据 + 2K 个 (μ, σ)（4 位小数）+ Levene、SW、Overall（4 位小数）+ raw、Q 列（Tukey 分支 Q 写空）
- [x] 6.2 重写 `_process_rows(ws, specs, rng)`：每行调度 generate→compute_levene→compute_shapiro_min→compute_overall→compute_pairwise→write_row

## 7. CLI 与日志

- [x] 7.1 `main()` 保持 `--input/--output/--sheet/--seed/--verbose` 五参数；启动日志增加"识别到位数列在第 X 列、K=Y"
- [x] 7.2 输入文件不存在时返回退出码 2；表头未声明位数列（`LookupError`）时记录 ERROR 并返回退出码 3

## 8. 端到端测试

- [x] 8.1 重写 `tests/test_cli.py`：
  - `test_cli_writes_expected_layout_k2`：用 `tests/fixtures/sample-k2.xlsx` + `--seed 42`，断言数据从 J 列起、A-H 列保持原样、I 列（备注）保持原样、统计列共 9 列且数值与 numpy/scipy 直接计算一致
  - `test_cli_writes_expected_layout_k3`：用 K=3 fixture，断言统计列 15 列、数据起始列 = M（13）、AG-AR 是第 3 组数据
  - `test_cli_branch_tukey_vs_welch`：构造一组「齐+正态」与一组「不齐」，断言 Q 列分别为 None 与 raw×C(K,2)
  - `test_cli_seed_reproducible`：保留 v1 的可复现性测试
  - `test_cli_requires_input_argument`：保留 v1 的 argparse 必填测试
- [x] 8.2 用 `--seed 42` 实际跑一次原 `20260509-随机数生成.xlsx`，肉眼复核生成结果（输出到 `/tmp/out-v2.xlsx`）

## 9. 文档与质量门

- [x] 9.1 重写 `scripts/README.md`：新列布局说明（A-位数列输入区 / 数据起始 = 位数列+2 / K 组动态 / 9 个或更多统计列）；删除 v1 的"H 备注"叙述；增加"K=2 / K=3 列布局图示"
- [x] 9.2 运行 `ruff format scripts tests && ruff check scripts tests`，全部通过
- [x] 9.3 运行 `pytest -v`，14+ 个 v1 测试全部重写为 v2 版本，全部 PASS
- [x] 9.4 `openspec validate extend-to-multi-group-stats --strict` 通过；`openspec status --change extend-to-multi-group-stats` 显示 4/4 done

## 10. 收尾

- [x] 10.1 把本 tasks.md 全部 `- [x]` 改成 `- [x]`
- [x] 10.2 `git status` / `git diff --stat` 确认改动只涉及：`scripts/` / `tests/` / `openspec/changes/extend-to-multi-group-stats/` / `openspec/specs/excel-random-data-generator/`（archive 阶段同步）
- [x] 10.3 提示用户运行 `/opsx:archive extend-to-multi-group-stats` 进入归档（合并 spec deltas 至 specs/，并同步 architecture.md）
