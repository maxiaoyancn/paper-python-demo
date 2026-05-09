## 1. 字母映射与边界

- [x] 1.1 在 `scripts/generate_excel_random_data.py` 加 `_letter(i: int) -> str`：i=1..26 → 'A'..'Z'；i>26 抛 `ValueError`。直接用 `chr(ord("A") + i - 1)` 实现，单行
- [x] 1.2 在 `read_specs` 推导出 K 后立即检查 `K > 26`，若超限抛 `LookupError(f"K={k} 超过 26 上限")`；让 `main()` 现有的 `LookupError` 兜底自动转为退出码 3
- [x] 1.3 单元测试 `tests/test_read_specs.py` 加 `test_read_specs_k_gt_26_raises`：现场构造 27 组列布局的 xlsx（位数列在 col 83），断言 `read_specs` 抛 `LookupError` 含"超过 26"

## 2. 数据列表头：从 `<count>（Gi）` → `<letter><idx>`

- [x] 2.1 重写 `_write_group_data_headers(ws, group_cols, group_sizes)`：循环每组 i（1-based），letter=`_letter(i)`；每列 col 写入 `f"{letter}{idx}"`（idx 1-based 对应组内索引）
- [x] 2.2 调整 `tests/test_cli.py` 的 K=2 端到端断言：
  - `ws_out.cell(row=1, column=col_S).value == "10（G1）"` → `== "A10"`
  - `ws_out.cell(row=1, column=col_AF).value == "12（G2）"` → `== "B12"`
  - 同时新增对中间列断言：`ws_out.cell(row=1, column=col_J).value == "A1"`、`col_J + 4 → "A5"` 等
- [x] 2.3 调整 K=3 端到端断言：`ws.cell(row=1, column=col_T).value` 从 `"8（G1）"` → `"A8"`；`col_AE` 从 `"10（G2）"` → `"B10"`；`col_AR` 从 `"12（G3）"` → `"C12"`

## 3. 统计列：扩 StatColIndices + 列布局 +2

- [x] 3.1 修改 `scripts/generate_excel_random_data.py` 中 `StatColIndices` dataclass：在 `levene` 后增加 `levene_flag: int`，在 `shapiro_min` 后增加 `normality_flag: int`；同时调整 `compute_layout` 在 levene 后 +1 给 levene_flag、SW 后 +1 给 normality_flag，再继续 overall / pairwise（即 pairwise 列号相对 v2 向右挪 +2）
- [x] 3.2 修改 `tests/test_layout.py` 三个 case：
  - K=2：v2 的 `levene == AL=38`；v3 在 38 后插入 `levene_flag=AM=39`、`shapiro_min=AN=40`、`normality_flag=AO=41`、`overall=AP=42`、`pairwise_raw=(AQ=43,)`、`pairwise_q=(AR=44,)`
  - K=3：v2 的 `levene=52`；v3 = `levene=52, levene_flag=53, shapiro_min=54, normality_flag=55, overall=56, pairwise_raw=(57,58,59), pairwise_q=(60,61,62)`
  - K=4：n_stat 计算改为 `2*4 + 5 + 2*6 = 25`（含 levene_flag 与 normality_flag 两列）

## 4. 统计列表头：换字母 + 加"是否方差齐"/"是否正态"两列

- [x] 4.1 在 `scripts/generate_excel_random_data.py` 顶部把 `LEVENE_HEADER` / `SHAPIRO_HEADER` / `OVERALL_HEADER` 保留不变；新增 `LEVENE_FLAG_HEADER = "是否方差齐"`、`NORMALITY_FLAG_HEADER = "是否正态"`
- [x] 4.2 重写 `_write_stat_headers(ws, stat_cols, k)`：
  - mean/SD 表头从 `f"第{i}组均值"` → `f"{_letter(i)} 组均值"`、`f"第{i}组SD值"` → `f"{_letter(i)} 组SD值"`
  - 在 `LEVENE_HEADER` 后插入 `(stat_cols.levene_flag, LEVENE_FLAG_HEADER)`
  - 在 `SHAPIRO_HEADER` 后插入 `(stat_cols.normality_flag, NORMALITY_FLAG_HEADER)`
  - pair_labels 从 `f"G{i+1}-G{j+1}"` → `f"{_letter(i+1)}-{_letter(j+1)}"`
- [x] 4.3 修改 `write_row` 在 Levene p 后写 levene_flag、SW min p 后写 normality_flag：
  - `ws.cell(row=row.row_index, column=stat_cols.levene_flag).value = "Y" if equal_var else "N"`
  - `ws.cell(row=row.row_index, column=stat_cols.normality_flag).value = "Y" if all_normal else "N"`
- [x] 4.4 修改 `write_row` 函数签名：`write_row(ws, row, generated, layout, levene_p, equal_var, sw_min_p, all_normal, overall_p, raw_ps, q_values)`（分别在 levene_p 与 sw_min_p 后插入对应布尔参数）；`_process_rows` 把 `compute_levene` 的 `equal_var` 与 `compute_shapiro_min` 的 `all_normal` 一并传入

## 5. 端到端测试

- [x] 5.1 调整 `tests/test_cli.py::test_cli_writes_expected_layout_k2`：
  - col_AL（Levene）=38、**col_AM（是否方差齐）=39**、col_AN（SW）=40、**col_AO（是否正态）=41**、col_AP（整体 p）=42、col_AQ（A-B raw）=43、col_AR（A-B Q）=44
  - 断言 `ws_out.cell(row=2, column=col_AL + 1).value in ("Y", "N")`（levene_flag）
  - 断言 `ws_out.cell(row=2, column=col_AL + 3).value in ("Y", "N")`（normality_flag）
  - overall p 断言移到 `col_AL + 4`
- [x] 5.2 调整 `tests/test_cli.py::test_cli_writes_expected_layout_k3`：
  - stat_start = 46 不变；mean/SD 列共 6（46-51）；Levene 52、**levene_flag 53**、SW 54、**normality_flag 55**、Overall 56、raw 起 57、Q 起 60
  - 断言"是否方差齐"与"是否正态"列内容都是 Y 或 N
- [x] 5.3 新增 `tests/test_cli.py::test_cli_flags_consistent_with_p_values`：用 `sample-k2.xlsx` + `--seed 42` 跑后，对每个数据行：
  - 断言 `levene_flag == "Y"` 当且仅当 `Levene p` cell > 0.05
  - 断言 `normality_flag == "Y"` 当且仅当 `Shapiro-Wilk min p` cell > 0.05

- [x] 5.4 跑 `.venv/bin/pytest -v`，全部 PASS

## 6. 文档与冒烟

- [x] 6.1 重写 `scripts/README.md` 列布局段：
  - 数据列表头格式从 `<count>（Gi）` 改为 `<letter><idx>`
  - 统计列表格里每组从 `第i组均值` → `<letter> 组均值`
  - 新增"是否方差齐"行（紧跟 Levene p 后）与"是否正态"行（紧跟 SW min p 后）
  - 更新 K=2/3/4 列数表格：9→**11**、15→**17**、23→**25**
- [x] 6.2 实际跑 K=2 / K=3 sample 验证：`python scripts/generate_excel_random_data.py --input <fixture> --output /tmp/out-v3.xlsx --seed 42 -v`，肉眼复核字母表头、是否方差齐 Y/N、是否正态 Y/N 三处
- [x] 6.3 `.venv/bin/ruff format scripts tests && .venv/bin/ruff check scripts tests` 通过

## 7. 收尾

- [x] 7.1 把本 tasks.md 全部 `- [x]` 改成 `- [x]`
- [x] 7.2 `git status` / `git diff --stat` 确认改动只涉及：`scripts/` / `tests/` / `openspec/changes/use-group-letters-and-normality-flag/`
- [x] 7.3 提示用户运行 `/opsx:archive use-group-letters-and-normality-flag` 完成归档（合并 spec deltas 至 specs/，同步 architecture.md）
