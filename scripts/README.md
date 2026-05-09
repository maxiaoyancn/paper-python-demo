# Excel 随机数据生成脚本

## 用途

读取一份 Excel 表（第 1 行为表头，第 2 行起每行一个指标），按表头描述的 K 组（K≥2）目标统计量（每组的样本量 N、均值、SD）生成符合正态分布的原始数据，并回算实际 mean/SD、Levene 方差齐性 p、Shapiro-Wilk 正态性 p（每组取最小）、整体差异 p（ANOVA / Kruskal-Wallis）、两两对比 raw p 与多重比较校正 Q-value（Tukey HSD / Welch+Bonferroni）。

## 列布局（动态，按 K 与每组 N 计算，不写死）

**输入区**（脚本不修改）：

- A 列：指标名
- 紧跟 K 个 (N、μ、σ) 三元组：第 1 组在 B-D、第 2 组在 E-G、第 3 组在 H-J、…
- 第 K 组之后**紧跟一列"原始数据小数点后位数"**（脚本通过扫描第 1 行表头中含"位数"的列定位它，K 由位置反推）
- 位数列后**1 列空白**（可选填备注，脚本不读不写）

**数据区**（脚本写入，覆盖原内容）：

- 数据起始列 = 位数列 + 2
- 第 i 组占 N_i 列，组间空 1 列
- 第 i 组最后一格的第 1 行表头会被覆写为 `<N_i>（G<i>）`

**统计区**（脚本写入，紧跟数据区 + 1 空列）：

- 第 i 组均值、第 i 组 SD（共 2K 列，4 位小数）
- Levene p（中心 median；> 0.05 视为方差齐）
- Shapiro-Wilk min p（每组分别 SW 取最小；> 0.05 视为全部正态）
- 整体 p（齐 → ANOVA；不齐 → Kruskal-Wallis）
- 两两 raw p × C(K,2) 列（按 (1,2)(1,3)(2,3)... 顺序）
- 两两 Q-value × C(K,2) 列（齐+全正态走 Tukey 时为空；否则为 Welch+Bonferroni 校正值）

**示例**：

| K | N 列表 | 位数列 | 数据区 | 统计区起 | 统计列数 |
|---|---|---|---|---|---|
| 2 | (10, 12) | H (8) | J-S, U-AF | AH (34) | 9 |
| 3 | (8, 10, 12) | K (11) | M-T, V-AE, AG-AR | AT (46) | 15 |
| 4 | (6, 8, 10, 12) | N (14) | P-U, W-AD, AF-AO, AQ-BB | BD (56) | 23 |

## 命令

`--input` 是**必填参数**，必须显式传入 Excel 路径，避免相对路径踩 CWD 的坑、也避免误改某个"默认文件"。`--output` 不传则默认 = `--input`（就地修改），建议先备份。

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# 写到旁路文件（推荐）
python scripts/generate_excel_random_data.py \
    --input "20260509-随机数生成.xlsx" \
    --output /tmp/out.xlsx \
    --seed 42

# 就地修改（先 cp 备份）
cp "20260509-随机数生成.xlsx" "20260509-随机数生成.bak.xlsx"
python scripts/generate_excel_random_data.py \
    --input "20260509-随机数生成.xlsx" \
    --seed 42
```

## CLI 参数

- `--input PATH` **必填**，输入 Excel 文件路径
- `--output PATH` 默认与 `--input` 相同（**就地修改源文件**，请先备份）
- `--sheet NAME` 默认第一个 sheet
- `--seed INT` 不传则随机；INFO 日志会打印实际种子，便于复现
- `-v, --verbose` 把日志级别从 INFO 调到 DEBUG（输出每次 SD 容差重试 / Levene·SW 内部值 / 分支选择等）

## 退出码

- `0` 正常完成
- `2` argparse 参数错误（如缺 `--input`）
- `3` 输入文件结构无法解析（找不到位数列 / K<2）

## 统计方法说明

| 步骤 | 方法 | 阈值 |
|---|---|---|
| 方差齐性 | Levene's test, `center='median'` (Brown-Forsythe 变种) | p > 0.05 视为齐 |
| 正态性 | Shapiro-Wilk per group, 取最小 p | p > 0.05 视为全部正态 |
| 整体差异 | 齐 → ANOVA (`f_oneway`)；不齐 → Kruskal-Wallis | — |
| 两两对比 | 齐+全正态 → Tukey HSD（无 Q-value）；否则 → Welch t + Bonferroni | — |
| Bonferroni | `Q = min(1.0, raw × C(K,2))` | 封顶 1.0 |

> SD 容差：生成的原始数据按行级 `decimals` 四舍五入后，实际 SD 与目标 SD 的相对误差需在 ±10% 内；超出会换种子重试，最多 5 次仍失败抛 `RuntimeError`。
