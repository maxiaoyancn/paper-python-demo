# Excel 随机数据生成脚本

## 用途

读取一份 Excel 表（第 1 行为表头，第 2 行起每行一个指标），按 A-G 列描述的两组目标统计量（样本量 N/M、均值、SD）生成两组正态分布原始数据，并回算实际 mean/SD 与 Levene→Student/Welch 双尾 p 值。

## 列布局（动态，按 N、M 计算，不写死）

- A-H 列：原表头与备注（脚本不修改）。
- I 列起：第一组数据，长度 N。最后一格的第 1 行表头会被覆写为 `<N>（N）`。
- 第一组之后空 1 列。
- 紧跟第二组数据，长度 M。最后一格的第 1 行表头会被覆写为 `<M>（M）`。
- 第二组之后空 1 列。
- 接着 5 列统计：第一组均值 / 第一组SD值 / 第二组均值 / 第二组SD值 / 两组的p值。

当 N=10、M=12 时，三块区域恰好对应 I-R / T-AE / AG-AK。

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
- `-v, --verbose` 把日志级别从 INFO 调到 DEBUG
