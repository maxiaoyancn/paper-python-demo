## Why

v1 的 `excel-random-data-generator` 只支持 2 组（B-G 列固定描述两组、I 列起固定布局）。论文场景常需 ≥3 组对比（含整体差异 + 两两对比 + 多重比较校正），且不同指标的"小数位数"也不同。当前需手动改脚本或改文件结构来扩展，违反外科手术式改动原则。本次扩展把脚本升级为通用的"K 组（K≥2）随机数据生成 + 多组统计分析"流水线，并把"小数位数"作为输入数据的一部分。

## What Changes

- **BREAKING：输入区列布局重定义**
  - A 列：指标名
  - 紧跟 K 个 (N、μ、σ) 三元组：第 i 组 N 在 col(2 + 3(i-1))、μ 在 col(3 + 3(i-1))、σ 在 col(4 + 3(i-1))
  - 第 K 组之后**紧跟一列"原始数据小数点后位数"**（位置 col(2 + 3K)）；K 由读取时遇到位数列即停止动态确定
  - 位数列识别策略：行 1 表头**必须**包含子串 `位数`（兼容 `原始数据小数点后位数`、`位数` 等表头）
  - 位数列后**空 1 列**（用户原备注列），再开始数据区
- **BREAKING：原始数据表头样本量标注由 `<count>（N）/（M）` 改为 `<count>（G1/G2/G3...）`**，所有 K 统一新格式（`G` 代表 Group）
- **BREAKING：每个数值的小数位数由动态识别的"位数列"决定，不再硬编码 4 位**——位数列位置 = `col(2 + 3K)`，随 K 自动后挪：K=2 时落在 H 列、K=3 时 K 列、K=4 时 N 列…脚本通过扫描第 1 行表头中含子串"位数"的列定位它，**不绑定特定字母列**
- **新增：每行 K 由数据自适应**——同一文件内不同行可有不同 K
- **新增：统计输出列布局**（在数据区结束后空 1 列起，按顺序）：
  1. K 组的 (μ, σ) 对，共 2K 列（μ₁、σ₁、μ₂、σ₂、…、μ_K、σ_K）
  2. **Levene p 值**（真实数值，4 位小数；中心使用 median）—— 决定后续分支
  3. **Shapiro-Wilk 最小 p 值**（对每组分别做 SW 取最小，4 位小数）—— 与 Levene 联合决定后续分支
  4. **整体 p 值**：方差齐（Levene p > 0.05）→ ANOVA；不齐 → Kruskal-Wallis
  5. **两两 raw p**，按组对索引升序：(1,2)(1,3)(2,3)(1,4)... 共 C(K,2) 列
     - 方差齐 **且** 全组正态（min SW p > 0.05）→ Tukey HSD
     - 否则 → Welch t（独立两两）
  6. **两两校正后 p（Q-value）**，列序与 raw 对应，共 C(K,2) 列
     - Tukey 分支：留空（无 Bonferroni 校正语义）
     - Welch 分支：raw_p × C(K,2)，封顶 1.0（标准 Bonferroni）
- 回算的 μ、σ 与 p 值**全部 round 到 4 位小数**
- **保留**：CLI 接口（`--input` 必填、`--output` / `--sheet` / `--seed` / `--verbose`）、`±10%` SD 容差重试机制、就地 vs 旁路写入

## Capabilities

### New Capabilities
<!-- 无新 capability。 -->

### Modified Capabilities
- `excel-random-data-generator`: 全面改写读取/生成/统计/写入四层的需求以支持 K 组（K≥2）+ 行级位数 + Levene/SW/ANOVA/KW/Tukey/Welch+Bonferroni 流水线。

## Impact

- 代码：重写 [scripts/generate_excel_random_data.py](../../../scripts/generate_excel_random_data.py) 的数据结构（`RowSpec.groups: tuple[GroupSpec, ...]` + 新 `decimals: int` 字段）、布局算法（动态识别位数列）、生成器（按 decimals 而非固定 4 位 round）、统计层（新增 Levene/SW/ANOVA/KW/Tukey/Welch+Bonferroni）、写入层（K 组 + 新统计列布局）。
- 测试：[tests/](../../../tests/) 全部用例（read/generate/stats/layout/cli）需重写，覆盖 K=2、K=3、K=4；新增对 Tukey 与 Welch+Bonferroni 双分支的覆盖。
- 数据：用户已把 [20260509-随机数生成.xlsx](../../../20260509-随机数生成.xlsx) 调整为 v2 输入示例（K=2 场景下位数列正好落在 H 列，表头 `'原始数据小数点后位数'`、值 = 4，数据从 J 列开始）。**位数列是按 K 动态定位的，不绑定 H**——K=3 时位数列移到 K 列、数据起始移到 M 列。脚本会**完全重写** "位数列+2" 至最末统计列范围；输入区与备注列保留。
- 文档：[scripts/README.md](../../../scripts/README.md) 列布局段需重写；新版 docstring。
- 依赖：复用现有 `scipy`（已有 `levene` / `shapiro` / `f_oneway` / `kruskal` / `tukey_hsd` / `ttest_ind`），无需新增第三方依赖。
- 兼容性：**v2 与 v1 输出列布局不兼容**；v1 旧 xlsx（H 列填备注、统计列只有 5 列）需用户先按新布局调整。本次 archive 后 v1 历史保留在 `openspec/changes/archive/2026-05-09-add-excel-random-data-generator/`。
