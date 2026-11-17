# skills-dwg-change-plan-pdf

**正式版 v1.3.2** — 核心流程已实现并经多样本验证（黑白矢量导出、索引排版裁切、加字、单项工程合并、A4 自动分幅、尺寸判定）。

**v1.2.0 重点**：新增 `--frame-print` 高分辨率图框打印模式（推荐）——按图框定位、放大打印后再裁 A4，从根源解决 SHX 文字黑块/线条过粗；不再做宋体替换，保留原始 SHX 大字体。

版本历史：v0.1.0-test（初始）→ v0.2.0-test（A4 分幅、尺寸判定、按编号配对）→ v1.0.0（正式版）→ v1.1.0（拼版布局：图框定位工具 + 手工 clip + A4 单幅输出）→ v1.2.0（高分辨率图框打印，保留 SHX 字体，解决黑块）→ v1.3.2（`--plan-only` 只输出平面图，双下划线自动定位）。

DWG 施工图 + 竣工图 → **变更方案附图 PDF** 的 WorkBuddy 技能。

把两张 DWG 的平面图导出为黑白矢量 A1 横向 PDF，加"设计方案：/变更方案："文字说明，并按单项工程合并为 `工程编号+工程名称-变更方案附图.pdf`。

## 功能

- 黑白矢量 A1 横向出图（也支持 A2/A3/A4）
- **成品默认 A4 分幅**：大图自动切为多张 A4（A1→8、A2→4、A3→2），每张标注"设计方案1/2/3…""变更方案1/2/3…"（见 `scripts/split_to_a4.py`）
- 自动检测竣工图的"索引排版"，裁出平面图框、去掉材料表
- **整套图纸拼版布局**：自动检测大框≥3 / "请勿打印"大 X，警告并支持用 `inspect_frames.py` 定位候选图框、`--sj-clip/--jg-clip` 精确裁切、`--a4-single` 整幅出 A4（详见 SKILL.md「整套图纸拼版布局」）
- 自动修复竣工图缺失 SHX 大字体导致的**文字糊成黑块**问题
- 平面图右侧加可自定义文字说明
- 施工图 + 竣工图合并为两页 PDF（设计方案在前、变更方案在后）
- 多张图按单项工程（以竣工图为准）分别合并

## 安装

将本目录放到 WorkBuddy 技能目录：

```bash
# 用户级（推荐）
git clone https://github.com/lonyii/skills-dwg-change-plan-pdf.git ~/.workbuddy/skills/dwg-change-plan-pdf
```

## 使用

单对图纸（推荐 v1.2.0 高分辨率模式）：

```bash
python scripts/change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "输出目录" --frame-print --scale 4
```

传统模式（非拼版图纸）：

```bash
python scripts/change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "输出目录"
```

多对图纸（按单项工程分别合并）：

```bash
python scripts/change_plan_pdf.py --pairs pairs.json --out-dir "输出目录"
# 或只给竣工图，按编号自动在施工图目录配对：
python scripts/change_plan_pdf.py --pairs pairs.json --sj-dir "施工图目录" --out-dir "输出目录"
```

一份施工图 DWG 含多个工程图纸时，以竣工图工程编号为准定位对应施工图/布局：
- 施工图文件名含相同编号 → 自动匹配；
- 布局名含该编号 → 自动选该布局；
- 都不含 → 在 pairs.json 里为该条目指定 `"sj_layout": "布局名"`。

pairs.json 可写：`[{"jg": "竣工图NC25E096-xxx.dwg", "sj_layout": "总平面图2"}]`

也可直接在 WorkBuddy 对话中提供施工图与竣工图 DWG，说"出变更方案附图"，技能会自动触发。

## 依赖

- Windows + 浩辰 GStarCAD / 中望 / AutoCAD（COM 接口）
- Python 3.8+：`pywin32`、`pymupdf`、`numpy`、`PIL`
- 中文字体 `C:\Windows\Fonts\simhei.ttf`

## 文件

- `SKILL.md` — 技能说明与触发条件
- `scripts/change_plan_pdf.py` — 主程序（CLI）
- `scripts/inspect_frames.py` — 拼版布局的候选图框定位与预览
- `scripts/split_to_a4.py` — A4 分幅切割
- `references/technical_notes.md` — COM 接口、纸张介质、字体诊断等排错参考
- `references/naming_patterns.md`、`references/a4_split_standard.md` — 样本实测规律与判定标准
- `assets/fonts/` — hztxt.shx / gbcbig.shx / stedi.shx 备查（实测浩辰 PDF 驱动仍无法渲染 SHX，出图一律走宋体替换，见技术文档）

## License

MIT
