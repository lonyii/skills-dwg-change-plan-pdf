---
name: dwg-change-plan-pdf
description: 从施工图与竣工图 DWG 导出平面图，黑白 A1 横向 PDF，加"设计方案：/变更方案："文字说明，并按单项工程合并为「工程编号+工程名称-变更方案附图.pdf」。当用户提供施工图与竣工图 DWG，要求出变更方案附图、平面图 PDF、黑白打印、裁剪平面图、加设计方案/变更方案标注、或多张图按单项工程合并 PDF 时使用本技能。
agent_created: true
---

# 变更方案附图 PDF（DWG 平面图出图）

## 概述

把施工图/竣工图 DWG 的平面图导出为**黑白矢量 A1 横向 PDF**，在平面图右侧加文字说明（默认施工图"设计方案："、竣工图"变更方案："），最终合并为一份两页 PDF。适用于燃气/市政等工程的变更方案附图出图。

## 环境依赖

- Windows；浩辰 GStarCAD / 中望 / AutoCAD 其一（本机浩辰 GStarCAD 3.05，COM ProgID `GStarCAD.Application`）
- Python 3.8+，依赖：`pywin32`、`pymupdf`、`numpy`、`PIL`
- 中文字体：`C:\Windows\Fonts\simhei.ttf`（脚本默认）

## 一键使用

单对图纸（施工图 + 竣工图）：

```bash
python scripts/change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "输出目录"
```

多对图纸（按单项工程分别合并）：

```bash
python scripts/change_plan_pdf.py --pairs pairs.json --out-dir "输出目录"
```

`pairs.json`：

```json
[
  {"sj": "施工图A.dwg", "jg": "竣工图NC25E096-xxx.dwg"},
  {"sj": "施工图B.dwg", "jg": "竣工图NC25E097-xxx.dwg"}
]
```

仅给竣工图、施工图按编号在同目录自动配对时，用 `--sj-dir 施工图目录` 并省略 sj 字段。

常用选项：

| 选项 | 说明 |
|---|---|
| `--paper A1/A2/A3/A4` | 纸张，默认 A1 横向 |
| `--sj-text` / `--jg-text` | 两页文字内容 |
| `--preview` | 同时输出低分辨率预览 PNG |
| `--verbose` | 打印详细过程 |

## 输出口径（强制）

1. **单对图纸**：施工图 + 竣工图合并为一份 PDF——第 1 页设计方案（施工图）、第 2 页变更方案（竣工图）。
2. **多张图/多个文件**：**以竣工图为准、按单项工程分别合并 PDF**。
   - 单项工程的划分、工程编号、工程名称一律取竣工图（文件名前缀编号，如 NC25E096）。
   - 每个单项工程各自配对其施工图与竣工图，单独出一份合并 PDF，不跨工程混拼。
3. **命名**：`工程编号+工程名称-变更方案附图.pdf`
   （如 `NC25E096-江西省中西医结合医院蒸汽热源及锅炉燃气工程（室外）-变更方案附图.pdf`）。

## 关键实现要点（踩坑记录）

详细技术说明见 `references/technical_notes.md`。核心提醒：

- **纸张 A1 在 CAD 里不叫 "A1"**：`DWG To PDF.pc3` 介质名是 User 编号，A1 横向 = `User65737` 且 `PlotRotation=0`（用 1 会变竖版）。脚本已内置映射。
- **`PlotType=1`（范围）是浩辰唯一可靠值**；`PlotType=4`（窗口）会导出空白 888B 文件。
- **`PlotToFile` 是异步的**，调用后必须等待 8 秒再读取文件。
- **竣工图文字糊成黑块 = 缺 SHX 大字体**（hztxt.shx/gbcbig.shx/nchz.shx 等）。导出前须把所有文字样式替换为系统宋体：
  `SetFont("宋体",False,False,134,49)` + `fontFile="simsun.ttc"` + `BigFontFile=""`，仅在内存修改，`Close(False)` 不保存 DWG。**施工图字体正常，不要替换**。
- **竣工图布局常是"索引排版"**（平面图框 + 材料表框上下排列）。脚本自动检测，裁出上方平面图框铺满 A1，去掉材料表。施工图直接整张导出。
- 文字位置按纸张比例从 A3 基准 (916,686)/(908,703) 自动换算；需要精确指定时改 `default_text_pos()`。

## 资源

- `scripts/change_plan_pdf.py`：主程序，封装导出/字体修复/裁切/加字/合并全流程，可直接 CLI 调用。
- `references/technical_notes.md`：浩辰 COM 打印接口、介质映射、字体诊断与排错参考。
