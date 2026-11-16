# skills-dwg-change-plan-pdf

DWG 施工图 + 竣工图 → **变更方案附图 PDF** 的 WorkBuddy 技能。

把两张 DWG 的平面图导出为黑白矢量 A1 横向 PDF，加"设计方案：/变更方案："文字说明，并按单项工程合并为 `工程编号+工程名称-变更方案附图.pdf`。

## 功能

- 黑白矢量 A1 横向出图（也支持 A2/A3/A4）
- 自动检测竣工图的"索引排版"，裁出平面图框、去掉材料表
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

单对图纸：

```bash
python scripts/change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "输出目录"
```

多对图纸（按单项工程分别合并）：

```bash
python scripts/change_plan_pdf.py --pairs pairs.json --out-dir "输出目录"
```

也可直接在 WorkBuddy 对话中提供施工图与竣工图 DWG，说"出变更方案附图"，技能会自动触发。

## 依赖

- Windows + 浩辰 GStarCAD / 中望 / AutoCAD（COM 接口）
- Python 3.8+：`pywin32`、`pymupdf`、`numpy`、`PIL`
- 中文字体 `C:\Windows\Fonts\simhei.ttf`

## 文件

- `SKILL.md` — 技能说明与触发条件
- `scripts/change_plan_pdf.py` — 主程序（CLI）
- `references/technical_notes.md` — COM 接口、纸张介质、字体诊断等排错参考

## License

MIT
