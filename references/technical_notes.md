# 技术参考（浩辰 GStarCAD COM + PDF 出图）

## 1. COM 接口结构

用 `win32com.client.gencache.EnsureDispatch("GStarCAD.Application")` 得到带类型库的包装，接口来自浩辰的 Gcad 类型库（GUID 02316B95-30D8-446F-BE6E-B2D6103DDDA4）。

关键对象：

| 对象 | 位置 | 说明 |
|---|---|---|
| `Application.Documents` | app | 打开/关闭 DWG |
| `Document.Layouts.Item(i)` | doc | 布局，**打印属性都在这里** |
| `Document.Plot` | doc | 只有 `PlotToFile/PlotToDevice/SetLayoutsToPlot` 等精简方法 |
| `Document.TextStyles` | doc | 文字样式，`SetFont`/`fontFile`/`BigFontFile` |

`doc.Plot` 上**没有** `ConfigName`/`StyleSheet`/`PlotType` 等属性——这些在 Layout 上。

Layout 常用属性：
`ConfigName`, `StyleSheet`, `CanonicalMediaName`, `PlotRotation`, `PlotType`, `StandardScale`,
`CenterPlot`, `ModelType`, `Name`, `GetPaperSize()`, `GetCanonicalMediaNames()`,
`GetPlotDeviceNames()`, `GetPlotStyleTableNames()`, `RefreshPlotDeviceInfo()`。

## 2. 纸张介质映射（DWG To PDF.pc3）

`GetCanonicalMediaNames()` 返回 113 项，全是 `UserNNNNN`，直接搜 "A1" 搜不到，须用
`GetLocaleMediaName("UserNNNNN")` 反查。实测：

| 介质名 | 本地名 | 尺寸(mm) | 建议旋转 |
|---|---|---|---|
| User200 | ISO A0 | 841×1189 | 0 |
| **User65737** | **ISO A1** | **841×594** | **0（横向）** |
| User202 | ISO A1 | 594×841 | 0（竖向） |
| **User65739** | ISO A2 | 594×420 | 0（横向） |
| User204 | ISO A2 | 420×594 | 0（竖向） |
| User76 | ISO A3 | 420×297 | 1 |
| User77 | ISO A4 | 297×210 | 1 |

注意：User65737 本身已是横向尺寸（宽>高），再设 `PlotRotation=1` 会转成竖版
（输出 1683.78×2383.94pt）。正确是 `PlotRotation=0`，输出 2383.94×1683.78pt。

其它驱动：`Microsoft Print to PDF` 支持 A2/A3/A4；`Adobe PDF` 只有 A4；
`pdfFactory Pro` 在本机设置时会报"输入无效"。

## 3. PlotType 取值（浩辰实测）

| PlotType | 结果 |
|---|---|
| 1（范围/Extents） | ✅ 完整导出，矢量对象数 4.5万~12万 |
| 4（窗口/Window） | ❌ 导出 888B 空白，即使 `SetWindowToPlot` 正确 |
| 5（布局） | ❌ 934B 空白 |
| 0（显示） | 内容错乱/缩小 |

固定用 `PlotType=1` + `StandardScale=0` + `CenterPlot=True`。

`PlotToFile(path)` 是**异步**的，立即返回后文件尚未写完；必须 `time.sleep(8)` 再读取。

## 4. 字体糊成黑块的诊断与修复

**现象**：导出的 PDF 中文字全是实心黑块，放大也无笔画。

**根因**：图纸用了 CAD 的 SHX 大字体（常见 hztxt.shx、gbcbig.shx、nchz.shx、
gbenor.shx、@0hzdx.shx、stedi.shx 等），浩辰的 PDF 驱动**无法把 SHX 字形正确转成 PDF 轮廓**，
于是把文字画成实心块。施工图若用 TrueType（宋体/黑体）则正常。

**重要实测结论（2026-11-16 对比实验）**：即使本机已安装这些 SHX 字体
（浩辰 CAD2019 的 Fonts 目录里确实有 hztxt.shx / gbcbig.shx / stedi.shx），
保持原字体导出的 PDF **文字仍是黑块**；替换为宋体后文字笔画分明、清晰可读。
对比数据：原字体导出 708KB/64416 绘图对象但文字全黑块；宋体替换 1187KB/27382 对象、文字清晰。
→ **结论：不要依赖 SHX 字体文件，一律用宋体替换方案**。字体文件仅作备查资源
（见 `assets/fonts/`），在未来换用其他 CAD/PDF 驱动时或许有用，当前环境不解决黑块问题。

**诊断**：

```python
for i in range(doc.TextStyles.Count):
    s = doc.TextStyles.Item(i)
    print(s.Name, s.BigFontFile)
```

大量 `*.shx` 即为风险。也可放大导出 PDF 局部，看文字是否只剩黑块。

**修复**（导出前，内存操作，不保存 DWG）：

```python
for i in range(doc.TextStyles.Count):
    s = doc.TextStyles.Item(i)
    s.SetFont("宋体", False, False, 134, 49)  # 参数是 TypeFace 名，不是路径
    s.fontFile = "simsun.ttc"
    s.BigFontFile = ""
doc.Close(False)  # 关键：不保存
```

`SetFont` 签名：`(TypeFace, Bold, Italic, Charset, PitchAndFamily)`。
修复后 PDF 体积会明显增大（字形轮廓替代黑块），属正常。

只对需要的图纸（通常竣工图）修复；施工图保持原字体。

## 5. 索引排版检测与裁切

竣工图的 `总平面图` 布局可能是"图纸索引页"：一张纸上排两个等大 A1 图框，上框放平面图、
下框放材料表，中间有引出线和"A1图框"标注。若直接导出，PDF 会包含材料表。

检测：用 pymupdf 读 `page.get_drawings()`，找宽高接近、纵向排列的两个大矩形
（宽度在页面 20%~95%、高度在页面 20%~95%、x0>10、y0>10），且 `b.y0 - a.y1 > 页高*0.05`。

裁切（以上框为 clip，铺满新 A1 页）：

```python
src = fitz.open(raw_pdf)
out = fitz.open()
page = out.new_page(width=2383.937, height=1683.780)
page.show_pdf_page(fitz.Rect(20,20,W-20,H-20), src, 0,
                   clip=fitz.Rect(x0,y0,x1,y1), keep_proportion=True)
out.save(dst, garbage=3, deflate=True)
```

## 5b. 整套图纸拼版布局（v1.1.0）

**现象**：默认参数出图，施工图页是封面（"XX工程施工图设计"）、竣工图页是"请勿打印"大 X 索引页 + 一堆材料表缩略图，完全没有管道平面图。

**根因**：DWG 布局里塞了整本图纸（封面、目录、材料表、设计说明、索引页、A2 小图框），`is_index_layout()` 的「两个等大框纵向排列」判据不匹配，脚本直接整页导出。

**识别**（`change_plan_pdf.py::looks_composite()` / `inspect_frames.py`）：
- 页面中「大矢量矩形框」数量 ≥ 3（标准索引排版只有 2 个）；或
- 存在贯穿页面 ≥15% 宽高的长斜线（"请勿打印"大 X）。

**处理**：
1. `python scripts/inspect_frames.py sj_raw.pdf jg_raw.pdf --out-dir 预览目录` → 输出每个候选框坐标与 `frame_N_*.png` 放大预览、`overview_*.png` 编号整页图。
2. 肉眼/AI 在预览里找到平面图所在框；若该框内还并排多张子图（如左目录右平面图），再用 pymupdf 按像素细裁到平面图本身（NC26E789：右上框右半 `(1500,48,1714,430)`；竣工 A2 小框左半 `(957,1312,1045,1455)`，页面 2384×1684 pt）。
3. 重跑主程序：`--a4-single --sj-clip "..." --jg-clip "..."`。`--a4-single` 把裁出的小图整幅缩放进单张 A4；**不要**对这类小图跑 `split_to_a4.py`——它按 1:1 把 A1 切成 8 张 A4，会把一张平面图切碎。
4. 出图后用 `--preview` 或渲染 PDF 首页目视复核，确认平面图铺满可用区、标题中文正常、无相邻表格残影。

**回归**：NC26E789 用上述流程生成的 2 页 A4 成品与人工精裁版本逐页像素完全一致。

## 5c. SHX 文字黑块 / 线条过粗的真正根因（v1.2.0 修正）

**用户口径（强制）**：不要替换宋体，保留原始 SHX 大字体。黑块是因为**打印尺寸太小、分辨率不够**，不是字体本身的问题。

**实证**（NC26E789 与 J 盘 IC25E339 对比）：
- 同一张竣工图，1:1 打印时图框只占 A1 页面一角 → SHX 笔画糊成黑块；
- `StandardScale=4` 放大打印后，同一 SHX 字体输出为清晰矢量轮廓，线条粗细正常、文字可读；
- 宋体替换版虽然不黑，但文字与线条明显偏粗偏黑——故 v1.2.0 起默认不再替换。

**正确流程**（`--frame-print`）：
1. 1:1 先探一次；竣工图黑块风险高，直接用 `StandardScale` 放大（默认 4）打印。
2. 放大后其它图框被推出页面；用 `get_drawings()` 找 A 系列横版矩形，并按「框内 `re` 矩形实体数」区分平面图框（有建筑轮廓/表格，re>0）与设计说明页（几乎只有文字，re=0）。
3. 施工图字体本就正常，取 1:1；竣工图若放大后找不到平面图框（如 J 盘样本图框本就够大），退回 1:1。
4. 裁出图框 → 放单张 A4，直接合并输出。

**注意**：`PlotType=4`（窗口）在浩辰下导出 888B 空白，不可用；`ZoomWindow` 也不影响 `PlotToFile` 的打印范围。唯一可靠放大手段是 `StandardScale`。

## 5d. 只输出平面图（--plan-only，v1.3.0）

用户要求变更方案附图**只要平面图本身**（建筑轮廓+燃气管道+标题"室内外低压燃气管道平面图1:100"+图例），
不要材料表、设计说明、系统图。

**定位方法**：在高分辨率打印的 PDF 里，平面图标题下方有两条间距 3~10pt、长度 50~300pt、
左右端对齐的水平线（标题下划线）——这是平面图独有的矢量特征（材料表/说明页没有）。
`detect_plan_title_underline()` 检测这对双线后，按经验比例向上取建筑平面图、向下取图例，
得到竖版平面图裁切框。NC26E789 实测：下划线 (1033,1022)-(1120,1027)，
裁切框 (1013,872)-(1225,1212)，即得用户要求的纯平面图。

**注意**：建筑轮廓在图纸空间、标题图例也在图纸空间，但整本图纸拼版时它们分散在不同区域；
直接从 PDF 高分辨率打印结果按标题下划线定位子区域最可靠，比在 DWG 里 Wblock 提取更稳
（Wblock 只能拿到图纸空间实体，模型空间轮廓会丢失）。

## 6. 文字标注

用 pymupdf 插入中文：

```python
page.insert_font(fontname="SH", fontfile=r"C:\Windows\Fonts\simhei.ttf")
page.insert_text((x, y), "设计方案：", fontname="SH", fontsize=26, color=(0,0,0))
```

文字落点用 PIL+numpy 积分图找纯空白窗口：把页面渲染为 PNG，二值化（灰度<128 为黑），
用积分图统计候选矩形内黑像素数，为 0 即为空白区，再按比例换算回 PDF 坐标。

字号按纸张等比：A3 基准 16pt，A1=32pt（脚本内 A1 用 26pt，经验值）。

## 7. 常见错误对照

| 现象 | 原因 | 处理 |
|---|---|---|
| PDF 只有 888B | PlotType=4 或 5 | 改 PlotType=1 |
| PDF 空白 | PlotToFile 异步未完成 | sleep 8 秒后再读 |
| 文字黑块 | 缺 SHX 字体 | 按第 4 节替换字体 |
| A1 变竖版 | User65737 用了旋转 1 | 改 PlotRotation=0 |
| 设 A2/A1 报"输入无效" | 直接传 "A2"/"A1" | 用 User 介质名 |
| 竣工图含材料表 | 索引排版未裁切 | 按第 5 节裁切 |

## 5e. "请勿打印"索引页的打印目标框定位（v1.3.1 修正）

多份竣工图 DWG 把整本图纸拼在一张 A1 上，带大 X 的"请勿打印"索引页在**上半页**，
真正要打印的图框（带 ✓ 和"打印N张"标记）在**下半页**。此前按矢量矩形选最大框会
选到材料表/索引框（E188 裁到请勿打印施工图、E339 设计方案裁到材料表）。

**正确做法**（`detect_print_target_frame()`）：栅格投影法——取页面下半 (y>0.5H) 的暗像素，
按 x 投影找最大连续暗块（目标图框是"实"的一块，比大 X 的细线更密），再取其行范围。
实测 E309/E188/E339 均准确定位到 ✓ 标记的目标框。

**平面图子区域**（`detect_plan_title_underline()`）：
- 标题下划线是**孤立的一对水平线**（间距 1~10pt、宽 35~300pt），材料表网格线是密集平行线
  （同 x 范围 40pt 内有 >4 条）——用邻线数量过滤。
- 标题在框**左半**（x < 框左+0.58宽），排除右侧材料表的窄表格线。
- 优先选**框中部**的线对（abs(y-框中心y) 最小），而非框底的材料表行。
- E339 标题下划线宽仅 47pt，故下限设为 35pt。

**残留**：少数样本平面图边缘会带一点相邻图框的线（如 E188 右侧系统图一角），
属可接受的裁切余量；如用户要求纯净可再收紧。
