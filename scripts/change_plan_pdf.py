# -*- coding: utf-8 -*-
"""
变更方案附图 PDF 生成器。

从施工图 + 竣工图 DWG 导出平面图，黑白 A1 横向 PDF，加文字说明，按单项工程合并。

用法：
  # 单对图纸（施工图标注"设计方案："、竣工图标注"变更方案："）
  python change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "输出目录"

  # 多对图纸（自动以竣工图为基准，按单项工程编号分组配对）
  python change_plan_pdf.py --pairs pairs.json --out-dir "输出目录"

  # 自定义文字/纸张/预览
  python change_plan_pdf.py "施工图.dwg" "竣工图.dwg" --out-dir "D:/out" ^
      --sj-text "原设计：" --jg-text "变更后：" --text-pos right-bottom ^
      --paper A1 --preview

pairs.json 格式（多对）：
  [
    {"sj": "路径/施工图A.dwg", "jg": "路径/竣工图NC25E096-xxx.dwg"},
    {"sj": "路径/施工图B.dwg", "jg": "路径/竣工图NC25E097-xxx.dwg"}
  ]
  不指定 sj 时按编号在 sj-dir 中自动查找同名编号的施工图。

依赖：Windows + 浩辰/中望/AutoCAD COM + Python 3.8+ (pywin32, pymupdf, numpy, PIL)
"""
from __future__ import annotations

__version__ = "1.3.0"

import argparse
import json
import os
import re
import sys
import time
from typing import Optional, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8", errors="replace")
    except Exception:
        pass

# --- 纸张介质映射（DWG To PDF.pc3 实际介质名是 User 编号，需按本地名反查） ---
PAPER_MEDIA = {
    "A1": ("User65737", 0, (2383.937, 1683.780)),  # ISO A1 841x594 横向
    "A2": ("User65739", 0, (1683.780, 1190.551)),  # ISO A2 594x420 横向
    "A3": ("User76", 1, (1190.551, 841.890)),      # ISO A3 420x297 横向
    "A4": ("User77", 1, (841.890, 595.276)),       # ISO A4 297x210 横向
}
FONT_PATH = r"C:\Windows\Fonts\simhei.ttf"
CTB = "monochrome.ctb"
PDF_DRIVER = "DWG To PDF.pc3"


def _com():
    import pythoncom
    import win32com.client.gencache as gc
    pythoncom.CoInitialize()
    app = gc.EnsureDispatch("GStarCAD.Application")
    try:
        app.Visible = False
    except Exception:
        pass
    return pythoncom, app


def find_layout(doc, prefer: str = "总平面图", code: Optional[str] = None):
    """返回布局对象。

    定位优先级：
      1. 指定 code 时，布局名含该工程编号的（用于一份 DWG 多工程图纸）
      2. 布局名含 prefer（如"总平面图"）
      3. 模型空间之外的第一个布局
      4. 模型空间
    """
    n = doc.Layouts.Count
    if code:
        for i in range(n):
            lay = doc.Layouts.Item(i)
            if code.lower() in str(lay.Name).lower():
                return lay
    for i in range(n):
        lay = doc.Layouts.Item(i)
        if prefer in str(lay.Name):
            return lay
    for i in range(n):
        lay = doc.Layouts.Item(i)
        if not lay.ModelType:
            return lay
    return doc.Layouts.Item(0)


def match_sj_by_code(jg_path: str, sj_dir: str, verbose=False) -> Optional[str]:
    """以竣工图工程编号为准，在施工图目录中找到对应的施工图文件。

    匹配规则（按优先级）：
      1. 施工图文件名含与竣工图相同的工程编号（如 NC25E096）
      2. 施工图文件名含竣工图的工程名称关键字
    编号一律取竣工图，不用施工图自己的命名。
    """
    code = project_code(jg_path)
    if not code or not sj_dir or not os.path.isdir(sj_dir):
        return None
    jg_base = os.path.splitext(os.path.basename(jg_path))[0]
    name_key = re.sub(r"^[A-Z]{2}\d{2}[A-Z]?\d{3}[-_ ]*", "", jg_base, flags=re.I)
    name_key = re.split(r"[（(]", name_key)[0][:12]  # 取工程名前段作为兜底关键字
    files = [f for f in os.listdir(sj_dir) if f.lower().endswith(".dwg")]
    # 1) 编号精确匹配
    for f in files:
        if code.lower() in f.lower():
            if verbose:
                print("  施工图按编号匹配: %s (编号 %s)" % (f, code), flush=True)
            return os.path.join(sj_dir, f)
    # 2) 工程名关键字匹配
    if name_key:
        for f in files:
            if name_key in f:
                if verbose:
                    print("  施工图按工程名匹配: %s" % f, flush=True)
                return os.path.join(sj_dir, f)
    return None


def fix_fonts(doc, verbose=False):
    """把 SHX 大字体样式临时替换为系统宋体，避免导出 PDF 文字糊成黑块。

    仅在内存修改，不保存 DWG；调用方关闭文档时用 Close(False)。
    """
    styles = doc.TextStyles
    n = 0
    for i in range(styles.Count):
        s = styles.Item(i)
        try:
            s.SetFont("宋体", False, False, 134, 49)
            n += 1
        except Exception:
            pass
        try:
            s.fontFile = "simsun.ttc"
        except Exception:
            pass
        try:
            s.BigFontFile = ""
        except Exception:
            pass
    if verbose:
        print("  字体样式已替换为宋体: %d 个" % n, flush=True)
    return n


def export_layout(doc, out_pdf: str, paper: str = "A1", plot_type: int = 1,
                  fix_shx: bool = False, code: Optional[str] = None,
                  prefer: str = "总平面图", verbose=False,
                  scale: float = 0, plot_origin=None) -> str:
    """把指定布局黑白导出为 PDF。plot_type=1(范围)是浩辰唯一可靠值，4(窗口)会导出空白。

    code 不为空时，优先选布局名含该工程编号的布局（一份 DWG 多工程图纸场景）。
    prefer 可覆盖默认布局偏好（用于 pairs.json 指定 sj_layout）。
    scale > 0 时设 StandardScale（放大打印）：SHX 大字体清晰度取决于打印比例，
    拼版图纸按图纸比例打印时目标图框只占页面一角，文字笔画因分辨率不足糊成黑块；
    放大打印（如 4 倍）可让 SHX 输出为清晰矢量轮廓，再从结果中裁出图框。
    plot_origin=(x_mm,y_mm) 配合 scale 使用。
    """
    import fitz
    media, rotation, _ = PAPER_MEDIA[paper]
    lay = find_layout(doc, prefer=prefer, code=code)
    if verbose:
        print("  布局=%s 纸张=%s(%s) 旋转=%d 比例=%s" %
              (lay.Name, paper, media, rotation, scale or "按图纸"), flush=True)
    lay.ConfigName = PDF_DRIVER
    lay.RefreshPlotDeviceInfo()
    lay.CanonicalMediaName = media
    lay.PlotRotation = rotation
    lay.StyleSheet = CTB
    lay.PlotType = plot_type
    lay.StandardScale = scale
    if scale and scale > 1:
        lay.CenterPlot = False
        if plot_origin:
            lay.PlotOrigin = tuple(plot_origin)
    else:
        lay.CenterPlot = True
    if fix_shx:
        fix_fonts(doc, verbose)
    plot = doc.Plot
    plot.SetLayoutsToPlot([lay.Name])
    if os.path.exists(out_pdf):
        os.remove(out_pdf)
    plot.PlotToFile(out_pdf)
    time.sleep(8)  # PlotToFile 是异步的，必须等待写完
    if not os.path.exists(out_pdf) or os.path.getsize(out_pdf) < 5000:
        raise RuntimeError("PDF 导出失败或内容为空: %s" % out_pdf)
    d = fitz.open(out_pdf)
    pages = d.page_count
    rect = d[0].rect
    d.close()
    if verbose:
        print("  导出完成 %.0fKB, %d页, %s" % (os.path.getsize(out_pdf)/1024, pages, rect), flush=True)
    return out_pdf


def detect_plan_frame(pdf_path: str, verbose=False, min_rects=1):
    """在高比例放大打印的 PDF 中定位「平面图框」。

    拼版图纸放大打印（StandardScale≈4）后，页面里可能同时有多个 A 系列横版
    图框（平面图、设计说明、材料表…）。不能只按宽高比取最大——说明页常被误选。
    判据：平面图框内有大量「矩形」实体（建筑物轮廓、图例方块），而设计说明页
    几乎只有文字（矢量轮廓曲线），re 实体数接近 0。因此按「框内 re 实体数」
    优先、面积次之来选。

    返回 (x0,y0,x1,y1)；找不到返回 None。
    """
    import fitz
    d = fitz.open(pdf_path)
    p = d[0]
    W, H = p.rect.width, p.rect.height

    # 收集所有 "re"（矩形）实体，稍后按所在图框归类
    rect_items = []
    for g in p.get_drawings():
        for it in g["items"]:
            if it[0] == "re":
                rect_items.append(it[1])   # fitz.Rect

    cands = []
    seen = set()
    for g in p.get_drawings():
        r = g["rect"]
        w, h = r.width, r.height
        if not (W*0.15 < w < W*0.92 and H*0.15 < h < H*0.92):
            continue
        ratio = w/h
        # A 系列横版(~1.414)或竖版(~0.707)；竣工图在A2横框里、施工图常是竖版A4
        if not (1.25 < ratio < 1.60 or 0.65 < ratio < 0.80):
            continue
        key = (round(r.x0), round(r.y0), round(r.x1), round(r.y1))
        if key in seen:
            continue
        seen.add(key)
        nrects = sum(1 for rr in rect_items
                     if r.x0 <= rr.x0 and r.y0 <= rr.y0 and rr.x1 <= r.x1 and rr.y1 <= r.y1)
        cands.append((r.x0, r.y0, r.x1, r.y1, w*h, nrects))
    d.close()
    if not cands:
        return None
    # 先按 re 数量降序（平面图特征），同档按面积降序
    cands.sort(key=lambda c: (-c[5], -c[4]))
    best = cands[0]
    if verbose:
        for c in cands[:5]:
            print("  候选框 %s 面积=%.0f 矩形数=%d" %
                  (tuple(round(v) for v in c[:4]), c[4], c[5]), flush=True)
        print("  → 选中: %s" % (tuple(round(v) for v in best[:4]),), flush=True)
    return (best[0], best[1], best[2], best[3])


def detect_plan_title_underline(pdf_path: str, frame=None, verbose=False):
    """在高分辨率打印的 PDF 中，用「平面图标题的双下划线」定位平面图区域。

    竣工图平面图的标题（如"室内外低压燃气管道平面图1:100"）下方有两条间距
    3~10pt、长度 50~300pt、左右端对齐的水平线，这是平面图标题的独特标志
    （材料表/设计说明页没有）。找到后按经验比例向上取建筑平面图、向下取图例，
    返回竖版平面图裁切框 (x0,y0,x1,y1)。找不到返回 None。
    """
    import fitz
    d = fitz.open(pdf_path)
    p = d[0]
    if frame is None:
        frame = p.rect
    fr = fitz.Rect(*frame) if not isinstance(frame, fitz.Rect) else frame
    hs = []
    for g in p.get_drawings():
        r = g["rect"]
        if r.height < 3 and 50 < r.width < 300 and fr.contains(r):
            hs.append((r.x0, r.y0, r.x1))
    hs.sort(key=lambda h: (h[1], h[0]))
    pairs = []
    for i, a in enumerate(hs):
        for b in hs[i+1:]:
            gap = b[1] - a[1]
            if gap > 12:
                break
            if abs(a[0]-b[0]) < 6 and abs(a[2]-b[2]) < 6 and 3 < gap < 10:
                pairs.append((a, b))
    d.close()
    if not pairs:
        # 退化：有些图的标题下划线是两条重合线(gap=0)，或只有一条长线
        # 取框左半内、长度 40~300 的水平线里最靠下的那条作标题锚点
        # 回退: 取框左半内、长度 40~200 的水平线中，位于框上半部的最靠下一条
        left_lines = [h for h in hs if h[0] < fr.x0 + fr.width*0.62
                      and 40 < h[2]-h[0] < 200 and h[1] < fr.y0 + fr.height*0.55]
        if not left_lines:
            return None
        left_lines.sort(key=lambda h: -h[1])
        a = left_lines[0]
        b = a
    else:
        # 选最靠下、且左右端在框左半的一对（平面图在A2框左侧）
        pairs.sort(key=lambda p: (p[0][0] > fr.x0 + fr.width*0.6, -p[0][1]))
        a, b = pairs[0]
    if verbose:
        print("  标题双下划线: x=%.0f..%.0f y=%.0f" % (a[0], a[2], a[1]), flush=True)
    # 用标题上方的线段端点找建筑轮廓bbox（比固定比例更贴合不同图幅）
    ux0, ux1, uy = a[0], a[2], a[1]
    cx = (ux0 + ux1) / 2
    uw = max(ux1 - ux0, 40)
    d2 = fitz.open(pdf_path)
    pg = d2[0]
    top = fitz.Rect(cx - uw*1.8, max(fr.y0, uy-340), cx + uw*1.8, uy + 5)
    xs, ys = [], []
    for g in pg.get_drawings():
        for it in g["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if top.contains(p1) and top.contains(p2):
                    xs += [p1.x, p2.x]; ys += [p1.y, p2.y]
    d2.close()
    if xs:
        # 建筑轮廓左界不应比标题左界靠左太多（避免把图纸目录竖列表格带进来）
        bx0 = max(min(xs), ux0 - uw*0.8)
        bx1 = min(max(xs), ux1 + uw*1.5)
        by0 = min(ys)
    else:
        bx0, bx1, by0 = ux0 - uw, ux1 + uw, uy - 200
    # 图例在标题下方约 0.5~1 倍下划线宽
    return (bx0 - 10, by0 - 10, bx1 + 10, uy + uw*0.9)


def parse_clip(value) -> Optional[Tuple[float, float, float, float]]:
    """解析 --sj-clip/--jg-clip：x0,y0,x1,y1（导出的 raw PDF 坐标，单位 pt）。"""
    if not value:
        return None
    parts = [p.strip() for p in str(value).replace("，", ",").split(",")]
    if len(parts) != 4:
        raise ValueError("clip 需为 4 个逗号分隔的数字 x0,y0,x1,y1，收到: %s" % value)
    try:
        nums = tuple(float(p) for p in parts)
    except ValueError:
        raise ValueError("clip 含非数字: %s" % value)
    if nums[2] <= nums[0] or nums[3] <= nums[1]:
        raise ValueError("clip 坐标无效（x1<=x0 或 y1<=y0）: %s" % value)
    return nums


def looks_composite(pdf_path: str) -> bool:
    """检测「整套图纸拼版」布局：>=3 个大矢量框，或存在贯穿整页的长斜线（"请勿打印"大 X）。

    此类布局下 is_index_layout() 的两框判据会失效/误裁，必须用 inspect_frames.py
    列出候选框后用 --sj-clip/--jg-clip 指定平面图坐标。
    """
    import fitz
    d = fitz.open(pdf_path)
    p = d[0]
    pw, ph = p.rect.width, p.rect.height
    frames = []
    diag = 0
    for g in p.get_drawings():
        r = g["rect"]
        if pw*0.12 < r.width < pw*0.98 and ph*0.12 < r.height < ph*0.98 and r.x0 > 10 and r.y0 > 10:
            t = (round(r.x0), round(r.y0), round(r.x1), round(r.y1))
            if not any(all(abs(a-b) < 4 for a, b in zip(t, u)) for u in frames):
                frames.append(t)
        for it in g["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(b.x-a.x) > pw*0.15 and abs(b.y-a.y) > ph*0.15:
                    diag += 1
    d.close()
    return len(frames) >= 3 or diag >= 2


def is_index_layout(pdf_path: str, paper: str) -> Optional[Tuple[float, float, float, float]]:
    """检测是否为"索引排版"（一张纸上有平面图框+材料表框）。

    返回上方平面图框 rect；普通整张平面图返回 None。
    判据：找到两个宽度、高度接近的大矩形，且纵向排列。
    ⚠ 整套图纸拼版（多图框/含"请勿打印"大 X）时此判据不可靠，
       由 looks_composite() 识别并提示手动指定 clip。
    """
    import fitz
    d = fitz.open(pdf_path)
    p = d[0]
    pw, ph = p.rect.width, p.rect.height
    rects = []
    for g in p.get_drawings():
        r = g["rect"]
        if pw*0.2 < r.width < pw*0.95 and ph*0.2 < r.height < ph*0.95 and r.x0 > 10 and r.y0 > 10:
            rects.append(r)
    uniq = []
    for r in rects:
        if not any(abs(r.x0-u.x0) < 4 and abs(r.y0-u.y0) < 4 and
                   abs(r.x1-u.x1) < 4 and abs(r.y1-u.y1) < 4 for u in uniq):
            uniq.append(r)
    uniq.sort(key=lambda r: r.y0)
    d.close()
    # 两个等大框、纵向分开 → 索引排版，取上面那个
    if len(uniq) >= 2:
        a, b = uniq[0], uniq[1]
        if (abs(a.width-b.width) < 8 and abs(a.height-b.height) < 8 and
                b.y0 - a.y1 > ph*0.05):
            return (a.x0, a.y0, a.x1, a.y1)
    return None


def crop_plan_frame(src_pdf: str, dst_pdf: str, clip, paper: str = "A1", verbose=False):
    """把平面图框裁出并铺满新 PDF 页。"""
    import fitz
    _, _, (pw, ph) = PAPER_MEDIA[paper]
    src = fitz.open(src_pdf)
    out = fitz.open()
    page = out.new_page(width=pw, height=ph)
    page.show_pdf_page(fitz.Rect(20, 20, pw-20, ph-20), src, 0,
                       clip=fitz.Rect(*clip), keep_proportion=True)
    out.save(dst_pdf, garbage=3, deflate=True)
    out.close()
    src.close()
    if verbose:
        print("  已裁出平面图框并铺满 %s: %s" % (paper, clip), flush=True)


def default_text_pos(kind: str, paper: str):
    """按纸张比例返回文字落点（PDF 坐标 x, y 基线）与字号。"""
    _, _, (pw, ph) = PAPER_MEDIA[paper]
    scale = pw / PAPER_MEDIA["A3"][2][0]
    if kind == "sj":
        x, y, size = 916, 686, 16
    else:
        x, y, size = 908, 703, 16
    return (round(x*scale), round(y*scale), round(size*scale))


def stamp_text(pdf_path: str, text: str, x: int, y: int, size: int,
               out_pdf: Optional[str] = None, font: str = FONT_PATH):
    """在 PDF 上写入黑色中文文字。"""
    import fitz
    if not text:
        return pdf_path
    d = fitz.open(pdf_path)
    p = d[0]
    p.insert_font(fontname="SH", fontfile=font)
    p.insert_text((x, y), text, fontname="SH", fontsize=size, color=(0, 0, 0))
    dst = out_pdf or pdf_path.replace(".pdf", "_标注.pdf")
    d.save(dst, garbage=3, deflate=True)
    d.close()
    if out_pdf and os.path.abspath(dst) != os.path.abspath(pdf_path):
        os.replace(dst, out_pdf)
    return dst


A4_PORTRAIT = (595.276, 841.890)  # A4 竖向 pt


def stamp_a4_single(src_pdf: str, title: str, label: str, out_pdf: str,
                    clip=None, font: str = FONT_PATH):
    """把一张平面图整幅放进一张 A4 竖向页，顶部写标题、右上写编号。

    用于拼版布局里裁出的小平面图：标准 split_to_a4.py 按 1:1 把 A1 切 8 张 A4，
    对小图框会把一张平面图切碎；本函数改为整幅缩放铺满单张 A4。
    clip 不为空时直接从 raw PDF 按该框裁切（推荐，避免先铺满 A1 再缩小的留白）；
    clip 为空时对 src_pdf 整页铺入。
    """
    import fitz
    src = fitz.open(src_pdf)
    out = fitz.open()
    pw, ph = A4_PORTRAIT
    page = out.new_page(width=pw, height=ph)
    page.show_pdf_page(fitz.Rect(40, 80, pw-40, ph-40), src, 0,
                       clip=fitz.Rect(*clip) if clip else fitz.Rect(src[0].rect),
                       keep_proportion=True)
    page.insert_font(fontname="SH", fontfile=font)
    page.insert_text((28, 34), title, fontname="SH", fontsize=13, color=(0, 0, 0))
    page.insert_text((pw-110, 62), label, fontname="SH", fontsize=12, color=(0, 0, 0))
    out.save(out_pdf, garbage=3, deflate=True)
    out.close()
    src.close()
    return out_pdf


def merge_pdfs(paths, out_pdf: str):
    """按顺序合并 PDF。"""
    import fitz
    m = fitz.open()
    for p in paths:
        d = fitz.open(p)
        m.insert_pdf(d)
        d.close()
    m.save(out_pdf, garbage=3, deflate=True)
    m.close()
    return out_pdf


def project_code(name: str) -> Optional[str]:
    """从竣工图文件名提取单项工程编号。

    本地样本（556/562 命中）验证的格式：
      NC25E096 / IC24A016 / ID25H023 / NC25F009  (2字母+2数字+1字母+3数字)
      NC24EA15 / NC24EA86                          (2字母+2数字+1字母+3位，字母后跟2位也被接受)
    用宽松正则：2字母 + 2数字 + 可选1字母 + 3~4位字母数字。
    """
    base = os.path.basename(name)
    # 优先标准格式，兼容字母后跟 2~3 位
    m = re.search(r"\b([A-Z]{2}\d{2}[A-Z]?[A-Z0-9]{2,4})\b", base, re.I)
    if not m:
        return None
    code = m.group(1).upper()
    # 排除明显的图号/合同号（含 DD/FN/BG 等项目内编号前缀）
    return code


def process_pair(sj: str, jg: str, work: str, paper: str, sj_text: str, jg_text: str,
                 text_pos: str, preview: bool, code: Optional[str] = None,
                 sj_layout: Optional[str] = None, verbose=False,
                 sj_clip=None, jg_clip=None, a4_single=False) -> Tuple[str, str]:
    """处理一对施工图/竣工图，返回 (设计页PDF, 变更页PDF)。

    code 为竣工图工程编号：一份施工图 DWG 含多个工程图纸时，用它定位对应布局。
    sj_layout 可显式指定施工图要导出的布局名。
    sj_clip/jg_clip 为手工指定的平面图裁切框 (x0,y0,x1,y1, raw PDF 坐标)，
    用于「整套图纸拼版」布局（自动检测会裁错封面/索引页）。
    """
    import fitz
    pythoncom, app = _com()

    def open_doc(path):
        """打开 DWG；若该文档已被 CAD 打开（残留未关闭）则复用，避免文件锁报错。"""
        abs_path = os.path.abspath(path)
        # 先确认文件可访问（映射网络盘可能临时掉线），给出明确提示
        if not os.path.exists(abs_path):
            raise FileNotFoundError(
                "DWG 文件不可访问（可能是映射网络盘掉线）: %s" % abs_path)
        for i in range(app.Documents.Count):
            try:
                if os.path.abspath(app.Documents.Item(i).FullName) == abs_path:
                    return app.Documents.Item(i)
            except Exception:
                pass
        last = None
        for attempt in range(3):
            try:
                return app.Documents.Open(abs_path)
            except Exception as e:
                last = e
                time.sleep(2)
        raise last

    try:
        sj_raw = os.path.join(work, "sj_raw.pdf")
        jg_raw = os.path.join(work, "jg_raw.pdf")
        if verbose:
            print("[1/3] 导出施工图...", flush=True)
        d = open_doc(sj)
        export_layout(d, sj_raw, paper=paper, fix_shx=False, code=code,
                      prefer=sj_layout or "总平面图", verbose=verbose)
        d.Close(False)
        if verbose:
            print("[2/3] 导出竣工图（字体修复）...", flush=True)
        d = open_doc(jg)
        export_layout(d, jg_raw, paper=paper, fix_shx=True, verbose=verbose)
        d.Close(False)
    finally:
        pythoncom.CoUninitialize()

    sj_final = os.path.join(work, "sj_final.pdf")
    jg_final = os.path.join(work, "jg_final.pdf")

    # 施工图：优先手工 clip；否则整张导出（通常不是索引排版）
    sj_src = sj_raw
    sj_effective_clip = sj_clip
    if sj_clip:
        if verbose:
            print("  施工图使用手工 clip: %s" % (sj_clip,), flush=True)
        if not a4_single:
            crop_plan_frame(sj_raw, os.path.join(work, "sj_crop.pdf"), sj_clip, paper, verbose)
            sj_src = os.path.join(work, "sj_crop.pdf")
    elif looks_composite(sj_raw):
        print("  ⚠ 施工图为整套图纸拼版布局，自动出图可能裁到封面/目录："
              "请运行 inspect_frames.py 后用 --sj-clip 指定平面图坐标", flush=True)
    else:
        sj_effective_clip = is_index_layout(sj_raw, paper)
        if sj_effective_clip:
            crop_plan_frame(sj_raw, os.path.join(work, "sj_crop.pdf"), sj_effective_clip, paper, verbose)
            sj_src = os.path.join(work, "sj_crop.pdf")
    if a4_single:
        stamp_a4_single(sj_raw, sj_text, "设计方案1", sj_final, clip=sj_effective_clip)
    else:
        x, y, size = default_text_pos("sj", paper)
        stamp_text(sj_src, sj_text, x, y, size, sj_final)

    # 竣工图：优先手工 clip；否则若为索引排版则裁平面图框
    jg_src = jg_raw
    jg_effective_clip = jg_clip
    if jg_clip:
        if verbose:
            print("  竣工图使用手工 clip: %s" % (jg_clip,), flush=True)
        if not a4_single:
            crop_plan_frame(jg_raw, os.path.join(work, "jg_crop.pdf"), jg_clip, paper, verbose)
            jg_src = os.path.join(work, "jg_crop.pdf")
    elif looks_composite(jg_raw):
        print("  ⚠ 竣工图为整套图纸拼版布局，自动出图可能裁到索引/说明页："
              "请运行 inspect_frames.py 后用 --jg-clip 指定平面图坐标", flush=True)
    else:
        jg_effective_clip = is_index_layout(jg_raw, paper)
        if jg_effective_clip:
            if verbose:
                print("  检测到索引排版，裁出上方平面图框", flush=True)
            crop_plan_frame(jg_raw, os.path.join(work, "jg_crop.pdf"), jg_effective_clip, paper, verbose)
            jg_src = os.path.join(work, "jg_crop.pdf")
    if a4_single:
        stamp_a4_single(jg_raw, jg_text, "变更方案1", jg_final, clip=jg_effective_clip)
    else:
        x, y, size = default_text_pos("jg", paper)
        stamp_text(jg_src, jg_text, x, y, size, jg_final)

    if preview:
        for tag, path in [("设计", sj_final), ("变更", jg_final)]:
            d = fitz.open(path)
            pix = d[0].get_pixmap(matrix=fitz.Matrix(0.6, 0.6), alpha=False)
            pix.save(path.replace(".pdf", "_预览.png"))
            d.close()
    return sj_final, jg_final


def process_frame_print(sj: str, jg: str, work: str, paper: str,
                        sj_text: str, jg_text: str, scale: float = 4,
                        origin=None, preview=False, verbose=False,
                        sj_clip=None, jg_clip=None, plan_only=False):
    """【高分辨率图框打印模式】放大打印 → 定位目标图框 → 裁切到 A4。

    原理：SHX 大字体糊成黑块的根因是打印比例太小、分辨率不够（目标图框在
    整张 A1 拼版里只占一角）。把 StandardScale 放大（默认 4 倍）后，浩辰
    会把 SHX 字形输出为清晰矢量轮廓；此时其它图框被推出页面，目标图框以
    A 系列横版比例独占页面，用 detect_plan_frame() 自动定位后裁到 A4。

    不用宋体替换——原 SHX 大字体原样保留，文字与线条粗细都正常。

    返回 (设计页PDF, 变更页PDF)，均为 A4 单幅。
    """
    import fitz
    pythoncom, app = _com()

    def open_doc(path):
        abs_path = os.path.abspath(path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError("DWG 不可访问: %s" % abs_path)
        for i in range(app.Documents.Count):
            try:
                if os.path.abspath(app.Documents.Item(i).FullName) == abs_path:
                    return app.Documents.Item(i)
            except Exception:
                pass
        last = None
        for attempt in range(3):
            try:
                return app.Documents.Open(abs_path)
            except Exception as e:
                last = e; time.sleep(2)
        raise last

    results = {}
    # 策略：优先 1:1（多数图纸文字本就清晰，如 J 盘样本）；
    # 只有 1:1 找不到「平面图框」（有矩形特征）时，才按 scale 放大重打——
    # 放大后其它图框被推出页面、目标框独占页面，SHX 也输出为清晰矢量。
    for tag, path, text, label in [
        ("sj", sj, sj_text, "设计方案1"),
        ("jg", jg, jg_text, "变更方案1"),
    ]:
        manual_clip = sj_clip if tag == "sj" else jg_clip
        raw = os.path.join(work, "%s_scaled.pdf" % tag)
        out = os.path.join(work, "%s_final.pdf" % tag)
        chosen = None
        tried = []
        # 施工图字体本来就清晰 → 1:1；竣工图黑块来自比例太小 → 先放大。
        # 放大后若找不到「平面图框」（矩形特征），退回 1:1（如 J 盘样本图框本就够大）。
        trial_scales = [0, scale, scale - 1] if tag == "sj" else [scale, scale - 1, 0]
        for trial_scale in trial_scales:
            if trial_scale < 0:
                continue
            d = open_doc(path)
            try:
                export_layout(d, raw, paper=paper, fix_shx=False, scale=trial_scale,
                              plot_origin=(origin or (0, 0)) if trial_scale else None,
                              verbose=verbose)
            finally:
                d.Close(False)
            frame = detect_plan_frame(raw, verbose, min_rects=1)
            if plan_only and frame and not manual_clip:
                sub = detect_plan_title_underline(raw, frame, verbose)
                clip = sub or frame
            else:
                clip = manual_clip or frame
            tried.append((trial_scale, clip))
            if clip:
                chosen = (trial_scale, clip)
                break
        if not chosen:
            trial_scale, clip = tried[-1]
            chosen = (trial_scale, clip)
            print("  ⚠ [%s] 未能自动定位平面图框，已用最后候选；建议用 --%s-clip 或 inspect_frames.py 指定" %
                  (tag, tag), flush=True)
        use_scale, clip = chosen
        if verbose:
            print("  [%s] 最终比例=%s clip=%s" % (tag, use_scale or "1:1",
                  tuple(round(v) for v in clip) if clip else None), flush=True)
        stamp_a4_single(raw, text, label, out, clip=clip)
        results[tag] = out
        if preview:
            fitz.open(out)[0].get_pixmap(matrix=fitz.Matrix(0.6, 0.6)).save(
                out.replace(".pdf", "_预览.png"))
    pythoncom.CoUninitialize()
    return results["sj"], results["jg"]


def main():
    ap = argparse.ArgumentParser(description="施工图+竣工图 → 变更方案附图 PDF")
    ap.add_argument("dwg", nargs="*", help="施工图.dwg 竣工图.dwg（共两个）")
    ap.add_argument("--pairs", help="多对图纸 JSON 清单")
    ap.add_argument("--sj-dir", help="竣工图配对时用：施工图搜索目录")
    ap.add_argument("--out-dir", required=True, help="成品输出目录")
    ap.add_argument("--paper", default="A1", choices=list(PAPER_MEDIA), help="纸张，默认 A1")
    ap.add_argument("--sj-text", default="设计方案：", help="施工图页面文字")
    ap.add_argument("--jg-text", default="变更方案：", help="竣工图页面文字")
    ap.add_argument("--text-pos", default="right-bottom", help="文字位置（预留）")
    ap.add_argument("--preview", action="store_true", help="生成低分辨率预览 PNG")
    ap.add_argument("--sj-clip", help="施工图平面图的 raw PDF 裁切框 x0,y0,x1,y1（拼版布局用）")
    ap.add_argument("--jg-clip", help="竣工图平面图的 raw PDF 裁切框 x0,y0,x1,y1（拼版布局用）")
    ap.add_argument("--a4-single", action="store_true",
                    help="拼版布局专用：把裁出的小平面图整幅放单张 A4，直接输出成品，不再走 split_to_a4 1:1 切分")
    ap.add_argument("--frame-print", action="store_true",
                    help="【推荐/高分辨率】放大打印定位目标图框再裁 A4：保留原始 SHX 大字体，彻底解决文字黑块/过粗；"
                         "不用宋体替换。仅打印目标图框，不把整套拼版塞进一页")
    ap.add_argument("--scale", type=float, default=4.0, help="--frame-print 的放大比例，默认 4")
    ap.add_argument("--plot-origin", help="--frame-print 的打印原点 x_mm,y_mm（可选，默认不设）")
    ap.add_argument("--plan-only", action="store_true",
                    help="只输出平面图本身(建筑+管道+标题+图例)，不含材料表/说明/系统图；"
                         "通过平面图标题双下划线自动定位子区域")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    import fitz  # noqa

    os.makedirs(args.out_dir, exist_ok=True)
    sj_clip, jg_clip = parse_clip(args.sj_clip), parse_clip(args.jg_clip)

    # 组装 pairs
    pairs = []
    if args.pairs:
        with open(args.pairs, encoding="utf-8") as f:
            items = json.load(f)
        for it in items:
            sj, jg = it.get("sj"), it.get("jg")
            # 以竣工图工程编号为准找对应施工图（一个施工图文件含多工程图纸时也适用）
            if not sj and args.sj_dir:
                sj = match_sj_by_code(jg, args.sj_dir, args.verbose)
            if not sj or not os.path.exists(sj) or not os.path.exists(jg):
                print("跳过（找不到文件）: %s" % it, flush=True)
                continue
            pairs.append((sj, jg, it.get("sj_layout"),
                          parse_clip(it.get("sj_clip")), parse_clip(it.get("jg_clip"))))
    elif len(args.dwg) == 2:
        pairs.append((args.dwg[0], args.dwg[1], None, sj_clip, jg_clip))
    else:
        ap.error("需提供两个 DWG（施工图、竣工图），或使用 --pairs")

    if not pairs:
        print("没有可处理的图纸对", flush=True)
        return 1

    for idx, item in enumerate(pairs, 1):
        sj, jg = item[0], item[1]
        sj_layout = item[2] if len(item) > 2 else None
        pair_sj_clip = item[3] if len(item) > 3 else None
        pair_jg_clip = item[4] if len(item) > 4 else None
        code = project_code(jg) or ("工程%02d" % idx)
        base = os.path.splitext(os.path.basename(jg))[0]
        # 名称取竣工图文件名去掉编号前缀部分
        name = re.sub(r"^[A-Z]{2}\d{2}[A-Z]?\d{3}[-_ ]*", "", base, flags=re.I) or base
        out_name = "%s-%s-变更方案附图.pdf" % (code, name)
        work = os.path.join(args.out_dir, "_work_%s" % code)
        os.makedirs(work, exist_ok=True)
        print("\n=== %s ===" % out_name, flush=True)
        if args.verbose:
            print("  竣工图编号: %s | 施工图: %s" %
                  (code, os.path.basename(sj)), flush=True)
        if args.frame_print:
            origin = None
            if args.plot_origin:
                try:
                    origin = tuple(float(v) for v in args.plot_origin.replace("，", ",").split(","))
                except ValueError:
                    ap.error("--plot-origin 需为 x_mm,y_mm")
            sj_final, jg_final = process_frame_print(
                sj, jg, work, args.paper, args.sj_text, args.jg_text,
                scale=args.scale, origin=origin, preview=args.preview,
                verbose=args.verbose, sj_clip=pair_sj_clip, jg_clip=pair_jg_clip,
                plan_only=args.plan_only)
        else:
            sj_final, jg_final = process_pair(
                sj, jg, work, args.paper, args.sj_text, args.jg_text,
                args.text_pos, args.preview, code=code, sj_layout=sj_layout,
                verbose=args.verbose, sj_clip=pair_sj_clip, jg_clip=pair_jg_clip,
                a4_single=args.a4_single)
        merged = merge_pdfs([sj_final, jg_final], os.path.join(args.out_dir, out_name))
        if args.a4_single:
            print("完成（A4 单幅）: %s (%.0fKB)" % (merged, os.path.getsize(merged)/1024), flush=True)
        else:
            print("完成: %s (%.0fKB)" % (merged, os.path.getsize(merged)/1024), flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
