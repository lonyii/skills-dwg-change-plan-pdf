# -*- coding: utf-8 -*-
"""
把大幅面 PDF（合并版变更方案附图）切割为多张 A4，每张加编号标注。

分割规则（按 ISO 面积比，1:1 矢量平铺，不缩放）：
  A1 -> 8 张 A4 (4列x2行)    A2 -> 4 张 (2x2)
  A3 -> 2 张 (2x1)          A4 -> 1 张

页序：第 1 组 PDF 的 1..N 页（标注"设计方案1..N"），
      第 2 组 PDF 的 1..N 页（标注"变更方案1..N"）。

用法：
  python split_to_a4.py "设计方案.pdf" "变更方案.pdf" --out "变更方案附图.pdf"
  python split_to_a4.py "变更方案附图.pdf" --out "变更方案附图_A4.pdf"  # 单文件模式：各页按序标注
"""
from __future__ import annotations

__version__ = "0.2.0-test"

import argparse, math, os, sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8", errors="replace")
    except Exception:
        pass

import fitz

A4_W, A4_H = 595.276, 841.890       # A4 竖向 pt
A4_LW, A4_LH = 841.890, 595.276     # A4 横向 pt
FONT = "china-s"                    # fitz 内置中文字体，无需外部字体文件


def split_spec(w: float, h: float):
    """返回 (cols, rows, a4_w, a4_h, n)：按面积比求 A4 网格，选宽高比最接近的排布。"""
    n = max(1, round((w * h) / (A4_W * A4_H)))
    best = None
    for cols in range(1, n + 1):
        if n % cols:
            continue
        rows = n // cols
        cw, ch = w / cols, h / rows
        dv = abs(math.log((cw / ch) / (A4_W / A4_H)))
        dl = abs(math.log((cw / ch) / (A4_LW / A4_LH)))
        d, tw, th = (dv, A4_W, A4_H) if dv <= dl else (dl, A4_LW, A4_LH)
        if best is None or d < best[0]:
            best = (d, cols, rows, tw, th)
    _, cols, rows, tw, th = best
    return cols, rows, tw, th, n


def stamp(page, text: str, size: int = 13):
    """在页面右上角放带白底的黑色编号标注（内置中文字体）。"""
    w = page.rect.width
    tw = fitz.get_text_length(text, fontname=FONT, fontsize=size)
    x, y = w - tw - 16, 25
    page.draw_rect(fitz.Rect(x - 5, y - size - 3, w - 8, y + 5),
                   fill=(1, 1, 1), color=None, overlay=True)
    page.insert_text((x, y), text, fontname=FONT, fontsize=size, color=(0, 0, 0))


def content_bbox(page, threshold=200, pad_pt=14):
    """渲染缩略图找非白像素边界，映射回 PDF 坐标。

    返回 (fitz.Rect, used_flag)。图纸内容偏在一角时，据此裁掉大片空白，
    让内容铺满 A4、文字保持原始大小（清晰）；内容本就铺满时返回近整页。
    """
    import numpy as np
    from PIL import Image
    zoom = 0.5
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    gray = img.mean(axis=2)
    ys, xs = np.where(gray < threshold)
    if len(xs) == 0:
        return page.rect, False
    x0, x1 = xs.min()/zoom, xs.max()/zoom
    y0, y1 = ys.min()/zoom, ys.max()/zoom
    r = fitz.Rect(max(0, x0-pad_pt), max(0, y0-pad_pt),
                  min(page.rect.width, x1+pad_pt), min(page.rect.height, y1+pad_pt))
    # 只有内容显著小于整页（面积 < 55%）时才认为需要裁空白；
    # 否则一律按整页纸张分幅，避免边缘留白导致张数判定偏小（如 A1 被误判成 7 张）。
    area_ratio = (r.width * r.height) / (page.rect.width * page.rect.height)
    used = area_ratio < 0.55
    return (r if used else page.rect), used


def cut_pdf(src_pdf: str, label: str, out_pdf: str, verbose=False) -> int:
    """把单个 PDF 的第一页按内容区域切成多张 A4，逐张标注 label+序号。

    1:1 矢量平铺、不缩放；自动裁掉内容外的大片空白，保证文字原始大小。
    """
    src = fitz.open(src_pdf)
    sp = src[0]
    bbox, trimmed = content_bbox(sp)
    if verbose and trimmed:
        print("  检测到内容边界 %s（原页 %s），裁掉外围空白" %
              (tuple(round(v) for v in bbox), tuple(round(v) for v in sp.rect)), flush=True)
    cols, rows, aw, ah, n = split_spec(bbox.width, bbox.height)
    if verbose:
        print("  %s -> %d 张 A4 (%dx%d, %.0fx%.0f)" %
              (label, n, cols, rows, aw, ah), flush=True)
    out = fitz.open()
    bw, bh = bbox.width, bbox.height
    idx = 0
    for row in range(rows):
        for col in range(cols):
            clip = fitz.Rect(bbox.x0 + col*bw/cols, bbox.y0 + row*bh/rows,
                             bbox.x0 + (col+1)*bw/cols, bbox.y0 + (row+1)*bh/rows)
            page = out.new_page(width=aw, height=ah)
            page.show_pdf_page(fitz.Rect(0, 0, aw, ah), src, 0,
                               clip=clip, keep_proportion=True)
            idx += 1
            stamp(page, "%s%d" % (label, idx))
    out.save(out_pdf, garbage=3, deflate=True)
    out.close(); src.close()
    return n


def main():
    ap = argparse.ArgumentParser(description="大幅面 PDF 切割为 A4 并加编号标注")
    ap.add_argument("pdf", nargs="+", help="1个合并PDF，或2个PDF(设计/变更)")
    ap.add_argument("--out", required=True, help="输出 PDF 路径")
    ap.add_argument("--sj-label", default="设计方案", help="第一组标注前缀")
    ap.add_argument("--jg-label", default="变更方案", help="第二组标注前缀")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    import tempfile
    work = tempfile.mkdtemp(prefix="a4split_")
    parts = []
    if len(args.pdf) == 2:
        specs = [(args.pdf[0], args.sj_label), (args.pdf[1], args.jg_label)]
    else:
        # 单文件模式：按页序，前半设计方案后半变更方案无法确定，统一用 sj-label
        specs = [(args.pdf[0], args.sj_label)]
    for i, (pdf, label) in enumerate(specs):
        part = os.path.join(work, "part%d.pdf" % i)
        cut_pdf(pdf, label, part, args.verbose)
        parts.append(part)

    m = fitz.open()
    for p in parts:
        d = fitz.open(p); m.insert_pdf(d); d.close()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    m.save(args.out, garbage=3, deflate=True)
    print("输出: %s (%d 页)" % (args.out, m.page_count), flush=True)
    m.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
