# -*- coding: utf-8 -*-
"""
图框定位工具：在导出的 raw PDF（整套图纸拼版）中列出候选图框并生成编号预览图。

适用场景：DWG 的布局是「整套图纸拼版」——封面、图纸目录、材料表、设计说明、
"请勿打印"索引页、A2 小图框等挤在同一张 A1 上，change_plan_pdf.py 的自动
索引排版检测会裁错图框。本工具把候选框逐一列出并编号，便于肉眼/AI 选定后
把坐标传给 change_plan_pdf.py 的 --sj-clip / --jg-clip。

用法：
  python inspect_frames.py "sj_raw.pdf" "jg_raw.pdf"
  python inspect_frames.py "某.pdf" --out-dir "预览目录"

输出：
  - 终端：页面尺寸、每个候选图框的序号与 PDF 坐标 (x0,y0,x1,y1)、
          是否检测到整页 X 标记（"请勿打印"索引页特征）、拼版布局判定；
  - --out-dir 目录：frames.json（机器可读坐标）、
          frame_N.png（单框放大预览）、overview.png（整页画框编号）。
"""
from __future__ import annotations

__version__ = "1.4.1"

import argparse
import json
import os
import sys
from typing import List, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8", errors="replace")
    except Exception:
        pass


def frame_candidates(page, min_frac=0.12, max_frac=0.98) -> List[Tuple[float, float, float, float]]:
    """返回页面中足够大的矢量矩形框（去重复），按 y0,x0 排序。"""
    pw, ph = page.rect.width, page.rect.height
    rects = []
    for g in page.get_drawings():
        r = g["rect"]
        if (pw * min_frac < r.width < pw * max_frac and
                ph * min_frac < r.height < ph * max_frac and
                r.x0 > 10 and r.y0 > 10):
            rects.append(r)
    uniq: List[Tuple[float, float, float, float]] = []
    for r in rects:
        t = (round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1))
        if not any(all(abs(a - b) < 4 for a, b in zip(t, u)) for u in uniq):
            uniq.append(t)
    uniq.sort(key=lambda t: (t[1], t[0]))
    return uniq


def long_diagonals(page, frac=0.15):
    """统计贯穿页面的长斜线（"请勿打印"索引页的大 X）。"""
    pw, ph = page.rect.width, page.rect.height
    n = 0
    for g in page.get_drawings():
        for it in g["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(b.x - a.x) > pw * frac and abs(b.y - a.y) > ph * frac:
                    n += 1
    return n


def inspect_pdf(pdf_path: str, out_dir: str, verbose=False) -> dict:
    import fitz
    d = fitz.open(pdf_path)
    p = d[0]
    pw, ph = p.rect.width, p.rect.height
    frames = frame_candidates(p)
    diag = long_diagonals(p)
    # 拼版判定：>=3 个大框，或有贯穿整页的 X 标记
    composite = len(frames) >= 3 or diag >= 2
    info = {
        "pdf": os.path.basename(pdf_path),
        "page": [round(pw, 1), round(ph, 1)],
        "frames": [list(map(lambda v: round(v, 1), r)) for r in frames],
        "long_diagonals": diag,
        "composite_layout": composite,
    }
    print("=== %s ===" % os.path.basename(pdf_path))
    print("  页面尺寸: %.1f x %.1f pt" % (pw, ph))
    print("  候选图框: %d 个 | 贯穿斜线: %d 条" % (len(frames), diag))
    if composite:
        print("  ⚠ 判定为整套图纸拼版布局：自动套框不可靠，请用下面的坐标指定 --sj-clip / --jg-clip")
    for i, r in enumerate(frames):
        print("    frame_%d: %s  (宽%.0f 高%.0f)" % (
            i, ",".join("%.1f" % v for v in r), r[2] - r[0], r[3] - r[1]))

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        # 整页画框编号
        ov = fitz.open()
        ov.insert_pdf(d)
        ovp = ov[0]
        for i, r in enumerate(frames):
            ovp.draw_rect(fitz.Rect(*r), color=(1, 0, 0), width=3)
            ovp.insert_text((r[0] + 4, r[1] + 22), "frame_%d" % i,
                            fontsize=16, color=(1, 0, 0))
        ov[0].get_pixmap(matrix=fitz.Matrix(0.8, 0.8)).save(
            os.path.join(out_dir, "overview_%s.png" % os.path.splitext(os.path.basename(pdf_path))[0]))
        # 逐框放大预览
        for i, r in enumerate(frames):
            clip = fitz.Rect(*r)
            pix = p.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), clip=clip)
            pix.save(os.path.join(out_dir, "frame_%d_%s.png" % (
                i, os.path.splitext(os.path.basename(pdf_path))[0])))
        with open(os.path.join(out_dir, "frames.json"), "a", encoding="utf-8") as f:
            f.write(json.dumps(info, ensure_ascii=False) + "\n")
    d.close()
    return info


def main():
    ap = argparse.ArgumentParser(description="列出 PDF 中的候选图框，辅助选定平面图裁切坐标")
    ap.add_argument("pdf", nargs="+", help="已导出的 raw PDF（可多个）")
    ap.add_argument("--out-dir", help="预览图与 frames.json 输出目录")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()
    if args.out_dir:
        jf = os.path.join(args.out_dir, "frames.json")
        if os.path.exists(jf):
            os.remove(jf)
    for path in args.pdf:
        if not os.path.exists(path):
            print("跳过（不存在）: %s" % path, flush=True)
            continue
        inspect_pdf(path, args.out_dir, args.verbose)
    return 0


if __name__ == "__main__":
    sys.exit(main())
