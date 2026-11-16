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
                  prefer: str = "总平面图", verbose=False) -> str:
    """把指定布局黑白导出为 PDF。plot_type=1(范围)是浩辰唯一可靠值，4(窗口)会导出空白。

    code 不为空时，优先选布局名含该工程编号的布局（一份 DWG 多工程图纸场景）。
    prefer 可覆盖默认布局偏好（用于 pairs.json 指定 sj_layout）。
    """
    import fitz
    media, rotation, _ = PAPER_MEDIA[paper]
    lay = find_layout(doc, prefer=prefer, code=code)
    if verbose:
        print("  布局=%s 纸张=%s(%s) 旋转=%d" % (lay.Name, paper, media, rotation), flush=True)
    lay.ConfigName = PDF_DRIVER
    lay.RefreshPlotDeviceInfo()
    lay.CanonicalMediaName = media
    lay.PlotRotation = rotation
    lay.StyleSheet = CTB
    lay.PlotType = plot_type
    lay.StandardScale = 0
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


def is_index_layout(pdf_path: str, paper: str) -> Optional[Tuple[float, float, float, float]]:
    """检测是否为"索引排版"（一张纸上有平面图框+材料表框）。

    返回上方平面图框 rect；普通整张平面图返回 None。
    判据：找到两个宽度、高度接近的大矩形，且纵向排列。
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
                 sj_layout: Optional[str] = None, verbose=False) -> Tuple[str, str]:
    """处理一对施工图/竣工图，返回 (设计页PDF, 变更页PDF)。

    code 为竣工图工程编号：一份施工图 DWG 含多个工程图纸时，用它定位对应布局。
    sj_layout 可显式指定施工图要导出的布局名。
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

    # 施工图：整张导出（通常不是索引排版）
    clip = is_index_layout(sj_raw, paper)
    sj_src = sj_raw
    if clip:
        crop_plan_frame(sj_raw, os.path.join(work, "sj_crop.pdf"), clip, paper, verbose)
        sj_src = os.path.join(work, "sj_crop.pdf")
    x, y, size = default_text_pos("sj", paper)
    stamp_text(sj_src, sj_text, x, y, size, sj_final)

    # 竣工图：若是索引排版，裁平面图框
    clip = is_index_layout(jg_raw, paper)
    jg_src = jg_raw
    if clip:
        if verbose:
            print("  检测到索引排版，裁出上方平面图框", flush=True)
        crop_plan_frame(jg_raw, os.path.join(work, "jg_crop.pdf"), clip, paper, verbose)
        jg_src = os.path.join(work, "jg_crop.pdf")
    x, y, size = default_text_pos("jg", paper)
    stamp_text(jg_src, jg_text, x, y, size, jg_final)

    if preview:
        for tag, path in [("设计", sj_final), ("变更", jg_final)]:
            d = fitz.open(path)
            pix = d[0].get_pixmap(matrix=fitz.Matrix(0.6, 0.6), alpha=False)
            pix.save(path.replace(".pdf", "_预览.png"))
            d.close()
    return sj_final, jg_final


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
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    import fitz  # noqa

    os.makedirs(args.out_dir, exist_ok=True)

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
            pairs.append((sj, jg, it.get("sj_layout")))
    elif len(args.dwg) == 2:
        pairs.append((args.dwg[0], args.dwg[1], None))
    else:
        ap.error("需提供两个 DWG（施工图、竣工图），或使用 --pairs")

    if not pairs:
        print("没有可处理的图纸对", flush=True)
        return 1

    for idx, item in enumerate(pairs, 1):
        sj, jg = item[0], item[1]
        sj_layout = item[2] if len(item) > 2 else None
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
        sj_final, jg_final = process_pair(
            sj, jg, work, args.paper, args.sj_text, args.jg_text,
            args.text_pos, args.preview, code=code, sj_layout=sj_layout,
            verbose=args.verbose)
        merged = merge_pdfs([sj_final, jg_final], os.path.join(args.out_dir, out_name))
        print("完成: %s (%.0fKB)" % (merged, os.path.getsize(merged)/1024), flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
