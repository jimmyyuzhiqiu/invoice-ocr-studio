# -*- coding: utf-8 -*-
"""
多页PDF电子发票：按固定ROI裁剪 -> RapidOCR离线识别 -> 提取[发票号码, 开票日期, 价税合计] -> Excel

新增：
- --all_pages：处理PDF所有页（每页一行）
- --max_pages：限制最多处理前N页
- 输出增加“页码”列（从1开始）
"""

import os
import re
import json
import argparse
from pathlib import Path

import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import cv2

from rapidocr import RapidOCR

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
PDF_EXT = ".pdf"


def rotate_img(img, rotate: str):
    rotate = (rotate or "0").lower()
    if rotate in ["0", "none"]:
        return img
    if rotate in ["cw90", "90", "right", "r"]:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    if rotate in ["ccw90", "-90", "left", "l"]:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    if rotate in ["180", "flip"]:
        return cv2.rotate(img, cv2.ROTATE_180)
    raise ValueError(f"不支持的rotate参数: {rotate}")


def render_pdf_page_to_bgr(doc: fitz.Document, page_index: int, dpi: int):
    """渲染PDF指定页为BGR ndarray（不依赖poppler），逐页 page.get_pixmap(dpi=...)[1](https://juejin.cn/post/7510925068987252755)[2](https://pypi.org/project/rapidocr-onnxruntime/)"""
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img


def imread_unicode(path: str):
    p = Path(path)
    data = np.fromfile(str(p), dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        img = cv2.imread(str(p))
    return img


def crop_by_norm(img, norm_box):
    H, W = img.shape[:2]
    x1 = int(norm_box["x1"] * W)
    y1 = int(norm_box["y1"] * H)
    x2 = int(norm_box["x2"] * W)
    y2 = int(norm_box["y2"] * H)
    x1, y1 = max(0, x1), max(0, y1)
    x2, y2 = min(W, x2), min(H, y2)
    if x2 <= x1 or y2 <= y1:
        return None
    return img[y1:y2, x1:x2]


def upscale_if_small(img_bgr, min_h=60):
    if img_bgr is None or img_bgr.size == 0:
        return img_bgr
    h, w = img_bgr.shape[:2]
    if h < min_h:
        scale = max(2.0, min_h / max(h, 1))
        img_bgr = cv2.resize(img_bgr, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
    return img_bgr


def light_preprocess(img_bgr):
    """轻量预处理：增强对比度但不二值化，避免det找不到框"""
    if img_bgr is None or img_bgr.size == 0:
        return img_bgr
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)


def get_txt_from_rapid_output(out):
    """
    RapidOCR新版返回 RapidOCROutput，可通过 out.txts 访问文本[3](https://pymupdftest.readthedocs.io/en/stable/recipes-images.html)[4](https://sqlpey.com/python/solved-how-to-extract-a-pdf-page-as-a-jpeg/)
    """
    if out is None:
        return ""
    if hasattr(out, "txts"):
        txts = out.txts
        if not txts:
            return ""
        return "".join([t for t in txts if isinstance(t, str)]).strip()

    # 旧版兜底
    if isinstance(out, tuple) and len(out) == 2:
        maybe = out[0]
        if hasattr(maybe, "txts"):
            return "".join(maybe.txts or []).strip()
        if isinstance(maybe, list):
            return "".join([r[1] for r in maybe if len(r) >= 2]).strip()

    if isinstance(out, str):
        return out.strip()
    return ""


def ocr_text(engine: RapidOCR, img_bgr):
    """
    固定ROI建议先 use_det=False 跳过检测直接识别；为空再回退 use_det=True。[5](https://stackoverflow.com/questions/69643954/converting-pdf-to-png-with-python-without-pdf2image)[3](https://pymupdftest.readthedocs.io/en/stable/recipes-images.html)
    """
    if img_bgr is None or img_bgr.size == 0:
        return ""

    img_bgr = upscale_if_small(img_bgr)
    img_bgr = light_preprocess(img_bgr)

    # 1) 先不做检测（ROI很小更稳）
    try:
        out = engine(img_bgr, use_det=False, use_cls=False, use_rec=True)
        text = get_txt_from_rapid_output(out)
        if text:
            return text
    except Exception:
        pass

    # 2) 回退：做检测+识别（阈值稍放宽）
    try:
        out = engine(img_bgr, use_det=True, use_cls=False, use_rec=True, box_thresh=0.3, text_score=0.3)
        return get_txt_from_rapid_output(out)
    except Exception:
        return ""


def normalize_date(s: str) -> str:
    if not s:
        return s
    s = s.strip().replace(" ", "")
    s = s.replace("年", "-").replace("月", "-").replace("日", "")
    s = re.sub(r"[./]", "-", s)
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if not m:
        return s
    y, mo, d = m.groups()
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"


def normalize_amount(s: str) -> str:
    if not s:
        return s
    s = s.strip().replace(",", "").replace("￥", "").replace("¥", "").replace("RMB", "").replace(" ", "")
    m = re.search(r"(\d+(?:\.\d{1,2})?)", s)
    return m.group(1) if m else s

def extract_invoice_no(text: str):
    """
    电子发票常见：发票代码12位 + 发票号码8位 = 20位票号。[1](https://www.zlq.gov.cn/zlq/ztzl/zcwdk/ns/2023092617185722268/index.shtml)[2](https://zhidao.baidu.com/question/2276102211575308708.html)
    目标：优先返回20位；否则返回12位或8位兜底。
    """
    if not text:
        return None

    # 1) 先把非数字全部去掉（处理换行/空格/冒号等）
    digits = re.sub(r"\D", "", text)

    # 2) 优先：如果能拼出20位，直接取前20位
    #    （ROI里通常只有代码+号码，不太会多出其它数字）
    m20 = re.search(r"\d{20}", digits)
    if m20:
        return m20.group(0)

    # 3) 次优：分别找12位代码与8位号码，拼成20位
    m12 = re.search(r"\d{12}", digits)
    m8  = re.search(r"\d{8}", digits)
    if m12 and m8:
        # 如果8位号码在12位代码后面更合理，就优先取“12后面的8”
        tail8 = re.search(r"\d{12}(\d{8})", digits)
        if tail8:
            return m12.group(0) + tail8.group(1)
        return m12.group(0) + m8.group(0)

    # 4) 兜底：只有12位或只有8位就直接返回
    if m12:
        return m12.group(0)
    if m8:
        return m8.group(0)

    # 5) 最后兜底：任意数字串
    m = re.search(r"\d+", digits)
    return m.group(0) if m else None

def extract_date(text: str):
    if not text:
        return None
    m = re.search(r"(\d{4}[年/\-\.]\d{1,2}[月/\-\.]\d{1,2}日?)", text)
    return normalize_date(m.group(1)) if m else None


def extract_amount(text: str):
    if not text:
        return None
    return normalize_amount(text)


def walk_files(root: str, only_pdf: bool = False):
    p = Path(root)
    if p.is_file():
        return [str(p)]

    files = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            ext = Path(fn).suffix.lower()
            if only_pdf:
                if ext == PDF_EXT:
                    files.append(str(Path(dp) / fn))
            else:
                if ext == PDF_EXT or ext in IMAGE_EXTS:
                    files.append(str(Path(dp) / fn))
    return sorted(files)


def process_one_image(img_bgr, cfg, engine, rotate, debug_dir: Path | None, stem: str, page_no: int | None):
    """对一张BGR图（已是某页）按ROI提取"""
    img_bgr = rotate_img(img_bgr, rotate)

    inv_roi = crop_by_norm(img_bgr, cfg["invoice_no"])
    date_roi = crop_by_norm(img_bgr, cfg["invoice_date"])
    amt_roi = crop_by_norm(img_bgr, cfg["total_amount"])

    inv_text = ocr_text(engine, inv_roi)
    date_text = ocr_text(engine, date_roi)
    amt_text = ocr_text(engine, amt_roi)

    if debug_dir:
        debug_dir.mkdir(parents=True, exist_ok=True)
        tag = f"{stem}_p{page_no:02d}" if page_no is not None else stem
        if inv_roi is not None:
            cv2.imencode(".png", inv_roi)[1].tofile(str(debug_dir / f"{tag}_inv.png"))
        if date_roi is not None:
            cv2.imencode(".png", date_roi)[1].tofile(str(debug_dir / f"{tag}_date.png"))
        if amt_roi is not None:
            cv2.imencode(".png", amt_roi)[1].tofile(str(debug_dir / f"{tag}_amt.png"))

    return {
        "发票号码": extract_invoice_no(inv_text),
        "开票日期": extract_date(date_text),
        "价税合计": extract_amount(amt_text),
    }


def main():
    ap = argparse.ArgumentParser(description="按固定ROI离线识别发票号码/开票日期/价税合计并导出Excel（支持PDF多页）")
    ap.add_argument("input_path", help="发票PDF/图片 或 文件夹")
    ap.add_argument("roi_config", help="roi_config.json（由 calibrate_roi.py 生成）")
    ap.add_argument("output_excel", help="输出Excel路径 .xlsx")

    ap.add_argument("--with_filename", action="store_true", help="结果包含文件名列")
    ap.add_argument("--debug_dir", default=None, help="调试：保存ROI裁剪图到该目录（可选）")
    ap.add_argument("--only_pdf", action="store_true", help="只处理PDF，忽略png/jpg等")

    ap.add_argument("--page_index", type=int, default=0, help="单页模式：处理指定页（从0开始），默认0")
    ap.add_argument("--all_pages", action="store_true", help="处理PDF所有页（每页一行）")
    ap.add_argument("--max_pages", type=int, default=0, help="最多处理前N页（0表示不限制）")

    args = ap.parse_args()

    cfg = json.loads(Path(args.roi_config).read_text(encoding="utf-8"))
    dpi = int(cfg.get("dpi", 300))
    rotate = cfg.get("rotate", "0")

    engine = RapidOCR()  # 返回 RapidOCROutput[3](https://pymupdftest.readthedocs.io/en/stable/recipes-images.html)[4](https://sqlpey.com/python/solved-how-to-extract-a-pdf-page-as-a-jpeg/)
    debug_dir = Path(args.debug_dir) if args.debug_dir else None

    rows = []
    files = walk_files(args.input_path, only_pdf=args.only_pdf)
    if not files:
        print("未找到可处理文件。")
        return

    for fp in files:
        fp_path = Path(fp)
        suf = fp_path.suffix.lower()

        try:
            if suf == PDF_EXT:
                doc = fitz.open(str(fp_path))
                total = len(doc)
                if args.all_pages:
                    page_indices = list(range(total))
                    if args.max_pages and args.max_pages > 0:
                        page_indices = page_indices[: args.max_pages]
                else:
                    page_indices = [max(0, min(args.page_index, total - 1))]

                for idx in page_indices:
                    img = render_pdf_page_to_bgr(doc, idx, dpi=dpi)
                    fields = process_one_image(
                        img, cfg, engine, rotate,
                        debug_dir=debug_dir,
                        stem=fp_path.stem,
                        page_no=idx + 1
                    )
                    row = {"页码": idx + 1, **fields}
                    if args.with_filename:
                        row = {"文件名": fp_path.name, **row}
                    rows.append(row)

                doc.close()

            else:
                # 图片文件（不分多页）
                img = imread_unicode(str(fp_path))
                fields = process_one_image(
                    img, cfg, engine, rotate,
                    debug_dir=debug_dir,
                    stem=fp_path.stem,
                    page_no=None
                )
                row = {"页码": 1, **fields}
                if args.with_filename:
                    row = {"文件名": fp_path.name, **row}
                rows.append(row)

            print("[OK]", fp_path.name, "pages processed" if suf == PDF_EXT and args.all_pages else "")

        except Exception as e:
            # 失败也占位
            base = {"页码": None, "发票号码": None, "开票日期": None, "价税合计": None}
            if args.with_filename:
                base = {"文件名": fp_path.name, **base}
            rows.append(base)
            print("[FAIL]", fp_path.name, "->", e)

    df = pd.DataFrame(rows)

    # 列顺序
    cols = ["页码", "发票号码", "开票日期", "价税合计"]
    if args.with_filename:
        cols = ["文件名"] + cols
    df = df[cols]

    out = Path(args.output_excel)
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="发票提取")

    print("完成：", out)


if __name__ == "__main__":
    main()