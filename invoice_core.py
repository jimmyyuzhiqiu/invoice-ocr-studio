# -*- coding: utf-8 -*-
"""
核心提取逻辑（离线）：
- PDF逐页渲染 -> 旋转 -> ROI裁剪 -> OCR -> 导出Excel
- 固定ROI配置路径：C:\\Users\\MY43DN\\Documents\\ocr\\roi_config.json
- 票号：严格只提取 20 位纯数字（不拼接、不退化）
- 日期：YYYYMMDD
- 新增：progress_hook(current_page, total_pages) 回调，用于UI进度条
"""

import re
import json
from pathlib import Path
import os

import fitz  # PyMuPDF
import numpy as np
import cv2
import pandas as pd
from rapidocr import RapidOCR

ROI_CONFIG_PATH = r"C:\Users\MY43DN\Documents\ocr\roi_config.json"


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
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
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


def upscale_if_small(img_bgr, min_h=70):
    if img_bgr is None or img_bgr.size == 0:
        return img_bgr
    h, w = img_bgr.shape[:2]
    if h < min_h:
        scale = max(2.0, min_h / max(h, 1))
        img_bgr = cv2.resize(img_bgr, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
    return img_bgr


def light_preprocess(img_bgr):
    if img_bgr is None or img_bgr.size == 0:
        return img_bgr
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)


def get_txt_from_rapid_output(out):
    if out is None:
        return ""
    if hasattr(out, "txts") and out.txts:
        return "".join([t for t in out.txts if isinstance(t, str)]).strip()
    if isinstance(out, str):
        return out.strip()
    return ""


def extract_no20_only(text: str) -> str | None:
    """严格只提取20位纯数字票号"""
    if not text:
        return None
    digits = re.sub(r"\D", "", text)
    m = re.search(r"\d{20}", digits)
    return m.group(0) if m else None


def ocr_inv_text(engine: RapidOCR, img_bgr) -> str:
    """票号区域：强制 use_det=True 更稳"""
    if img_bgr is None or img_bgr.size == 0:
        return ""
    img1 = upscale_if_small(img_bgr, min_h=70)
    try:
        out = engine(img1, use_det=True, use_cls=False, use_rec=True, box_thresh=0.3, text_score=0.3)
        return get_txt_from_rapid_output(out)
    except Exception:
        return ""


def ocr_text_simple(engine: RapidOCR, img_bgr) -> str:
    """日期/金额：优先不检测直接识别，失败再det重试"""
    if img_bgr is None or img_bgr.size == 0:
        return ""
    img1 = upscale_if_small(img_bgr, min_h=60)
    img1 = light_preprocess(img1)
    try:
        out = engine(img1, use_det=False, use_cls=False, use_rec=True)
        t = get_txt_from_rapid_output(out)
        if t:
            return t
    except Exception:
        pass
    try:
        out = engine(img1, use_det=True, use_cls=False, use_rec=True, box_thresh=0.3, text_score=0.3)
        return get_txt_from_rapid_output(out)
    except Exception:
        return ""


def normalize_date_to_yyyymmdd(s: str):
    if not s:
        return None
    s = s.strip().replace(" ", "").replace("年", "").replace("月", "").replace("日", "")
    s = re.sub(r"[./\-]", "", s)
    m = re.search(r"(\d{8})", s)
    return m.group(1) if m else None


def normalize_amount(s: str):
    if not s:
        return None
    s = s.strip().replace(",", "").replace("￥", "").replace("¥", "").replace("RMB", "").replace(" ", "")
    m = re.search(r"(\d+\.\d{2})", s)
    if m:
        return m.group(1)
    m = re.search(r"(\d+(?:\.\d{1,2})?)", s)
    return m.group(1) if m else None


def load_roi_config(path=ROI_CONFIG_PATH):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"ROI配置不存在：{path}")
    return json.loads(p.read_text(encoding="utf-8"))


def extract_pdf_to_rows(pdf_path: str, debug_dir: str | None = None, progress_hook=None):
    """
    progress_hook: callable(current_page:int, total_pages:int)
    """
    cfg = load_roi_config()
    dpi = int(cfg.get("dpi", 300))
    rotate = cfg.get("rotate", "0")

    engine = RapidOCR()
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    rows = []

    dbg = Path(debug_dir) if debug_dir else None
    if dbg:
        dbg.mkdir(parents=True, exist_ok=True)

    for i in range(total_pages):
        img = render_pdf_page_to_bgr(doc, i, dpi=dpi)
        img = rotate_img(img, rotate)

        inv_roi = crop_by_norm(img, cfg["invoice_no"])
        date_roi = crop_by_norm(img, cfg["invoice_date"])
        amt_roi = crop_by_norm(img, cfg["total_amount"])

        inv_text = ocr_inv_text(engine, inv_roi)
        ticket20 = extract_no20_only(inv_text)
        if ticket20 is None and inv_roi is not None:
            inv_text2 = ocr_inv_text(engine, light_preprocess(inv_roi))
            ticket20 = extract_no20_only(inv_text2)

        date_text = ocr_text_simple(engine, date_roi)
        amt_text = ocr_text_simple(engine, amt_roi)

        rows.append({
            "文件名": Path(pdf_path).name,
            "页码": i + 1,
            "票号20位": ticket20,
            "票号完整": "Y" if (ticket20 and len(ticket20) == 20) else "N",
            "开票日期": normalize_date_to_yyyymmdd(date_text),
            "价税合计": normalize_amount(amt_text)
        })

        # debug保存ROI图
        if dbg:
            stem = Path(pdf_path).stem
            tag = f"{stem}_p{i+1:02d}"
            if inv_roi is not None:
                cv2.imencode(".png", inv_roi)[1].tofile(str(dbg / f"{tag}_inv.png"))
            if date_roi is not None:
                cv2.imencode(".png", date_roi)[1].tofile(str(dbg / f"{tag}_date.png"))
            if amt_roi is not None:
                cv2.imencode(".png", amt_roi)[1].tofile(str(dbg / f"{tag}_amt.png"))

        # 进度回调
        if progress_hook:
            try:
                progress_hook(i + 1, total_pages)
            except Exception:
                pass

    doc.close()
    return rows


def export_rows_to_excel(rows, excel_path: str):
    df = pd.DataFrame(rows)
    cols = ["文件名", "页码", "票号20位", "开票日期", "价税合计", "票号完整"]
    df = df[cols]
    out = Path(excel_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="发票提取")
    return str(out)


def open_file_windows(path: str):
    try:
        os.startfile(path)  # noqa
    except Exception:
        pass