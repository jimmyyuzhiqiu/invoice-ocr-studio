# -*- coding: utf-8 -*-
"""
ROI校准工具：
- 支持指定PDF页码校准：--page_index (0-based)
- 支持强制旋转：--rotate cw90/ccw90/180/0
- 显示自适应缩放：窗口显示缩小图，但保存坐标映射回原图（相对坐标）
输出：roi_config.json（含 rotate、dpi、3个ROI相对坐标）
"""

import json
import cv2
import fitz  # PyMuPDF
import numpy as np
from pathlib import Path
import argparse


def render_page(pdf_path: str, dpi: int = 300, page_index: int = 0):
    doc = fitz.open(pdf_path)
    if page_index < 0 or page_index >= len(doc):
        page_index = 0
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)  # PyMuPDF渲染[4](https://blog.csdn.net/qq_41866626/article/details/116710899)[5](https://www.php.cn/faq/2067550.html)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img


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


def fit_to_window(img, max_w=1400, max_h=900):
    h, w = img.shape[:2]
    scale = min(max_w / w, max_h / h, 1.0)  # 只缩小不放大
    if scale < 1.0:
        disp = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    else:
        disp = img.copy()
    return disp, scale


def select_roi_scaled(win_name, disp_img):
    r = cv2.selectROI(win_name, disp_img, fromCenter=False, showCrosshair=True)
    x, y, w, h = map(int, r)
    return x, y, w, h


def box_disp_to_orig(box_disp, scale):
    x, y, w, h = box_disp
    if scale <= 0:
        return x, y, w, h
    xo = int(round(x / scale))
    yo = int(round(y / scale))
    wo = int(round(w / scale))
    ho = int(round(h / scale))
    return xo, yo, wo, ho


def to_norm(box_orig, W, H):
    x, y, w, h = box_orig
    return {
        "x1": x / W, "y1": y / H,
        "x2": (x + w) / W, "y2": (y + h) / H
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="用于校准的PDF文件路径")
    ap.add_argument("--dpi", type=int, default=300, help="渲染DPI")
    ap.add_argument("--page_index", type=int, default=0, help="用于校准的PDF页码（从0开始）")
    ap.add_argument("--out", default="roi_config.json", help="输出roi_config.json路径")
    ap.add_argument("--rotate", default="0", help="旋转：0/cw90/ccw90/180（强制）")
    ap.add_argument("--max_w", type=int, default=1400)
    ap.add_argument("--max_h", type=int, default=900)
    args = ap.parse_args()

    img = render_page(args.pdf, dpi=args.dpi, page_index=args.page_index)
    img = rotate_img(img, args.rotate)

    H, W = img.shape[:2]
    disp, scale = fit_to_window(img, max_w=args.max_w, max_h=args.max_h)

    cv2.namedWindow("invoice", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("invoice", min(args.max_w, disp.shape[1]), min(args.max_h, disp.shape[0]))

    print("请依次框选：票号(20位纯数字) -> 开票日期 -> 价税合计")
    print("提示：框选后按 SPACE/ENTER 确认；按 c 取消本次框选。")

    inv_disp = select_roi_scaled("invoice - 1) 票号(20位)", disp)
    date_disp = select_roi_scaled("invoice - 2) 开票日期", disp)
    amt_disp = select_roi_scaled("invoice - 3) 价税合计", disp)

    inv_box = box_disp_to_orig(inv_disp, scale)
    date_box = box_disp_to_orig(date_disp, scale)
    amt_box = box_disp_to_orig(amt_disp, scale)

    cfg = {
        "dpi": args.dpi,
        "rotate": args.rotate,
        "page_index_for_calibration": args.page_index,
        "invoice_no": to_norm(inv_box, W, H),
        "invoice_date": to_norm(date_box, W, H),
        "total_amount": to_norm(amt_box, W, H)
    }

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保存：{out}")

    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()