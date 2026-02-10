# -*- coding: utf-8 -*-
"""
ROI覆盖预览（离线）：
- 读取固定 ROI 配置：C:\\Users\\MY43DN\\Documents\\ocr\\roi_config.json
- 将PDF每页渲染为图片（PyMuPDF）
- 按ROI画框（票号/日期/金额）
- 输出 overlay 图片 + index.html（浏览器快速查看）
- index.html 使用绝对 file:/// 路径，避免相对路径导致图片不显示

用法：
python roi_preview_cli.py "xxx.pdf"
"""

import json
from pathlib import Path
from datetime import datetime
import urllib.parse

import fitz  # PyMuPDF
import numpy as np
import cv2

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


def render_page(doc: fitz.Document, page_index: int, dpi: int):
    page = doc[page_index]
    pix = page.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img


def norm_to_abs(norm_box, W, H):
    x1 = int(norm_box["x1"] * W)
    y1 = int(norm_box["y1"] * H)
    x2 = int(norm_box["x2"] * W)
    y2 = int(norm_box["y2"] * H)
    return x1, y1, x2, y2


def draw_box(img, box, color, label):
    x1, y1, x2, y2 = box
    cv2.rectangle(img, (x1, y1), (x2, y2), color, 3)
    (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.8, 2)
    cv2.rectangle(img, (x1, max(0, y1 - th - 10)), (x1 + tw + 10, y1), color, -1)
    cv2.putText(img, label, (x1 + 5, y1 - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)


def to_file_uri(p: Path) -> str:
    # Windows 路径转 file:///C:/... 并做URL编码（空格等）
    s = p.resolve().as_posix()  # C:/Users/...
    return "file:///" + urllib.parse.quote(s)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="输入PDF")
    ap.add_argument("--max_pages", type=int, default=0, help="最多预览前N页（0=全部）")
    args = ap.parse_args()

    cfg_path = Path(ROI_CONFIG_PATH)
    if not cfg_path.exists():
        raise FileNotFoundError(f"找不到ROI配置：{ROI_CONFIG_PATH}，请先校准ROI。")

    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    dpi = int(cfg.get("dpi", 300))
    rotate = cfg.get("rotate", "0")

    pdf = Path(args.pdf)
    doc = fitz.open(str(pdf))
    total_pages = len(doc)
    n = total_pages
    if args.max_pages and args.max_pages > 0:
        n = min(n, args.max_pages)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = pdf.parent / f"roi_preview_{pdf.stem}_{ts}"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_imgs = []
    for i in range(n):
        img = render_page(doc, i, dpi=dpi)
        img = rotate_img(img, rotate)
        H, W = img.shape[:2]

        inv_box = norm_to_abs(cfg["invoice_no"], W, H)
        date_box = norm_to_abs(cfg["invoice_date"], W, H)
        amt_box = norm_to_abs(cfg["total_amount"], W, H)

        overlay = img.copy()
        draw_box(overlay, inv_box, (0, 128, 255), "票号(20位)")
        draw_box(overlay, date_box, (0, 200, 0), "开票日期")
        draw_box(overlay, amt_box, (200, 0, 200), "价税合计")

        out_img = out_dir / f"page_{i+1:03d}_overlay.png"
        cv2.imencode(".png", overlay)[1].tofile(str(out_img))
        out_imgs.append(out_img)

    doc.close()

    # 生成HTML（用绝对 file:/// 路径）
    html = []
    html.append("<html><head><meta charset='utf-8'>")
    html.append("<title>ROI预览</title>")
    html.append("<style>")
    html.append("body{font-family:Arial,'Microsoft YaHei';}")
    html.append(".page{margin:18px 0; padding:12px; border:1px solid #ddd;}")
    html.append(".img{max-width:98%; border:1px solid #ccc;}")
    html.append("</style>")
    html.append("</head><body>")
    html.append(f"<h2>ROI预览：{pdf.name}</h2>")
    html.append(f"<p>rotate={rotate} dpi={dpi} pages={n}/{total_pages}</p>")
    html.append("<hr>")

    for idx, img_path in enumerate(out_imgs, start=1):
        img_uri = to_file_uri(img_path)
        html.append("<div class='page'>")
        html.append(f"<h3>第 {idx} 页</h3>")
        html.append(f"<div>{img_uri}</div>")
        html.append("</div>")

    html.append("</body></html>")

    index_path = out_dir / "index.html"
    index_path.write_text("\n".join(html), encoding="utf-8")

    print(str(index_path))  # 给UI读取并打开


if __name__ == "__main__":
    main()