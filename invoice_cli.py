# -*- coding: utf-8 -*-
import argparse
from pathlib import Path
import sys
import invoice_core as core

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="输入PDF路径")
    ap.add_argument("--out", default=None, help="输出Excel路径（可选）")
    ap.add_argument("--debug_dir", default=None, help="保存ROI调试截图目录（可选）")
    args = ap.parse_args()

    pdf = Path(args.pdf)
    if not pdf.exists():
        raise FileNotFoundError(str(pdf))

    def hook(cur, total):
        # 给UI解析用：PROGRESS cur total
        print(f"PROGRESS {cur} {total}", flush=True)

    rows = core.extract_pdf_to_rows(str(pdf), debug_dir=args.debug_dir, progress_hook=hook)

    out_xlsx = args.out if args.out else str(pdf.parent / f"{pdf.stem}_extract.xlsx")
    out = core.export_rows_to_excel(rows, out_xlsx)

    # 给UI解析用：RESULT path
    print(f"RESULT {out}", flush=True)

if __name__ == "__main__":
    main()