# -*- coding: utf-8 -*-
import sys
import json
import subprocess
from pathlib import Path
from datetime import datetime

import numpy as np
import cv2
import fitz  # PyMuPDF

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QPixmap, QImage, QIcon, QPainter, QPainterPath, QColor
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QMessageBox, QCheckBox, QTextEdit, QHBoxLayout, QInputDialog, QSpinBox,
    QDialog, QScrollArea, QFrame, QGraphicsDropShadowEffect, QProgressBar
)

ROI_CONFIG_PATH = r"C:\Users\MY43DN\Documents\ocr\roi_config.json"
APP_ICON_PATH = r"C:\Users\MY43DN\Documents\ocr\app.ico"
LOGO_PATH = r"C:\Users\MY43DN\Documents\ocr\ing-logo.png"


# ------------------ UI视觉工具 ------------------
def make_shadow(widget, blur=24, dx=0, dy=10, color=QColor(0, 0, 0, 90)):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(dx, dy)
    shadow.setColor(color)
    widget.setGraphicsEffect(shadow)


def rounded_square_pixmap(pix: QPixmap, size: int = 48, radius: int = 12) -> QPixmap:
    """把logo变成“圆角正方形”"""
    if pix.isNull():
        return pix
    s = min(pix.width(), pix.height())
    x = (pix.width() - s) // 2
    y = (pix.height() - s) // 2
    cropped = pix.copy(x, y, s, s).scaled(size, size, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)

    out = QPixmap(size, size)
    out.fill(Qt.transparent)
    painter = QPainter(out)
    painter.setRenderHint(QPainter.Antialiasing, True)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, radius, radius)
    painter.setClipPath(path)
    painter.drawPixmap(0, 0, cropped)
    painter.end()
    return out


# ------------------ 预览相关工具 ------------------
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


def norm_to_abs(norm_box, W, H):
    x1 = int(norm_box["x1"] * W)
    y1 = int(norm_box["y1"] * H)
    x2 = int(norm_box["x2"] * W)
    y2 = int(norm_box["y2"] * H)
    return x1, y1, x2, y2


def draw_box(img, box, color, label):
    x1, y1, x2, y2 = box
    cv2.rectangle(img, (x1, y1), (x2, y2), color, 3)
    (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.85, 2)
    cv2.rectangle(img, (x1, max(0, y1 - th - 12)), (x1 + tw + 12, y1), color, -1)
    cv2.putText(img, label, (x1 + 6, y1 - 7),
                cv2.FONT_HERSHEY_SIMPLEX, 0.85, (255, 255, 255), 2)


def bgr_to_qpixmap(img_bgr):
    rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
    h, w, ch = rgb.shape
    bytes_per_line = ch * w
    qimg = QImage(rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
    return QPixmap.fromImage(qimg)


# ------------------ 子进程 Worker：OCR（实时进度） ------------------
class RunOcrWorker(QThread):
    progress_text = pyqtSignal(str)
    progress_value = pyqtSignal(int, int)  # cur, total
    finished = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, pdf_path: str, debug_dir: str | None):
        super().__init__()
        self.pdf_path = pdf_path
        self.debug_dir = debug_dir
        self._proc = None

    def cancel(self):
        try:
            if self._proc and self._proc.poll() is None:
                self._proc.terminate()
        except Exception:
            pass

    def run(self):
        try:
            self.progress_text.emit(f"开始处理：{self.pdf_path}")
            pdf = Path(self.pdf_path)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_xlsx = str(pdf.parent / f"invoice_extract_{ts}.xlsx")

            cli = Path(__file__).parent / "invoice_cli.py"
            if not cli.exists():
                self.failed.emit(f"缺少文件：{cli}")
                return

            cmd = [sys.executable, str(cli), self.pdf_path, "--out", out_xlsx]
            if self.debug_dir:
                cmd += ["--debug_dir", self.debug_dir]

            self._proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1
            )

            last_result = None
            while True:
                line = self._proc.stdout.readline() if self._proc.stdout else ""
                if not line:
                    if self._proc.poll() is not None:
                        break
                    continue
                line = line.strip()

                if line.startswith("PROGRESS "):
                    parts = line.split()
                    if len(parts) >= 3:
                        try:
                            cur = int(parts[1])
                            total = int(parts[2])
                            self.progress_value.emit(cur, total)
                            self.progress_text.emit(f"处理中：{cur}/{total}")
                        except Exception:
                            self.progress_text.emit(line)
                    else:
                        self.progress_text.emit(line)
                    continue

                if line.startswith("RESULT "):
                    last_result = line.replace("RESULT ", "", 1).strip()
                    self.progress_text.emit(f"导出完成：{last_result}")
                    continue

                self.progress_text.emit(line)

            rc = self._proc.poll()
            if rc != 0:
                self.failed.emit(f"子进程退出码={rc}（可能被取消或出错）")
                return

            if not last_result:
                last_result = out_xlsx

            self.finished.emit(last_result)

        except Exception as e:
            self.failed.emit(str(e))


# ------------------ 预览窗口 ------------------
class PreviewDialog(QDialog):
    def __init__(self, overlay_paths: list[Path], parent=None):
        super().__init__(parent)
        self.setWindowTitle("ROI 覆盖预览")
        self.resize(1100, 820)

        self.overlay_paths = overlay_paths
        self.total = len(overlay_paths)
        self.idx = 0

        self.setStyleSheet("""
            QDialog { background: #101418; }
            QLabel { color: #EAECEF; }
            QPushButton {
                background: #1B2430; color: #EAECEF;
                border: 1px solid #2C3A4A;
                border-radius: 10px; padding: 8px 14px;
            }
            QPushButton:hover { background: #223041; }
            QPushButton:pressed { background: #15202B; }
            QSpinBox {
                background: #121A22; color: #EAECEF;
                border: 1px solid #2C3A4A; border-radius: 10px;
                padding: 6px 10px;
            }
        """)

        self.lbl_info = QLabel("")
        self.lbl_info.setStyleSheet("font-size: 14px; color: #C9D1D9;")

        self.btn_prev = QPushButton("上一页")
        self.btn_next = QPushButton("下一页")
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next.clicked.connect(self.next_page)

        self.spin_page = QSpinBox()
        self.spin_page.setRange(1, max(1, self.total))
        self.spin_page.valueChanged.connect(self.jump_page)

        top = QHBoxLayout()
        top.addWidget(self.btn_prev)
        top.addWidget(self.btn_next)
        top.addWidget(QLabel("跳转："))
        top.addWidget(self.spin_page)
        top.addStretch(1)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        scroll.setWidget(self.image_label)

        layout = QVBoxLayout()
        layout.addWidget(self.lbl_info)
        layout.addLayout(top)
        layout.addWidget(scroll)
        self.setLayout(layout)

        self.refresh()

    def refresh(self):
        if self.total <= 0:
            self.lbl_info.setText("没有预览图片。")
            return
        p = self.overlay_paths[self.idx]
        self.lbl_info.setText(f"第 {self.idx+1}/{self.total} 页：{p.name}")
        self.spin_page.blockSignals(True)
        self.spin_page.setValue(self.idx + 1)
        self.spin_page.blockSignals(False)

        data = np.fromfile(str(p), dtype=np.uint8)
        img = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if img is None:
            self.image_label.setText("图片读取失败")
            return

        pix = bgr_to_qpixmap(img)
        w = max(860, self.width() - 80)
        self.image_label.setPixmap(pix.scaledToWidth(w, Qt.SmoothTransformation))

    def prev_page(self):
        if self.idx > 0:
            self.idx -= 1
            self.refresh()

    def next_page(self):
        if self.idx < self.total - 1:
            self.idx += 1
            self.refresh()

    def jump_page(self, v: int):
        self.idx = max(0, min(self.total - 1, v - 1))
        self.refresh()


# ------------------ 拖拽卡片（支持点击选择PDF） ------------------
class DropCard(QFrame):
    file_selected = pyqtSignal(str)   # 点击选择后触发
    file_dropped = pyqtSignal(str)    # 拖拽触发

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setObjectName("DropCard")
        self.setMinimumHeight(170)

        title = QLabel("拖拽 PDF 到这里")
        title.setObjectName("DropTitle")

        clickable = QLabel("<u>点击这里选择PDF</u>")
        clickable.setObjectName("DropSub")
        clickable.setTextInteractionFlags(Qt.NoTextInteraction)
        clickable.setCursor(Qt.PointingHandCursor)
        self._clickable = clickable

        hint = QLabel("支持多页PDF：将逐页导出到 Excel（带进度条）")
        hint.setObjectName("DropHint")

        box = QVBoxLayout()
        box.addStretch(1)
        box.addWidget(title, alignment=Qt.AlignCenter)
        box.addWidget(clickable, alignment=Qt.AlignCenter)
        box.addSpacing(6)
        box.addWidget(hint, alignment=Qt.AlignCenter)
        box.addStretch(1)
        self.setLayout(box)

        make_shadow(self, blur=28, dy=12)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # 点击卡片任意位置都可选择PDF
            self._open_file_dialog()
        super().mousePressEvent(event)

    def _open_file_dialog(self):
        pdf, _ = QFileDialog.getOpenFileName(self, "选择PDF", "", "PDF Files (*.pdf)")
        if pdf:
            self.file_selected.emit(pdf)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        if path.lower().endswith(".pdf"):
            self.file_dropped.emit(path)
        else:
            QMessageBox.warning(self, "提示", "只支持拖入 PDF 文件。")


# ------------------ 主窗口（高级UI + 进度条） ------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        if Path(APP_ICON_PATH).exists():
            self.setWindowIcon(QIcon(APP_ICON_PATH))

        self.setWindowTitle("Invoice OCR Studio")
        self.resize(1020, 760)
        self.setObjectName("Root")

        self.setStyleSheet("""
            #Root {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #0E1116, stop:1 #0B0F14);
                font-family: "Segoe UI", "Microsoft YaHei";
            }
            #HeaderCard {
                background: rgba(18, 24, 31, 0.92);
                border: 1px solid rgba(44, 58, 74, 0.85);
                border-radius: 18px;
            }
            QLabel { color: #EAECEF; }
            #Title { font-size: 22px; font-weight: 700; }
            #SubTitle { color: #9FB0C3; font-size: 12px; }

            QPushButton {
                background: #1B2430; color: #EAECEF;
                border: 1px solid #2C3A4A;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 600;
            }
            QPushButton:hover { background: #223041; }
            QPushButton:pressed { background: #15202B; }

            QCheckBox { color: #C9D1D9; }
            QSpinBox {
                background: #121A22;
                color: #EAECEF;
                border: 1px solid #2C3A4A;
                border-radius: 10px;
                padding: 6px 10px;
            }

            QTextEdit {
                background: rgba(15, 20, 26, 0.95);
                color: #D6DEE7;
                border: 1px solid rgba(44, 58, 74, 0.9);
                border-radius: 14px;
                padding: 10px;
            }

            QProgressBar {
                border: 1px solid rgba(44, 58, 74, 0.9);
                border-radius: 10px;
                background: rgba(15, 20, 26, 0.95);
                text-align: center;
                color: #EAECEF;
                height: 18px;
            }
            QProgressBar::chunk {
                border-radius: 10px;
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #D4AF37, stop:1 #8A6F1E);
            }

            #DropCard {
                background: rgba(18, 24, 31, 0.92);
                border: 1px dashed rgba(90, 120, 160, 0.9);
                border-radius: 18px;
            }
            #DropTitle { color: #EAECEF; font-size: 18px; font-weight: 700; }
            #DropSub { color: #A7B6C8; font-size: 12px; }
            #DropHint { color: #7FA0C0; font-size: 11px; }

            #Footer {
                color: #D4AF37;
                font-size: 10px;
                letter-spacing: 0.3px;
            }
        """)

        # Header
        header = QFrame()
        header.setObjectName("HeaderCard")
        make_shadow(header, blur=28, dy=12)

        logo_lbl = QLabel()
        logo_pix = QPixmap(LOGO_PATH) if Path(LOGO_PATH).exists() else QPixmap()
        logo_lbl.setPixmap(rounded_square_pixmap(logo_pix, size=52, radius=14))
        logo_lbl.setFixedSize(QSize(56, 56))

        title = QLabel("Invoice OCR Studio")
        title.setObjectName("Title")
        subtitle = QLabel("Offline OCR • Page Progress • ROI Calibrate & Preview")
        subtitle.setObjectName("SubTitle")

        title_box = QVBoxLayout()
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        self.btn_select = QPushButton("选择PDF并识别")
        self.btn_select.clicked.connect(self.select_pdf_and_run)

        self.btn_calibrate = QPushButton("重新校准ROI")
        self.btn_calibrate.clicked.connect(self.recalibrate_roi)

        self.btn_preview = QPushButton("预览ROI覆盖")
        self.btn_preview.clicked.connect(self.preview_roi_in_ui)

        self.btn_open_folder = QPushButton("打开当前目录")
        self.btn_open_folder.clicked.connect(self.open_current_folder)

        self.btn_cancel = QPushButton("取消任务")
        self.btn_cancel.clicked.connect(self.cancel_task)
        self.btn_cancel.setEnabled(False)

        self.chk_debug = QCheckBox("保存debug_roi截图")
        self.chk_debug.setChecked(True)

        self.lbl_max_pages = QLabel("预览最多页数(0=全部)：")
        self.lbl_max_pages.setStyleSheet("color:#9FB0C3; font-size:12px;")
        self.spin_max_pages = QSpinBox()
        self.spin_max_pages.setRange(0, 9999)
        self.spin_max_pages.setValue(0)

        self.status = QLabel("就绪")
        self.status.setStyleSheet("color:#9FB0C3; font-size:12px;")
        self.pbar = QProgressBar()
        self.pbar.setRange(0, 100)
        self.pbar.setValue(0)

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.btn_select)
        btn_row.addWidget(self.btn_calibrate)
        btn_row.addWidget(self.btn_preview)
        btn_row.addWidget(self.btn_open_folder)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addStretch(1)
        btn_row.addWidget(self.lbl_max_pages)
        btn_row.addWidget(self.spin_max_pages)
        btn_row.addWidget(self.chk_debug)

        header_layout = QVBoxLayout()
        header_top = QHBoxLayout()
        header_top.addWidget(logo_lbl)
        header_top.addSpacing(10)
        header_top.addLayout(title_box)
        header_top.addStretch(1)
        header_layout.addLayout(header_top)
        header_layout.addSpacing(10)
        header_layout.addLayout(btn_row)
        header_layout.addWidget(self.status)
        header_layout.addWidget(self.pbar)
        header.setLayout(header_layout)

        # Drop Card (clickable)
        self.drop_card = DropCard()
        self.drop_card.file_dropped.connect(self.run_ocr)
        self.drop_card.file_selected.connect(self.run_ocr)

        # Log
        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # Footer
        footer = QLabel("Designed by 余智秋 in Shanghai.")
        footer.setObjectName("Footer")
        footer.setAlignment(Qt.AlignCenter)

        root = QVBoxLayout()
        root.setContentsMargins(18, 18, 18, 14)
        root.setSpacing(14)
        root.addWidget(header)
        root.addWidget(self.drop_card)
        root.addWidget(QLabel("运行日志：", styleSheet="color:#9FB0C3; font-size:12px;"))
        root.addWidget(self.log, stretch=1)
        root.addWidget(footer)
        self.setLayout(root)

        self.ocr_worker = None

        self.append_log(f"Python解释器：{sys.executable}")
        self.append_log(f"ROI配置固定路径：{ROI_CONFIG_PATH}")
        if not Path(ROI_CONFIG_PATH).exists():
            self.append_log("⚠️ 未发现 ROI 配置，请先点击“重新校准ROI”。")

    def append_log(self, msg: str):
        self.log.append(msg)

    def set_controls_enabled(self, enabled: bool):
        self.btn_select.setEnabled(enabled)
        self.btn_calibrate.setEnabled(enabled)
        self.btn_preview.setEnabled(enabled)
        self.btn_open_folder.setEnabled(enabled)
        self.chk_debug.setEnabled(enabled)
        self.spin_max_pages.setEnabled(enabled)
        self.btn_cancel.setEnabled(not enabled)

    def cancel_task(self):
        if self.ocr_worker:
            self.append_log("已请求取消任务...")
            self.status.setText("正在取消...")
            self.ocr_worker.cancel()

    def open_current_folder(self):
        try:
            import os
            os.startfile(str(Path.cwd()))
        except Exception:
            pass

    def check_roi_exists_or_warn(self) -> bool:
        if not Path(ROI_CONFIG_PATH).exists():
            QMessageBox.warning(self, "ROI配置缺失", f"找不到：\n{ROI_CONFIG_PATH}\n请先校准ROI。")
            return False
        return True

    # ---------- OCR ----------
    def select_pdf_and_run(self):
        pdf, _ = QFileDialog.getOpenFileName(self, "选择PDF", "", "PDF Files (*.pdf)")
        if pdf:
            self.run_ocr(pdf)

    def run_ocr(self, pdf_path: str):
        if not self.check_roi_exists_or_warn():
            return

        debug_dir = None
        if self.chk_debug.isChecked():
            debug_dir = str(Path(pdf_path).parent / "debug_roi")

        self.pbar.setValue(0)
        self.pbar.setRange(0, 100)
        self.status.setText("启动中...")

        self.set_controls_enabled(False)

        self.ocr_worker = RunOcrWorker(pdf_path, debug_dir)
        self.ocr_worker.progress_text.connect(self.on_progress_text)
        self.ocr_worker.progress_value.connect(self.on_progress_value)
        self.ocr_worker.finished.connect(self.on_ocr_finished)
        self.ocr_worker.failed.connect(self.on_ocr_failed)
        self.ocr_worker.start()

    def on_progress_text(self, text: str):
        self.append_log(text)
        if text.startswith("处理中："):
            self.status.setText(text)

    def on_progress_value(self, cur: int, total: int):
        self.pbar.setRange(0, total)
        self.pbar.setValue(cur)

    def on_ocr_finished(self, excel_path: str):
        self.append_log("完成。将自动打开Excel。")
        self.status.setText("完成")
        try:
            import os
            os.startfile(excel_path)
        except Exception:
            pass
        self.set_controls_enabled(True)

    def on_ocr_failed(self, err: str):
        self.append_log(f"[失败] {err}")
        self.status.setText("失败/已取消")
        reply = QMessageBox.question(
            self, "识别失败",
            f"识别失败：\n{err}\n\n是否现在重新校准ROI？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.recalibrate_roi()
        self.set_controls_enabled(True)

    # ---------- 校准 ROI ----------
    def recalibrate_roi(self):
        pdf, _ = QFileDialog.getOpenFileName(self, "选择一份代表性PDF用于校准ROI", "", "PDF Files (*.pdf)")
        if not pdf:
            return

        page_str, ok = QInputDialog.getText(self, "校准页码", "请输入用于校准的页码（从1开始）：", text="1")
        if not ok:
            return
        try:
            page_index = max(0, int(page_str) - 1)
        except Exception:
            page_index = 0

        rotate, ok = QInputDialog.getItem(
            self, "旋转方向", "选择旋转方向（内容横着一般选 cw90）：",
            ["0", "cw90", "ccw90", "180"], 1, False
        )
        if not ok:
            return

        script_path = Path(__file__).parent / "calibrate_roi.py"
        if not script_path.exists():
            QMessageBox.critical(self, "缺少文件", f"找不到：{script_path}")
            return

        self.append_log(f"开始校准ROI：page_index={page_index} rotate={rotate}")
        self.append_log("请依次框选：票号(20位纯数字) / 日期 / 价税合计")

        try:
            cmd = [
                sys.executable, str(script_path), pdf,
                "--page_index", str(page_index),
                "--rotate", rotate,
                "--out", ROI_CONFIG_PATH
            ]
            subprocess.run(cmd, check=True)
            self.append_log(f"ROI校准完成：{ROI_CONFIG_PATH}")
            QMessageBox.information(self, "完成", "ROI校准完成！现在可以预览ROI覆盖或拖入PDF识别。")
        except Exception as e:
            self.append_log(f"[校准失败] {e}")
            QMessageBox.critical(self, "校准失败", f"校准执行失败：\n{e}")

    # ---------- ROI 预览（可用） ----------
    def preview_roi_in_ui(self):
        if not self.check_roi_exists_or_warn():
            return

        pdf_path, _ = QFileDialog.getOpenFileName(self, "选择PDF进行ROI覆盖预览", "", "PDF Files (*.pdf)")
        if not pdf_path:
            return

        try:
            cfg = json.loads(Path(ROI_CONFIG_PATH).read_text(encoding="utf-8"))
            dpi = int(cfg.get("dpi", 300))
            rotate = cfg.get("rotate", "0")
        except Exception as e:
            QMessageBox.critical(self, "ROI配置错误", f"无法读取ROI配置：\n{e}")
            return

        max_pages = int(self.spin_max_pages.value())

        try:
            pdf = Path(pdf_path)
            doc = fitz.open(str(pdf))
            total = len(doc)
            n = total if max_pages == 0 else min(total, max_pages)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = pdf.parent / f"roi_preview_{pdf.stem}_{ts}"
            out_dir.mkdir(parents=True, exist_ok=True)

            self.append_log(f"ROI预览：{pdf.name} rotate={rotate} dpi={dpi} pages={n}/{total}")

            overlay_paths = []
            for i in range(n):
                img = render_pdf_page_to_bgr(doc, i, dpi=dpi)
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
                overlay_paths.append(out_img)

            doc.close()

            if not overlay_paths:
                QMessageBox.information(self, "预览", "没有生成任何预览图片。")
                return

            dlg = PreviewDialog(overlay_paths, parent=self)
            dlg.exec_()

        except Exception as e:
            self.append_log(f"[预览失败] {e}")
            QMessageBox.critical(self, "预览失败", str(e))


def main():
    app = QApplication(sys.argv)
    if Path(APP_ICON_PATH).exists():
        app.setWindowIcon(QIcon(APP_ICON_PATH))
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()