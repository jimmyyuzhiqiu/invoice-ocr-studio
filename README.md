

# Invoice OCR Studio

Invoice OCR Studio 是一个**离线 PDF 发票 OCR 工具**：把扫描版/图片型 PDF 发票拖进来，程序会**逐页识别**并把关键字段导出到 **Excel**，导出后自动打开 Excel。

---

## Screenshots

### Main UI
<p align="center">
  docs/images/ui.png
</p>

### ROI Overlay Preview (UI 内翻页预览)
<p align="center">
  docs/images/ui-overlay.png
</p>

---

## 它能做什么
- **拖拽 PDF 一键识别**（支持多页 PDF）
- **导出 Excel**（逐页一行）
- 支持 **ROI 校准**：指定用第几页做校准 + 选择旋转方向（cw90/ccw90/180/0）
- 支持 **ROI 覆盖预览**：在 UI 内翻页查看每页框选位置是否正确
- 支持 **真实进度条**：按页显示进度（cur/total），并可取消任务

> 输出字段：票号（20位纯数字）、开票日期（YYYYMMDD）、价税合计、页码、文件名、票号完整(Y/N)。

---

## 使用方法（推荐流程）

### 1) 安装依赖（Windows）
建议使用 venv：
```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -U pip
pip install PyQt5 PyMuPDF opencv-python numpy pandas openpyxl rapidocr onnxruntime
````

### 2) 启动

```bash
python invoice_ui.py
```

### 3) 首次使用：先校准 ROI

1.  点击 **重新校准ROI**
2.  选择一份代表性 PDF
3.  输入 **校准页码（从 1 开始）**
4.  选择旋转方向（内容横着一般选 **cw90**）
5.  按提示依次框选：
    *   票号（20位纯数字）
    *   开票日期
    *   价税合计

### 4) 预览 ROI 覆盖（强烈建议）

点击 **预览ROI覆盖**，在弹出的预览窗口里翻页确认每一页框选位置是否正确。  
如果有偏移，用偏移明显的页重新校准一次即可。

### 5) 识别导出

*   直接把 PDF **拖拽到界面**，或点击拖拽区域选择 PDF
*   等待进度条跑完
*   自动导出 Excel 并打开

***


***

<p align="center">
  <sub><b>Designed by 余智秋 in Shanghai.</b></sub>
</p>
