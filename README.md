
# Invoice OCR Studio

**Invoice OCR Studio** 是一个面向扫描版/图片型 PDF 发票的**离线 OCR 桌面工具**：  
把 PDF 拖进来 → 程序逐页识别 → 自动导出 Excel 并打开。

---

## Screenshots


### Main UI
!Invoice OCR Studio Main UI

### ROI Overlay Preview (UI 内翻页预览)
!ROI Overlay Preview UI

---

## 软件能做什么

- **拖拽 PDF 一键识别**（支持多页 PDF）
- **导出 Excel**（逐页一行）
- **真实进度条**：按页更新（cur/total），支持取消任务
- **ROI 校准**：可指定校准页码 + 强制旋转（cw90/ccw90/180/0）
- **ROI 覆盖预览**：UI 内翻页查看每页 ROI 框是否覆盖正确位置

> 导出字段（默认）：文件名、页码、票号20位、开票日期(YYYYMMDD)、价税合计、票号完整(Y/N)

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
如发现偏移，用偏移明显的页重新校准一次即可。

### 5) 识别导出

*   拖拽 PDF 到界面，或点击拖拽区域选择 PDF
*   等待进度条跑完
*   自动导出 Excel 并打开

***

## 文件说明（需在同目录）

运行需要以下文件在同一目录：

*   `invoice_ui.py`（主界面）
*   `invoice_cli.py`（子进程执行识别/输出进度）
*   `invoice_core.py`（核心识别逻辑）
*   `calibrate_roi.py`（ROI 校准工具）
*   `app.ico`、`ing-logo.png`（UI 资源）
*   `docs/images/ui.png`、`docs/images/ui-overlay.png`（README 截图）

***

<p align="center">
  <sub><b>Designed by 余智秋 in Shanghai.</b></sub>
</p>

