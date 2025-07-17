+++
date = '2025-07-17T09:39:15+08:00'
draft = false
title = 'PPT_to_Excel'
+++
# PPT_to_Excel & IMG 工具
<!--more-->

此 Python 程式用於自動化處理資料夾及子資料夾內的所有 PPT 檔，
根據「工作項目說明」至「工作計畫說明」的頁面範圍，
將文字擷取至 Excel，並將該範圍投影片匯出為 JPG 圖片。

### 特色功能

- ✅ **遞迴遍歷**：自動遍歷指定資料夾及其子資料夾的 `.ppt` / `.pptx` 檔案
- ✅ **範圍偵測**：以「工作項目說明」為起始頁，「工作計畫說明」為終止頁，自動定位範圍
- ✅ **Excel 輸出**：擷取文字內容，移除不支援的控制字元，輸出至 `.xlsx`
- ✅ **圖片匯出**：將範圍內的每頁投影片批次匯出為 JPG 圖檔
- ✅ **統一參數區**：程式頂端集中管理所有可調整參數，方便日後維護

---

## 1. 引用函式庫

```python
import os
import sys
import zipfile
import re
from pathlib import Path
import win32com.client
from pptx import Presentation
from openpyxl import Workbook
```

- **os, sys**：檔案與錯誤處理
- **zipfile**：偵測破損 PPT 檔案時拋出例外
- **re**：正規表達式，擷取檔名中的編號
- **Path**：統一路徑處理
- **win32com.client**：COM 自動化介面匯出圖片
- **python-pptx**：操作 PPT 結構與文字擷取
- **openpyxl**：建立與輸出 Excel 檔案

---

## 2. 參數設定

```python
# 輸入
INPUT_FOLDER = None  # 包含 PPT 的根資料夾路徑
# 輸出
OUTPUT_ROOT = r"C:\Users\USER\Desktop\ppt_2_excel\excel_output"
# 圖片格式
IMG_FORMAT = 'jpg'
# 偵測範圍關鍵字
START_KEYWORD = "工作項目說明"
END_KEYWORD   = "工作計畫說明"
```

- **INPUT_FOLDER**：執行時透過 `input()` 填入，程式將遞迴處理此資料夾
- **OUTPUT_ROOT**：所有結果將依 PPT 編號（檔名前綴數字）存入此根目錄下的子資料夾
- **IMG_FORMAT**：匯出圖片格式，預設 `jpg`
- **START_KEYWORD** / **END_KEYWORD**：文字範圍定位依據

---

## 3. 核心函式

### sanitize_text()
移除 Excel 不支援的控制字元，確保寫入不會失敗。  
```python
def sanitize_text(text):
    return ''.join(
        c for c in text
        if c in ('\t', '\n', '\r') or ord(c) >= 32
    )
```

### find_slide_range()
偵測包含起始與終止關鍵字的頁碼，返回 `(start, end)`，若失敗回傳 `(None, None)`。
```python
def find_slide_range(ppt_path, start_kw, end_kw):
    try:
        prs = Presentation(ppt_path)
    except:
        return None, None
    start_idx = end_idx = None
    for idx, slide in enumerate(prs.slides, start=1):
        text = "".join(s.text for s in slide.shapes if hasattr(s, 'text'))
        if start_idx is None and start_kw in text:
            start_idx = idx
        if start_idx and end_kw in text:
            end_idx = idx
            break
    return start_idx, end_idx
```

### extract_text_to_excel()
將指定範圍文字擷取並輸出至 Excel，若無內容則填 `無`。
```python
def extract_text_to_excel(ppt_path, excel_path, start_slide, end_slide):
    prs = Presentation(ppt_path)
    wb = Workbook()
    ws = wb.active
    # ...（標題列與迴圈寫入）
    wb.save(excel_path)
```

### export_images()
透過 COM 介面將投影片範圍批次匯出為圖片。
```python
def export_images(ppt_path, output_dir, img_format, start_slide, end_slide):
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(ppt_path, ReadOnly=1, WithWindow=0)
    # ...（匯出迴圈）
    pres.Close()
    app.Quit()
```

---

## 4. 主程式流程

1. 輸入 `INPUT_FOLDER` 路徑，確認資料夾存在  
2. 遞迴搜尋 `.ppt` / `.pptx` 檔案  
3. 依檔名前綴數字建立子資料夾  
4. 偵測頁碼範圍：若失敗僅輸出 Excel `無`  
5. 執行 `extract_text_to_excel()` 與 `export_images()`  
6. 完成後於 `OUTPUT_ROOT` 查看結果

---

## 5. 程式執行

```bash
python p2e_3.py  # 執行後依提示輸入包含 PPT 的資料夾路徑
```
執行後過程會類似像這樣
```bash
請輸入包含 PPT 檔案的資料夾絕對路徑："C:\Users\USER\Desktop\ppt_2_excel\八里廠ppt"
[11301] 偵測到範圍：第18頁 → 第30頁
✔ 已儲存文字 Excel: C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\11301.xlsx
✔ 匯出圖片 第18頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_18.jpg
✔ 匯出圖片 第19頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_19.jpg
✔ 匯出圖片 第20頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_20.jpg
✔ 匯出圖片 第21頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_21.jpg
✔ 匯出圖片 第22頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_22.jpg
✔ 匯出圖片 第23頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_23.jpg
✔ 匯出圖片 第24頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_24.jpg
✔ 匯出圖片 第25頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_25.jpg
✔ 匯出圖片 第26頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_26.jpg
✔ 匯出圖片 第27頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_27.jpg
✔ 匯出圖片 第28頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_28.jpg
✔ 匯出圖片 第29頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_29.jpg
✔ 匯出圖片 第30頁 → C:\Users\USER\Desktop\ppt_2_excel\八里廠_維修項目整理\11301\slide_30.jpg
```
---

## 完整程式碼檔案
[🔗ppt_to_excel.py](https://github.com/Bgb941207/work-log/blob/master/static/PPT_to_Excel.py)
