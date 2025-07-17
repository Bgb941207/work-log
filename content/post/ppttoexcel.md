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
- ✅ **Excel 輸出**：擷取文字內容，移除不支援的控制字元，輸出至 `.xlsx`，並於 C 欄填入對應截圖的絕對路徑
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
- **win32com.client**：COM 自動化匯出圖片
- **python-pptx**：純 Python 讀取 PPT 結構與文字
- **openpyxl**：建立與操作 Excel 檔案

---

## 2. 參數設定

```python
# 輸入資料夾（執行時由 input 填入）
INPUT_FOLDER = None

# 統一管理區
OUTPUT_ROOT   = r"C:\Users\USER\Desktop\ppt_2_excel\excel_output"
IMG_FORMAT    = 'jpg'            # 匯出圖片格式，可調為 'png'
START_KEYWORD = "工作項目說明"   # 範圍起始關鍵字
END_KEYWORD   = "工作計畫說明"   # 範圍結束關鍵字
```


---

## 3. 核心函式

### sanitize_text()

移除 Excel 不支援的控制字元 (0x00-0x1F，保留 `\t`,`\n`,`\r`)，避免寫入錯誤。

```python
def sanitize_text(text):
    return ''.join(
        c for c in text
        if c in ('\t','\n','\r') or ord(c) >= 32
    )
```

---

### find_slide_range()

偵測第一個含起始關鍵字的頁碼為 `start_idx`，再往下尋找結束關鍵字為 `end_idx`，
若任一未找到回傳 `(None, None)`。

```python
def find_slide_range(ppt_path, start_kw, end_kw):
    try:
        prs = Presentation(ppt_path)
    except Exception:
        return None, None

    start_idx = end_idx = None
    for idx, slide in enumerate(prs.slides, start=1):
        text = "".join(
            shape.text for shape in slide.shapes
            if hasattr(shape, 'text') and shape.text
        )
        if start_idx is None and start_kw in text:
            start_idx = idx
        if start_idx and end_kw in text:
            end_idx = idx
            break
    return start_idx, end_idx
```

---

### extract_text_to_excel()

將指定頁碼範圍內的文字與對應截圖路徑輸出至 Excel。
- 欄位：A=Slide Number，B=Text Content，C=Image Path
- 若解析失敗或無內容，輸出 `無`

```python
def extract_text_to_excel(
    ppt_path, excel_path,
    start_slide, end_slide,
    image_dir=None, img_format='jpg'
):
    try:
        prs = Presentation(ppt_path)
    except Exception:
        wb = Workbook(); ws = wb.active
        ws.cell(1,1,"Result"); ws.cell(2,1,"無"); wb.save(excel_path)
        return

    wb = Workbook(); ws = wb.active; ws.title="Slides_Text"
    ws.append(["Slide Number","Text Content","Image Path"])

    for idx, slide in enumerate(prs.slides, start=1):
        if idx < start_slide or idx > end_slide:
            continue

        texts = []
        for shape in slide.shapes:
            if getattr(shape,'has_text_frame',False):
                txt = shape.text.strip()
                if txt:
                    texts.append(txt)
        content = sanitize_text("\n".join(texts)) or "無文字"

        row = [idx, content]
        if image_dir:
            img = f"slide_{idx}.{img_format}"
            row.append(os.path.abspath(os.path.join(image_dir,img)))

        ws.append(row)
    wb.save(excel_path)
```

---

### export_images()

使用 PowerPoint COM API 批次匯出指定範圍投影片為圖片。

```python
def export_images(
    ppt_path, output_dir,
    img_format, start_slide, end_slide
):
    path = os.path.abspath(ppt_path)
    if not os.path.exists(path):
        print(f"[Error] 檔案不存在: {path}"); return

    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(path, ReadOnly=1, WithWindow=0)
    os.makedirs(output_dir,exist_ok=True)

    total = pres.Slides.Count
    end_idx = min(end_slide, total)
    for i in range(start_slide, end_idx+1):
        dest = os.path.join(output_dir,f"slide_{i}.{img_format}")
        pres.Slides.Item(i).Export(dest,img_format)
    pres.Close(); app.Quit()
```

---

## 4. 主程式流程

```python
if __name__=='__main__':
    INPUT_FOLDER = input("請輸入 PPT 資料夾路徑：").strip().strip('"\'')
    folder = Path(INPUT_FOLDER)
    if not folder.is_dir():
        print(f"[Error] 資料夾不存在: {INPUT_FOLDER}"); sys.exit(1)

    for ppt in folder.rglob('*.ppt*'):
        num = re.match(r'^(\d+)',ppt.stem)
        base = num.group(1) if num else ppt.stem
        out_dir = Path(OUTPUT_ROOT)/base; out_dir.mkdir(parents=True,exist_ok=True)
        excel = out_dir/f"{base}.xlsx"

        s,e = find_slide_range(str(ppt),START_KEYWORD,END_KEYWORD)
        if not s or not e:
            wb=Workbook();ws=wb.active;ws.cell(1,1,"Result");ws.cell(2,1,"無");wb.save(excel);continue

        extract_text_to_excel(str(ppt),str(excel),s,e,image_dir=str(out_dir),img_format=IMG_FORMAT)
        export_images(str(ppt),str(out_dir),IMG_FORMAT,s,e)
    print("全部檔案處理完成，輸出於:",OUTPUT_ROOT)
```

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
