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

## 參數設定

程式頂端統一配置：

```python
# 由使用者在執行時輸入
INPUT_FOLDER = None  # PPT 檔案根資料夾

# 輸出設定
OUTPUT_ROOT   = r" "  # 最終輸出根目錄，改成想存檔的資料夾
IMG_FORMAT    = 'jpg'          # 匯出圖片格式：'jpg' / 'png'

# 偵測範圍關鍵字(可以自行更改)
START_KEYWORD = "工作項目說明" #本月工作項目報告
END_KEYWORD   = "工作計畫說明" #下個月工作計劃說明
``` 

---

## 核心函式

### 1. `sanitize_text(text: str) -> str`
移除 Excel 不支援的控制字元（ASCII 0x00-0x1F，保留 `\t`,`\n`,`\r`）。

```python
def sanitize_text(text):
    return ''.join(
        c for c in text
        if c in ('\t','\n','\r') or ord(c) >= 32
    )
``` 

### 2. `find_slide_range(ppt_path, start_kw, end_kw) -> (int, int)`
掃描簡報文字，找到首個含 `start_kw` 的頁碼作為 `start_idx`，
接著尋找 `end_kw` 作為 `end_idx`；若任一關鍵字缺失，回傳 `(None, None)`。

```python
def find_slide_range(ppt_path, start_kw, end_kw):
    try:
        prs = Presentation(ppt_path)
    except Exception:
        return None, None

    start_idx = end_idx = None
    for idx, slide in enumerate(prs.slides, start=1):
        text = ''.join(
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

### 3. `extract_text_to_excel(...)`
將指定範圍的文字匯出至 Excel，並在 C 欄寫入對應截圖路徑：

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
        ws.cell(1,1,'Result'); ws.cell(2,1,'無'); wb.save(excel_path)
        return

    wb = Workbook(); ws = wb.active; ws.title = 'Slides_Text'
    ws.append(['Slide Number', 'Text Content', 'Image Path'])

    for idx, slide in enumerate(prs.slides, start=1):
        if idx < start_slide or idx > end_slide:
            continue

        texts = []
        for shape in slide.shapes:
            if getattr(shape, 'has_text_frame', False):
                txt = shape.text.strip()
                if txt:
                    texts.append(txt)
        content = sanitize_text('\n'.join(texts)) or '無文字'

        row = [idx, content]
        if image_dir:
            img = f'slide_{idx}.{img_format}'
            row.append(os.path.abspath(os.path.join(image_dir, img)))

        ws.append(row)
    wb.save(excel_path)
``` 

### 4. `export_images(...)`
透過 [**PowerPoint COM API**](https://developer.microsoft.com/en-us/powerpoint)，將範圍內的投影片匯出為圖片：

```python
def export_images(
    ppt_path, output_dir,
    img_format, start_slide, end_slide
):
    path = os.path.abspath(ppt_path)
    if not os.path.exists(path):
        print(f"[Error] 檔案不存在: {path}")
        return

    app = win32com.client.Dispatch('PowerPoint.Application')
    pres = app.Presentations.Open(path, ReadOnly=1, WithWindow=0)
    os.makedirs(output_dir, exist_ok=True)

    total = pres.Slides.Count
    end_idx = min(end_slide, total)
    for i in range(start_slide, end_idx+1):
        dest = os.path.join(output_dir, f'slide_{i}.{img_format}')
        pres.Slides.Item(i).Export(dest, img_format)
    pres.Close(); app.Quit()
``` 

---

## 主程式流程

```python
if __name__ == '__main__':
    INPUT_FOLDER = input('請輸入 PPT 資料夾路徑：').strip().strip('"\'')
    folder = Path(INPUT_FOLDER)
    if not folder.is_dir():
        print(f"[Error] 資料夾不存在: {INPUT_FOLDER}")
        sys.exit(1)

    for ppt in folder.rglob('*.ppt*'):
        num = re.match(r'^(\\d+)', ppt.stem)
        base = num.group(1) if num else ppt.stem
        out_dir = Path(OUTPUT_ROOT) / base
        out_dir.mkdir(parents=True, exist_ok=True)
        excel = out_dir / f"{base}.xlsx"

        s, e = find_slide_range(str(ppt), START_KEYWORD, END_KEYWORD)
        if not s or not e:
            wb = Workbook(); ws = wb.active
            ws.cell(1,1,'Result'); ws.cell(2,1,'無'); wb.save(excel)
            continue

        extract_text_to_excel(str(ppt), str(excel), s, e, image_dir=str(out_dir), img_format=IMG_FORMAT)
        export_images(str(ppt), str(out_dir), IMG_FORMAT, s, e)
    print('全部檔案處理完成，輸出於:', OUTPUT_ROOT)
``` 

---

## 執行指令
於**Terminal**中輸入以下指令:

```bash
python ppt_to_excel.py
```

輸入資料夾後，程式將依序顯示偵測範圍與匯出進度，範例如下:

```
請輸入 PPT 資料夾路徑："C:\Users\USER\Desktop\ppts"
[11301] 偵測到範圍：第18頁 → 第30頁
✔ 已儲存 Excel: C:\...\11301\11301.xlsx
✔ 匯出圖片 第18頁 → C:\...\11301\slide_18.jpg
...
```
---

## 完整程式碼檔案
[🔗ppt_to_excel.py](https://github.com/Bgb941207/work-log/blob/master/static/PPT_to_Excel.py)
