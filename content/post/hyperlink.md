+++
date = '2025-07-18T15:43:55+08:00'
draft = false
title = 'Hyperlink'
+++
# Excel 超連結批次新增工具
<!--more-->

此 Python 程式用於自動化處理指定資料夾及其子資料夾內所有 Excel 檔（`.xlsx` / `.xlsm`），
將每張工作表 C 欄的文字自動設為可點擊的超連結，並套用預設超連結樣式。

### 特色功能

- ✅ **遞迴遍歷**：自動搜尋指定根資料夾及其所有子目錄的 Excel 檔  
- ✅ **檔案過濾**：跳過暫存檔（以 `~$` 開頭）與非 `.xlsx`/`.xlsm` 檔案  
- ✅ **批次新增**：將 C 欄中非空字串設為超連結，並套用「Hyperlink」樣式  
- ✅ **變更檢測**：只有在檔案確實被修改後才重新儲存，減少不必要 I/O  
- ✅ **集中參數區**：程式頂端統一設定根目錄路徑，簡易維護  

---

## 參數設定

於程式最上方調整以下變數即可指定待轉換檔案之根資料夾：

```python
# ======= 在這裡指定你的根資料夾路徑 =======
ROOT_FOLDER = r"C:\Users\USER\Desktop\ppt_2_excel_相對路徑\全興廠"
# =====================================
```

---

## 核心函式

### `add_hyperlinks_in_c(root_dir: str) -> None`
遞迴搜尋 `root_dir` 下所有 Excel 檔案，並將每個檔案中每張工作表 C 欄的文字加入超連結。

```python
def add_hyperlinks_in_c(root_dir):
    '''
    遞迴搜尋 root_dir 下所有 .xlsx/.xlsm 檔案，
    並把每個檔案中每張表格 C 欄的文字，設為超連結。
    '''
    for dirpath, _, filenames in os.walk(root_dir):
        for fn in filenames:
            # 1. 跳過 Excel 暫存檔與非 .xlsx/.xlsm
            if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xlsm")):
                continue

            full_path = os.path.join(dirpath, fn)
            print(f"Processing: {full_path}")
            wb = load_workbook(full_path)
            changed = False

            # 2. 遍歷所有工作表與每列 C 欄
            for ws in wb.worksheets:
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=3)  # 第3欄 (C)
                    addr = cell.value
                    if isinstance(addr, str) and addr.strip():
                        # 3. 若尚未設定同地址的超連結，則加入
                        if getattr(cell, "hyperlink", None) != addr:
                            cell.hyperlink = addr
                            cell.style = "Hyperlink"
                            changed = True

            # 4. 若有修改，則儲存
            if changed:
                wb.save(full_path)
                print("  → Saved with hyperlinks.")
            else:
                print("  → No changes needed.")
    print("Done.")
```

---

## 主程式流程

1. 載入模組與設定 `ROOT_FOLDER`。  
2. 呼叫 `add_hyperlinks_in_c(ROOT_FOLDER)` 遞迴處理所有符合條件的檔案。  

```python
if __name__ == "__main__":
    add_hyperlinks_in_c(ROOT_FOLDER)
```

---

## 執行指令

於 **Terminal** 中輸入：

```bash
python hyperlink.py
```

輸入後程式將依序列出處理中的檔案路徑與儲存狀況，直至顯示 `Done.`。

---

## 完整程式碼檔案

[🔗hyperlink.py](https://github.com/Bgb941207/work-log/blob/master/static/hyperlink.py)
