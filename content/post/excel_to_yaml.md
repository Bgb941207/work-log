+++
date = '2025-07-04T13:50:24+08:00'
draft = false
title = 'Excel_to_yaml.py'
+++

# excel轉成yaml檔的程式
<!--more-->
這份 Python 程式設計用來將 Excel 中的結構化資料轉換成 YAML 格式。特色如下：

- ✅ 支援 Excel 多欄位資料選擇與欄位轉換
- ✅ 自動合併設備編號與名稱為單一欄位
- ✅ 支援某些欄位轉為 YAML list 格式
- ✅ 自訂 YAML 字串格式（加單引號）、縮排與輸出格式
---


## 1.引用函式庫
匯入程式所需的標準與第三方套件：
```python
import os
import re
import pandas as pd
import yaml
```
- os: 操作系統功能（如路徑處理）
- re: 正規表達式處理字串
- pandas: 讀取與處理 Excel 資料
- yaml: 將資料結構轉成 YAML 格式輸出
---

## 2.參數設定
定義檔案路徑、要讀取的工作表與欄位，以及 YAML 的輸出格式設定:
```python
EXCEL_PATH = "re_2025.xlsx"
SHEET_NAME = 2  # 要選第幾個工作表(從0開始)
SELECTED_COLUMNS = ["設備編號", "設備名稱", "設備類型", "循環系統"]
KEY_MAP = {
    "設備名稱": "name",
    "設備類型": "machineType",
    "循環系統": "machineSystems"
}
LIST_FIELDS = ["machineSystems"]
OUTPUT_SHEET_NAME = "觀音廠"
YAML_INDENT = 3  # 縮排參數
```
---

## 3.格式設定與縮排
為了讓輸出的 YAML 檔案符合特定格式需求，我們加入以下客製化設定：
```python
# 字串加單引號
class SingleQuoted(str):
    pass

def representer(dumper, data):
    return dumper.represent_scalar('tag:yaml.org,2002:str', data, style="'")

yaml.add_representer(SingleQuoted, representer)
```
- 這段程式會讓所有字串輸出時自動加上單引號 '...'，例如 'P-0601.主機A'
```python
# 縮排
class IndentDumper(yaml.Dumper):
    def increase_indent(self, flow=False, indentless=False):
        return super().increase_indent(flow, indentless=False)
```
- 這段確保 YAML 檔案縮排正確，避免列表格式錯亂問題
---
## 4.資料轉換函式
這個函式會逐筆讀取 Excel 每一列資料，轉成 dictionary 並處理欄位轉換與格式化。
```python
# 資料轉換
def rows2dict(data_rows, key_map, list_fields):
    result = []

    for row in data_rows:
        item = {}
        id = str(row.get("設備編號", "")).strip()
        name = str(row.get("設備名稱", "")).strip()
```
- id 是設備編號，例如 P-0601
- name 是設備名稱，例如 P-0601主機A，我們會試著從中拆出和設備編號重複的部分。

```python
        for col_name, value in row.items():
            if col_name == "設備編號":
                continue

            key = key_map.get(col_name, col_name)

            # 處理 name 欄位
            if key == "name":
            # 嘗試找到 id 的前綴在 name 中出現的位置
                for i in range(len(id), 0, -1):
                    prefix = id[:i]
                    if name.startswith(prefix):
                        suffix = name[len(prefix):]
                        name_suffix = re.sub(r'^[\s\-:]+', '', suffix).strip()
                        item[key] = SingleQuoted(f"{id}.{name_suffix}")
                        break
                else:
                    # 如果完全沒有前綴相符，就用原始 name
                    item[key] = SingleQuoted(f"{id}.{name}")
                continue
```
- 這段會試圖從名稱中把編號前綴抽出，只保留剩下的「設備名稱」部分。
- 效果如下:

| 設備編號 | 設備名稱 | 輸出name |
|:---:|:---:|:---:|
|A-1234|設備A|A-1234.設備A|
|B-2345|B-2345設備B|B-2345.設備B|
|C-3456|#1設備C|C-3456.#1設備C|
```python
            value_str = str(value).strip()
            if key in list_fields:
                item[key] = [SingleQuoted(value_str)]
            else:
                item[key] = SingleQuoted(value_str)

        result.append(item)

    return result
```

- machineSystems 等欄位會以 YAML 陣列格式輸出，例如：
machineSystems: '油壓系統'
---
## 5.主函式
這個主程式會：
- 開啟 Excel
- 篩選需要的欄位
- 呼叫 rows2dict() 處理資料
- 將結果輸出為 YAML 檔案
```python
def excel2yaml(
    excel_path,
    sheet_name=0,
    selected_columns=None,
    key_map=None,
    list_fields=None,
    output_name=None
):
    # 讀取
    xls = pd.ExcelFile(excel_path)
    title = xls.sheet_names[sheet_name] if isinstance(sheet_name, int) else sheet_name
    df = pd.read_excel(xls, sheet_name=sheet_name)

    if selected_columns:
        df = df[selected_columns]

    data_rows = df.to_dict(orient="records")
    transformed_data = rows2dict(
        data_rows, key_map or {}, list_fields or []
    )

    # 輸出結構
    dict = {output_name or title: transformed_data}
    output_file = f"{output_name or title}.yaml"

    with open(output_file, "w", encoding="utf-8") as f:
        yaml.dump(
            dict,
            f,
            allow_unicode=True,
            sort_keys=False,
            width=float("inf"),
            indent=YAML_INDENT,
            Dumper=IndentDumper
        )

    print(f"輸出完成：{output_file}")
```
---
## 6.程式執行
當程式被直接執行時，會呼叫 excel2yaml() 並帶入我們預先設定的參數。
```python
# 執行程式
if __name__ == "__main__":
    excel2yaml(
        excel_path=EXCEL_PATH,
        sheet_name=SHEET_NAME,
        selected_columns=SELECTED_COLUMNS,
        key_map=KEY_MAP,
        list_fields=LIST_FIELDS,
        output_name=OUTPUT_SHEET_NAME
    )
```
---
## 以下為轉換後的 YAML 範例（含單引號與縮排）：

```python
觀音廠:
   -  name: 'B-0117.#1M系收塵風機'
      machineType: 'M系風車'
      machineSystems:
         - '研磨系統#1#2'
```
---
## 完整程式碼檔案:
[🔗 excel2yaml.py]( https://github.com/Bgb941207/work-log/blob/master/content/post/excel2yaml.py )