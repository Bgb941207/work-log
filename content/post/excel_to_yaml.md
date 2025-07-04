+++
date = '2025-07-04T13:50:24+08:00'
draft = false
title = 'Excel_to_yaml.py'
+++

這是一個可以把excel中的工廠圖控資料轉成yaml檔的程式

<!--more-->
```python
import os
import re
import pandas as pd
import yaml

# 參數設定
EXCEL_PATH = "2025_new.xlsx"
SHEET_NAME = 0  # 要選第幾個工作表(從0開始)
SELECTED_COLUMNS = ["設備編號", "設備名稱", "設備類型", "循環系統"]
KEY_MAP = {
    "設備名稱": "name",
    "設備類型": "machineType",
    "循環系統": "machineSystems"
}
LIST_FIELDS = ["machineSystems"]
OUTPUT_SHEET_NAME = "觀音廠_new"
YAML_INDENT = 3  # 縮排參數

# 字串加單引號
class SingleQuoted(str):
    pass

def representer(dumper, data):
    return dumper.represent_scalar('tag:yaml.org,2002:str', data, style="'")

yaml.add_representer(SingleQuoted, representer)

# 縮排
class IndentDumper(yaml.Dumper):
    def increase_indent(self, flow=False, indentless=False):
        return super().increase_indent(flow, indentless=False)

# 資料轉換
def rows2dict(data_rows, key_map, list_fields):
    result = []

    for row in data_rows:
        item = {}
        id = str(row.get("設備編號", "")).strip()
        name = str(row.get("設備名稱", "")).strip()

        for col_name, value in row.items():
            if col_name == "設備編號":
                continue

            key = key_map.get(col_name, col_name)

            # 類型=>其他
            if key == "machineType":
                item[key] = SingleQuoted("其他")
                continue

            # 處理 name 欄位
            if key == "name":
                pattern = re.escape(id)
                match = re.search(pattern, name)
                if match:
                    suffix = name[match.end():]
                    # ✅ 只移除空白、破折號或冒號，保留 #
                    name_suffix = re.sub(r'^[\s\-:]+', '', suffix).strip()
                else:
                    name_suffix = name
                item[key] = SingleQuoted(f"{id}.{name_suffix}")
                continue

            value_str = str(value).strip()
            if key in list_fields:
                item[key] = [SingleQuoted(value_str)]
            else:
                item[key] = SingleQuoted(value_str)

        result.append(item)

    return result

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
