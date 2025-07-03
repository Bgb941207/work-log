import os
import pandas as pd
import yaml

EXCEL_PATH = "2025.xlsx"
SHEET_NAME = 6
SELECTED_COLUMNS = ["設備編號", "設備名稱", "設備類型", "循環系統"]
KEY_MAP = {
    "設備名稱": "name",
    "設備類型": "machineType",
    "循環系統": "machineSystems"
}
LIST_FIELDS = ["machineSystems"]
SHEET_NAME_EN = "觀音廠"
Indent = 3

# 單引號
class SingleQuoted(str): pass

def single_quoted(dumper, data):
    return dumper.represent_scalar('tag:yaml.org,2002:str', data, style="'")

yaml.add_representer(SingleQuoted, single_quoted)

# 縮排
class IndentDumper(yaml.Dumper):
    def increase_indent(self, flow=False, indentless=False):
        return super().increase_indent(flow, indentless=False)

# 資料轉換
def listed_dict(data_list, key_map, list_fields):
    result = []
    for row in data_list:
        translated_row = {}

        編號 = str(row.get("設備編號", "")).strip()
        名稱 = str(row.get("設備名稱", "")).strip()

        for k, v in row.items():
            new_key = key_map.get(k, k)

            if k == "設備編號":
                continue

            if new_key == "machineType":
                translated_row[new_key] = SingleQuoted("其他")
                continue

            if new_key == "name":
                combined = f"{編號}.{名稱}"
                translated_row[new_key] = SingleQuoted(combined)
                continue

            val = str(v)
            if new_key in list_fields:
                translated_row[new_key] = [SingleQuoted(val)]
            else:
                translated_row[new_key] = SingleQuoted(val)
        result.append(translated_row)
    return result

def e2y (excel_path, sheet_name=0, selected_columns=None, key_map=None, list_fields=None, output_name=None):
    xls = pd.ExcelFile(excel_path)
    sheet_name_str = xls.sheet_names[sheet_name] if isinstance(sheet_name, int) else sheet_name
    df = pd.read_excel(xls, sheet_name=sheet_name)

    if selected_columns:
        df = df[selected_columns]

    data = df.to_dict(orient="records")
    translated_data = listed_dict(data, key_map or {}, list_fields or [])

    output_dict = {output_name or sheet_name_str: translated_data}
    output_path = f"{output_name or sheet_name_str}.yaml"

    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(
            output_dict,
            f,
            allow_unicode=True,
            sort_keys=False,
            width=float("inf"),
            indent=Indent,  
            Dumper=IndentDumper
        )

    print(f"✅ YAML 輸出完成：{output_path}")

e2y(
    excel_path=EXCEL_PATH,
    sheet_name=SHEET_NAME,
    selected_columns=SELECTED_COLUMNS,
    key_map=KEY_MAP,
    list_fields=LIST_FIELDS,
    output_name=SHEET_NAME_EN
)
