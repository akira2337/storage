import glob
import json
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

# ==========================================
# 1. 設定ファイル (JSON) の読み込み
# ==========================================
CONFIG_FILE = "config.json"

try:
  with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    config = json.load(f)
except FileNotFoundError:
  raise FileNotFoundError(
      f"設定ファイル '{CONFIG_FILE}' が見つかりません。同じ階層に配置してください。"
  )

# キーの存在を前提として直接取得（スッキリ版）
INPUT_FOLDER = config["input_folder"]
OUTPUT_FILE = config["output_file"]
HEADER_ROW = config["header_row"]
TOOL_COL_IDX = config["tool_col_idx"]
TEXT_COL_INDICES = config["text_col_indices"]

CIRCLED_MAPPING = config["mapping"]
CIRCLED_NUMS = config["circled_nums"]
CIRCLED_TO_IDX = {num: i for i, num in enumerate(CIRCLED_NUMS)}


# ==========================================
# 2. 丸数字標準化・解析関数
# ==========================================
def normalize_circled_number(char):
  return CIRCLED_MAPPING.get(char, char)


def parse_cell_to_blocks(text):
  if pd.isna(text) or not str(text).strip():
    return {}, []

  lines = [line.strip() for line in str(text).splitlines() if line.strip()]
  mapping = {}
  common_texts = []
  all_circled_chars = r"[①-㉚❶-❿➀-➉➊-➓⓵-⓾]"

  for line in lines:
    match = re.match(rf"^([{all_circled_chars}0-9\-〜〜、,\s]+)([:：].*)?$", line)
    if not match:
      common_texts.append(line)
      continue

    prefix = match.group(1).strip()
    body = match.group(2) if match.group(2) else ""
    body = re.sub(r"^[:：]\s*", "", body).strip()

    if not body and not re.search(rf"{all_circled_chars}", prefix):
      common_texts.append(line)
      continue
    elif not body:
      body = line

    expanded_nums = set()
    parts = re.split(r"[、,]", prefix)

    for part in parts:
      part = part.strip()
      if not part:
        continue

      range_match = re.search(
          rf"({all_circled_chars})\s*[〜〜-]\s*({all_circled_chars})", part
      )
      if range_match:
        start_char, end_char = range_match.groups()
        norm_start = normalize_circled_number(start_char)
        norm_end = normalize_circled_number(end_char)
        start_idx = CIRCLED_TO_IDX.get(norm_start, 0)
        end_idx = CIRCLED_TO_IDX.get(norm_end, 0)
        if start_idx <= end_idx:
          for idx in range(start_idx, end_idx + 1):
            if idx < len(CIRCLED_NUMS):
              expanded_nums.add(CIRCLED_NUMS[idx])
        continue

      found_in_part = re.findall(rf"{all_circled_chars}", part)
      if found_in_part:
        for num_char in found_in_part:
          norm_char = normalize_circled_number(num_char)
          if norm_char in CIRCLED_TO_IDX:
            expanded_nums.add(norm_char)

    if expanded_nums:
      for num in expanded_nums:
        mapping[num] = body
    else:
      common_texts.append(line)

  return mapping, common_texts


# ==========================================
# 3. 指定フォルダ内のExcelファイルを自動検索・処理
# ==========================================
if not os.path.exists(INPUT_FOLDER):
  os.makedirs(INPUT_FOLDER)
  raise FileNotFoundError(
      f"'{INPUT_FOLDER}' フォルダが存在しないため作成しました。"
      "対象のExcelファイルを格納して再度実行してください。"
  )

search_pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
excel_files = glob.glob(search_pattern)

if not excel_files:
  raise FileNotFoundError(
      f"'{INPUT_FOLDER}' フォルダ内に対象のExcelファイル(.xlsx)が見つかりません。"
  )

input_file = excel_files[0]
print(f"処理対象ファイル: {input_file}")

# ------------------------------------------
# データ加工処理 (Pandas + with文)
# ------------------------------------------
with pd.ExcelFile(input_file) as xls:
  df = pd.read_excel(xls, header=HEADER_ROW)

new_rows = []

for _, row in df.iterrows():
  tool_cell = row.iloc[TOOL_COL_IDX]
  tool_dict, tool_common = parse_cell_to_blocks(tool_cell)

  if not tool_dict:
    new_row = row.copy()
    for col_idx in range(len(new_row)):
      val = new_row.iloc[col_idx]
      if pd.isna(val) or str(val).strip() == "":
        new_row.iloc[col_idx] = "ー"
    new_rows.append(new_row)
    continue

  text_dicts = {}
  text_commons = {}
  for t_idx in TEXT_COL_INDICES:
    if t_idx < len(row):
      t_dict, t_comm = parse_cell_to_blocks(row.iloc[t_idx])
      text_dicts[t_idx] = t_dict
      text_commons[t_idx] = t_comm
    else:
      text_dicts[t_idx] = {}
      text_commons[t_idx] = []

  for num, tool_text in tool_dict.items():
    new_row = row.copy()
    new_row.iloc[TOOL_COL_IDX] = tool_text

    for t_idx in TEXT_COL_INDICES:
      if t_idx < len(row):
        t_dict = text_dicts[t_idx]
        t_comm = text_commons[t_idx]

        if num in t_dict:
          body_text = t_dict[num]
          new_row.iloc[t_idx] = f"{num}：{body_text}"
        elif t_comm:
          new_row.iloc[t_idx] = "\n".join(t_comm)
        else:
          new_row.iloc[t_idx] = "ー"

    for col_idx in range(len(new_row)):
      val = new_row.iloc[col_idx]
      if pd.isna(val) or str(val).strip() == "":
        new_row.iloc[col_idx] = "ー"

    new_rows.append(new_row)

df_result = pd.DataFrame(new_rows)

# ------------------------------------------
# フォーマット維持の書き出し (Openpyxl + with文)
# ------------------------------------------
with openpyxl.load_workbook(input_file) as wb:
  ws = wb.active

  START_ROW = 6
  max_row = ws.max_row
  for row in ws.iter_rows(min_row=START_ROW, max_row=max_row):
    for cell in row:
      cell.value = None

  for i, row_data in df_result.iterrows():
    current_row = START_ROW + i
    for col_idx, val in enumerate(row_data):
      cell = ws.cell(row=current_row, column=col_idx + 1)
      cell.value = val

      prev_cell = ws.cell(
          row=max(START_ROW, current_row - 1), column=col_idx + 1
      )
      if prev_cell.font:
        cell.font = Font(
            name=prev_cell.font.name,
            size=prev_cell.font.size,
            bold=prev_cell.font.bold,
            color=prev_cell.font.color,
        )
      if prev_cell.alignment:
        cell.alignment = Alignment(
            horizontal=prev_cell.alignment.horizontal,
            vertical=prev_cell.alignment.vertical,
            wrap_text=prev_cell.alignment.wrap_text,
        )
      if prev_cell.border:
        cell.border = prev_cell.border
      if prev_cell.fill and prev_cell.fill.fill_type:
        cell.fill = prev_cell.fill

  wb.save(OUTPUT_FILE)