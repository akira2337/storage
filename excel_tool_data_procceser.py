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

INPUT_FOLDER = config["input_folder"]
OUTPUT_FOLDER = config["output_folder"]
TARGET_SHEET_NAME = config["target_sheet_name"]
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


def expand_prefix_to_nums(prefix):
  """プレフィックス（例: '①〜③、⑤' や '①②③'）から

  該当する丸数字のリスト（重複なし・順序維持）を抽出・展開する関数
  """
  expanded = []
  parts = re.split(r"[、,]", prefix)
  all_circled_chars = r"[①-㉚❶-❿➀-➉➊-➓⓵-⓾]"

  for part in parts:
    part = part.strip()
    if not part:
      continue

    # 範囲指定のチェック（例: ①〜③）
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
            num_char = CIRCLED_NUMS[idx]
            if num_char not in expanded:
              expanded.append(num_char)
      continue

    # 個別または連続する丸数字のチェック
    found_in_part = re.findall(rf"{all_circled_chars}", part)
    if found_in_part:
      for num_char in found_in_part:
        norm_char = normalize_circled_number(num_char)
        if norm_char in CIRCLED_TO_IDX:
          if norm_char not in expanded:
            expanded.append(norm_char)

  return expanded


def parse_cell_to_blocks(text):
  """セル内のテキストを解析し、

  {丸数字: [その丸数字が属するテキスト項目のリスト]} と、
  共通テキストを抽出する
  """
  if pd.isna(text) or not str(text).strip():
    return {}, []

  lines = [line.strip() for line in str(text).splitlines() if line.strip()]
  # 各丸数字に紐づくテキストブロックのリストを保持する辞書
  mapping = {num: [] for num in CIRCLED_NUMS}
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

    nums_in_line = expand_prefix_to_nums(prefix)

    if nums_in_line:
      for num in nums_in_line:
        mapping[num].append(body)
    else:
      common_texts.append(line)

  return mapping, common_texts


# ==========================================
# 3. 指定フォルダ内の複数ファイルを順次一括処理
# ==========================================
if not os.path.exists(INPUT_FOLDER):
  os.makedirs(INPUT_FOLDER)
  raise FileNotFoundError(
      f"'{INPUT_FOLDER}' フォルダが存在しないため作成しました。"
      "対象のExcelファイルを格納して再度実行してください。"
  )

if not os.path.exists(OUTPUT_FOLDER):
  os.makedirs(OUTPUT_FOLDER)

search_pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
excel_files = glob.glob(search_pattern)

if not excel_files:
  raise FileNotFoundError(
      f"'{INPUT_FOLDER}' フォルダ内に対象のExcelファイル(.xlsx)が見つかりません。"
  )

print(f"対象ファイル数: {len(excel_files)}件")
print(f"ターゲットシート名: {TARGET_SHEET_NAME}\n")

# すべてのファイルを1つずつループ処理
for input_file in excel_files:
  base_name = os.path.basename(input_file)
  output_file = os.path.join(OUTPUT_FOLDER, base_name)
  print(f"処理中: {base_name} ...")

  try:
    # ------------------------------------------
    # A. データ加工処理 (Pandas)
    # ------------------------------------------
    with pd.ExcelFile(input_file) as xls:
      if TARGET_SHEET_NAME not in xls.sheet_names:
        print(
            f"  -> スキップ: シート '{TARGET_SHEET_NAME}' が存在しません。"
        )
        continue
      df = pd.read_excel(xls, sheet_name=TARGET_SHEET_NAME, header=HEADER_ROW)

    new_rows = []

    for _, row in df.iterrows():
      tool_cell = row.iloc[TOOL_COL_IDX]
      tool_dict, tool_common = parse_cell_to_blocks(tool_cell)

      # ツール一覧に属する有効な丸数字を抽出（順番通り）
      active_tool_nums = []
      for num in CIRCLED_NUMS:
        if num in tool_dict and tool_dict[num]:
          active_tool_nums.append(num)

      # ツール一覧に丸数字がない場合のフォールバック
      if not active_tool_nums:
        new_row = row.copy()
        for col_idx in range(len(new_row)):
          val = new_row.iloc[col_idx]
          if pd.isna(val) or str(val).strip() == "":
            new_row.iloc[col_idx] = "ー"
        new_rows.append(new_row)
        continue

      # 各テキスト列のセルを事前に解析
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

      # ツール一覧の各丸数字をベースに行を展開
      for num in active_tool_nums:
        new_row = row.copy()
        # ツール一覧の値をセット（例: ①ツールA）
        tool_body = tool_dict[num][0]
        new_row.iloc[TOOL_COL_IDX] = f"{num}{tool_body}"

        # 各テキスト列の処理
        for t_idx in TEXT_COL_INDICES:
          if t_idx < len(row):
            t_dict = text_dicts[t_idx]
            t_comm = text_commons[t_idx]

            if num in t_dict and t_dict[num]:
              # 画像12のように、同じ丸数字に紐づくテキストを自動結合して整形
              # 例: 連続・複数の場合は ①②③：テキスト1 や ①〜③：テキスト1 のようにまとめる
              bodies = t_dict[num]
              # ここでは該当する丸数字そのものを先頭につけて綺麗に結合
              combined_body = "\n".join(bodies)
              new_row.iloc[t_idx] = f"{num}：{combined_body}"
            elif t_comm:
              new_row.iloc[t_idx] = "\n".join(t_comm)
            else:
              new_row.iloc[t_idx] = "ー"

        # その他の空白セルを「ー」で埋める
        for col_idx in range(len(new_row)):
          val = new_row.iloc[col_idx]
          if pd.isna(val) or str(val).strip() == "":
            new_row.iloc[col_idx] = "ー"

        new_rows.append(new_row)

    df_result = pd.DataFrame(new_rows)

    # ------------------------------------------
    # B. フォーマット維持の書き出し (Openpyxl + with文)
    # ------------------------------------------
    with openpyxl.load_workbook(input_file) as wb:
      ws = wb[TARGET_SHEET_NAME]

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

      wb.save(output_file)

    print(f"  -> 成功: 保存先 [{output_file}]")

  except Exception as e:
    print(f"  -> エラー発生 ({base_name}): {e}")

print("\n=== すべてのファイルの処理が完了しました！ ===")
