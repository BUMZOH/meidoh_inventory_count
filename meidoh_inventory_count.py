from pathlib import Path

import pandas as pd


# ================================================
#   Settings
# ================================================
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "光和工業生産管理.xlsm"


# ================================================
#   Start message
# ================================================
print("=" * 60)
print("メイドー 棚卸数量集計")
print("=" * 60)
print()
print("集計対象のExcelファイルがフォルダ内にあることを確認してください。")
print(f"対象ファイル: {EXCEL_PATH.name}")
print()
input("確認できたら Enter キーを押してください...")
print()


# ================================================
#   Check Excel file
# ================================================
if not EXCEL_PATH.exists():
    raise FileNotFoundError(f"Excel file not found: {EXCEL_PATH}")


# ================================================
#   Load data
# ================================================
df = pd.read_excel(
    EXCEL_PATH,
    sheet_name="Data",
)


# ================================================
#   Filter data
# ================================================
filtered_df = df[
    (df["客先名"] == "メイドー")
    & (df["出荷"].isna())
]

if filtered_df.empty:
    print("集計対象のデータがありません。")
    input("ENTER キーを押して終了してください。")
    raise SystemExit


# ================================================
#   Aggregate data
# ================================================
result_df = (
    filtered_df.groupby("品番", as_index=False)["数量"]
    .sum()
    .sort_values("品番", ascending=True)
)


# ================================================
#   Display result
# ================================================
print(result_df.to_string(index=False))


# ================================================
#   Copy to clipboard
# ================================================
result_df.to_clipboard(index=False, header=False)

print()
print("集計結果をクリップボードにコピーしました。")
print("貼り付け先のExcelで Ctrl + V してください。")

