# make_pivot_procedural_clean.py
# pip install pandas xlsxwriter openpyxl

import pandas as pd
import numpy as np
from pathlib import Path

IN_XLSX  = Path("AP_08082025.xlsx")              # your weekly workbook
OUT_XLSX = Path("AP_08082025_with_pivot.xlsx")   # output workbook
PRODUCT_NOT_APPROPRIATE = "Product Not Appropriate"
VALID_FLAG = "Valid"

# ---------- READ ----------
auto_df = pd.read_excel(IN_XLSX, sheet_name="AutoComplete")

# standardize column names just in case
auto_df = auto_df.rename(columns={c: c.strip() for c in auto_df.columns})

required = {"segment3", "segment4", "SnapDate", "Volume"}
missing = required - set(auto_df.columns)
if missing:
    raise ValueError(f"Missing required columns in AutoComplete: {missing}")

# ensure SnapDate is datetime (drop time/nanos to avoid warnings)
auto_df["SnapDate"] = pd.to_datetime(auto_df["SnapDate"], errors="coerce").dt.normalize()

# ---------- BUILD SUMMARY (no groupby.apply) ----------
d = auto_df.loc[auto_df["segment3"] == PRODUCT_NOT_APPROPRIATE].copy()

if d.empty:
    pivot_df = pd.DataFrame(columns=[
        "Event Ending Week", "Valid_SumOfVolume", "Valid_%OfVolume",
        "Total_SumOfVolume", "Total_%"
    ])
else:
    # totals per week
    total_sum = d.groupby("SnapDate", dropna=False)["Volume"].sum()

    # valid per week (filter first, then group)
    valid_sum = (
        d.loc[d["segment4"] == VALID_FLAG]
          .groupby("SnapDate", dropna=False)["Volume"].sum()
          .reindex(total_sum.index, fill_value=0)
    )

    pivot_df = pd.DataFrame({
        "Valid_SumOfVolume": valid_sum.astype(float),
        "Total_SumOfVolume": total_sum.astype(float)
    })
    pivot_df["Valid_%OfVolume"] = np.where(
        pivot_df["Total_SumOfVolume"] > 0,
        pivot_df["Valid_SumOfVolume"] / pivot_df["Total_SumOfVolume"],
        0.0
    )
    pivot_df["Total_%"] = 1.0

    pivot_df = (
        pivot_df
        .reset_index()
        .rename(columns={"SnapDate": "Event Ending Week"})
        .sort_values("Event Ending Week")
        .reset_index(drop=True)
    )

# ---------- WRITE (dates are clean, no nanos) ----------
with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter", datetime_format="mm/dd/yyyy") as writer:
    auto_df.to_excel(writer, sheet_name="AutoComplete", index=False)

    wb = writer.book
    ws = wb.add_worksheet("Pivot")

    title_fmt = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#9CCB19", "align": "left"})
    sub_fmt   = wb.add_format({"bold": True})
    hdr_fmt   = wb.add_format({"bold": True, "align": "center", "bg_color": "#DAEEF3", "border": 1})
    int_fmt   = wb.add_format({"border": 1, "align": "center"})
    pct_fmt   = wb.add_format({"num_format": "0.0%", "border": 1, "align": "center"})
    date_fmt  = wb.add_format({"num_format": "mm/dd/yyyy", "border": 1, "align": "center"})

    ws.merge_range(0, 0, 0, 7, "Text Rationale Validity of Product Not Appropriate", title_fmt)

    ws.write(2, 0, "Product Appropriateness Result", sub_fmt)
    ws.write(2, 1, PRODUCT_NOT_APPROPRIATE)

    ws.write(4, 0, "Event Ending Week", hdr_fmt)
    ws.merge_range(4, 1, 4, 2, "Valid", hdr_fmt)
    ws.merge_range(4, 3, 4, 4, "Total", hdr_fmt)

    ws.write(5, 1, "Sum of Volume", hdr_fmt)
    ws.write(5, 2, "% of Volume", hdr_fmt)
    ws.write(5, 3, "Total Sum of Volume", hdr_fmt)
    ws.write(5, 4, "Total %", hdr_fmt)

    start_row = 6
    for i, row in pivot_df.iterrows():
        # row["Event Ending Week"] is midnight-normalized; safe for Excel date
        ws.write_datetime(start_row + i, 0, pd.to_datetime(row["Event Ending Week"]).to_pydatetime(), date_fmt)
        ws.write_number(start_row + i, 1, float(row["Valid_SumOfVolume"]), int_fmt)
        ws.write_number(start_row + i, 2, float(row["Valid_%OfVolume"]), pct_fmt)
        ws.write_number(start_row + i, 3, float(row["Total_SumOfVolume"]), int_fmt)
        ws.write_number(start_row + i, 4, float(row["Total_%"]), pct_fmt)

    gt = start_row + len(pivot_df)
    ws.write(gt, 0, "Grand Total", hdr_fmt)
    if len(pivot_df) > 0:
        ws.write_formula(gt, 1, f"=SUM(B{start_row+1}:B{gt})", int_fmt)
        ws.write_formula(gt, 2, f"=IF(SUM(D{start_row+1}:D{gt})=0,0,SUM(B{start_row+1}:B{gt})/SUM(D{start_row+1}:D{gt}))", pct_fmt)
        ws.write_formula(gt, 3, f"=SUM(D{start_row+1}:D{gt})", int_fmt)
        ws.write_formula(gt, 4, f"=IF(D{gt+1}=0,0,D{gt+1}/D{gt+1})", pct_fmt)

    ws.set_column(0, 0, 18)
    ws.set_column(1, 4, 20)

print(f"Done. Wrote: {OUT_XLSX.resolve()}")
