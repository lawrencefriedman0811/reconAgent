"""
Schedule K to K-2 reconciliation.

Python port of the Power Query (M) workbook queries found in "Recon Mcode.txt".
It reproduces the following M queries as pandas DataFrames:

    part2               -> K-2 Part II amounts (unpivoted, signed by REF_CODE)
    partx               -> K-2 Part X amounts  (unpivoted, signed by REF_CODE)
    excludePart2Table   -> K3 codes to exclude from the Part II reconciliation
    excludePartXTable   -> K3 codes to exclude from the Part X reconciliation
    schK                -> Schedule K amounts (unpivoted from the wide control grid)
    schKToK2Part2Diff   -> schK + part2 stacked, exclusions removed
    schKToK2PartXDiff   -> schK + partx stacked, exclusions removed
    grossIncome         -> Part II vs Part X gross-income comparison by K3 code

Data is pulled from the named ranges / tables in the workbook:

    k2Part2        'K2II-Part II'!A14:W3008
    k2PartX        'K2X-Part X'!A8:U2496
    part2Exclude   '(Z.00) Dynamic Dropdown List'!FN2:FO79
    partXExclude   '(Z.00) Dynamic Dropdown List'!FQ2:FR78
    allocRec       '(K.01) Sch K to K2 Control'!C10:GH1802
"""

from __future__ import annotations

import warnings
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string

warnings.simplefilter("ignore")

DEFAULT_WORKBOOK = (
    "2772 WM ALTERNATIVES TPG RISE_2025_Actual Filing_Federal Workpaper (1).xlsm"
)

# Named ranges (sheet, range) discovered from the workbook's defined names / tables.
RANGES = {
    "k2Part2": ("K2II-Part II", "A14:W3008"),
    "k2PartX": ("K2X-Part X", "A8:U2496"),
    "part2Exclude": ("(Z.00) Dynamic Dropdown List", "FN2:FO79"),
    "partXExclude": ("(Z.00) Dynamic Dropdown List", "FQ2:FR78"),
    "allocRec": ("(K.01) Sch K to K2 Control", "C10:GH1802"),
}


# --------------------------------------------------------------------------- #
# Excel helpers
# --------------------------------------------------------------------------- #
def _read_range(ws, ref) -> list[list]:
    """Return the cell values of a worksheet range as a list of rows."""
    return [[c.value for c in row] for row in ws[ref]]


def _range_to_df(ws, ref, promote_headers: bool = True) -> pd.DataFrame:
    """Read a range and (optionally) promote the first row to column headers."""
    rows = _read_range(ws, ref)
    if not rows:
        return pd.DataFrame()
    if promote_headers:
        header = rows[0]
        return pd.DataFrame(rows[1:], columns=header)
    return pd.DataFrame(rows)


def _to_number(series: pd.Series) -> pd.Series:
    """Coerce to numeric (non-convertible -> NaN), mirroring M's `type number`."""
    return pd.to_numeric(series, errors="coerce")


def _is_number(value) -> bool:
    """Replicate M's `(try Number.From(x) otherwise null) <> null`."""
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return not (isinstance(value, float) and np.isnan(value))
    try:
        float(str(value).strip())
        return True
    except (ValueError, TypeError):
        return False


# --------------------------------------------------------------------------- #
# Queries
# --------------------------------------------------------------------------- #
def build_part2(ws) -> pd.DataFrame:
    """M query: part2  (K-2 Part II)."""
    df = _range_to_df(ws, RANGES["k2Part2"][1])

    df = df.rename(
        columns={
            "K3_CODE_ID": "K3_Code",
            "Activity Number": "Activity_Num",
            "Name of Activity": "Activity_Name",
            "Activity Allocation": "Activity_Allocation",
            "ACTIVITY_ALLOCATION_ID": "Allocation_ID",
        }
    )
    df["Source"] = "K-2 Part II"
    df["KDescription"] = None

    # Columns that stay fixed during the unpivot (everything except the
    # category amount columns gets carried through).
    id_cols = [
        "Schedule K Equivelant", "REF_CODE", "K3_Code", "INSERT_TS", "USER_ID",
        "SOURCE_FILENAME", "Allocation_ID", "Activity_Num", "Activity_Name",
        "Actual vs. Estimate", "PY True-Up (Yes/No)", "Activity_Allocation",
        "Section", "Line", "Country Code (See Note)", "Detail", "Source",
        "KDescription",
    ]
    value_cols = [c for c in df.columns if c not in id_cols]

    melted = df.melt(
        id_vars=id_cols, value_vars=value_cols,
        var_name="Attribute", value_name="Value",
    )
    # M's Unpivot drops null cells.
    melted = melted[melted["Value"].notna()].copy()

    melted["Value"] = _to_number(melted["Value"])
    melted["REF_CODE"] = _to_number(melted["REF_CODE"])
    melted["Part 2"] = melted["REF_CODE"] * melted["Value"]

    keep = ["K3_Code", "Activity_Num", "Activity_Name", "Activity_Allocation",
            "Line", "Source", "Attribute", "Part 2"]
    melted = melted[keep].rename(columns={"Line": "Part 2 Line"})
    return melted.reset_index(drop=True)


def build_partx(ws) -> pd.DataFrame:
    """M query: partx  (K-2 Part X)."""
    df = _range_to_df(ws, RANGES["k2PartX"][1])

    # Filter to data rows: Activity Number must be numeric.
    df = df[df["Activity Number"].apply(_is_number)].copy()

    df = df.rename(
        columns={
            "K3_CODE_ID": "K3_Code",
            "Activity Number": "Activity_Num",
            "Name of Activity": "Activity_Name",
            "Activity Allocation": "Activity_Allocation",
            "ACTIVITY_ALLOCATION_ID": "Allocation_ID",
            "Line": "Part X Line",
        }
    )
    df["Activity_Num"] = _to_number(df["Activity_Num"]).astype("Int64")
    df["Allocation_ID"] = _to_number(df["Allocation_ID"]).astype("Int64")

    df["Source"] = "K-2 Part X"
    df["KDescription"] = None

    id_cols = [
        "Schedule K Equivelant", "REF_CODE", "K3_Code", "INSERT_TS", "USER_ID",
        "SOURCE_FILENAME", "Allocation_ID", "Activity_Num", "Activity_Name",
        "Actual vs. Estimate", "PY True-Up (Yes/No)", "Activity_Allocation",
        "Section", "Part X Line", "Details", "Source", "KDescription",
    ]
    value_cols = [c for c in df.columns if c not in id_cols]

    melted = df.melt(
        id_vars=id_cols, value_vars=value_cols,
        var_name="Attribute", value_name="Value",
    )
    melted = melted[melted["Value"].notna()].copy()

    melted["Value"] = _to_number(melted["Value"])
    melted["REF_CODE"] = _to_number(melted["REF_CODE"])
    melted["Part X"] = melted["REF_CODE"] * melted["Value"]

    keep = ["K3_Code", "Activity_Num", "Activity_Name", "Activity_Allocation",
            "Part X Line", "Source", "Attribute", "Part X"]
    melted = melted[keep]
    return melted.reset_index(drop=True)


def build_exclude(ws, key) -> pd.DataFrame:
    """M queries: excludePart2Table / excludePartXTable."""
    df = _range_to_df(ws, RANGES[key][1])
    df["excludeCode"] = _to_number(df["excludeCode"]).astype("Int64")
    df["exclude"] = df["exclude"].astype("object")
    return df.reset_index(drop=True)


def build_schk(ws) -> pd.DataFrame:
    """M query: schK  (Schedule K from the wide control grid)."""
    rows = _read_range(ws, RANGES["allocRec"][1])
    grid = pd.DataFrame(rows)  # positional columns 0..N-1

    META_COL_COUNT = 13
    n_cols = grid.shape[1]
    data_col_idx = list(range(META_COL_COUNT, n_cols))  # amount columns by position

    # --- Header band (first 8 rows of the range) ---
    def header_row(idx):
        return list(grid.iloc[idx].values)

    row10 = header_row(0)  # Activity Number
    row11 = header_row(1)  # Activity Name
    row12 = header_row(2)  # Activity Allocation
    row14 = header_row(4)  # D/O flag
    row15 = header_row(5)  # Allocation ID

    header = pd.DataFrame(
        {
            "AmtCol": data_col_idx,
            "Activity_Num": [row10[i] for i in data_col_idx],
            "Activity_Name": [row11[i] for i in data_col_idx],
            "Activity_Allocation": [row12[i] for i in data_col_idx],
            "DO_Flag": [row14[i] for i in data_col_idx],
            "Allocation_ID": [row15[i] for i in data_col_idx],
        }
    )

    # Keep header columns whose Activity_Num is numeric...
    header = header[header["Activity_Num"].apply(_is_number)]
    # ...and whose Activity_Name is not blank / "0".
    def _valid_name(v):
        if v is None:
            return False
        s = str(v).strip()
        return s != "" and s != "0"

    header = header[header["Activity_Name"].apply(_valid_name)].copy()
    header = header.drop_duplicates(subset=["AmtCol"])
    valid_amt_cols = list(header["AmtCol"].unique())

    # --- Data rows (RowIdx >= 8 within the range) ---
    data = grid.iloc[8:].copy()

    meta_names = [
        "SchKLineCode", "K3_Code", "SchKDetailCode", "Line", "KDescription",
        "Active", "FootnoteKey", "NaturalSign", "Part2", "PartX",
        "Part 2 Line", "Part X Line", "Total",
    ]
    rename_map = {i: meta_names[i] for i in range(META_COL_COUNT)}
    data = data.rename(columns=rename_map)

    # Filter to rows where K3_Code is numeric (skip totals / blank rows).
    data = data[data["K3_Code"].apply(_is_number)].copy()

    keep_meta = ["SchKLineCode", "K3_Code", "SchKDetailCode", "Line",
                 "KDescription", "Active", "NaturalSign", "Part2", "PartX",
                 "Part 2 Line", "Part X Line"]
    keep_cols = keep_meta + valid_amt_cols
    data = data[[c for c in keep_cols if c in data.columns]]

    # Unpivot the valid amount columns.
    melted = data.melt(
        id_vars=keep_meta, value_vars=valid_amt_cols,
        var_name="AmtCol", value_name="Amount",
    )

    # Join header info onto each amount row (1:1 on AmtCol).
    melted = melted.merge(
        header[["AmtCol", "Activity_Num", "Activity_Name",
                "Activity_Allocation", "DO_Flag", "Allocation_ID"]],
        on="AmtCol", how="inner",
    )

    # Types + drop rows that fail conversion (Table.RemoveRowsWithErrors).
    melted["K3_Code"] = _to_number(melted["K3_Code"])
    melted["NaturalSign"] = _to_number(melted["NaturalSign"])
    melted["Amount"] = _to_number(melted["Amount"])
    melted["Allocation_ID"] = _to_number(melted["Allocation_ID"])
    melted = melted.dropna(subset=["K3_Code", "NaturalSign", "Amount",
                                   "Allocation_ID"])
    melted["K3_Code"] = melted["K3_Code"].astype("Int64")

    # Non-zero amounts only.
    melted = melted[melted["Amount"] != 0].copy()

    melted["Source"] = "Sch K"
    melted["Amt"] = melted["NaturalSign"] * melted["Amount"]

    # SchKLine = "Line-KDescription".
    def _combine(line, desc):
        a = "" if line is None or (isinstance(line, float) and np.isnan(line)) else str(line)
        b = "" if desc is None or (isinstance(desc, float) and np.isnan(desc)) else str(desc)
        return f"{a}-{b}"

    melted["SchKLine"] = [
        _combine(l, d) for l, d in zip(melted["Line"], melted["KDescription"])
    ]

    drop_cols = ["SchKLineCode", "SchKDetailCode", "Part2", "PartX", "AmtCol",
                 "DO_Flag", "Allocation_ID", "Amount", "NaturalSign", "Line",
                 "KDescription"]
    melted = melted.drop(columns=[c for c in drop_cols if c in melted.columns])
    melted = melted.rename(columns={"Amt": "Schedule K"})
    return melted.reset_index(drop=True)


def _combine_tables(top: pd.DataFrame, bottom: pd.DataFrame) -> pd.DataFrame:
    """Mimic Table.Combine: union of columns, preserving column order."""
    cols = list(top.columns) + [c for c in bottom.columns if c not in top.columns]
    return pd.concat([top, bottom], ignore_index=True).reindex(columns=cols)


def build_diff(schk: pd.DataFrame, other: pd.DataFrame,
               part_col: str, exclude: pd.DataFrame) -> pd.DataFrame:
    """M queries: schKToK2Part2Diff / schKToK2PartXDiff."""
    combined = _combine_tables(schk, other)

    # Replace nulls in label / amount columns.
    for col in ("SchKLine", "Part 2 Line"):
        if col in combined.columns:
            combined[col] = combined[col].where(combined[col].notna(), "")
    for col in ("Schedule K", part_col):
        if col in combined.columns:
            combined[col] = combined[col].where(combined[col].notna(), 0)

    combined[part_col] = _to_number(combined[part_col]).fillna(0).astype("int64")

    # Full-outer join with the exclusion list on K3_Code = excludeCode, then
    # keep only rows that did NOT match (exclude is null).
    merged = combined.merge(
        exclude[["excludeCode", "exclude"]],
        left_on="K3_Code", right_on="excludeCode", how="outer",
    )
    merged = merged[merged["exclude"].isna()].copy()
    merged = merged.drop(columns=["exclude", "excludeCode"])

    # Activity_Num -> text, nulls -> "".
    def _txt(v):
        if v is None or (isinstance(v, float) and np.isnan(v)) or pd.isna(v):
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    merged["Activity_Num"] = merged["Activity_Num"].apply(_txt)
    return merged.reset_index(drop=True)


def build_gross_income(part2: pd.DataFrame, partx: pd.DataFrame) -> pd.DataFrame:
    """M query: grossIncome."""
    combined = _combine_tables(part2, partx)

    for col in ("Part 2", "Part X"):
        combined[col] = _to_number(combined.get(col)).fillna(0).astype("int64")

    def _classify(r):
        if r["Part 2"] > 0:
            return "X"
        if r["Part X"] > 0:
            return "X"
        return "Y"

    combined["grossIncome"] = combined.apply(_classify, axis=1)

    def _grouped(flag):
        sub = combined[combined["grossIncome"] == flag]
        g = (sub.groupby("K3_Code", as_index=False)
                .agg({"Part 2": "sum", "Part X": "sum"}))
        g["Difference"] = g["Part 2"] - g["Part X"]
        return g

    return pd.concat([_grouped("X"), _grouped("Y")], ignore_index=True)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def run(workbook_path: str | Path = DEFAULT_WORKBOOK) -> dict[str, pd.DataFrame]:
    """Build every query and return them keyed by name."""
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)

    ws_part2 = wb[RANGES["k2Part2"][0]]
    ws_partx = wb[RANGES["k2PartX"][0]]
    ws_excl = wb[RANGES["part2Exclude"][0]]
    ws_alloc = wb[RANGES["allocRec"][0]]

    part2 = build_part2(ws_part2)
    partx = build_partx(ws_partx)
    exclude_p2 = build_exclude(ws_excl, "part2Exclude")
    exclude_px = build_exclude(ws_excl, "partXExclude")
    schk = build_schk(ws_alloc)

    diff_p2 = build_diff(schk, part2, "Part 2", exclude_p2)
    diff_px = build_diff(schk, partx, "Part X", exclude_px)
    gross = build_gross_income(part2, partx)

    wb.close()
    return {
        "part2": part2,
        "partx": partx,
        "excludePart2Table": exclude_p2,
        "excludePartXTable": exclude_px,
        "schK": schk,
        "schKToK2Part2Diff": diff_p2,
        "schKToK2PartXDiff": diff_px,
        "grossIncome": gross,
    }


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Schedule K to K-2 reconciliation")
    parser.add_argument("workbook", nargs="?", default=DEFAULT_WORKBOOK,
                        help="Path to the .xlsm workbook")
    parser.add_argument("-o", "--output", default="recon_output.xlsx",
                        help="Output .xlsx file")
    args = parser.parse_args()

    results = run(args.workbook)

    with pd.ExcelWriter(args.output, engine="openpyxl") as xl:
        for name, df in results.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)

    print(f"Wrote {args.output}")
    for name, df in results.items():
        print(f"  {name:22s} {df.shape[0]:6d} rows x {df.shape[1]} cols")


if __name__ == "__main__":
    main()
