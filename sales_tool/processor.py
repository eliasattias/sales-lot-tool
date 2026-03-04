import re
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

# ===================== CONFIG =====================
# Default used only when running as a script without arguments.
DEFAULT_INPUT_FILE = Path("input.xlsx")

SHEET = 0
COL_LOT_TEXT = "Lot Numbers(QTY)"
COL_QTY_SHIPPED = "Qty Shipped"
COL_EXT_SALES = "Ext Sales"
COL_SALES_PRICE = "Sales Price"
COL_EXT_COST = "Ext Cost"
COL_WEIGHT = "Weight"
COL_FEE = "Distributor Channel Fee"
COL_MPN = "Manufacturer Part Number"
COL_DIVISION = "Division"  # only referenced if present

# Referenced only if present:
COL_BUYER = "Buyer"
COL_CATEGORY = "Category"

# Optional/expected semantic columns (reserved for future)
COL_DATE = "Date"
COL_PO_ENTRY = "PO entry date"
COL_CUST_PO = "Customer PO#"
COL_SHIP_ZIP = "Shipping Zip"

TOKEN_RE = re.compile(r"(?P<code>[A-Za-z0-9\-_\/]+)\s*\((?P<qty>-?\d+(?:\.\d+)?)\)")


# ===================== HELPERS =====================
def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[:\\/*?\[\]/]", "-", str(name))
    return cleaned[:31]


def parse_lots(text):
    if not isinstance(text, str) or not text.strip():
        return []
    return [
        {"LotNumber": m.group("code"), "LotQty": float(m.group("qty"))}
        for m in TOKEN_RE.finditer(text)
    ]


def coerce_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.strip()
        .str.replace(r"[\$,]", "", regex=True)
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True),
        errors="coerce",
    ).fillna(0.0)


# numeric coercions we want on outputs (apply if present)
NUMERIC_COLS = [
    "Qty Shipped",
    "Sales Price",
    "Ext Sales",
    "Ext Cost",
    "Weight",
    "Distributor Channel Fee",
    "LotQty Cases",
    "Unit",
    "Ext Sales $ (By Case)",
    "Total per lot",
]


def force_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in NUMERIC_COLS:
        if c in df.columns:
            df[c] = coerce_numeric(df[c])
    return df


def force_ext_sales_numeric(df: pd.DataFrame) -> pd.DataFrame:
    if COL_EXT_SALES in df.columns:
        df[COL_EXT_SALES] = coerce_numeric(df[COL_EXT_SALES])
    return df


# Remove generic Total/Grand Total rows from processed data before splitting
_TOTAL_PATTERN = r"^\s*(grand\s+)?totals?\s*$"


def drop_total_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    mask = df.astype(str).apply(
        lambda col: col.str.match(_TOTAL_PATTERN, case=False, na=False)
    ).any(axis=1)
    return df.loc[~mask].copy()


def drop_division_total(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or COL_DIVISION not in df.columns:
        return df
    return df[
        ~df[COL_DIVISION]
        .astype(str)
        .str.strip()
        .str.casefold()
        .eq("total")
    ].copy()


# ---------- TEXT / ROW-GUARDS ----------
def as_text(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()


def is_explicit_zero_text(s: pd.Series) -> pd.Series:
    return as_text(s).str.fullmatch(r"-?0+(?:\.0+)?", na=False)


def ensure_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col in df.columns:
        return df[col]
    return pd.Series([""] * len(df), index=df.index)


def any_nonempty(*cols: pd.Series) -> pd.Series:
    flags = [as_text(c).ne("") for c in cols]
    return pd.concat(flags, axis=1).any(axis=1)


def build_masks(df: pd.DataFrame):
    price_txt = as_text(ensure_series(df, COL_SALES_PRICE))
    ext_txt = as_text(ensure_series(df, COL_EXT_SALES))

    price_is_zero_txt = is_explicit_zero_text(price_txt)
    ext_is_zero_txt = is_explicit_zero_text(ext_txt)
    SAMPLE_PRICE_TEXTS = {"0.01", "0.001"}

    qty_txt = ensure_series(df, COL_QTY_SHIPPED)
    row_has_content = any_nonempty(qty_txt, ext_txt, price_txt)

    qty_num = coerce_numeric(qty_txt)
    ext_num = coerce_numeric(ext_txt)

    mask_credit = qty_num < 0
    mask_samples_only = (~mask_credit) & row_has_content & (
        price_txt.isin(SAMPLE_PRICE_TEXTS)
        | price_is_zero_txt
        | ext_is_zero_txt
    )
    mask_sales = (
        (~mask_credit)
        & row_has_content
        & (~mask_samples_only)
        & (ext_num != 0.01)
    )

    return {"credit": mask_credit, "samples": mask_samples_only, "sales": mask_sales}


# ---------- Expand/compute ----------
def expand_and_compute(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Explode by lot and compute:
      - Qty Shipped: allocated per lot (in cases; sums back to original)
      - Ext Sales: allocated per lot (sums back to original)
      - Ext Cost: allocated per lot (sums back to original)
      - Sales Price: Ext Sales / Qty Shipped (price per case)
      - Ext Sales $ (By Case): same as allocated Ext Sales
      - LotQty Cases & Unit (units-per-case mapping: 6 or 14)
      - Total per lot: Ext Sales per lot / units_per_case (no rounding)
      - Distributor Channel Fee: allocated per lot based on units (Unit)
    Filters to only 'real' rows first (any of Division/Buyer/Category non-blank if present).
    """
    if df_in.empty:
        return df_in.copy()

    df = df_in.copy().reset_index(drop=True)

    # Real-row guard
    division_txt = as_text(ensure_series(df, COL_DIVISION))
    buyer_txt = as_text(ensure_series(df, COL_BUYER))
    category_txt = as_text(ensure_series(df, COL_CATEGORY))
    row_is_real = any_nonempty(division_txt, buyer_txt, category_txt)
    df = df.loc[row_is_real].reset_index(drop=True)
    if df.empty:
        return df

    df["_RowID"] = range(len(df))

    if COL_MPN in df.columns:
        df[COL_MPN] = df[COL_MPN].astype(str).str.strip().str.upper()

    # ---------- explode lots ----------
    lots_series = df[COL_LOT_TEXT].apply(parse_lots)
    qty_shipped_num = coerce_numeric(df[COL_QTY_SHIPPED])
    ensured_lots = []

    for i, lots_row in enumerate(lots_series):
        raw_text = (
            ""
            if pd.isna(df.iloc[i][COL_LOT_TEXT])
            else str(df.iloc[i][COL_LOT_TEXT]).strip()
        )
        fq = float(qty_shipped_num.iloc[i]) if qty_shipped_num.iloc[i] else 1.0
        if lots_row:
            ensured_lots.append(lots_row)
        else:
            m = re.search(r"[A-Za-z0-9\-_\/]+", raw_text)
            lot_label = m.group(0) if m else raw_text
            ensured_lots.append([{"LotNumber": lot_label, "LotQty": fq}])

    tmp = df.copy()
    tmp["_Lots"] = ensured_lots
    tmp = tmp.explode("_Lots", ignore_index=True)
    tmp["LotNumber"] = tmp["_Lots"].apply(lambda d: d.get("LotNumber"))
    tmp["LotQty"] = tmp["_Lots"].apply(lambda d: d.get("LotQty"))
    tmp.drop(columns=["_Lots"], inplace=True)
    tmp["LotNumber"] = tmp["LotNumber"].astype(str).str.split("-").str[0]

    # ---------- proportional shares ----------
    total_qty = tmp.groupby("_RowID")["LotQty"].transform("sum").replace(0, pd.NA)
    tmp["LotShare"] = (tmp["LotQty"] / total_qty).fillna(0.0)

    # Original numeric values (same on each exploded row for a given _RowID)
    tmp["_QtyOrig"] = coerce_numeric(tmp[COL_QTY_SHIPPED])
    tmp["_ExtSalesOrig"] = coerce_numeric(tmp[COL_EXT_SALES])
    if COL_EXT_COST in tmp.columns:
        tmp["_ExtCostOrig"] = coerce_numeric(tmp[COL_EXT_COST])
    else:
        tmp["_ExtCostOrig"] = 0.0

    if COL_WEIGHT in tmp.columns:
        tmp["_WeightOrig"] = coerce_numeric(tmp[COL_WEIGHT])
    else:
        tmp["_WeightOrig"] = 0.0

    # Allocate Qty Shipped (in cases) directly by lot qty (this will be integer in your data)
    tmp[COL_QTY_SHIPPED] = tmp["LotQty"]

    # Allocate Ext Sales
    tmp["Ext Sales Alloc"] = tmp["_ExtSalesOrig"] * tmp["LotShare"]

    def adjust_group_sales(g: pd.DataFrame) -> pd.DataFrame:
        src_total = float(g["_ExtSalesOrig"].iloc[0])
        lot_sum = float(g["Ext Sales Alloc"].sum())
        if lot_sum == 0 or src_total == 0:
            return g
        scale = src_total / lot_sum
        g["Ext Sales Alloc"] = (g["Ext Sales Alloc"] * scale).round(2)
        diff = round(src_total - g["Ext Sales Alloc"].sum(), 2)
        if abs(diff) >= 0.01:
            idx_max = g["Ext Sales Alloc"].idxmax()
            g.loc[idx_max, "Ext Sales Alloc"] += diff
        return g

    tmp = tmp.groupby("_RowID", group_keys=False).apply(adjust_group_sales)

    # Allocate Ext Cost similarly
    tmp["Ext Cost Alloc"] = tmp["_ExtCostOrig"] * tmp["LotShare"]

    if COL_EXT_COST in df_in.columns:

        def adjust_group_cost(g: pd.DataFrame) -> pd.DataFrame:
            src_total = float(g["_ExtCostOrig"].iloc[0])
            lot_sum = float(g["Ext Cost Alloc"].sum())
            if lot_sum == 0 or src_total == 0:
                return g
            scale = src_total / lot_sum
            g["Ext Cost Alloc"] = (g["Ext Cost Alloc"] * scale).round(2)
            diff = round(src_total - g["Ext Cost Alloc"].sum(), 2)
            if abs(diff) >= 0.01:
                idx_max = g["Ext Cost Alloc"].idxmax()
                g.loc[idx_max, "Ext Cost Alloc"] += diff
            return g

        tmp = tmp.groupby("_RowID", group_keys=False).apply(adjust_group_cost)

    # ----- Allocate Weight proportionally to lots -----
    if COL_WEIGHT in df_in.columns:
        tmp["Weight Alloc"] = tmp["_WeightOrig"] * tmp["LotShare"]

        def adjust_group_weight(g: pd.DataFrame) -> pd.DataFrame:
            src_total = float(g["_WeightOrig"].iloc[0])
            lot_sum = float(g["Weight Alloc"].sum())
            if lot_sum == 0 or src_total == 0:
                return g
            # scale + round so sum matches original
            scale = src_total / lot_sum
            g["Weight Alloc"] = (g["Weight Alloc"] * scale).round(2)
            diff = round(src_total - g["Weight Alloc"].sum(), 2)
            if abs(diff) >= 0.01:
                idx_max = g["Weight Alloc"].idxmax()
                g.loc[idx_max, "Weight Alloc"] += diff
            return g

        tmp = tmp.groupby("_RowID", group_keys=False).apply(adjust_group_weight)

        # write back into the real Weight column
        tmp[COL_WEIGHT] = tmp["Weight Alloc"]

    # Write allocated values back to the standard columns
    tmp["Ext Sales $ (By Case)"] = tmp["Ext Sales Alloc"]
    tmp[COL_EXT_SALES] = tmp["Ext Sales Alloc"]
    if COL_EXT_COST in df_in.columns:
        tmp[COL_EXT_COST] = tmp["Ext Cost Alloc"]

    # ---------- Cases & Units ----------
    tmp.rename(columns={"LotQty": "LotQty Cases"}, inplace=True)
    if COL_MPN in tmp.columns:
        tmp[COL_MPN] = tmp[COL_MPN].astype(str).str.strip().str.upper()
        units_per_case = np.select(
            [
                tmp[COL_MPN].isin(["V7871800002", "7871800002"]),  # 2000 mL → 6
                tmp[COL_MPN].isin(["V7871800000", "7871800000"]),  # 1000 mL → 14
            ],
            [6, 14],
            default=np.nan,
        )
    else:
        units_per_case = np.nan

    unit_values = units_per_case * pd.to_numeric(
        tmp["LotQty Cases"], errors="coerce"
    )
    tmp.insert(tmp.columns.get_loc("LotQty Cases") + 1, "Unit", unit_values)

    # Total per lot = Ext Sales per lot / units_per_case (no rounding)
    by_case_num = pd.to_numeric(tmp["Ext Sales $ (By Case)"], errors="coerce")
    tmp["Total per lot"] = np.where(
        tmp["Unit"] > 0, tmp["Ext Sales $ (By Case)"] / tmp["Unit"], np.nan
    )

    # Recompute Sales Price as Ext Sales / Qty Shipped (price per case)
    tmp[COL_QTY_SHIPPED] = pd.to_numeric(tmp[COL_QTY_SHIPPED], errors="coerce")
    tmp[COL_SALES_PRICE] = np.where(
        tmp[COL_QTY_SHIPPED] != 0,
        tmp[COL_EXT_SALES] / tmp[COL_QTY_SHIPPED],
        0.0,
    )

    # ---------- Allocate Distributor Channel Fee by units ----------
    if COL_FEE in tmp.columns:
        fee_raw = coerce_numeric(df_in[COL_FEE]).reset_index(drop=True)
        tmp["OriginalFee"] = tmp["_RowID"].map(fee_raw)

        total_units = tmp.groupby("_RowID")["Unit"].transform("sum")
        total_units = total_units.replace(0, np.nan)

        unit_share = tmp["Unit"] / total_units
        fee_share = unit_share.fillna(tmp["LotShare"])
        tmp[COL_FEE] = tmp["OriginalFee"] * fee_share

        def fix_fee_group(g: pd.DataFrame) -> pd.DataFrame:
            raw_total = float(g["OriginalFee"].iloc[0])
            if raw_total == 0:
                return g
            g[COL_FEE] = g[COL_FEE].round(2)
            diff = round(raw_total - g[COL_FEE].sum(), 2)
            if abs(diff) >= 0.01:
                idx = g[COL_FEE].idxmax()
                g.loc[idx, COL_FEE] += diff
            return g

        tmp = tmp.groupby("_RowID", group_keys=False).apply(fix_fee_group)
        tmp.drop(columns=["OriginalFee"], inplace=True)

    # cleanup helpers
    tmp.drop(
        columns=[
            "LotShare",
            "_RowID",
            "_QtyOrig",
            "_ExtSalesOrig",
            "_ExtCostOrig",
            "_WeightOrig",
            "Ext Sales Alloc",
            "Ext Cost Alloc",
            "Weight Alloc",
        ],
        inplace=True,
        errors="ignore",
    )

    return tmp


# ---------- Build Summary Sales (row-level per unique lot) ----------
def build_summary_sales(real_sales: pd.DataFrame) -> pd.DataFrame:
    """
    One row per unique (MPN, LotNumber) with sums.
    Includes both 'Ext Sales $ (By Case)' and 'Total per lot'.
    """
    if real_sales.empty:
        return real_sales.copy()

    needed = [
        COL_MPN,
        "LotNumber",
        "LotQty Cases",
        "Unit",
        "Ext Sales $ (By Case)",
        "Total per lot",
        COL_FEE,
    ]
    for c in needed:
        if c not in real_sales.columns:
            real_sales[c] = 0.0 if c not in [COL_MPN, "LotNumber"] else ""

    grp = (
        real_sales.groupby([COL_MPN, "LotNumber"], dropna=False, as_index=False)
        .agg(
            {
                "LotQty Cases": "sum",
                "Unit": "sum",
                "Ext Sales $ (By Case)": "sum",
                "Total per lot": "sum",
                COL_FEE: "sum",
            }
        )
        .sort_values([COL_MPN, "LotNumber"], kind="stable")
        .reset_index(drop=True)
    )

    return grp


# ---------- Append a Grand Total row that sums all numeric columns ----------
def append_totals_row(df: pd.DataFrame, label_value: str = "Grand Total") -> pd.DataFrame:
    if df.empty:
        return df

    # Identify numeric columns and sum all of them INCLUDING Sales Price
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()

    totals = {c: pd.to_numeric(df[c], errors="coerce").sum() for c in num_cols}

    # Put label on Division (if exists) otherwise first column
    row = {**totals}
    if COL_DIVISION in df.columns:
        row[COL_DIVISION] = label_value
    else:
        first_col = df.columns[0]
        row[first_col] = label_value

    # ensure all columns exist in the totals row
    for c in df.columns:
        if c not in row:
            row[c] = ""

    return pd.concat([df, pd.DataFrame([row], columns=df.columns)], ignore_index=True)


# ===================== CORE PROCESSOR =====================
def process_workbook(input_file: Path) -> Path:
    """
    Process a monthly sales/inventory Excel workbook in-place.

    - Reads the original sheet
    - Explodes and allocates lots
    - Builds derived tabs (Real Sales, Samples, Credit Memo, Summary Sales, etc.)
    - Rewrites/creates the derived sheets in the same workbook.
    """
    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    # Read original as strings
    df_raw = pd.read_excel(input_file, sheet_name=SHEET, dtype=str)

    # Processed copy (drop generic totals)
    df_proc = drop_total_rows(df_raw.copy())

    # Required columns present?
    required_cols = [COL_LOT_TEXT, COL_QTY_SHIPPED, COL_EXT_SALES, COL_SALES_PRICE]
    missing = [c for c in required_cols if c not in df_proc.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    # Masks on ORIGINAL (unexpanded sales-only tab)
    masks_raw = build_masks(df_proc)
    original_sales_no_samples = df_proc.loc[masks_raw["sales"]].copy()
    original_sales_no_samples = force_ext_sales_numeric(original_sales_no_samples)

    # Masks on PROCESSED (derived tabs)
    masks = build_masks(df_proc)
    credit_memo_raw = df_proc.loc[masks["credit"]].copy()
    samples_raw = df_proc.loc[masks["samples"]].copy()
    sales_raw = df_proc.loc[masks["sales"]].copy()

    # Expand/compute
    real_sales = expand_and_compute(sales_raw)
    samples = expand_and_compute(samples_raw)
    credit_memo = expand_and_compute(credit_memo_raw)

    # Coerce numerics on all outputs
    real_sales = force_numeric(real_sales)
    samples = force_numeric(samples)
    credit_memo = force_numeric(credit_memo)
    original_sales_no_samples = force_numeric(original_sales_no_samples)

    # Remove Division == Total (if any)
    real_sales = drop_division_total(real_sales)
    samples = drop_division_total(samples)
    credit_memo = drop_division_total(credit_memo)

    # Build Summary Sales from detail (no total row yet)
    summary_sales = build_summary_sales(real_sales)

    # Append Grand Total row on derived sheets
    original_sales_no_samples = append_totals_row(original_sales_no_samples)
    real_sales = append_totals_row(real_sales)
    samples = append_totals_row(samples)
    credit_memo = append_totals_row(credit_memo)
    summary_sales = append_totals_row(summary_sales)

    # Write tabs and freeze header row
    with pd.ExcelWriter(
        input_file, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        # Original Sales (No Samples/Credits)
        sheet_name_os = safe_sheet_name("Original Sales (No Samples/Credits)")
        original_sales_no_samples.to_excel(
            writer, index=False, sheet_name=sheet_name_os
        )
        ws = writer.sheets[sheet_name_os]
        ws.freeze_panes = "A2"

        # Real Sales
        real_sales.to_excel(writer, index=False, sheet_name="Real Sales")
        ws = writer.sheets["Real Sales"]
        ws.freeze_panes = "A2"

        # Samples
        samples.to_excel(writer, index=False, sheet_name="Samples")
        ws = writer.sheets["Samples"]
        ws.freeze_panes = "A2"

        # Credit Memo
        credit_memo.to_excel(writer, index=False, sheet_name="Credit Memo")
        ws = writer.sheets["Credit Memo"]
        ws.freeze_panes = "A2"

        # Summary Sales
        summary_sales.to_excel(writer, index=False, sheet_name="Summary Sales")
        ws = writer.sheets["Summary Sales"]
        ws.freeze_panes = "A2"

    print(f"✅ Updated workbook: {input_file.name}")
    print(
        f"{safe_sheet_name('Original Sales (No Samples/Credits)')}: {len(original_sales_no_samples)} | "
        f"Real Sales: {len(real_sales)} | "
        f"Samples: {len(samples)} | "
        f"Credit Memo: {len(credit_memo)} | "
        f"Summary Sales: {len(summary_sales)}"
    )

    return input_file


def main(argv: Optional[list[str]] = None) -> None:
    """
    CLI entrypoint:

    python -m sales_tool.processor path\\to\\file.xlsx
    """
    import argparse

    parser = argparse.ArgumentParser(description="Process a sales/inventory Excel file.")
    parser.add_argument(
        "file",
        nargs="?",
        default=None,
        help=f"Path to Excel workbook (default: {DEFAULT_INPUT_FILE})",
    )
    args = parser.parse_args(argv)

    path = Path(args.file) if args.file is not None else DEFAULT_INPUT_FILE
    process_workbook(path)


if __name__ == "__main__":
    main()

