import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------

EXCEL_EPOCH = date(1899, 12, 30)

# Output column order — mirrors the manual MILL_PLAN sheet header row.
PLAN_COLUMNS = [
    "Date", "Batch", "Planning/ SO No.", "Thick", "Width", "Weight", "RT",
    "Customer", "Prod.Code", "Quality Code", "TDC No", "Plant", "Storage Loc",
    "Planning Remark", "Current Work Center", "NEXT Work Center", "Route",
    "Finish", "Age",
]

# Column widths (character units) — matches reference file layout.
COLUMN_WIDTHS = {
    "A": 8,  "B": 10, "C": 14, "D": 6,  "E": 6,  "F": 8,  "G": 5,
    "H": 14, "I": 8,  "J": 8,  "K": 6,  "L": 5,  "M": 8,  "N": 20,
    "O": 16, "P": 14, "Q": 30, "R": 8,  "S": 5,
}

# Section keys — also defines the fixed display order on every sheet.
SECTION_ORDER = [
    "CRCA_FINISH_CRM04",
    "HT_FINISH_CRM04",
    "FIRST_ROLLING_CRM06",
    "RR_ROLLING_CRM06",
    "SKIN_PASS_SUPER_BRIGHT_CRM04",
    "SKIN_PASS_CHROME_CRM04",
    "TUBE_FH",
    "SKIN_PASS_HEAVY_MATT_CRM06",
]

# Header text + fill colour per section.
SECTION_META = {
    "CRCA_FINISH_CRM04": (
        "CRCA FINISH ON BRIGHT ROLLS --------------- AT CRM-04 "
        "----------- APPLY R.P.OIL",
        "DCE6F1",  # light blue
    ),
    "HT_FINISH_CRM04": (
        "H&T FINISH ON BRIGHT ROLLS --------------- AT CRM-04 "
        "------------ DO NOT APPLY R.P.OIL",
        "DCE6F1",
    ),
    "FIRST_ROLLING_CRM06": (
        "1ST ROLLING ON LIGHT MATT ROLLS --------------- AT CRM-06",
        "EBF1DE",  # light green
    ),
    "RR_ROLLING_CRM06": (
        "R/R ROLLING ON LIGHT MATT ROLLS --------------- AT CRM-06",
        "EBF1DE",
    ),
    "SKIN_PASS_SUPER_BRIGHT_CRM04": (
        "SKIN-PASS ON SUPER BRIGHT ROLLS --------------- AT CRM-04 "
        "----------- APPLY R.P.OIL",
        "FFFF99",  # light yellow
    ),
    "SKIN_PASS_CHROME_CRM04": (
        "SKIN-PASS ON CHROMEPLATED ROLLS --------------- AT CRM-04 "
        "----------- APPLY R.P.OIL",
        "FFFF99",
    ),
    "TUBE_FH": (
        "TUBE FH ON BRIGHT ROLLS --------------- AT CRM-04/06 "
        "----------- APPLY R.P.OIL",
        "FCE4D6",  # light orange
    ),
    "SKIN_PASS_HEAVY_MATT_CRM06": (
        "SKIN-PASS ON HEAVY MATT ROLLS --------------- AT CRM-06 "
        "----------- APPLY R.P.OIL",
        "FFFF99",
    ),
}

# Customer name abbreviations (extend as needed from planner conventions).
CUSTOMER_ABBREV = {
    "L.G BALAKRISHNAN": "LG BALA",
    "SAHIBABAD TUBE PLANT": "TUBE",
    "BIJOY TRADING": "BIJOY",
    "SPECIAL STEEL": "SPECIAL STEEL",
    "BANDSAW STRIP": "BANDSAW",
    "ANCHOR": "ANCHOR",
}

# Display style for known process routes (pass-through for the rest).
ROUTE_DISPLAY = {
    "M->RW->B->S->QA->R->QA->PACK": "M->RW->B->S->QA->R->QA->PACK",
    "M->B->M->R->QA->PACK":         "M->B->M->R->QA->PACK",
}

# SAP Work Center membership lookup — most reliable mill assignment signal.
CRM04_WORK_CENTERS = {"SNCRMM04", "SNCRS13", "SNCRS14", "SNCRS11", "SNCRS10", "SWCRS1"}
CRM06_WORK_CENTERS = {"SNCRMM06", "SNCRS15", "SNRWL06", "SNANN02"}


# ---------------------------------------------------------------------------
# DATA LOADING & FILTERING
# ---------------------------------------------------------------------------

def load_wip(filepath: str) -> pd.DataFrame:
    """Load Sheet1 of the WIP file, normalise text columns."""
    df = pd.read_excel(filepath, sheet_name="Sheet1")

    # Strip whitespace from text columns (customer names, codes, etc.).
    # Explicit dtype list keeps us forward-compatible with pandas 3.x.
    text_dtypes = df.select_dtypes(include=["object", "string"]).columns
    for col in text_dtypes:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": "", "None": ""})

    # Coerce numerics safely.
    for col in ("Actual Thick", "Actual Width", "Input Coil Weight",
                "Plan Rolling Thick 1", "Cust Width", "Cust Thick",
                "Plan Weight", "Balance Coil Weight",
                "Coil Age(# Days)", "Stage Age(# Days)"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def filter_rolling_coils(df: pd.DataFrame, plan_date: date) -> tuple[pd.DataFrame, dict]:
    """
    Apply the eligibility filters. Returns (eligible_df, exclusion_counts).
    The plan_date currently scopes nothing (we work off the live WIP pool),
    but is accepted for forward compatibility.
    """
    counts = {"low_weight": 0, "fg_or_pack": 0, "rt_zero_or_hold": 0,
              "wrong_plant": 0, "not_rolling": 0}

    mask_plant = df.get("Production Plant", "").astype(str).str.contains("760")
    counts["wrong_plant"] = (~mask_plant).sum()
    df = df[mask_plant]

    next_stage = df["Next Stage"].astype(str).str.upper()
    last_stage = df.get("Last Production Stage", "").astype(str).str.upper()
    current_stage = df["Current Stage"].astype(str).str.upper()

    # Must be at rolling mill (current OR last completed stage).
    mask_rolling = (current_stage.str.contains("ROLLING MILL") |
                    last_stage.str.contains("ROLLING MILL"))
    counts["not_rolling"] = (~mask_rolling).sum()
    df = df[mask_rolling]
    next_stage = next_stage.loc[df.index]

    # Exclude finished / palletisation.
    mask_done = next_stage.str.contains("11-FG") | next_stage.str.contains("PALLETIZ")
    counts["fg_or_pack"] = mask_done.sum()
    df = df[~mask_done]

    # Exclude tiny remnants.
    mask_low = df["Input Coil Weight"] < 0.5
    counts["low_weight"] = mask_low.sum()
    df = df[~mask_low]

    # Exclude RT=0 / unassigned.
    rt = df["Plan Rolling Thick 1"].fillna(0)
    mask_rt_zero = rt <= 0
    counts["rt_zero_or_hold"] = mask_rt_zero.sum()
    df = df[~mask_rt_zero]

    return df.copy(), counts


# ---------------------------------------------------------------------------
# SECTION ASSIGNMENT
# ---------------------------------------------------------------------------

def _is_crm04(wc: str) -> bool:
    return any(code in wc for code in CRM04_WORK_CENTERS) or "04" in wc

def _is_crm06(wc: str) -> bool:
    return any(code in wc for code in CRM06_WORK_CENTERS) or "06" in wc


def assign_section(row: pd.Series) -> str:
    """
    Map one WIP row to a plan section key.

    Order matters — narrower / more specific rules are tested first so that
    a Tube FH coil isn't accidentally pulled into the generic CRCA bucket.
    """
    wc = str(row.get("Work Center", "")).upper()
    quality = str(row.get("Actual Quality", "")).upper()
    prod_code = str(row.get("Product Code", "")).upper()
    route = str(row.get("Process Route", "")).upper()
    storage = str(row.get("Storage Location", "")).upper()
    next_stage = str(row.get("Next Stage", "")).upper()
    tdc = str(row.get("Cust TDC", "")).upper()
    rt = row.get("Plan Rolling Thick 1") or 0
    try:
        rt = float(rt)
    except (TypeError, ValueError):
        rt = 0.0

    # 1. TUBE FH — most specific (quality + product combo)
    if quality == "TATFHC" and prod_code == "C09":
        return "TUBE_FH"

    # 2. SKIN PASS HEAVY MATT (CRM06)
    if quality == "TATBID" or "S-SPM" in next_stage:
        # Confirm CRM06 context where possible
        if _is_crm06(wc) or storage in ("RNM6", "RC01") or "S-SPM" in next_stage:
            return "SKIN_PASS_HEAVY_MATT_CRM06"

    # 3. SKIN PASS routes at CRM-04 — distinguish chrome vs super-bright by TDC.
    skin_pass_route = ("S->QA" in route or "R->QA->PACK" in route)
    if prod_code == "C01" and skin_pass_route and storage == "R034":
        if "D012" in tdc and rt < 2.0:
            return "SKIN_PASS_CHROME_CRM04"
        if quality in ("TATXXD", "TATD12") or "T012" in tdc:
            return "SKIN_PASS_SUPER_BRIGHT_CRM04"

    # 4. CRCA FINISH (CRM04) — bright rolls with RP oil
    if prod_code == "C09" and ("TSBF" in quality or "TSBH" in quality):
        if _is_crm04(wc) or storage in ("R034", "R032", "R116"):
            return "CRCA_FINISH_CRM04"

    # 5. 1ST ROLLING CRM06 — heavy gauge first pass headed to anneal
    if (_is_crm06(wc) or storage == "R116") and "B-ANNEALING" in next_stage:
        return "FIRST_ROLLING_CRM06"

    # 6. R/R ROLLING CRM06 — rewinding after rolling
    if (_is_crm06(wc) or storage == "RNM6") and "RW-REWINDING" in next_stage:
        return "RR_ROLLING_CRM06"

    # 7. H&T FINISH CRM04 — default B28 bucket
    if prod_code == "B28" and _is_crm04(wc):
        return "HT_FINISH_CRM04"

    # 8. Fallback: try to keep the coil on its mill family
    if prod_code == "B28":
        return "HT_FINISH_CRM04"
    if _is_crm06(wc):
        return "FIRST_ROLLING_CRM06"

    return "OTHER"


def assign_sections(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["section"] = df.apply(assign_section, axis=1)
    return df


def sort_within_section(section_df: pd.DataFrame) -> pd.DataFrame:
    """Oldest first, then heaviest, then thickest — matches planner intuition."""
    return section_df.sort_values(
        by=["Coil Age(# Days)", "Input Coil Weight", "Actual Thick"],
        ascending=[False, False, False],
        kind="mergesort",
    )


# ---------------------------------------------------------------------------
# ROW BUILDING — translate one WIP row into a plan row
# ---------------------------------------------------------------------------

def _abbreviate_customer(name: str) -> str:
    if not name:
        return ""
    upper = name.upper()
    for key, short in CUSTOMER_ABBREV.items():
        if key in upper:
            return short
    return name[:12]


def _route_display(route: str) -> str:
    return ROUTE_DISPLAY.get(route, route)


def _plant_short(plant: str) -> str:
    """Strip leading zero — '0760' -> '760'."""
    s = str(plant).lstrip("0")
    return s or "0"


def _excel_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def build_plan_row(wip_row: pd.Series, plan_date: date, age_offset: int = 0) -> list:
    """Project one WIP row into the plan-sheet column order."""
    serial = _excel_serial(plan_date)
    age = wip_row.get("Coil Age(# Days)")
    try:
        age_int = int(age) + age_offset if pd.notna(age) else age_offset
    except (TypeError, ValueError):
        age_int = age_offset

    finish = str(wip_row.get("Surface Finish") or "").strip()
    if not finish:
        finish = str(wip_row.get("Edge") or "").strip()

    return [
        serial,                                                # Date (Excel serial)
        str(wip_row.get("Coil Number", "")),               # Batch (now using Coil Number)
        str(wip_row.get("SO No", "")),                        # Planning/SO No.
        float(wip_row.get("Actual Thick") or 0),
        int(wip_row.get("Actual Width") or 0),
        float(wip_row.get("Input Coil Weight") or 0),
        float(wip_row.get("Plan Rolling Thick 1") or 0),
        _abbreviate_customer(str(wip_row.get("Customer Desc", ""))),
        str(wip_row.get("Product Code", "")),
        str(wip_row.get("Actual Quality", "")),
        str(wip_row.get("Cust TDC", "")),
        _plant_short(wip_row.get("Production Plant", "")),
        str(wip_row.get("Storage Location", "")),
        str(wip_row.get("Planning Remark", "")),
        str(wip_row.get("Current Stage", "")),
        str(wip_row.get("Next Stage", "")),
        _route_display(str(wip_row.get("Process Route", ""))),
        finish,
        age_int,
    ]


# ---------------------------------------------------------------------------
# EXCEL WRITING
# ---------------------------------------------------------------------------

THIN = Side(border_style="thin", color="999999")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="000000")
SECTION_FONT = Font(name="Calibri", size=10, bold=True)
DATA_FONT = Font(name="Calibri", size=9)
SUBTOTAL_FONT = Font(name="Calibri", size=10, bold=True)
TITLE_FONT = Font(name="Calibri", size=12, bold=True)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

HEADER_FILL = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
TITLE_FILL = PatternFill("solid", start_color="305496", end_color="305496")


def _apply_column_widths(ws: Worksheet) -> None:
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width


def _write_sheet_title(ws: Worksheet, plan_date: date, row: int) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=19)
    cell = ws.cell(row=row, column=1)
    cell.value = (
        f"PLANNING FOR MILL-------------- AS ON  "
        f"{plan_date.strftime('%d-%m-%Y')}"
    )
    cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    cell.fill = TITLE_FILL
    cell.alignment = CENTER
    ws.row_dimensions[row].height = 22
    return row + 1


def _write_header_row(ws: Worksheet, row: int) -> int:
    for col_idx, name in enumerate(PLAN_COLUMNS, start=1):
        cell = ws.cell(row=row, column=col_idx, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = CELL_BORDER
    ws.row_dimensions[row].height = 30
    return row + 1


def _write_section_header(ws: Worksheet, label: str, fill_hex: str,
                          row: int) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=19)
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = SECTION_FONT
    cell.fill = PatternFill("solid", start_color=fill_hex, end_color=fill_hex)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = CELL_BORDER
    ws.row_dimensions[row].height = 18
    return row + 1


def _write_data_row(ws: Worksheet, values: list, row: int) -> int:
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col_idx, value=val)
        cell.font = DATA_FONT
        cell.border = CELL_BORDER
        # Numeric alignment + formats by column position.
        if col_idx in (1, 2):                          # Date, Batch
            cell.number_format = "0"
            cell.alignment = CENTER
        elif col_idx in (4, 7):                        # Thick, RT
            cell.number_format = "0.000"
            cell.alignment = RIGHT
        elif col_idx == 5:                             # Width
            cell.number_format = "0"
            cell.alignment = RIGHT
        elif col_idx == 6:                             # Weight
            cell.number_format = "0.000"
            cell.alignment = RIGHT
        elif col_idx == 19:                            # Age
            cell.number_format = "0"
            cell.alignment = CENTER
        else:
            cell.alignment = LEFT
    return row + 1


def _write_subtotal_row(ws: Worksheet, first_data_row: int,
                       last_data_row: int, row: int) -> int:
    if last_data_row >= first_data_row:
        # SUM formula keeps the model live for planners to edit.
        ws.cell(row=row, column=6).value = (
            f"=SUM(F{first_data_row}:F{last_data_row})"
        )
    else:
        ws.cell(row=row, column=6).value = 0
    for col_idx in range(1, 20):
        cell = ws.cell(row=row, column=col_idx)
        cell.font = SUBTOTAL_FONT
        cell.border = CELL_BORDER
        cell.fill = PatternFill("solid", start_color="F2F2F2",
                                end_color="F2F2F2")
        if col_idx == 6:
            cell.number_format = "0.000"
            cell.alignment = RIGHT
        elif col_idx == 3:
            cell.value = "SUB-TOTAL"
            cell.alignment = RIGHT
    return row + 1


def _write_grand_total(ws: Worksheet, subtotal_rows: list[int],
                       row: int) -> int:
    if subtotal_rows:
        formula = "=" + "+".join(f"F{r}" for r in subtotal_rows)
    else:
        formula = "=0"
    for col_idx in range(1, 20):
        cell = ws.cell(row=row, column=col_idx)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="305496",
                                end_color="305496")
        cell.border = CELL_BORDER
    ws.cell(row=row, column=3, value="GRAND TOTAL").alignment = RIGHT
    gt = ws.cell(row=row, column=6, value=formula)
    gt.number_format = "0.000"
    gt.alignment = RIGHT
    return row + 1


# ---------------------------------------------------------------------------
# SHEET BUILDER
# ---------------------------------------------------------------------------

@dataclass
class SectionResult:
    key: str
    coils: int
    weight: float


def write_day_sheet(wb: Workbook, df: pd.DataFrame, plan_date: date,
                    age_offset: int = 0) -> list[SectionResult]:
    """Write one day's sheet. Returns per-section coil/weight stats."""
    sheet_name = plan_date.strftime("%d-%m-%Y")
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    _apply_column_widths(ws)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    row = _write_sheet_title(ws, plan_date, 1)
    row = _write_header_row(ws, row)

    df = assign_sections(df)
    results: list[SectionResult] = []
    subtotal_rows: list[int] = []

    for section_key in SECTION_ORDER:
        label, fill = SECTION_META[section_key]
        section_df = df[df["section"] == section_key]
        section_df = sort_within_section(section_df)

        row = _write_section_header(ws, label, fill, row)

        first_data_row = row
        for _, wip_row in section_df.iterrows():
            plan_row_values = build_plan_row(wip_row, plan_date, age_offset)
            row = _write_data_row(ws, plan_row_values, row)
        last_data_row = row - 1

        # Subtotal row even if section is empty (keeps layout consistent).
        subtotal_rows.append(row)
        row = _write_subtotal_row(ws, first_data_row, last_data_row, row)

        results.append(SectionResult(
            key=section_key,
            coils=len(section_df),
            weight=float(section_df["Input Coil Weight"].sum() or 0),
        ))

    _write_grand_total(ws, subtotal_rows, row)
    return results


# ---------------------------------------------------------------------------
# VALIDATION PRINTOUT
# ---------------------------------------------------------------------------

SECTION_LABELS = {
    "CRCA_FINISH_CRM04":            "CRCA Finish (CRM04)",
    "HT_FINISH_CRM04":              "H&T Finish  (CRM04)",
    "FIRST_ROLLING_CRM06":          "1st Rolling (CRM06)",
    "RR_ROLLING_CRM06":             "R/R Rolling (CRM06)",
    "SKIN_PASS_SUPER_BRIGHT_CRM04": "Skin Pass Super Bright (CRM04)",
    "SKIN_PASS_CHROME_CRM04":       "Skin Pass Chrome       (CRM04)",
    "TUBE_FH":                      "Tube FH     (CRM04/06)",
    "SKIN_PASS_HEAVY_MATT_CRM06":   "Skin Pass Heavy Matt   (CRM06)",
}


def print_validation(plan_date: date, results: list[SectionResult],
                     exclusions: dict, other_count: int) -> None:
    total_coils = sum(r.coils for r in results)
    total_weight = sum(r.weight for r in results)
    print(f"\n=== MILL PLAN VALIDATION: {plan_date.strftime('%d-%m-%Y')} ===")
    print(f"Total coils planned:         {total_coils}")
    print(f"Total weight (MT):           {total_weight:,.2f}")
    print()
    for r in results:
        print(f"  {SECTION_LABELS[r.key]:<34}  "
              f"{r.coils:>4} coils, {r.weight:>9.2f} MT")
    if other_count:
        print(f"  {'Unassigned (OTHER bucket)':<34}  {other_count:>4} coils")
    print()
    print(f"Coils excluded (low weight):  {exclusions['low_weight']}")
    print(f"Coils excluded (FG/Pack):     {exclusions['fg_or_pack']}")
    print(f"Coils excluded (RT=0/HOLD):   {exclusions['rt_zero_or_hold']}")
    print(f"Coils excluded (not rolling): {exclusions['not_rolling']}")
    print(f"Coils excluded (wrong plant): {exclusions['wrong_plant']}")


# ---------------------------------------------------------------------------
# ORCHESTRATION
# ---------------------------------------------------------------------------

def generate_daily_plan(wip_file: str, plan_date: date, output_file: str,
                        days: int = 1) -> None:
    print(f"Loading WIP file: {wip_file}")
    wip = load_wip(wip_file)
    print(f"  rows loaded: {len(wip)}")

    eligible, exclusions = filter_rolling_coils(wip, plan_date)
    print(f"  rows eligible for rolling: {len(eligible)}")

    wb = Workbook()
    # Remove default sheet — we'll create dated sheets explicitly.
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    for day_idx in range(days):
        d = plan_date + timedelta(days=day_idx)
        results = write_day_sheet(wb, eligible, d, age_offset=day_idx)
        assigned = assign_sections(eligible)
        other_count = int((assigned["section"] == "OTHER").sum())
        print_validation(d, results, exclusions, other_count)

    wb.save(output_file)
    print(f"\nWritten: {output_file}")
    return output_file

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate Tata Steel CRM Narrow Complex daily mill plan.")
    parser.add_argument("--wip_file", required=True,
                        help="Path to Narrow_Data_Coil_Stage.xlsx")
    parser.add_argument("--plan_date", required=True, type=_parse_date,
                        help="Planning date as YYYY-MM-DD")
    parser.add_argument("--output", required=True,
                        help="Output mill plan .xlsx path")
    parser.add_argument("--days", type=int, default=1,
                        help="Number of consecutive days to generate "
                             "(default: 1)")
    args = parser.parse_args(argv)

    try:
        generate_daily_plan(args.wip_file, args.plan_date,
                            args.output, args.days)
    except FileNotFoundError as e:
        print(f"ERROR: file not found — {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        raise

    return 0


if __name__ == "__main__":
    example_args = [
        "--wip_file", "/content/Narrow Data_Coil Stage.xlsx", # Adjust this path
        "--plan_date", "2023-10-27", # Adjust this date
        "--output", "/content/mill_plan_output.xlsx", # Adjust this path
        "--days", "3" # Optional: generate for multiple days
    ]
    sys.exit(main(example_args))
