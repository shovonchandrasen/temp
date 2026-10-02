"""
Core processing logic - adapted from logic_update_for_excel.py.
Handles filename classification, month/year extraction, validation,
unmatched-product detection, and the final Excel processing
with support for user-provided manual name mappings.
"""
import os
import re
import shutil
import pandas as pd
import openpyxl
from datetime import datetime
from collections import OrderedDict

# ---------------------------------------------------------------------------
# Month / year extraction
# ---------------------------------------------------------------------------
MONTH_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def extract_month_year_from_xlsx(name):
    base = os.path.basename(name)
    m = re.search(
        r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\D+(\d{2,4})",
        base, re.IGNORECASE,
    )
    if m:
        mon_str = m.group(1)[:3].lower()
        yr_raw = m.group(2)
        month = MONTH_MAP.get(mon_str)
        year = int(yr_raw)
        if year < 100:
            year += 2000
        if month:
            return month, year, datetime(year, month, 1).strftime("%B %Y")
    return None, None, None


def extract_month_year_from_csv(name):
    base = os.path.basename(name)
    m = re.search(r"from-(\d{4})-(\d{2})-\d{2}-to-(\d{4})-(\d{2})-\d{2}", base)
    if m:
        y1, mo1 = int(m.group(1)), int(m.group(2))
        y2, mo2 = int(m.group(3)), int(m.group(4))
        if y1 == y2 and mo1 == mo2:
            return mo1, y1, datetime(y1, mo1, 1).strftime("%B %Y")
        return mo1, y1, f"{datetime(y1, mo1, 1).strftime('%B %Y')} - {datetime(y2, mo2, 1).strftime('%B %Y')}"
    return extract_month_year_from_xlsx(name)


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------
def classify_file(filename):
    base = os.path.basename(filename).lower()
    if base.endswith(".csv") and ("store wise" in base or "stock report" in base):
        return "stock_csv"
    if base.endswith(".xlsx") and "closing stock" in base:
        return "closing"
    if base.endswith(".xlsx") and "inventory activity" in base:
        return "inventory"
    return None


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def run_checks(inv_path, close_path, csv_path):
    checks = OrderedDict()
    for label, p in [("Inventory Activity file present", inv_path),
                     ("Closing Stock file present", close_path),
                     ("Store Wise Stock CSV present", csv_path)]:
        checks[label] = (p is not None and os.path.exists(p), "")
    checks["Inventory Activity is .xlsx"] = (
        inv_path is not None and inv_path.lower().endswith(".xlsx"), ""
    )
    checks["Closing Stock is .xlsx"] = (
        close_path is not None and close_path.lower().endswith(".xlsx"), ""
    )
    checks["Store Wise Stock is .csv"] = (
        csv_path is not None and csv_path.lower().endswith(".csv"), ""
    )

    inv_cols_ok = close_cols_ok = csv_cols_ok = False
    inv_msg = close_msg = csv_msg = ""

    if inv_path and os.path.exists(inv_path):
        try:
            wb = openpyxl.load_workbook(inv_path, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
                if any(v is not None and "Product name" in str(v) for v in row):
                    inv_cols_ok = True; break
            wb.close()
            if not inv_cols_ok:
                inv_msg = "'Product name' column not found in Inventory Activity"
        except Exception as e:
            inv_msg = f"Cannot open Inventory Activity: {e}"

    if close_path and os.path.exists(close_path):
        try:
            wb = openpyxl.load_workbook(close_path, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
                if any(v is not None and "Product name" in str(v) for v in row):
                    close_cols_ok = True; break
            wb.close()
            if not close_cols_ok:
                close_msg = "'Product name' column not found in Closing Stock"
        except Exception as e:
            close_msg = f"Cannot open Closing Stock: {e}"

    if csv_path and os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path, nrows=1)
            cols = [str(c).strip() for c in df.columns]
            required = {"Item Name", "Opening Stock", "Goods Receipt", "Goods Return",
                        "Inventory Deducted", "Closing Stock"}
            missing = required - set(cols)
            csv_cols_ok = not missing
            if not csv_cols_ok:
                csv_msg = f"Missing columns: {', '.join(sorted(missing))}"
        except Exception as e:
            csv_msg = f"Cannot read CSV: {e}"

    checks["Inventory Activity has 'Product name' column"] = (inv_cols_ok, inv_msg)
    checks["Closing Stock has 'Product name' column"] = (close_cols_ok, close_msg)
    checks["CSV has required stock columns"] = (csv_cols_ok, csv_msg)

    m_inv = extract_month_year_from_xlsx(inv_path) if inv_path else (None, None, None)
    m_clo = extract_month_year_from_xlsx(close_path) if close_path else (None, None, None)
    m_csv = extract_month_year_from_csv(csv_path) if csv_path else (None, None, None)

    months_extracted = all(x[0] is not None for x in [m_inv, m_clo, m_csv])
    months_match = (months_extracted
                    and m_inv[0] == m_clo[0] == m_csv[0]
                    and m_inv[1] == m_clo[1] == m_csv[1])
    checks["Month/Year detected in all 3 filenames"] = (months_extracted, "")
    checks["All three files report the SAME month"] = (
        months_match,
        f"→ {m_inv[2]}" if months_match else f"Inv={m_inv[2]}, Close={m_clo[2]}, CSV={m_csv[2]}"
    )
    meta = {
        "inv_month": m_inv, "clo_month": m_clo, "csv_month": m_csv,
        "all_passed": all(v[0] for v in checks.values()),
        "month_label": m_inv[2] if months_match else "(mixed)",
    }
    return checks, meta


# ---------------------------------------------------------------------------
# Read product lists from each workbook for the mapping dropdowns
# ---------------------------------------------------------------------------
def _read_xlsx_products(path, header_row):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[header_row]]
    pidx = headers.index("Product name")
    cat_idx = headers.index("Category Name") if "Category Name" in headers else None
    prods = []
    seen = set()
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        name = row[pidx]
        if name is None: continue
        name_s = str(name).strip()
        if not name_s or name_s.lower() in seen: continue
        seen.add(name_s.lower())
        cat = str(row[cat_idx]).strip() if cat_idx is not None and row[cat_idx] is not None else ""
        prods.append({"name": name_s, "category": cat})
    wb.close()
    return prods


def get_inventory_products(path, csv_names=None):
    """
    Return list of product dicts {name, category} from the Inventory sheet.
    If csv_names (set of lowercased+stripped CSV Item Names) is provided,
    only products NOT already in the CSV (exact match) are returned — these
    are the only candidates that need to be offered as manual-map targets.
    """
    prods = _read_xlsx_products(path, header_row=4)
    if csv_names is not None:
        prods = [p for p in prods if p["name"].lower() not in csv_names]
    return prods


def get_closing_products(path, csv_names=None):
    prods = _read_xlsx_products(path, header_row=7)
    if csv_names is not None:
        prods = [p for p in prods if p["name"].lower() not in csv_names]
    return prods


# ---------------------------------------------------------------------------
# Find unmatched CSV rows (non-zero in E,F,G,L,Q) for manual mapping
# ---------------------------------------------------------------------------
# Columns (1-based in Excel, but in pandas 0-based we use names):
# E=Opening Stock, F=Goods Receipt, G=Goods Return, L=Inventory Deducted, Q=Closing Stock
UNMATCH_VALUE_COLS = ["Opening Stock", "Goods Receipt", "Goods Return",
                      "Inventory Deducted", "Closing Stock"]


def _to_num(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return 0.0
        return float(x)
    except Exception:
        return 0.0


def find_unmatched_csv_items(csv_path, inv_products, clo_products):
    """
    Return rows from the CSV where Item Name is NOT present (case-insensitive,
    stripped) in either the Inventory Activity product list OR the Closing Stock
    product list, AND at least one of columns E, F, G, L, Q is non-zero.
    """
    df = pd.read_csv(csv_path)
    df["Item Name"] = df["Item Name"].astype(str).str.strip()

    inv_set = {p["name"].lower() for p in inv_products}
    clo_set = {p["name"].lower() for p in clo_products}

    unmatched = []
    for _, r in df.iterrows():
        name = str(r["Item Name"]).strip()
        if not name:
            continue
        nl = name.lower()
        if nl in inv_set and nl in clo_set:
            continue
        vals = {c: _to_num(r.get(c)) for c in UNMATCH_VALUE_COLS}
        # Aggregate across rows if duplicate Item Names (sum)
        existing = next((u for u in unmatched if u["name"].lower() == nl), None)
        if existing:
            for c in UNMATCH_VALUE_COLS:
                existing["values"][c] += vals[c]
            # Mark whether still unmatched in each sheet
            existing["missing_inv"] = existing["missing_inv"] and (nl not in inv_set)
            existing["missing_clo"] = existing["missing_clo"] and (nl not in clo_set)
            continue
        entry = {
            "name": name,
            "category": str(r.get("Category", "")).strip(),
            "values": vals,
            "missing_inv": nl not in inv_set,
            "missing_clo": nl not in clo_set,
        }
        unmatched.append(entry)

    # Keep only rows that are unmatched AND have at least one non-zero value
    # in E, F, G, L, Q (per user requirement). All-zero unmatched rows are
    # irrelevant because there's nothing to transfer into the Excel files.
    def non_zero(e):
        return any(abs(v) > 1e-9 for v in e["values"].values())
    return [u for u in unmatched if non_zero(u)]


# ---------------------------------------------------------------------------
# Main processing (core logic preserved; manual mappings layered on top)
# ---------------------------------------------------------------------------
def process_files(inv_path, close_path, csv_path, output_dir,
                  inv_mapping=None, clo_mapping=None):
    """
    inv_mapping: dict  {csv_item_name: inventory_activity_product_name}
    clo_mapping: dict  {csv_item_name: closing_stock_product_name}
    These are ADDITIONAL mappings on top of exact-name matches.
    """
    inv_mapping = inv_mapping or {}
    clo_mapping = clo_mapping or {}
    log = []

    # 1. Load CSV reference (case-insensitive lookup to match validation/review)
    df_stock = pd.read_csv(csv_path)
    df_stock["Item Name"] = df_stock["Item Name"].astype(str).str.strip()
    stock_map = {
        str(k).strip().lower(): v
        for k, v in df_stock.set_index("Item Name").to_dict("index").items()
    }

    # 1b. Add the manual mapping aliases into stock_map so core logic below
    #     can still look up by Excel product name directly.
    for csv_name, xlsx_name in inv_mapping.items():
        csv_key = str(csv_name).strip().lower()
        if csv_key in stock_map and xlsx_name and xlsx_name.strip():
            stock_map[xlsx_name.strip().lower()] = stock_map[csv_key]
    for csv_name, xlsx_name in clo_mapping.items():
        csv_key = str(csv_name).strip().lower()
        if csv_key in stock_map and xlsx_name and xlsx_name.strip():
            # Closing-stock sheet uses 'Closing Stock' col which is already in stock_map
            stock_map.setdefault(xlsx_name.strip().lower(), stock_map[csv_key])

    # 2. Closing stock (Physical column) map from Closing workbook
    wb2 = openpyxl.load_workbook(close_path)
    ws2 = wb2.active
    header_row2 = 7
    headers2 = [c.value for c in ws2[header_row2]]
    idx2 = headers2.index("Product name")

    # Build base map from workbook (case-insensitive lookup)
    stock_closing_map = {}
    for r in ws2.iter_rows(min_row=8, values_only=True):
        if r[idx2] is None: continue
        stock_closing_map[str(r[idx2]).strip().lower()] = {"Physical": r[idx2 + 2]}

    # If user mapped a CSV name to a Closing Stock product that's not in the sheet,
    # there's nothing to fill; but if they mapped a Closing product whose Physical
    # cell is empty we leave it. If the CSV item had a value we don't override.

    # 3. Process Inventory Activity workbook (copy first, preserve originals)
    inv_out = os.path.join(output_dir, os.path.basename(inv_path))
    shutil.copy2(inv_path, inv_out)
    wb = openpyxl.load_workbook(inv_out)
    ws = wb.active

    header_row = 4
    headers = [cell.value for cell in ws[header_row]]
    prod_col_idx = headers.index("Product name") + 1

    inv_updates = 0
    for row_num in range(header_row + 1, ws.max_row + 1):
        prod_name = str(ws.cell(row=row_num, column=prod_col_idx).value or "").strip()
        if not prod_name:
            continue
        prod_key = prod_name.lower()
        if prod_key in stock_map:
            data = stock_map[prod_key]
            ws.cell(row=row_num, column=prod_col_idx + 1).value = data.get("Opening Stock")
            ws.cell(row=row_num, column=prod_col_idx + 2).value = data.get("Goods Receipt")
            ws.cell(row=row_num, column=prod_col_idx + 3).value = data.get("Goods Return")
            ws.cell(row=row_num, column=prod_col_idx + 8).value = data.get("Inventory Deducted")
            inv_updates += 1
        if prod_key in stock_closing_map:
            ws.cell(row=row_num, column=prod_col_idx + 9).value = stock_closing_map[prod_key].get("Physical")
    wb.save(inv_out)
    log.append(f"Inventory Activity: {inv_updates} rows matched & updated → {os.path.basename(inv_out)}")

    # 4. Process Closing Stock workbook
    close_out = os.path.join(output_dir, os.path.basename(close_path))
    shutil.copy2(close_path, close_out)
    wb2 = openpyxl.load_workbook(close_out)
    ws2 = wb2.active
    headers2 = [c.value for c in ws2[header_row2]]
    prod_col_idx2 = headers2.index("Product name") + 1
    clo_updates = 0

    # Build a reverse-lookup for manual mappings: if user said CSV X maps to Closing Y,
    # then when we reach row Y in Closing sheet, we want stock_map[X] (the CSV values).
    reverse_clo = {}
    for csv_name, xlsx_name in clo_mapping.items():
        csv_key = str(csv_name).strip().lower()
        if xlsx_name and csv_key in stock_map:
            reverse_clo[xlsx_name.strip().lower()] = stock_map[csv_key]

    for row_num in range(header_row2 + 1, ws2.max_row + 1):
        prod_name = str(ws2.cell(row=row_num, column=prod_col_idx2).value or "").strip()
        if not prod_name:
            continue
        prod_key = prod_name.lower()
        # Direct match
        if prod_key in stock_map:
            ws2.cell(row=row_num, column=prod_col_idx2 + 1).value = stock_map[prod_key].get("Closing Stock")
            clo_updates += 1
        # Manual mapped alias (e.g. CSV name "X" → Closing product "Y")
        elif prod_key in reverse_clo:
            ws2.cell(row=row_num, column=prod_col_idx2 + 1).value = reverse_clo[prod_key].get("Closing Stock")
            clo_updates += 1
    wb2.save(close_out)
    log.append(f"Closing Stock: {clo_updates} rows matched & updated → {os.path.basename(close_out)}")

    if inv_mapping or clo_mapping:
        log.append(f"Applied manual mappings: {len(inv_mapping)} for Inventory, {len(clo_mapping)} for Closing.")

    return {
        "outputs": [os.path.basename(inv_out), os.path.basename(close_out)],
        "log": "\n".join(log),
        "inv_updates": inv_updates,
        "clo_updates": clo_updates,
    }
