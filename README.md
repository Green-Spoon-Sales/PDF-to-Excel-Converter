# PDF-to-Excel-Converter
Multiple pdf to excel conversion tools for promo team to convert invoices from UNFI/KEHE to excel.



Western Region UNFI bulk

import os
import re
import glob
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import pytesseract
from pdf2image import convert_from_path
import pdfplumber

# --- OCR config --------------------------------------------------------------
POPPLER_PATH = None  # set to your poppler/bin path on Windows if needed
TESSERACT_CONFIG = r"--oem 3 --psm 6"

# --- Output schema -----------------------------------------------------------
COLUMNS = [
    "Week Ending",
    "Customer Info",
    "Brand",
    "Product",
    "Unit",
    "Description",
    "Invoice",
    "Ordered",
    "Shipped",
    "Whlse",
    "Total Discount %",
    "MCB %",
    "MCB",
]

# --- Helpers / cleaners ------------------------------------------------------
def clean_val(v: Optional[str]) -> str:
    if v is None:
        return ""
    return str(v).strip()

def clean_shipped(v: Optional[str]) -> str:
    v = clean_val(v)
    return v.rstrip(".")  # OCR sometimes adds a trailing period

def clean_invoice(v: Optional[str]) -> str:
    return clean_val(v)

def clean_money(v: Optional[str]) -> str:
    return clean_val(v)

def normalize_percent(text: Optional[str]) -> str:
    """
    Normalize weird OCR % like:
      "100 00"  -> "100.00%"
      "100 00%" -> "100.00%"
      "12.64%"  (already fine)
      "5"       -> "5%"
    """
    if not text:
        return ""
    t = text.strip()

    # already like "12.34%"
    if re.match(r"^\d{1,3}(?:\.\d{1,2})%$", t):
        return t

    # "100 00" -> "100.00%"
    m = re.match(r"^(\d{1,3})\s+(\d{2})$", t)
    if m:
        return f"{m.group(1)}.{m.group(2)}%"

    # "100 00%" -> "100.00%"
    m2 = re.match(r"^(\d{1,3})\s+(\d{2})%$", t)
    if m2:
        return f"{m2.group(1)}.{m2.group(2)}%"

    # "5" or "12.3" -> "...%"
    m3 = re.match(r"^(\d{1,3}(?:\.\d{1,2})?)$", t)
    if m3:
        return m3.group(1) + "%"

    # collapse internal spaces like "50 . 00 %"
    t2 = re.sub(r"\s+", "", t)
    if re.match(r"^\d{1,3}(?:\.\d{1,2})%$", t2):
        return t2

    return t

def fix_ocr_money(val: Optional[str]) -> str:
    """
    Clean up money-like values from OCR.
    Handles:
      - thousands commas: "2,210.30"
      - embedded spaces: "4,121 5" -> "4121.5" -> "4121.50"
      - one decimal place: "2210.3" -> "2210.30"
      - big ints with no decimal: "3151" -> "31.51"
    """
    v = clean_val(val)
    if not v:
        return v

    # Strip $, commas, and internal spaces
    v = v.replace("$", "")
    v = v.replace(",", "")
    v = re.sub(r"\s+", "", v)

    # Already like 123.45
    if re.match(r"^\d+\.\d{2}$", v):
        return v

    # One decimal place -> pad trailing zero
    m = re.match(r"^(\d+)\.(\d)$", v)
    if m:
        return f"{m.group(1)}.{m.group(2)}0"

    # All digits and reasonably large -> tuck decimal before last 2 digits
    if re.match(r"^\d{3,}$", v):
        return v[:-2] + "." + v[-2:]

    return v

def to_float_or_none(val: Optional[str]) -> Optional[float]:
    """
    Try to convert a cleaned string to float.
    Return None if it isn't a valid number.
    """
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None

def compute_mcb_from_components(
    shipped: Optional[str],
    whlse: Optional[str],
    mcb_pct: Optional[str],
) -> Optional[float]:
    """
    Compute MCB = Shipped * Whlse * (MCB% / 100) as a float,
    using Decimal internally for stable 2-decimal rounding.
    Used ONLY for validation / suspicious logging, not to override OCR.
    """
    try:
        shipped_s = clean_val(shipped)
        if not shipped_s:
            return None

        whlse_clean = fix_ocr_money(whlse)
        if not whlse_clean:
            return None

        pct_norm = normalize_percent(mcb_pct)
        if not pct_norm or not pct_norm.endswith("%"):
            return None
        pct_s = pct_norm.rstrip("%")

        shipped_dec = Decimal(shipped_s)
        whlse_dec = Decimal(whlse_clean)
        pct_dec = Decimal(pct_s)

        calc_dec = (whlse_dec * shipped_dec * pct_dec / Decimal("100")).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
        return float(calc_dec)
    except Exception:
        return None

# --- OCR cleanup helpers -----------------------------------------------------
def normalize_invoice_paren(line: str) -> str:
    """
    OCR sometimes inserts '(' right before the invoice number:
      '... BEV (016389141 1 1 39.96 ...' -> make it ' 016389141 ...'
    """
    return re.sub(r"\((\d{6,12})", r" \1", line)

def clean_ocr_line_for_parse(line: str) -> str:
    """
    Normalize tiny OCR/text glitches that break regex parsing:
    - Strip stray '= ' right before trailing money
    - Collapse weird punctuation right before invoice digits
    - Remove '$' so amounts become plain numbers
    - Fix 'T'→'7' product OCR
    - Collapse multiple spaces
    """
    s = line.strip()

    # drop all dollar signs so money tokens become plain numbers
    s = s.replace("$", "")

    s = re.sub(r'=\s+(\d+\.\d{2}\s*$)', r'\1', s)
    s = re.sub(r'(%\s+)=\s+(\d)', r'\1 \2', s)
    s = normalize_invoice_paren(s)

    # normalize stray punctuation before invoice digits
    s = re.sub(r'([A-Za-z0-9\)])\s*[=:\-–—]+\s*(\(?\d{6,12}\)?)', r'\1 \2', s)

    # map OCR 'T' to '7' in product-like tokens: "T4274" -> "74274"
    s = re.sub(r'\bT(?=\d{4,}\b)', '7', s)

    # collapse multi-spaces
    s = re.sub(r"\s{2,}", " ", s)
    return s

# --- Shared unit pattern -----------------------------------------------------
UNIT_PATTERN = r"\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?){1,2}\s*(?:OZ|Z|0Z)"

# --- Regex patterns ----------------------------------------------------------
ROW_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@\*\-\.'&/ ]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    (?P<maybeS>S\d{{4,5}}\s+)?                     
    [\s–—-]*
    (?P<invoice>\d{{9,12}})\s+
    (?P<ordered>\d+)\s+
    (?P<shipped>\d+\.?)\s+
    (?P<whlse>\d+\.\d{{2}})\s+
    (?P<total_disc>(?:\d{{1,3}}(?:\.\d{{1,2}})?)%)\s+
    (?P<mcb_pct>(?:\d{{1,3}}(?:\.\d{{1,2}})?)%)\s+
    (?P<mcb>\d+\.\d{{2}})\s*
    $
    """,
    re.VERBOSE,
)

ROW_VARIANT_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@\*\-\.'&/ ]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    (?P<maybeS>S\d{{4,5}}\s+)?                                    
    [\s–—-]*[=:]*\s*
    \(?(?P<invoice>\d{{6,12}})\)?\s+
    (?P<ordered>\d+)\s+
    (?P<shipped>\d+\.?)\s+
    (?P<whlse>\d+(?:\.\d{{2}}|\d{{3,}}))\s+
    (?P<total_disc>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s*
    (?P<mcb_pct>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s*
    =?\s*
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

ROW_FALLBACK2_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@][A-Za-z @&/\-'\.]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    [\s–—-]*
    (?P<invoice>\d{{6,12}})\s+
    (?P<ordered>\d+)\s+
    (?P<shipped>\d+\.?)\s+
    (?P<whlse>\d+(?:\.\d{{2}}|\d{{3,}}))\s+
    (?P<total_disc>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    (?P<mcb_pct>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    =?\s*
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

ROW_FALLBACK_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@\*\-\.'&/ ]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    (?P<maybeS>S\d{{4,5}}\s+)?               
    [\s–—-]*
    (?P<invoice>\d{{6,12}})\s+
    (?P<tail>.+?)\s+
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

ROW_ADJUST_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@\*\-\.'&/ ]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    (?P<maybeS>S\d{{4,5}})?\s*
    [\s–—-]*
    (?P<invoice>\d{{6,12}})?\s*
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

# ALL SALES MCB promo rows (no invoice)
ROW_PROMO_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@][A-Za-z @&/\-'\.]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    (?P<shipped>\d+)\s+
    (?P<whlse>\d+(?:\.\d{{2}}|\d{{3,}}))\s+
    (?P<total_disc>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    (?P<mcb_pct>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

ROW_SUPER_LOOSE_REGEX = re.compile(
    rf"""
    ^
    (?P<brand>[A-Za-z@][A-Za-z @&/\-'\.]+?)\s+
    (?P<product>\d{{3,}})\s+
    (?P<unit>{UNIT_PATTERN})\s+
    (?P<desc>.*?)\s+
    .{{0,6}}?
    \(?(?P<invoice>\d{{6,12}})\)?\s+
    (?P<ordered>\d+)\s+
    (?P<shipped>\d+\.?)\s+
    (?P<whlse>\d+(?:\.\d{{2}}|\d{{3,}}))\s+
    (?P<total_disc>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    (?P<mcb_pct>\d{{1,3}}(?:[\.\s]\d{{2}})?%?)\s+
    =?\s*
    (?P<mcb>[0-9,\.]+)\s*$
    """,
    re.VERBOSE,
)

HEADER_REGEX = re.compile(
    r"\bBrand\s+Product\s+Unit\s+Description\s+Invoice\s+Ordered\s+Shipped\s+Whlse",
    re.IGNORECASE,
)

# --- Extraction helpers ------------------------------------------------------
def extract_week_ending(text: str) -> Optional[str]:
    m = re.search(
        r"Week\s*ending\s*:?\s*([0-9]{2}/[0-9]{2}/[0-9]{4})",
        text,
        re.IGNORECASE,
    )
    return m.group(1) if m else None

def maybe_get_customer(line: str) -> Optional[str]:
    cust_match = re.search(
        r"C[uU]st[0o]mer\s*[:\-]?\s*(.+)",
        line,
        re.IGNORECASE,
    )
    if cust_match:
        return cust_match.group(1).strip()
    return None

def append_row(
    rows: List[Dict[str, str]],
    week_ending: str,
    current_customer: str,
    brand: str,
    product: str,
    unit: str,
    desc: str,
    invoice: str,
    ordered: str,
    shipped: str,
    whlse: str,
    total_disc: str,
    mcb_pct: str,
    mcb: str,
    suspicious_rows: Optional[List[Dict[str, str]]] = None,
):
    """
    PURE EXTRACTION FOR MCB:

    - Always use the MCB value printed on the PDF (OCR/text) when it parses.
    - Only compute MCB for validation:
        * If OCR vs calc differ by > $0.01, log to suspicious_rows.
    - Only fall back to calc if OCR MCB is missing/garbage.
    """
    # Clean money fields
    whlse_fixed = fix_ocr_money(whlse)
    mcb_fixed = fix_ocr_money(mcb)

    # Normalize percent and parse as float, if possible (for logging)
    pct_norm = normalize_percent(mcb_pct)

    # Validation-only calculation
    mcb_calc_val = compute_mcb_from_components(shipped, whlse, mcb_pct)

    # Parse the OCR MCB as float if possible
    mcb_fixed_float = to_float_or_none(mcb_fixed)
    mcb_calc_float = mcb_calc_val  # already float or None

    # FINAL MCB VALUE: default is OCR/text, not calc
    if mcb_fixed_float is not None:
        final_mcb_val = mcb_fixed_float
    else:
        # Only if OCR MCB is unusable do we fall back to calc
        final_mcb_val = mcb_calc_float

    # Log suspicious differences between OCR and calc (for review only)
    if (
        mcb_fixed_float is not None and
        mcb_calc_float is not None and
        suspicious_rows is not None and
        abs(mcb_calc_float - mcb_fixed_float) > 0.01
    ):
        suspicious_rows.append({
            "Week Ending": clean_val(week_ending),
            "Customer Info": clean_val(current_customer),
            "Brand": clean_val(brand),
            "Product": clean_val(product),
            "Invoice": clean_invoice(invoice),
            "Whlse": whlse_fixed,
            "MCB %": pct_norm,
            "MCB_OCR": mcb_fixed_float,
            "MCB_Calc": mcb_calc_float,
            "Diff": round(mcb_calc_float - mcb_fixed_float, 2),
        })

    final_mcb_str = ""
    if final_mcb_val is not None:
        final_mcb_str = f"{final_mcb_val:.2f}"

    rows.append({
        "Week Ending": clean_val(week_ending),
        "Customer Info": clean_val(current_customer),
        "Brand": clean_val(brand),
        "Product": clean_val(product),
        "Unit": clean_val(unit),
        "Description": clean_val(desc),
        "Invoice": clean_invoice(invoice),
        "Ordered": clean_val(ordered),
        "Shipped": clean_shipped(shipped),
        "Whlse": clean_money(whlse_fixed),
        "Total Discount %": pct_norm if total_disc else normalize_percent(total_disc),
        "MCB %": pct_norm,
        "MCB": clean_money(final_mcb_str),
    })

# --- NEW: token-based tail parser -------------------------------------------
def parse_tail_tokens(tail: str) -> Optional[Dict[str, str]]:
    """
    More flexible tail parser for fallback rows. Handles:
      - 1 1 74.76 100.00% 100.00%
      - 1 1 74.76 100.00 % 100.00 %
      - 1 1 7476 100 00 %
      - Optional trailing garbage like 'w/o', '='
    """
    tokens = tail.split()
    if len(tokens) < 3:
        return None

    i = 0

    # ordered
    if not re.match(r"^\d+$", tokens[i]):
        return None
    ordered = tokens[i]
    i += 1

    # shipped
    if i >= len(tokens) or not re.match(r"^\d+\.?$", tokens[i]):
        return None
    shipped = tokens[i].rstrip(".")
    i += 1

    # whlse (money-like)
    if i >= len(tokens) or not re.match(r"^\d+(?:\.\d{1,2}|\d{3,})?$", tokens[i]):
        return None
    whlse = tokens[i]
    i += 1

    def consume_percent(idx: int) -> Tuple[str, int]:
        if idx >= len(tokens):
            return "", idx
        tok = tokens[idx]

        # Case: token already has %
        if "%" in tok:
            return tok, idx + 1

        # Case: token is number and next is '%'
        if re.match(r"^\d+(?:[\.\s]\d{2})?$", tok) and idx + 1 < len(tokens) and tokens[idx + 1] == "%":
            return tok + "%", idx + 2

        # Generic: number with implied %
        if re.match(r"^\d+(?:[\.\s]\d{2})?$", tok):
            return tok + "%", idx + 1

        return "", idx

    # total_disc
    total_disc, i = consume_percent(i)

    # mcb_pct
    mcb_pct, i = consume_percent(i)

    if not total_disc and not mcb_pct:
        # If we couldn't parse any percents, give up and let caller fall back
        return None

    return {
        "ordered": ordered,
        "shipped": shipped,
        "whlse": whlse,
        "total_disc": total_disc,
        "mcb_pct": mcb_pct,
    }

def try_tail_split_and_append(
    rows: List[Dict[str, str]],
    week_ending: str,
    current_customer: str,
    fd: Dict[str, str],
    suspicious_rows: Optional[List[Dict[str, str]]] = None,
):
    tail = (fd.get("tail") or "").strip()

    parsed = parse_tail_tokens(tail)
    if parsed is not None:
        append_row(
            rows,
            week_ending or "",
            current_customer or "",
            fd.get("brand", ""),
            fd.get("product", ""),
            fd.get("unit", ""),
            fd.get("desc", ""),
            fd.get("invoice", ""),
            parsed["ordered"],
            parsed["shipped"],
            parsed["whlse"],
            parsed["total_disc"],
            parsed["mcb_pct"],
            fd.get("mcb", ""),
            suspicious_rows=suspicious_rows,
        )
        return True

    # worst case fallback/partial (no components – no calc or validation)
    append_row(
        rows,
        week_ending or "",
        current_customer or "",
        fd.get("brand", ""),
        fd.get("product", ""),
        fd.get("unit", ""),
        fd.get("desc", ""),
        fd.get("invoice", ""),
        "", "", "",
        "", "",
        fd.get("mcb", ""),
        suspicious_rows=suspicious_rows,
    )
    return True

# --- Line merge logic --------------------------------------------------------
def merge_wrapped_lines(raw_lines: List[str]) -> List[str]:
    merged = []
    skip_next = False

    invoice_tail_regex = re.compile(r"^\s*\(?\d{6,12}\s+\d+\s+\d+")

    for i in range(len(raw_lines)):
        if skip_next:
            skip_next = False
            continue

        line = raw_lines[i].strip()
        if i < len(raw_lines) - 1:
            nxt = raw_lines[i + 1].strip()

            test_line = normalize_invoice_paren(line)
            test_nxt = normalize_invoice_paren(nxt)

            if line.endswith("(") and invoice_tail_regex.match(test_nxt):
                merged.append(line + " " + nxt)
                skip_next = True
                continue

            looks_like_start = re.match(r"^[A-Za-z][A-Za-z ]+\s+\d{3,}\s+\d+/\d+", test_line)
            has_invoice_digits = re.search(r"\d{6,12}\s+\d+\s+\d+", test_line)
            if looks_like_start and not has_invoice_digits and invoice_tail_regex.match(test_nxt):
                merged.append(line + " " + nxt)
                skip_next = True
                continue

        merged.append(line)

    return merged

# --- Page parsing ------------------------------------------------------------
def parse_rows_from_page_text(
    text: str,
    page_number: int,
    brands_to_extract: Optional[List[str]] = None,
    unparsed_accumulator: Optional[List[str]] = None,
    suspicious_rows: Optional[List[Dict[str, str]]] = None,
) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    week_ending = extract_week_ending(text)
    current_customer = None

    brand_filter_upper = [b.upper() for b in brands_to_extract] if brands_to_extract else None

    physical_lines = [l for l in text.splitlines() if l.strip()]
    logical_lines = merge_wrapped_lines(physical_lines)

    # Debug preview
    print(f"--- PAGE {page_number} POST-MERGE LINES DEBUG ---")
    for idx, l in enumerate(logical_lines):
        if re.match(r"^[A-Za-z][A-Za-z ]+\s+\d{3,}\s+\d+/\d+", l.strip()):
            print(f"[p{page_number} l{idx}] {l.strip()}")
    print("------ END PAGE DEBUG ------\n")

    for raw in logical_lines:
        original_line = raw.strip()
        if not original_line:
            continue

        line = clean_ocr_line_for_parse(original_line)

        cust = maybe_get_customer(line)
        if cust:
            current_customer = cust
            continue

        # Skip headings / rollups
        if HEADER_REGEX.search(line):
            continue
        if (
            line.startswith("Customer Totals")
            or "MCB TOTALS PAGE" in line
            or "MCBs BY CUSTOMER" in line.upper()
            or "MCBs by Customer Summaries".upper() in line.upper()
            or re.search(r"\bPage\s+\d+\s+of\s+\d+", line, re.IGNORECASE)
            or re.search(r"\bDivision\b", line, re.IGNORECASE)
            or re.search(r"\bCategory\b", line, re.IGNORECASE)
            or re.search(r"\bTotals?\b", line, re.IGNORECASE)
            or re.search(r"Report total", line, re.IGNORECASE)
            or re.search(r"Total Deduction", line, re.IGNORECASE)
        ):
            continue

        if brand_filter_upper is not None:
            if not any(b in line.upper() for b in brand_filter_upper):
                continue

        # strict
        m = ROW_REGEX.match(line)
        if m:
            d = m.groupdict()
            append_row(
                rows, week_ending or "", current_customer or "",
                d["brand"], d["product"], d["unit"], d["desc"],
                d["invoice"], d["ordered"], d["shipped"], d["whlse"],
                d["total_disc"], d["mcb_pct"], d["mcb"],
                suspicious_rows=suspicious_rows,
            )
            continue

        vm = ROW_VARIANT_REGEX.match(line)
        if vm:
            vd = vm.groupdict()
            append_row(
                rows, week_ending or "", current_customer or "",
                vd.get("brand", ""), vd.get("product", ""), vd.get("unit", ""), vd.get("desc", ""),
                vd.get("invoice", ""), vd.get("ordered", ""), vd.get("shipped", ""),
                vd.get("whlse", ""), vd.get("total_disc", ""), vd.get("mcb_pct", ""), vd.get("mcb", ""),
                suspicious_rows=suspicious_rows,
            )
            continue

        hm = ROW_FALLBACK2_REGEX.match(line)
        if hm:
            hd = hm.groupdict()
            append_row(
                rows, week_ending or "", current_customer or "",
                hd.get("brand", ""), hd.get("product", ""), hd.get("unit", ""), hd.get("desc", ""),
                hd.get("invoice", ""), hd.get("ordered", ""), hd.get("shipped", ""),
                hd.get("whlse", ""), hd.get("total_disc", ""), hd.get("mcb_pct", ""), hd.get("mcb", ""),
                suspicious_rows=suspicious_rows,
            )
            continue

        # Promo / ALL SALES rows (no invoice)
        pm = ROW_PROMO_REGEX.match(line)
        if pm:
            pd = pm.groupdict()
            shipped = pd.get("shipped", "")
            append_row(
                rows,
                week_ending or "",
                current_customer or "",
                pd.get("brand", ""),
                pd.get("product", ""),
                pd.get("unit", ""),
                pd.get("desc", ""),
                "",  # no invoice
                shipped,  # ordered = shipped
                shipped,
                pd.get("whlse", ""),
                pd.get("total_disc", ""),
                pd.get("mcb_pct", ""),
                pd.get("mcb", ""),
                suspicious_rows=suspicious_rows,
            )
            continue

        sm = ROW_SUPER_LOOSE_REGEX.match(line)
        if sm:
            sd = sm.groupdict()
            append_row(
                rows, week_ending or "", current_customer or "",
                sd.get("brand", ""), sd.get("product", ""), sd.get("unit", ""), sd.get("desc", ""),
                sd.get("invoice", ""), sd.get("ordered", ""), sd.get("shipped", ""),
                sd.get("whlse", ""), sd.get("total_disc", ""), sd.get("mcb_pct", ""), sd.get("mcb", ""),
                suspicious_rows=suspicious_rows,
            )
            continue

        fm = ROW_FALLBACK_REGEX.match(line)
        if fm:
            fd = fm.groupdict()
            try_tail_split_and_append(
                rows,
                week_ending or "",
                current_customer or "",
                fd,
                suspicious_rows=suspicious_rows,
            )
            continue

        am = ROW_ADJUST_REGEX.match(line)
        if am:
            ad = am.groupdict()
            append_row(
                rows, week_ending or "", current_customer or "",
                ad.get("brand", ""), ad.get("product", ""), ad.get("unit", ""), ad.get("desc", ""),
                ad.get("invoice", ""), "", "", "",
                "", "", ad.get("mcb", ""),
                suspicious_rows=suspicious_rows,
            )
            continue

        # suspicious unparsed lines
        if re.search(r"\d+\.\d{2}", line) or re.search(r"\d{3,}\s*\d{2}\s*%", line):
            print(f"[UNPARSED p{page_number}] {original_line}")
            if unparsed_accumulator is not None:
                unparsed_accumulator.append(f"[p{page_number}] {original_line}")

    return rows

# --- TEXT-ONLY extraction core (preferred) -----------------------------------
def text_pdf_to_rows(
    pdf_path: Path,
    brands_to_extract: Optional[List[str]] = None,
) -> Tuple[List[Dict[str, str]], str, List[str], List[Dict[str, str]]]:
    """
    First attempt: use pdfplumber's extract_text (no OCR).
    """
    all_rows: List[Dict[str, str]] = []
    first_page_text: str = ""
    unparsed_lines: List[str] = []
    suspicious_rows: List[Dict[str, str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf_obj:
        total_pages = len(pdf_obj.pages)
        for p_idx, page in enumerate(pdf_obj.pages, start=1):
            txt = page.extract_text() or ""
            if p_idx == 1:
                first_page_text = txt

            if not txt.strip():
                print(f"[TEXT] Page {p_idx}: no text extracted, may need OCR.")
                continue

            page_rows = parse_rows_from_page_text(
                txt,
                page_number=p_idx,
                brands_to_extract=brands_to_extract,
                unparsed_accumulator=unparsed_lines,
                suspicious_rows=suspicious_rows,
            )

            if page_rows:
                # carry forward last non-empty customer when needed
                if all(not r["Customer Info"] for r in page_rows) and all_rows:
                    last_customer = next(
                        (r["Customer Info"] for r in reversed(all_rows) if r["Customer Info"]),
                        ""
                    )
                    if last_customer:
                        for r in page_rows:
                            if not r["Customer Info"]:
                                r["Customer Info"] = last_customer

                all_rows.extend(page_rows)

    return all_rows, first_page_text, unparsed_lines, suspicious_rows

# --- OCR core (fallback) -----------------------------------------------------
def ocr_pdf_to_rows(
    pdf_path: Path,
    brands_to_extract: Optional[List[str]] = None,
    dpi: int = 300,
) -> Tuple[List[Dict[str, str]], str, List[str], List[Dict[str, str]]]:
    all_rows: List[Dict[str, str]] = []
    first_page_ocr_text: str = ""
    unparsed_lines: List[str] = []
    suspicious_rows: List[Dict[str, str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf_obj:
        total_pages = len(pdf_obj.pages)

    for p in range(1, total_pages + 1):
        try:
            if POPPLER_PATH:
                imgs = convert_from_path(
                    str(pdf_path),
                    first_page=p,
                    last_page=p,
                    dpi=dpi,
                    poppler_path=POPPLER_PATH,
                )
            else:
                imgs = convert_from_path(
                    str(pdf_path),
                    first_page=p,
                    last_page=p,
                    dpi=dpi,
                )
            if not imgs:
                print(f"[WARN] Page {p}: no image returned")
                continue

            text = pytesseract.image_to_string(imgs[0], config=TESSERACT_CONFIG)

            if p == 1:
                first_page_ocr_text = text

            print(f"\n--- DEBUG OCR PAGE {p} (first ~5 lines) ---")
            shown = 0
            for dbg_line in text.splitlines():
                if dbg_line.strip():
                    print(dbg_line[:200])
                    shown += 1
                if shown >= 5:
                    break
            print("--- END DEBUG ---\n")

            page_rows = parse_rows_from_page_text(
                text,
                page_number=p,
                brands_to_extract=brands_to_extract,
                unparsed_accumulator=unparsed_lines,
                suspicious_rows=suspicious_rows,
            )

            if page_rows:
                # carry forward last non-empty customer when needed
                if all(not r["Customer Info"] for r in page_rows) and all_rows:
                    last_customer = next(
                        (r["Customer Info"] for r in reversed(all_rows) if r["Customer Info"]),
                        ""
                    )
                    if last_customer:
                        for r in page_rows:
                            if not r["Customer Info"]:
                                r["Customer Info"] = last_customer

                all_rows.extend(page_rows)

        except Exception as e:
            print(f"[WARN] Failed OCR on page {p}: {e}")
            continue

    return all_rows, first_page_ocr_text, unparsed_lines, suspicious_rows

# --- Combined: text first, OCR fallback --------------------------------------
def extract_rows_with_text_then_ocr(
    pdf_path: Path,
    brands_to_extract: Optional[List[str]] = None,
    dpi: int = 300,
    min_rows_for_text: int = 50,
) -> Tuple[List[Dict[str, str]], str, List[str], List[Dict[str, str]]]:
    """
    Preferred pipeline:
      1) Try pdfplumber text extraction (no OCR).
      2) If that yields enough rows (>= min_rows_for_text), use it.
      3) Otherwise, fall back to OCR/Tesseract.
    """
    print("=== TEXT EXTRACTION PASS (pdfplumber) ===")
    text_rows, first_page_text, text_unparsed, text_suspicious = text_pdf_to_rows(
        pdf_path,
        brands_to_extract=brands_to_extract,
    )
    print(f"[TEXT] Extracted {len(text_rows)} rows via pdfplumber text.")

    # If text extraction worked well enough, use it
    if len(text_rows) >= min_rows_for_text:
        print("[TEXT] Using text-based extraction results.")
        return text_rows, first_page_text, text_unparsed, text_suspicious

    # Otherwise, fall back to OCR
    print("[TEXT] Too few rows from text; falling back to OCR/Tesseract...")
    ocr_rows, first_page_ocr_text, ocr_unparsed, ocr_suspicious = ocr_pdf_to_rows(
        pdf_path,
        brands_to_extract=brands_to_extract,
        dpi=dpi,
    )

    # Combine unparsed + suspicious logs for debugging
    combined_unparsed = text_unparsed + ocr_unparsed
    combined_suspicious = text_suspicious + ocr_suspicious

    # For totals extraction, pdfplumber is used directly anyway, so
    # we don't strictly need first_page_text here; pass OCR text.
    return ocr_rows, first_page_ocr_text, combined_unparsed, combined_suspicious

# --- Total extraction --------------------------------------------------------
def extract_report_total_from_pdf(
    pdf_path: Path,
    first_page_text_hint: str
) -> Optional[float]:
    total_candidates = []

    def scan_text_for_totals(txt: str):
        # "Report total : MCBs by Customer ... $#,###.##"
        for m in re.finditer(
            r"Report\s+total\s*:?\s*MCBs?\s+by\s+Customer.*?\$([0-9,]+\.\d{2})",
            txt,
            re.IGNORECASE | re.DOTALL,
        ):
            total_candidates.append(float(m.group(1).replace(",", "")))

        # "Total Deduction : $#,###.##"
        for m in re.finditer(
            r"Total\s+Deduction\s*:?\s*\$([0-9,]+\.\d{2})",
            txt,
            re.IGNORECASE,
        ):
            total_candidates.append(float(m.group(1).replace(",", "")))

        # "Western Region totals ... $#,###.##"
        for m in re.finditer(
            r"Western\s+Region\s+totals?.*?\$([0-9,]+\.\d{2})",
            txt,
            re.IGNORECASE | re.DOTALL,
        ):
            total_candidates.append(float(m.group(1).replace(",", "")))

    fallback_max_money = None
    with pdfplumber.open(str(pdf_path)) as pdf_obj:
        for page in pdf_obj.pages:
            txt = page.extract_text() or ""
            if txt:
                scan_text_for_totals(txt)
                for mm in re.findall(r"\$([0-9,]+\.\d{2})", txt):
                    val = float(mm.replace(",", ""))
                    if (fallback_max_money is None) or (val > fallback_max_money):
                        fallback_max_money = val

    if total_candidates:
        return max(total_candidates)

    if fallback_max_money is not None:
        return fallback_max_money

    # Last-ditch: scan the hint text (from OCR or text pass)
    if first_page_text_hint:
        scan_text_for_totals(first_page_text_hint)
        if total_candidates:
            return max(total_candidates)

        money_vals = [
            float(m.replace(",", ""))
            for m in re.findall(r"\$([0-9,]+\.\d{2})", first_page_text_hint)
        ]
        if money_vals:
            return max(money_vals)

    return None

# --- Adjustment row helper (disabled by default) -----------------------------
def add_adjustment_row_if_needed(
    df: pd.DataFrame,
    week_ending_guess: str,
    customer_info_guess: str,
    diff: float,
    adjustment_threshold: float = 1e9,
) -> pd.DataFrame:
    """
    If there's a huge gap between extracted total and the PDF's official total,
    you *could* add a synthetic adjustment row. By default, threshold is 1e9
    so this is effectively disabled — no fudge rows.
    """
    if abs(diff) < adjustment_threshold:
        return df

    adj_row = {
        "Week Ending": week_ending_guess,
        "Customer Info": customer_info_guess,
        "Brand": "MCB ADJUSTMENT",
        "Product": "",
        "Unit": "",
        "Description": "Synthetic adjustment to tie Excel MCB sum to PDF total",
        "Invoice": "",
        "Ordered": "",
        "Shipped": "",
        "Whlse": "",
        "Total Discount %": "",
        "MCB %": "",
        "MCB": diff,
    }

    df = pd.concat([df, pd.DataFrame([adj_row])], ignore_index=True)
    return df

# --- Export / Totals / Excel -------------------------------------------------
def pdf_to_excel(
    pdf_path: str,
    output_excel_path: str,
    brands_to_extract: Optional[List[str]] = None,
    debug_unparsed_out: Optional[str] = None,
    mismatch_write_threshold: float = 50.0  # kept for signature compatibility
):
    pdf_path = Path(pdf_path)

    # text-first, OCR-fallback pipeline
    rows, first_page_text_hint, unparsed_lines, suspicious_rows = extract_rows_with_text_then_ocr(
        pdf_path,
        brands_to_extract=brands_to_extract,
        dpi=300,
        min_rows_for_text=50,
    )

    if not rows:
        print(f"⚠️ No data extracted from {pdf_path}. Check text/OCR or regex.")
        return

    df = pd.DataFrame(rows)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNS]

    df["Customer Info"] = df["Customer Info"].replace("", pd.NA).ffill()

    for col in ["Ordered", "Shipped", "Whlse", "MCB"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    extracted_total_mcb = float(df["MCB"].sum(skipna=True))

    report_total_guess = extract_report_total_from_pdf(pdf_path, first_page_text_hint)
    if report_total_guess is None:
        report_total_guess = 0.0

    diff = report_total_guess - extracted_total_mcb

    print(f"→ Extracted MCB sum: {extracted_total_mcb:.2f}")
    print(f"→ Report total in PDF (official): {report_total_guess:.2f}")
    print(f"→ Difference (pdf - extracted): {diff:.2f}")

    # Always write unparsed debug lines if we have any
    if debug_unparsed_out and unparsed_lines:
        try:
            Path(debug_unparsed_out).parent.mkdir(parents=True, exist_ok=True)
            with open(debug_unparsed_out, "w", encoding="utf-8") as f:
                f.write("\n".join(unparsed_lines))
            print(f"📝 Wrote {len(unparsed_lines)} suspicious unparsed lines to {debug_unparsed_out}")
        except Exception as e:
            print(f"[WARN] Failed writing debug lines to {debug_unparsed_out}: {e}")

    # write suspicious rows (calc vs OCR disagreed) if any
    if suspicious_rows:
        base_path = Path(output_excel_path)
        suspicious_path = base_path.with_name(base_path.stem + "_suspicious_rows.xlsx")
        try:
            suspicious_df = pd.DataFrame(suspicious_rows)
            suspicious_df.to_excel(suspicious_path, index=False)
            print(f"🧐 Wrote {len(suspicious_rows)} suspicious rows to {suspicious_path}")
        except Exception as e:
            print(f"[WARN] Failed writing suspicious rows to {suspicious_path}: {e}")

    # Adjustment row is effectively disabled (threshold=1e9)
    week_guess_series = df["Week Ending"].dropna().astype(str).replace("", pd.NA).dropna()
    week_guess = week_guess_series.iloc[0] if not week_guess_series.empty else ""

    cust_guess_series = df["Customer Info"].dropna().astype(str).replace("", pd.NA).dropna()
    cust_guess = cust_guess_series.iloc[0] if not cust_guess_series.empty else ""

    df = add_adjustment_row_if_needed(df, week_guess, cust_guess, diff, adjustment_threshold=1e9)

    Path(output_excel_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_excel_path, index=False)

    final_sum = float(df["MCB"].sum(skipna=True))
    qc_status = "OK" if abs(report_total_guess - final_sum) < 0.01 else "REVIEW"

    print(f"✅ Wrote {len(df)} rows to {output_excel_path}")
    print(f"   Final Extracted Total MCB: {final_sum:.2f}")
    print(f"   QC Status: {qc_status}")

# --- Batch runner ------------------------------------------------------------
def process_multiple_pdfs(
    input_dir: str,
    output_dir: str,
    brands_to_extract: Optional[List[str]] = None
):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = glob.glob(os.path.join(str(input_dir), '*.pdf')) + \
                glob.glob(os.path.join(str(input_dir), '*.PDF'))

    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        base = os.path.splitext(os.path.basename(pdf_file))[0]
        output_excel_file = output_dir / f"{base}_converted.xlsx"
        debug_txt = output_dir / f"{base}_unparsed_debug.txt"
        print(f"→ Processing: {pdf_file}")
        pdf_to_excel(
            pdf_file,
            str(output_excel_file),
            brands_to_extract=brands_to_extract,
            debug_unparsed_out=str(debug_txt),
            mismatch_write_threshold=50.0
        )

# --- Main --------------------------------------------------------------------
if __name__ == "__main__":
    input_dir = "/Users/joewilt/Desktop/Incoming Sovos files/"
    output_dir = "/Users/joewilt/Desktop/Converted Sovos files/"

    brands_to_extract = None  # None = all brands

    process_multiple_pdfs(input_dir, output_dir, brands_to_extract)


Sprouts KEHE Bulk

import re
import PyPDF2
import pandas as pd
import os

# Define the pattern for parsing the lines
pattern = re.compile(
    r"(?P<code>\w+)\s+(?P<upc>\d+)\s+(?P<description>.*?)\s+\$(?P<unit_price>[\d\.]+)\s+(?P<start_date>\d{2}/\d{2}/\d{4})\s+(?P<end_date>\d{2}/\d{2}/\d{4})\s+(?P<units>\d+)\s+(?P<quantity>\d+)\s+\$(?P<total_price>[\d\.]+)"
)


def process_line(line):
    match = pattern.match(line)
    if match:
        return match.groupdict()
    return None


def parse_pdf(file_path):
    # Extract text from PDF
    parsed_data = []
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            page_text = page.extract_text()
            lines = page_text.splitlines()

            # Process each line, including the first line of each page
            for line in lines:
                parsed_line = process_line(line)
                if parsed_line:
                    print(f"All parts matched successfully!\n{parsed_line}")
                    parsed_data.append(parsed_line)
                else:
                    print(f"No match for line: {line}")

    return parsed_data


def export_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Data exported successfully to {output_file}")


def process_pdfs_in_folder(input_folder, output_folder):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Process each PDF in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith('.pdf'):
            pdf_file_path = os.path.join(input_folder, file_name)
            parsed_data = parse_pdf(pdf_file_path)

            if parsed_data:
                # Create output Excel file path
                excel_file_name = file_name.replace('.pdf', '.xlsx')
                excel_file_path = os.path.join(output_folder, excel_file_name)
                export_to_excel(parsed_data, excel_file_path)
            else:
                print(f"No valid data was parsed from {file_name}.")


if __name__ == "__main__":
    # Set the input and output folder paths
    input_folder = "/Users/joewilt/Desktop/Incoming Sovos files"  # Input folder path
    output_folder = "/Users/joewilt/Desktop/Converted Sovos files"  # Output folder path

    # Process all PDFs in the input folder and export to the output folder
    process_pdfs_in_folder(input_folder, output_folder)


UNFI East

import os
import re
import glob
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd

# Optional OCR deps (only used if USE_OCR=True)
try:
    from pdf2image import convert_from_path
    import pytesseract
except Exception:
    convert_from_path = None
    pytesseract = None

# ----------------------------
# Config
# ----------------------------
USE_OCR = False                         # Set True if the PDF is scanned images (requires Poppler + Tesseract)
TESSERACT_CONFIG = r"--oem 3 --psm 6"   # Only used when USE_OCR=True

# Output columns (now including Customer)
COLUMNS = [
    "Customer",
    "Prod", "Brand", "Pack/Size", "Prod Description",
    "Inv #", "Whse", "Qty", "Chb%%", "Chb $$$$", "Inv $$$$", "Authorized",
]

# Regex tailored to rows like:
# 022963 PACIFC   12 32 FZ    SOUP,OG2,CRMY RSTD 83763562 MAN     1     7      3.83      54.65
ROW_REGEX = re.compile(
    r"""
    ^\s*
    (?P<prod>\d{5,6})\s+                             # Prod (5-6 digits to be safe)
    (?P<brand>[A-Z]+)\s+                              # Brand (caps; if you see brands with spaces, we can widen this)
    (?P<packsize>\d+\s+\d+(?:\.\d+)?\s+(?:FZ|OZ))\s+  # Pack/Size: "12 32 FZ" or "12 10.5 OZ"
    (?P<desc>.+?)\s+                                  # Description (lazy)
    (?P<invoice>\d{8})\s+                             # Inv # (8 digits in samples)
    (?P<whse>[A-Z]{3})\s+                             # Whse (MAN, HVA, DAY, CHE, etc.)
    (?P<qty>\d+)\s+                                   # Qty
    (?P<chb_pct>\d+(?:\.\d+)?)\s+                     # Chb%% (numeric)
    (?P<chb_dollars>\d+\.\d{2})\s+                    # Chb $$$$
    (?P<inv_dollars>\d+\.\d{2})                       # Inv $$$$
    (?:\s+(?P<authorized>.+))?                        # Authorized (optional; often blank)
    \s*$
    """,
    re.VERBOSE
)

# Phrases to ignore (headers, totals, boilerplate)
HEADER_MARKERS = (
    "Vendor Chargeback Report", "WKLY_CHBK_RPT",
    "Week Ending:", "Prod  Brand", "Customer Total",
    "CHARGEBACK CATEGORY", "Page", "United Natural Foods, Inc.",
    "----       --------- ----------", "(MCB Process Type"
)


def is_header_or_noise(line: str) -> bool:
    """Filter out headers, separators, boilerplate lines."""
    if not line.strip():
        return True
    # IMPORTANT: no longer filter out 'Customer:' here; we handle it in the parser
    if any(k in line for k in HEADER_MARKERS):
        return True
    return False


def parse_lines_to_rows(lines: List[str]) -> List[Dict[str, str]]:
    """
    Parse a flat list of lines into row dicts.
    Tracks the current 'Customer:' and attaches it to each product row.
    """
    rows: List[Dict[str, str]] = []
    current_customer: Optional[str] = None

    for raw in lines:
        line = raw.rstrip()
        stripped = line.strip()

        # Capture customer lines, e.g.:
        # "Customer: 3 BIG Y #36 PLAINVILLE, PLAINVILLE,CT (BYN)"
        if stripped.startswith("Customer:"):
            # Store full value after "Customer:"
            current_customer = stripped[len("Customer:"):].strip()
            # Nothing else to do with this line
            continue

        # Skip headers / noise
        if is_header_or_noise(line):
            continue

        # Try to match a product row
        m = ROW_REGEX.match(line)
        if m:
            d = m.groupdict()
            rows.append({
                "Customer": current_customer or "",
                "Prod": d["prod"],
                "Brand": d["brand"],
                "Pack/Size": d["packsize"],
                "Prod Description": d["desc"].strip(),
                "Inv #": d["invoice"],
                "Whse": d["whse"],
                "Qty": d["qty"],
                "Chb%%": d["chb_pct"],
                "Chb $$$$": d["chb_dollars"],
                "Inv $$$$": d["inv_dollars"],
                "Authorized": (d.get("authorized") or "").strip(),
            })

    return rows


def extract_text_pages_pdfplumber(pdf_path: Path) -> List[str]:
    """Native text extraction (fast, reliable for digital PDFs)."""
    import pdfplumber
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Mildly tighter tolerances help keep columns together
            text = page.extract_text(x_tolerance=1.5, y_tolerance=1.5) or ""
            pages_text.append(text)
    return pages_text


def extract_text_pages_ocr(
    pdf_path: Path,
    dpi: int = 300,
    first_page: Optional[int] = None,
    last_page: Optional[int] = None,
) -> List[str]:
    """OCR fallback for scanned/image PDFs."""
    if convert_from_path is None or pytesseract is None:
        raise RuntimeError("OCR requested but pdf2image/pytesseract not available. Install Poppler + Tesseract.")
    pages_text = []
    p = first_page or 1
    while True:
        if last_page is not None and p > last_page:
            break
        try:
            img = convert_from_path(str(pdf_path), first_page=p, last_page=p, dpi=dpi)[0]
        except Exception:
            break  # No more pages
        text = pytesseract.image_to_string(img, config=TESSERACT_CONFIG)
        pages_text.append(text)
        p += 1
    return pages_text


def pdf_to_rows(pdf_path: Path, debug_dump_dir: Optional[Path] = None) -> List[Dict[str, str]]:
    """
    Convert a single PDF into parsed row dicts using native text, then optional OCR.
    Now parses *across* pages in one pass so customer context carries across page breaks.
    """
    # 1) Native text
    pages_text = extract_text_pages_pdfplumber(pdf_path)

    # Optional: dump the raw text for debugging/training the regex
    if debug_dump_dir:
        debug_dump_dir.mkdir(parents=True, exist_ok=True)
        for i, t in enumerate(pages_text, 1):
            (debug_dump_dir / f"{pdf_path.stem}_page_{i:03d}.txt").write_text(t)

    # Flatten all pages into one list of lines so 'Customer:' state persists across pages
    all_lines: List[str] = []
    for text in pages_text:
        all_lines.extend(text.splitlines())

    all_rows = parse_lines_to_rows(all_lines)

    # 2) OCR fallback
    if not all_rows and USE_OCR:
        print(f"   ↳ No rows from native text; trying OCR for: {pdf_path.name}")
        pages_text = extract_text_pages_ocr(pdf_path)
        if debug_dump_dir:
            for i, t in enumerate(pages_text, 1):
                (debug_dump_dir / f"{pdf_path.stem}_page_ocr_{i:03d}.txt").write_text(t)

        all_lines = []
        for text in pages_text:
            all_lines.extend(text.splitlines())

        all_rows = parse_lines_to_rows(all_lines)

    return all_rows


def pdf_to_excel(pdf_path: str, output_excel_path: str, debug_dump_dir: Optional[str] = None):
    """Process one PDF → Excel."""
    pdf_path = Path(pdf_path)
    debug_dir = Path(debug_dump_dir) if debug_dump_dir else None

    print(f"→ Reading: {pdf_path}")
    rows = pdf_to_rows(pdf_path, debug_dump_dir=debug_dir)
    print(f"→ Parsed rows: {len(rows)}")

    if not rows:
        print(
            "⚠️ No data extracted. If this is a scanned PDF, set USE_OCR=True "
            "and make sure Poppler + Tesseract are installed."
        )
        return

    df = pd.DataFrame(rows, columns=COLUMNS)
    out_path = Path(output_excel_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, index=False)
    print(f"✅ Wrote {len(df)} rows → {out_path}")


def process_multiple_pdfs(input_dir: str, output_dir: str, debug_dump_dir: Optional[str] = None):
    """Process all PDFs in a folder (case-insensitive *.pdf)."""
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = glob.glob(os.path.join(str(input_dir), "*.pdf")) + \
                glob.glob(os.path.join(str(input_dir), "*.PDF"))

    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        try:
            base = os.path.splitext(os.path.basename(pdf_file))[0]
            output_excel_file = output_dir / f"{base}_converted.xlsx"
            print(f"→ Processing: {pdf_file}")
            pdf_to_excel(pdf_file, str(output_excel_file), debug_dump_dir=debug_dump_dir)
        except Exception as e:
            print(f"❌ Error processing {pdf_file}: {e}")


# ----------------------------
# RUN IT (update these paths if not using mac)
# ----------------------------
if __name__ == "__main__":
    process_multiple_pdfs(
        input_dir="/Users/joewilt/Desktop/Incoming Sovos files/",        # ← change to your input folder
        output_dir="/Users/joewilt/Desktop/Converted Sovos files/",       # ← change to your output folder
        debug_dump_dir="/Users/joewilt/Desktop/Converted Sovos files/_debug_text"  # optional; set None to skip
    )



Bulk UNFI West

import pdfplumber
import pandas as pd
import re
import os
import glob


# Function to extract data from a PDF file
def extract_data_from_pdf(pdf_path, brands_to_extract):
    data = []
    week_ending = None  # Initialize week_ending variable
    customer_info = None  # Initialize customer_info variable to persist across pages

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')

            # Extract the week ending date
            if not week_ending:
                for line in lines:
                    if "Week ending" in line:
                        week_ending = line.split(':')[-1].strip()
                        break

            # Extract customer information and the table data
            for line in lines:
                if line.startswith("Customer :"):
                    customer_info = line.split("Customer :")[1].strip()

                if any(brand in line.upper() for brand in brands_to_extract) and week_ending and customer_info:
                    # Debug print to check the line content
                    print(f"Processing line: {line}")

                    # Split the line by whitespace
                    columns = line.split()

                    # Find the indices for known columns to handle description extraction
                    try:
                        invoice_idx = next(i for i, col in enumerate(columns) if col.isdigit() and len(col) == 9)
                        ordered_idx = invoice_idx + 1
                        shipped_idx = ordered_idx + 1
                        whlse_idx = shipped_idx + 1
                        total_discount_idx = whlse_idx + 1
                        mcb_percentage_idx = total_discount_idx + 1
                        mcb_idx = mcb_percentage_idx + 1

                        brand = columns[0]
                        product = columns[1]
                        unit = columns[2]

                        # Capture the unit with OZ if it's part of the unit
                        if not unit.endswith("OZ") and len(columns) > 3:
                            unit += " " + columns[3]
                            description_start_index = 4
                        else:
                            description_start_index = 3

                        description = " ".join(columns[description_start_index:invoice_idx])
                        invoice = columns[invoice_idx]
                        ordered = columns[ordered_idx]
                        shipped = columns[shipped_idx]
                        whlse = columns[whlse_idx]
                        total_discount = columns[total_discount_idx]
                        mcb_percentage = columns[mcb_percentage_idx]
                        mcb = columns[mcb_idx]

                        data.append([
                            week_ending, customer_info, brand, product, unit,
                            description, invoice, ordered, shipped, whlse,
                            total_discount, mcb_percentage, mcb
                        ])

                        # Debug print to check extracted data
                        print(f"Extracted data: {data[-1]}")
                    except StopIteration:
                        print(f"Failed to parse line: {line}")

    return data


# Function to convert PDF to Excel
def pdf_to_excel(pdf_path, output_excel_path, brands_to_extract):
    data = extract_data_from_pdf(pdf_path, brands_to_extract)

    if not data:
        print(f"No data extracted from {pdf_path}. Please check the PDF content and extraction logic.")
        return

    # Create a DataFrame
    columns = ["Week Ending", "Customer Info", "Brand", "Product", "Unit", "Description", "Invoice",
               "Ordered", "Shipped", "Whlse", "Total Discount %", "MCB %", "MCB"]

    df = pd.DataFrame(data, columns=columns)

    # Save to Excel
    df.to_excel(output_excel_path, index=False)
    print(f"Data has been written to {output_excel_path}")


# Function to process multiple PDFs in a directory
def process_multiple_pdfs(input_dir, output_dir, brands_to_extract):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + glob.glob(os.path.join(input_dir, '*.PDF'))
    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")
        # Define the output Excel file path
        output_excel_file = os.path.join(output_dir,
                                         os.path.splitext(os.path.basename(pdf_file))[0] + '_converted.xlsx')
        pdf_to_excel(pdf_file, output_excel_file, brands_to_extract)


# Example usage
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'
brands_to_extract = ["RAOS", "NOOSA", "*NOOSA", "MICHAEL ANGELO'S", "PACIFIC", "LATE JULY", "PACIFIC", "@RAOS"]
process_multiple_pdfs(input_dir, output_dir, brands_to_extract)

KEHE Bulk

from pdfminer.high_level import extract_text
import pandas as pd
import re
import os
import glob

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = extract_text(pdf_path)
    return text

# Parsing function tailored to the provided text format
def parse_text_to_table(text):
    lines = text.split('\n')
    table_data = []
    payee = None
    invoice_number = None
    sold_to = ""
    sold_to_lines = []
    date = None

    for i, line in enumerate(lines):
        line = line.strip()
        if 'PAYEE:' in line and payee is None:
            payee = lines[i + 1].strip()
        elif 'INVOICE #' in line:
            invoice_number_match = re.search(r'INVOICE #(\d+)', line)
            if invoice_number_match:
                invoice_number = invoice_number_match.group(1).strip()
        elif 'SOLD TO:' in line:
            # Start capturing "SOLD TO" information
            sold_to_lines = [line.replace('SOLD TO:', '').strip()]
            # Continue to capture the next lines that belong to "SOLD TO"
            j = i + 1
            while j < len(lines) and lines[j].strip() and not re.match(r'^\d{12}', lines[j].strip()):
                sold_to_lines.append(lines[j].strip())
                j += 1
            sold_to = ' '.join(sold_to_lines).replace('TOL User : EMDROBOT', '').strip()
        elif re.match(r'^\d{1,2}/\d{2}/\d{2}$', line):  # Check for mm/dd/yy or m/dd/yy format
            date = line  # Assuming this line contains the date
        elif re.match(r'^\d{12}', line):  # Line starts with 12-digit UPC parts
            parts = re.split(r'\s+', line)
            if len(parts) < 10:
                continue  # Skip if the line does not have enough parts

            upc = parts[0]
            qty_ship = parts[1]
            # Determine where description ends by checking the remaining elements from the right
            description_end_idx = len(parts) - 7  # Description ends where the last 6 elements start
            description = ' '.join(parts[2:description_end_idx])
            nbr = parts[-6]
            date = parts[-5]
            comment = parts[-4]
            cost = parts[-3]
            disc = parts[-2]
            ext_cost = parts[-1]

            # Append the data for the current line to the table_data list
            table_data.append(
                [payee, sold_to, invoice_number, upc, qty_ship, description, nbr, date, comment, cost, disc, ext_cost])
        elif '-' in line:  # Handling lines with data separated by dashes
            parts = line.split('-')
            if len(parts) == 8:
                upc = parts[0].strip()
                qty_ship = parts[1].strip()
                description = parts[2].strip()
                nbr = parts[3].strip()
                date = parts[4].strip()
                comment = parts[5].strip()
                cost = parts[6].strip()
                disc = parts[7].strip()
                ext_cost = parts[8].strip()
                table_data.append(
                    [payee, sold_to, invoice_number, upc, qty_ship, description, nbr, date, comment, cost, disc, ext_cost])

    return table_data

# Main function to extract, parse, and write to Excel
def pdf_to_excel(pdf_path, output_excel_path):
    try:
        extracted_text = extract_text_from_pdf(pdf_path)
        parsed_data = parse_text_to_table(extracted_text)

        # Convert parsed data to a DataFrame
        df = pd.DataFrame(parsed_data,
                          columns=['Payee', 'Sold to', 'Invoice #', 'UPC#', 'QTY ship', 'Description', 'NBR', 'Date',
                                   'Comment', 'Cost', 'Disc $ or %', 'EXT-Cost'])

        # Write DataFrame to an Excel file
        df.to_excel(output_excel_path, index=False)
        print(f"Data has been written to {output_excel_path}")
    except Exception as e:
        print(f"Failed to process {pdf_path}: {e}")

# Function to process multiple PDFs in a directory
def process_multiple_pdfs(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + glob.glob(os.path.join(input_dir, '*.PDF'))
    print(f"Found {len(pdf_files)} PDF files in {input_dir}")

    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")
        # Define the output Excel file path
        output_excel_file = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_file))[0] + '_converted.xlsx')
        pdf_to_excel(pdf_file, output_excel_file)

# Example usage
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'
process_multiple_pdfs(input_dir, output_dir)


UNFI The Fresh Market

import os
import re
import glob
import pytesseract
import pandas as pd
from pdf2image import convert_from_path

# Set the Tesseract command path (adjust if necessary)
pytesseract.pytesseract.tesseract_cmd = '/opt/homebrew/bin/tesseract'

COLUMNS = [
    "Supplier Name",
    "Supplier Brand",
    "Customer Invoice",
    "Program Description",
    "Promo Period",
    "UPC Code",
    "Item Description",
    "Chargeback Quantity",
    "Chargeback Amount",
    "Total Chargeback",
    "Customer Notes"
]


def ocr_extract_text(file_path, dpi=300):
    """Convert PDF pages to images and perform OCR using Tesseract."""
    try:
        pages = convert_from_path(file_path, dpi=dpi, poppler_path='/opt/homebrew/bin')
        texts = []
        for page in pages:
            text = pytesseract.image_to_string(page, config="--psm 6")
            texts.append(text)
        return texts
    except Exception as e:
        print(f"Error converting PDF to images for {file_path}: {e}")
        return []


def parse_table_row(row_text):
    """Parse a single table row with debug output."""
    # Debug: Print the row being processed
    print(f"Processing row: {row_text}")

    # Remove extra whitespace and normalize spaces
    row_text = ' '.join(row_text.split())

    # Try to match the row pattern
    # Looking for: Supplier Name, Brand, Invoice, Program, Period, UPC, Description, Total, Notes
    pattern = r"""
        (RAO'S\s+SPECIALTY\s+FOODS[,\.]?\s*INC\.?)\s+  # Supplier Name
        (RAOS)\s+                                      # Brand
        (\d+)\s+                                       # Invoice Number
        (OTHER)\s+                                     # Program Description
        (\d{2}\.\d{2}\.\d{4})\s+                      # Promo Period
        (\d{11,12})\s+                                # UPC Code
        (.+?)\s+                                      # Item Description
        (\$\d+\.\d{2})\s*                            # Total Chargeback
        (.+)                                          # Customer Notes
    """

    match = re.match(pattern, row_text, re.VERBOSE)

    if match:
        print("Row matched pattern!")
        return [
            match.group(1).strip(),  # Supplier Name
            match.group(2).strip(),  # Brand
            match.group(3).strip(),  # Invoice
            match.group(4).strip(),  # Program
            match.group(5).strip(),  # Period
            match.group(6).strip(),  # UPC
            match.group(7).strip(),  # Description
            "",  # Chargeback Quantity (empty)
            "",  # Chargeback Amount (empty)
            match.group(8).strip(),  # Total
            match.group(9).strip()  # Notes
        ]

    print("Row did not match pattern")
    return None


def find_and_parse_tables(texts):
    """Find and parse tables from all pages of the PDF."""
    all_rows = []

    for page_num, text in enumerate(texts, 1):
        print(f"\nProcessing page {page_num}")
        print("=" * 50)
        print("Page content snippet:")
        print(text[:500])  # Print first 500 chars of the page

        lines = text.split('\n')
        table_start = -1

        # Find table header
        for i, line in enumerate(lines):
            if ('Supplier Name' in line and 'UPC Code' in line) or \
                    ('Supplier Name' in line and 'Customer Invoice' in line):
                table_start = i + 1
                print(f"Found table header at line {i}")
                break

        if table_start != -1:
            print("\nProcessing table rows:")
            # Process each line after the headers
            for line in lines[table_start:]:
                if line.strip():  # Only process non-empty lines
                    print(f"\nChecking line: {line}")
                    # More flexible row detection
                    if any(keyword in line for keyword in ["RAO", "SPECIALTY", "FOODS"]):
                        parsed_row = parse_table_row(line)
                        if parsed_row:
                            print("Successfully parsed row!")
                            all_rows.append(parsed_row)
                        else:
                            print("Failed to parse row")

    print(f"\nTotal rows parsed: {len(all_rows)}")
    return all_rows


def process_pdf_file(file_path, output_dir):
    """Process a single PDF file."""
    print(f"\nProcessing {file_path} with OCR...")
    texts = ocr_extract_text(file_path)
    if not texts:
        print(f"No text extracted from {file_path} after OCR.")
        return

    table_rows = find_and_parse_tables(texts)
    if not table_rows:
        print(f"No table rows detected in {file_path}. Check parsing logic.")
        return

    df = pd.DataFrame(table_rows, columns=COLUMNS)
    base = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(output_dir, f"{base}_converted.xlsx")

    try:
        df.to_excel(output_file, index=False)
        print(f"Excel file created: {output_file}")
    except Exception as e:
        print(f"Error writing {output_file}: {e}")


def process_multiple_pdfs(input_dir, output_dir):
    """Process all PDFs in the input directory."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf')) + \
                glob.glob(os.path.join(input_dir, '*.PDF'))

    print(f"Found {len(pdf_files)} PDF files in {input_dir}")
    for file_path in pdf_files:
        process_pdf_file(file_path, output_dir)


# Directory paths
input_dir = '/Users/joewilt/Desktop/Incoming Sovos files/'
output_dir = '/Users/joewilt/Desktop/Converted Sovos files/'

# Process all PDFs
process_multiple_pdfs(input_dir, output_dir)
