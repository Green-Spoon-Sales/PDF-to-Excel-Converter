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
