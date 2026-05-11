"""
invoice_parser.py
Hebrew/RTL Invoice OCR Text Parser
===================================
Takes raw OCR text → returns structured JSON.
Handles: ח.פ/ע.מ extraction, Hebrew keywords, RTL amounts, VAT detection.
"""

import re
import json
from typing import Optional, List, Dict, Any


# ══════════════════════════════════════════════════════════
#  UTILITIES
# ══════════════════════════════════════════════════════════

def clean_amount(raw: str) -> Optional[float]:
    """
    Strip currency symbols, RTL marks, commas → float.
    Handles: "193.52 ₪", "₪ 1,234.00", "‏2,500" etc.
    """
    if not raw:
        return None
    # Remove: ₪, Unicode directional marks (U+200E/F, U+202A-202E), spaces, commas
    cleaned = re.sub(r'[₪$€£\s,‎‏‪‫‬‭‮]', '', str(raw))
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return None


def first_hebrew_line(lines: List[str], stop_re: re.Pattern, max_lines: int = 15) -> Optional[str]:
    """Return first Hebrew-containing line before a stop keyword."""
    for line in lines[:max_lines]:
        if stop_re.search(line):
            break
        if len(line.strip()) < 3:
            continue
        if re.match(r'^[\d\s./:\-,₪%()]+$', line.strip()):
            continue
        if re.search(r'[֐-׿]', line):   # Hebrew char range
            return line.strip()
    return None


# ══════════════════════════════════════════════════════════
#  1. VENDOR INFO
# ══════════════════════════════════════════════════════════

def extract_vendor_info(text: str) -> Dict[str, Any]:
    """
    Extract:
      id   → 9-digit ח.פ / ע.מ number
      name → first descriptive Hebrew string in header
    """

    # ── Business ID ──────────────────────────────────────
    # Matches: ח.פ, ח"פ, ע.מ, עוסק מורשה (with optional spaces/colons)
    id_pattern = re.compile(
        r'(?:ח\.?פ\.?|ח"פ|ע\.?מ\.?|עוסק\s*מורשה|company\s*(?:no|id))[:\s#]*(\d{9})',
        re.IGNORECASE
    )
    id_match = id_pattern.search(text)
    vendor_id = id_match.group(1) if id_match else None

    # Fallback: bare 9-digit number that looks like a business ID
    if not vendor_id:
        bare = re.search(r'(?<!\d)(\d{9})(?!\d)', text)
        if bare:
            vendor_id = bare.group(1)

    # ── Vendor Name ──────────────────────────────────────
    stop_re = re.compile(
        r'חשבונית|מס[׳\']|תאריך|לכבוד|bill\s*to|invoice|date',
        re.IGNORECASE
    )
    lines = text.splitlines()
    name = first_hebrew_line(lines, stop_re)

    # Fallback: look for text on the same line as ח.פ
    if not name and id_match:
        surrounding = text[max(0, id_match.start()-120):id_match.start()]
        candidate = surrounding.strip().splitlines()
        for c in reversed(candidate):
            c = c.strip()
            if c and re.search(r'[֐-׿]', c):
                name = c
                break

    return {"id": vendor_id, "name": name}


# ══════════════════════════════════════════════════════════
#  2. INVOICE METADATA
# ══════════════════════════════════════════════════════════

def extract_invoice_metadata(text: str) -> Dict[str, str]:
    """
    Extract:
      doc_number → invoice serial number
      date       → normalized DD/MM/YYYY
    """

    # ── Document Number ──────────────────────────────────
    doc_pattern = re.compile(
        r'(?:'
        r'חשבונית\s*(?:מס[׳\']?|מספר|#)'   # חשבונית מס׳ / מספר
        r'|מס[׳\']?\s*(?:חשבונית|מסמך)'    # מס׳ חשבונית
        r'|invoice\s*(?:no\.?|number|#)'    # English
        r')[:\s#]*([A-Z0-9][A-Z0-9\-/]{1,29})',
        re.IGNORECASE
    )
    doc_match = doc_pattern.search(text)
    doc_number = doc_match.group(1).strip() if doc_match else None

    # Fallback: standalone alphanumeric ID pattern
    if not doc_number:
        fb = re.search(r'\b([A-Z]{1,4}[-/]?\d{4,12})\b', text)
        if fb:
            doc_number = fb.group(1)

    # ── Date ─────────────────────────────────────────────
    # Look near the keyword "תאריך" first, then anywhere
    date_re = re.compile(r'(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})')

    anchored = re.search(r'תאריך[:\s]*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})', text)
    raw_date = anchored.group(1) if anchored else None

    if not raw_date:
        dm = date_re.search(text)
        raw_date = dm.group(0) if dm else None

    date_str = None
    if raw_date:
        parts = re.split(r'[/.\-]', raw_date)
        if len(parts) == 3:
            d, m, y = parts
            y = ('20' + y) if len(y) == 2 else y
            date_str = f"{d.zfill(2)}/{m.zfill(2)}/{y}"

    return {"doc_number": doc_number, "date": date_str}


# ══════════════════════════════════════════════════════════
#  3. LINE ITEMS
# ══════════════════════════════════════════════════════════

def extract_line_items(text: str) -> List[Dict[str, Any]]:
    """
    Extract table rows between header keywords (תיאור/פירוט)
    and footer keywords (סה"כ/סכום).
    """
    items: List[Dict[str, Any]] = []

    header_kw = re.compile(r'(?:תיאור|פירוט|שירות|מוצר|description|item)', re.IGNORECASE)
    footer_kw = re.compile(r'(?:סה["״]כ|סכום\s*כולל|subtotal|total)', re.IGNORECASE)

    h_match = header_kw.search(text)
    f_match = footer_kw.search(text)

    start = h_match.start() if h_match else 0
    end   = f_match.start() if (f_match and f_match.start() > start) else len(text)
    table = text[start:end]

    skip_re = re.compile(
        r'(?:תיאור|פירוט|כמות|מחיר|description|qty|unit|price|amount)',
        re.IGNORECASE
    )

    # ── Full row: description  qty  unit_price  line_total ──
    row_re = re.compile(
        r'^(?P<description>.{3,60}?)\s{2,}'       # description
        r'(?P<qty>\d+(?:\.\d+)?)\s+'               # qty
        r'(?P<unit_price>[\d,]+(?:\.\d{1,2})?)\s+' # unit price
        r'(?P<line_total>[\d,]+(?:\.\d{1,2})?)$',  # line total
        re.MULTILINE
    )
    for m in row_re.finditer(table):
        desc = m.group('description').strip()
        if skip_re.search(desc):
            continue
        items.append({
            "description": desc,
            "qty":         float(m.group('qty')),
            "unit_price":  clean_amount(m.group('unit_price')),
            "line_total":  clean_amount(m.group('line_total')),
        })

    # ── Fallback: description + amount on same line ──────
    if not items:
        simple_re = re.compile(
            r'(?P<description>[^\n\d₪]{5,80?})\s+'
            r'(?P<amount>[\d,]+\.?\d{0,2})\s*₪?',
        )
        for m in simple_re.finditer(table):
            desc = m.group('description').strip()
            if skip_re.search(desc) or not re.search(r'[֐-׿A-Za-z]', desc):
                continue
            amt = clean_amount(m.group('amount'))
            if amt and amt > 0:
                items.append({
                    "description": desc,
                    "qty":         1,
                    "unit_price":  amt,
                    "line_total":  amt,
                })

    return items


# ══════════════════════════════════════════════════════════
#  4. FINANCIAL SUMMARY
# ══════════════════════════════════════════════════════════

def extract_financial_summary(text: str) -> Dict[str, Any]:
    """
    Extract subtotal, vat_rate (17/18/21/0), vat_amount, total_amount.
    RTL-safe: amounts can appear before or after labels.
    """

    def find_amount(patterns: List[str]) -> Optional[float]:
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
            if m:
                val = clean_amount(m.group(1))
                if val and val > 0:
                    return val
        return None

    # ── VAT rate ─────────────────────────────────────────
    # Detect 17%, 18%, 21%, 0%
    vat_rate = 18  # Israeli default
    vat_rate_re = re.compile(
        r'(?:מע["״]מ|vat)[:\s]*(\d{1,2})\s*%'   # מע"מ 18%
        r'|(\d{1,2})\s*%\s*(?:מע["״]מ|vat)',     # 18% מע"מ
        re.IGNORECASE
    )
    vr_match = vat_rate_re.search(text)
    if vr_match:
        raw = int(vr_match.group(1) or vr_match.group(2))
        if raw in (0, 17, 18, 21):
            vat_rate = raw

    # ── Subtotal (before VAT) ────────────────────────────
    subtotal = find_amount([
        r'(?:סה["״]כ\s*לפני\s*מע["״]מ|סכום\s*חייב|בסיס\s*מס)[:\s₪]*([\d,]+\.?\d*)',
        r'(?:subtotal|net\s*amount|amount\s*before\s*tax)[:\s]*([\d,]+\.?\d*)',
    ])

    # ── VAT amount ───────────────────────────────────────
    vat_amount = find_amount([
        r'(?:מע["״]מ|vat)[:\s₪]*([\d,]+\.?\d*)(?!\s*%)',
        r'(?:tax\s*amount)[:\s]*([\d,]+\.?\d*)',
    ])

    # ── Total (including VAT) ────────────────────────────
    total = find_amount([
        r'(?:סה["״]כ\s*לתשלום|סה["״]כ\s*כולל\s*מע["״]מ|לתשלום)[:\s₪]*([\d,]+\.?\d*)',
        r'(?:total\s*(?:amount\s*)?due|grand\s*total|amount\s*due)[:\s]*([\d,]+\.?\d*)',
        r'(?:סה["״]כ)[:\s₪]*([\d,]+\.?\d*)',
    ])

    # ── Cross-derivation ─────────────────────────────────
    if not subtotal and total and vat_rate:
        subtotal = round(total / (1 + vat_rate / 100), 2)
    if not vat_amount and subtotal and vat_rate:
        vat_amount = round(subtotal * vat_rate / 100, 2)
    if not total and subtotal and vat_amount:
        total = round(subtotal + vat_amount, 2)

    return {
        "subtotal":     subtotal,
        "vat_rate":     vat_rate,
        "vat_amount":   vat_amount,
        "total_amount": total,
    }


# ══════════════════════════════════════════════════════════
#  5. VENDOR DB UPSERT
# ══════════════════════════════════════════════════════════

def upsert_vendor(vendor_info: Dict, db: Dict) -> str:
    """
    Primary key = vendor_id (ח.פ).
    Returns: 'created' | 'updated' | 'skipped'
    """
    vid = vendor_info.get("id")
    if not vid:
        return "skipped"
    if vid in db:
        db[vid].update({k: v for k, v in vendor_info.items() if v is not None})
        return "updated"
    db[vid] = vendor_info.copy()
    return "created"


# ══════════════════════════════════════════════════════════
#  MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════

def parse_hebrew_invoice(
    ocr_text: str,
    vendor_db: Optional[Dict] = None
) -> Dict[str, Any]:
    """
    Parse raw OCR text from a Hebrew invoice.

    Args:
        ocr_text:   Raw text string from OCR engine.
        vendor_db:  Optional dict {vendor_id: vendor_info} — simulates your DB.
                    If provided, vendor is upserted (created or updated).

    Returns:
        {
          "vendor_info":       {id, name},
          "invoice_metadata":  {doc_number, date},
          "line_items":        [{description, qty, unit_price, line_total}, ...],
          "financial_summary": {subtotal, vat_rate, vat_amount, total_amount},
          "_meta":             {db_action, confidence}
        }
    """
    vendor_info = extract_vendor_info(ocr_text)
    metadata    = extract_invoice_metadata(ocr_text)
    line_items  = extract_line_items(ocr_text)
    financial   = extract_financial_summary(ocr_text)

    db_action = upsert_vendor(vendor_info, vendor_db) if vendor_db is not None else None

    # Confidence score: how many key fields were found
    found = sum([
        bool(vendor_info.get("id")),
        bool(vendor_info.get("name")),
        bool(metadata.get("doc_number")),
        bool(metadata.get("date")),
        bool(financial.get("total_amount")),
        bool(line_items),
    ])
    confidence = round(found / 6, 2)

    return {
        "vendor_info":       vendor_info,
        "invoice_metadata":  metadata,
        "line_items":        line_items,
        "financial_summary": financial,
        "_meta": {
            "db_action":  db_action,
            "confidence": confidence,   # 0.0 – 1.0
        }
    }


# ══════════════════════════════════════════════════════════
#  EXAMPLE / TEST
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    sample_ocr = """
    חברת הוק בע"מ
    ח.פ 512345678
    רח' הירקון 22, תל אביב

    חשבונית מס׳: HOK-2026-041
    תאריך: 10/05/2026

    תיאור                    כמות    מחיר יח'    סה"כ
    ארוחות בוקר – מאי 2026     1      27,966      27,966

    סה"כ לפני מע"מ:   27,966 ₪
    מע"מ 18%:          5,034 ₪
    סה"כ לתשלום:      33,000 ₪
    """

    vendor_database: Dict = {}   # your in-memory / sheet-backed store

    result = parse_hebrew_invoice(sample_ocr, vendor_db=vendor_database)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nVendor DB after upsert: {vendor_database}")
