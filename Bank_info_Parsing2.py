def parse_pdf(pdfpath):
    import re
    import pdfplumber
    from collections import defaultdict

    COLUMN_ORDER = ["Date", "Description", "Category", "Amount", "Current Balance"]
    HEADER_TARGETS = {"Date", "Description", "Category", "Amount", "Current"}
    DATE_RE = re.compile(r"^[A-Z][a-z]{2}\s+\d{1,2},\s+20\d{2}$")
    JUNK_RE = re.compile(r"^(https?://|\d+/\d+\s*$|\d+/\d+/\d+,\s*\d+:\d+|Bank Account Summary)")
    ROW_TOLERANCE = 3

    def find_column_bounds(page):
        """Use the header row on page 1 to derive column x-ranges."""
        words = page.extract_words()
        # Group header words by approximate top
        by_top = {}
        for w in words:
            if w["text"] in HEADER_TARGETS:
                by_top.setdefault(round(w["top"]), []).append(w)
        # Pick the top with the most matches (the real header row)
        header_top = max(by_top, key=lambda t: len(by_top[t]))
        header_words = sorted(by_top[header_top], key=lambda w: w["x0"])

        # Build (column_name, x_left) pairs, merging "Current"+"Balance"
        starts = []
        for w in header_words:
            starts.append(("Current Balance" if w["text"] == "Current" else w["text"], w["x0"]))

        # Convert to (lo, hi) ranges using next column's left edge
        bounds = {}
        for i, (name, x) in enumerate(starts):
            right = starts[i + 1][1] if i + 1 < len(starts) else float("inf")
            bounds[name] = (x - 2, right - 2)  # tiny nudge for safety
        return bounds

    def assign_column(x_center, bounds):
        for name, (lo, hi) in bounds.items():
            if lo <= x_center < hi:
                return name
        return None

    def cluster_rows(words, bounds):
        """Group words into row-bands by y-coordinate, with column assignments."""
        enriched = []
        for w in words:
            x_center = (w["x0"] + w["x1"]) / 2
            col = assign_column(x_center, bounds)
            if col:
                enriched.append({
                    "text": w["text"],
                    "x0": w["x0"],
                    "y": (w["top"] + w["bottom"]) / 2,
                    "col": col,
                })
        enriched.sort(key=lambda d: d["y"])

        rows = []
        current = []
        band_y = None
        for w in enriched:
            if band_y is None or abs(w["y"] - band_y) <= ROW_TOLERANCE:
                current.append(w)
                band_y = sum(d["y"] for d in current) / len(current)
            else:
                rows.append(current)
                current = [w]
                band_y = w["y"]
        if current:
            rows.append(current)

        # Convert each band into {column: text} and remember the row's y
        banded = []
        for band in rows:
            cells = {col: [] for col in COLUMN_ORDER}
            for w in band:
                cells[w["col"]].append((w["x0"], w["text"]))
            cells = {c: " ".join(t for _, t in sorted(parts)) for c, parts in cells.items()}
            cells["_y"] = sum(d["y"] for d in band) / len(band)
            banded.append(cells)
        return banded

    def parse_money(s):
        return float(s.replace("$", "").replace(",", ""))

    def assemble_transactions(rows):
        """Anchor on rows whose Date cell matches a date pattern. Each
        description-only row is assigned to its NEAREST date row by y-distance,
        so it never bleeds into the next transaction."""
        date_rows = [(i, r) for i, r in enumerate(rows) if DATE_RE.match(r["Date"].strip())]
        if not date_rows:
            return []

        # For each description-only row, find the nearest date row's index
        desc_assignments = {di: [] for di, _ in date_rows}  # date_row_idx -> [(y, text)]
        date_ys = [(di, r["_y"]) for di, r in date_rows]

        for i, row in enumerate(rows):
            if DATE_RE.match(row["Date"].strip()):
                continue
            text = row["Description"].strip()
            if not text or JUNK_RE.match(text):
                continue
            # Skip rows that have an Amount (they're probably the page header row, etc.)
            if row["Amount"].strip() or row["Current Balance"].strip():
                continue
            nearest_di, _ = min(date_ys, key=lambda dy: abs(row["_y"] - dy[1]))
            desc_assignments[nearest_di].append((row["_y"], text))

        txns = []
        for di, row in date_rows:
            amount_str = row["Amount"].strip()
            balance_str = row["Current Balance"].strip()
            if not amount_str or not balance_str:
                continue

            # Build description: rows above the date come first, then the date
            # row's own description cell, then rows below.
            date_y = row["_y"]
            above = [t for y, t in desc_assignments[di] if y < date_y]
            below = [t for y, t in desc_assignments[di] if y > date_y]
            own = [row["Description"].strip()] if row["Description"].strip() else []
            desc_parts = above + own + below

            txns.append({
                "date": row["Date"].strip(),
                "description": " | ".join(p for p in desc_parts if p),
                "category": row["Category"].strip(),
                "amount": parse_money(amount_str),
                "balance": parse_money(balance_str),
            })
        return txns

    transactions = []
    with pdfplumber.open(pdfpath) as pdf:
        bounds = find_column_bounds(pdf.pages[0])
        for page in pdf.pages:
            rows = cluster_rows(page.extract_words(), bounds)
            transactions.extend(assemble_transactions(rows))


    # repeats = []
    # for r in transactions:
    #     if r["description"] not in repeats:
    #         repeats.append(r["description"])

    # Aggregate by description: count, total spent, and individual amounts
    aggregated = defaultdict(lambda: {"count": 0, "total": 0.0, "amounts": []})
    for t in transactions:
        desc = t["description"]
        aggregated[desc]["count"] += 1
        aggregated[desc]["total"] += t["amount"]
        aggregated[desc]["amounts"].append(t["amount"])

    # Keep only duplicates (appearing more than once)
    duplicates = {desc: info for desc, info in aggregated.items() if info["count"] > 1}

    # Sort by count, highest first (optional but useful)
    duplicates = dict(sorted(duplicates.items(), key=lambda kv: kv[1]["count"], reverse=True))

    return duplicates

def save_File(duplicates, outputpath):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Duplicate Transactions"

    headers = ["Description", "Times Repeated", "Total Amount", "Average Amount", "All Amounts"]

    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 40

    row_num = 2
    for desc, info in duplicates.items():
        ws.cell(row=row_num, column=1, value=desc).alignment = Alignment(wrap_text=True)
        ws.cell(row=row_num, column=2, value=info["count"])
        ws.cell(row=row_num, column=3, value=round(info["total"], 2))
        ws.cell(row=row_num, column=4, value=round(info["total"] / info["count"], 2))
        ws.cell(row=row_num, column=5, value=", ".join(f"{a:.2f}" for a in info["amounts"]))
        row_num += 1

 
    wb.save(outputpath)