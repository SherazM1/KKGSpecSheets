import os
import re
import pdfplumber
import openpyxl
from io import BytesIO

# ---------- FIELD ALIASES ----------
field_order = [
    "customer", "design", "rev.", "part", "oppty/proj. #", "pieces per set",
    "board", "corr direction", "view", "project mngr.", "designer", "id",
    "area", "blank width", "blank height", "inches of rule", "date"
]

field_aliases = {
    "customer": ["customer"],
    "design": ["design"],
    "rev.": ["rev.", "revision"],
    "part": ["part"],
    "oppty/proj. #": ["oppty/proj. #", "opportunity", "project number", "proj #", "project #"],
    "pieces per set": ["pieces per set"],
    "board": ["board"],
    "corr direction": ["corr direction", "grain/corr", "corr/grain", "grain direction", "corrugation", "corr"],
    "view": ["view", "side shown"],
    "project mngr.": ["project mngr.", "project manager", "proj mngr"],
    "designer": ["designer", "engineer"],
    "id": ["id"],
    "area": ["area"],
    "blank width": ["blank width"],
    "blank height": ["blank height"],
    "inches of rule": ["inches of rule", "len. cutting rule", "len. other rule", "length cutting rule", "length other rule"],
    "date": ["date"]
}

def normalize(text: str) -> str:
    return re.sub(r'[^a-z0-9]', "", text.lower())

def build_field_map(aliases):
    return {
        normalize(a): canon
        for canon, vals in aliases.items()
        for a in vals
    }

def match_field(label, fmap):
    nl = normalize(label)
    if nl in fmap:
        return fmap[nl]
    # fallback on substring match
    for norm_alias, canon in fmap.items():
        if len(norm_alias) >= 3 and norm_alias in nl:
            return canon
    return None

# ---------- COLOR HELPER ----------
def rgb(color):
    # Always return exactly three components, padding with 0 if needed
    if isinstance(color, (tuple, list)):
        vals = list(color[:3])
        vals += [0] * (3 - len(vals))
        return tuple(vals)
    if isinstance(color, (int, float)):
        return (color, color, color)
    return (0, 0, 0)

def is_black(c):
    r, g, b = rgb(c)
    return r < 0.05 and g < 0.05 and b < 0.05

# (You can remove is_value_color entirely if you’re treating all non-black as values,
#  or keep it for more nuance.)
def is_value_color(c):
    r, g, b = rgb(c)
    gold1 = (0.86, 0.65, 0.0)
    gold2 = (0.94669, 0.78061, 0.0)
    return any(abs(r - x) < 0.13 and abs(g - y) < 0.13 and abs(b - z) < 0.13
               for (x, y, z) in (gold1, gold2))

# ---------- EXTRACTION LOGIC ----------
def extract_fields(page, fmap):
    words = page.extract_words(extra_attrs=["non_stroking_color"])
    lines = {}
    for w in words:
        y = round(w["top"])
        lines.setdefault(y, []).append(w)

    out = {}
    for y in sorted(lines):
        row = sorted(lines[y], key=lambda w: w["x0"])
        # Build color-chunk list
        chunks = []
        last_color = row[0]["non_stroking_color"]
        buf = [row[0]]
        for w in row[1:]:
            if w["non_stroking_color"] == last_color:
                buf.append(w)
            else:
                chunks.append((last_color, buf))
                last_color, buf = w["non_stroking_color"], [w]
        chunks.append((last_color, buf))

        i = 0
        # Uncomment for debugging if needed
        # print([(i, "black" if is_black(c) else "not-black", " ".join(w["text"] for w in ch)) for i, (c, ch) in enumerate(chunks)])
        while i < len(chunks):
            label_color, label_chunk = chunks[i]
            if is_black(label_color):
                # Collect ALL consecutive non-black chunks after this label
                value_words = []
                j = i + 1
                while j < len(chunks) and not is_black(chunks[j][0]):
                    value_words += [w["text"] for w in chunks[j][1] if w["text"].strip()]
                    j += 1

                if value_words:
                    raw_label = " ".join(w["text"] for w in label_chunk).replace(":", "").strip().lower()
                    field = match_field(raw_label, fmap)
                    if field:
                        value = re.sub(r'\s+', ' ', " ".join(value_words)).strip()
                        prev_val = out.get(field, "")
                        combined_val = (prev_val + " " + value).strip() if prev_val else value
                        out[field] = combined_val

                i = j
            else:
                i += 1

    return out

def extract_pdf_data(pdf_file, field_order, field_aliases):
    fmap = build_field_map(field_aliases)
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            data = extract_fields(page, fmap)
            rows.append(data)
    # return a list of dicts, each dict ordered by field_order
    return [
        {f: row.get(f, "") for f in field_order}
        for row in rows
    ]

def make_excel_file_from_data(data_rows, field_order, file_name="output.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(field_order)
    for row in data_rows:
        ws.append([row.get(f, "") for f in field_order])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

