import os
import re
import pdfplumber
import openpyxl
import platform
import subprocess
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
    if isinstance(color, (tuple, list)):
        return tuple(color[:3])
    if isinstance(color, (int, float)):
        return (color,)*3
    return (0,0,0)

def is_black(c): 
    r,g,b = rgb(c); return r<0.05 and g<0.05 and b<0.05

def is_value_color(c):
    r,g,b = rgb(c)
    # define your tolerances once
    gold1 = (0.86,0.65,0.0)
    gold2 = (0.94669,0.78061,0.0)
    return any(abs(r-x)<0.13 and abs(g-y)<0.13 and abs(b-z)<0.13
               for (x,y,z) in (gold1, gold2))

# ---------- EXTRACTION LOGIC ----------
def extract_fields(page, fmap):
    words = page.extract_words(extra_attrs=["non_stroking_color"])
    # group by approximate y
    lines = {}
    for w in words:
        y = round(w["top"])
        lines.setdefault(y, []).append(w)

    out = {}
    for y in sorted(lines):
        row = sorted(lines[y], key=lambda w: w["x0"])
        chunks = []
        last = row[0]["non_stroking_color"]
        buf = [row[0]]
        for w in row[1:]:
            if w["non_stroking_color"] == last:
                buf.append(w)
            else:
                chunks.append((last, buf))
                last, buf = w["non_stroking_color"], [w]
        chunks.append((last, buf))

        i = 0
        while i < len(chunks):
            col, chunk = chunks[i]
            if is_black(col):
                j = i+1
                vals = []
                while j < len(chunks) and is_value_color(chunks[j][0]):
                    vals += [w["text"] for w in chunks[j][1]]
                    j += 1
                if vals:
                    label = " ".join(w["text"] for w in chunk).rstrip(":")
                    key   = match_field(label, fmap)
                    if key:
                        out[key] = " ".join(vals)
                i = j
            else:
                i += 1
    return out

