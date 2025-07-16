import os
import re
import pdfplumber
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
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

def normalize(text):
    return re.sub(r'[^a-z0-9]', '', text.lower())

def build_normalized_field_map(field_aliases):
    return {
        normalize(alias): canonical
        for canonical, aliases in field_aliases.items()
        for alias in aliases
    }

def match_field(label, normalized_field_map):
    norm_label = normalize(label)
    if norm_label in normalized_field_map:
        return normalized_field_map[norm_label]
    for norm_alias, canonical in normalized_field_map.items():
        if len(norm_alias) < 3:
            continue
        if norm_alias in norm_label:
            return canonical
    return None

# ---------- COLOR-AWARE FIELD EXTRACTION ----------

def extract_fields_from_page_by_color(page, match_field_func):
    words = page.extract_words(extra_attrs=["non_stroking_color"])
    lines_by_y = {}
    for word in words:
        y = round(word['top'])
        lines_by_y.setdefault(y, []).append(word)

    fields = {}

    def get_rgb_tuple(color):
        if isinstance(color, (tuple, list)):
            return tuple(color[:3]) + (0,) * (3 - len(color[:3]))
        elif isinstance(color, (int, float)):
            return (color, color, color)
        return (0, 0, 0)

    def is_black(color):
        rgb = get_rgb_tuple(color)
        return rgb[0] < 0.05 and rgb[1] < 0.05 and rgb[2] < 0.05

    def is_value_color(color):
        rgb = get_rgb_tuple(color)
        # Gold/yellow shades seen in sample PDFs, tolerance for others
        return (
            abs(rgb[0] - 0.86) < 0.13 and
            abs(rgb[1] - 0.65) < 0.13 and
            abs(rgb[2] - 0.0) < 0.13
        ) or (
            abs(rgb[0] - 0.94669) < 0.13 and
            abs(rgb[1] - 0.78061) < 0.13 and
            abs(rgb[2] - 0.0) < 0.13
        )

    for y in sorted(lines_by_y):
        line_words = sorted(lines_by_y[y], key=lambda w: w['x0'])
        color_chunks = []
        if not line_words:
            continue
        last_color = line_words[0].get("non_stroking_color")
        chunk = [line_words[0]]
        for w in line_words[1:]:
            color = w.get("non_stroking_color")
            if color == last_color:
                chunk.append(w)
            else:
                color_chunks.append((last_color, chunk))
                chunk = [w]
                last_color = color
        color_chunks.append((last_color, chunk))

        i = 0
        while i < len(color_chunks):
            label_color, label_chunk = color_chunks[i]
            if is_black(label_color):
                # Merge all consecutive value chunks as value (robust for multi-word)
                value_texts = []
                j = i + 1
                while j < len(color_chunks) and is_value_color(color_chunks[j][0]):
                    value_texts.extend(w['text'] for w in color_chunks[j][1])
                    j += 1
                if value_texts:
                    raw_label = ' '.join(w['text'] for w in label_chunk).replace(':', '').strip()
                    raw_value = ' '.join(value_texts).strip()
                    field = match_field_func(raw_label)
                    if field:
                        fields[field] = raw_value
                i = j
            else:
                i += 1
    return fields

# ---------- DATA EXTRACTION FROM PDF ----------

def extract_pdf_data(pdf_file, field_order, field_aliases):
    """Takes file-like or path, returns list-of-dicts with consistent field ordering."""
    normalized_field_map = build_normalized_field_map(field_aliases)
    match_func = lambda label: match_field(label, normalized_field_map)
    all_rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            fields = extract_fields_from_page_by_color(page, match_func)
            all_rows.append(fields)
    return [
        {field: row.get(field, "") for field in field_order}
        for row in all_rows
    ]

# ---------- IN-MEMORY EXCEL WRITER ----------

def make_excel_file_from_data(data_rows, field_order, file_name="output.xlsx"):
    """Returns BytesIO ready for download, not saved to disk."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(field_order)
    for row in data_rows:
        ws.append([row.get(field, "") for field in field_order])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- END FILE ---