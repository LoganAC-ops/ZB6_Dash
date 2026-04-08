#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP ECC -> S4 Output Accrual Comparison Tool
Mondelez International | Accenture
─────────────────────────────────────────────────────────────────
Usage:
  Double-click to open with file picker GUI, or:
  python compare_sap.py <ecc_file.xlsx> <s4_file.xlsx>
"""

import sys
import os
import re
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ─── Brand Colors ─────────────────────────────────────────────────────────────
BANNER_BG  = "1E0A4C"   # Deep Mondelez dark purple
HDR_BG     = "5A1F8C"   # Mondelez purple — column headers
SEC_BG     = "2D0E6A"   # Section dividers

MATCH_BG   = "E8F5E9"   # Soft green   — present in both
MISSING_BG = "FFF3CD"   # Soft amber   — missing in S4
EXTRA_BG   = "E3F2FD"   # Soft blue    — extra/new in S4
ROW_ALT    = "F9F7FE"   # Very light purple — alternating row
WHITE      = "FFFFFF"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def _font(size=10, bold=False, color="212121", italic=False):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)

def _border():
    s = Side(style="thin", color="D0C8E8")
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _set(cell, value, bg=None, size=10, bold=False, color="212121",
         h="left", wrap=False, italic=False):
    cell.value = value
    if bg:
        cell.fill = _fill(bg)
    cell.font = _font(size=size, bold=bold, color=color, italic=italic)
    cell.alignment = _align(h=h, wrap=wrap)
    cell.border = _border()


# ─── Parsing ──────────────────────────────────────────────────────────────────

def parse_file(filepath):
    """
    Parse SAP XML/HTML output file into two flat lists:
      header_rows: one entry per HeaderFreeText (TextTypeCode + Text)
      line_rows:   one entry per ItemAmountsCharges within each LineItemInformation
    Namespace prefixes are stripped before parsing so tag names work regardless
    of which SAP system produced the file.
    """
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    # Strip XML namespace declarations and prefixes from tags
    content = re.sub(r'\s+xmlns(?::\w+)?="[^"]*"', '', content)
    content = re.sub(r'<([\w.-]+):([\w.-]+)', r'<\2', content)
    content = re.sub(r'</([\w.-]+):([\w.-]+)', r'</\2', content)

    root = ET.fromstring(content)

    def _text(node, tag):
        el = node.find(f".//{tag}")
        return el.text.strip() if el is not None and el.text else None

    doc_number  = _text(root, "DocumentNumber")
    header_rows = []
    line_rows   = []

    for hft in root.findall(".//HeaderFreeText"):
        text_type = _text(hft, "TextTypeCode")
        if text_type:
            header_rows.append({
                "field":   text_type,
                "value":   _text(hft, "Text") or "",
                "row_num": None,
            })

    for li in root.findall(".//LineItemInformation"):
        line_num     = _text(li, "LineItemNumber")
        material_num = _text(li, "MaterialNumber")
        material_desc= _text(li, "MaterialDescription")
        product_desc = _text(li, "ProductDescription")
        discount_amt = _text(li, "LineItemDiscountAmount")
        net_weight   = _text(li, "NetWeight")

        for iac in li.findall("ItemAmountsCharges"):
            if line_num:
                line_rows.append({
                    "line_num":      line_num,
                    "material_num":  material_num,
                    "material_desc": material_desc,
                    "product_desc":  product_desc,
                    "discount_amt":  discount_amt,
                    "charge_type":   _text(iac, "ChargeTypeCode"),
                    "description":   _text(iac, "Description"),
                    "amount":        _text(iac, "Amount"),
                    "net_weight":    net_weight,
                    "row_num":       None,
                })

    return header_rows, line_rows, doc_number


# ─── Date Validation ──────────────────────────────────────────────────────────

DATE_FIELDS = {"SODate", "DateAsoc", "From", "to"}

def check_dates(header_rows):
    """
    For header fields that should be YYYYMMDD, return a list of
    {field, value, source} dicts for any that fail validation.
    'source' is passed in by the caller so the UI can label ECC vs S4.
    """
    issues = []
    for r in header_rows:
        if r["field"] in DATE_FIELDS:
            val = r["value"]
            valid = bool(val and re.match(r"^\d{8}$", val))
            if valid:
                try:
                    datetime.strptime(val, "%Y%m%d")
                except ValueError:
                    valid = False
            if not valid:
                issues.append({"field": r["field"], "value": val or "(empty)"})
    return issues


# ─── Comparison Logic ─────────────────────────────────────────────────────────

def compare_headers(ecc_hdr, s4_hdr):
    """
    Compare all TextTypeCode fields across both files (flat, no grouping).
    ECC order first, then any S4-only fields appended at the end.
    Returns list of {field, ecc_value, s4_value, status}.
    """
    ecc_map = {r["field"]: r["value"] for r in ecc_hdr}
    s4_map  = {r["field"]: r["value"] for r in s4_hdr}

    ordered, seen = [], set()
    for r in ecc_hdr:
        if r["field"] not in seen:
            ordered.append(r["field"])
            seen.add(r["field"])
    for r in s4_hdr:
        if r["field"] not in seen:
            ordered.append(r["field"])
            seen.add(r["field"])

    results = []
    for field in ordered:
        in_ecc = field in ecc_map
        in_s4  = field in s4_map
        if in_ecc and in_s4:
            status = "match"
        else:
            status = "missing_in_s4" if in_ecc else "extra_in_s4"
        results.append({
            "field":     field,
            "ecc_value": ecc_map.get(field),
            "s4_value":  s4_map.get(field),
            "status":    status,
        })
    return results


def compare_line_items(ecc_lines, s4_lines):
    """
    Compare line item rows across both files (flat, no grouping by Serie).
    Match key: (line_num, charge_type).
    Returns list of {key, ecc, s4, status}.
    """
    def _key(r):
        return (str(r.get("line_num") or ""), str(r.get("charge_type") or ""))

    ecc_map, s4_map = {}, {}
    for r in ecc_lines:
        ecc_map.setdefault(_key(r), []).append(r)
    for r in s4_lines:
        s4_map.setdefault(_key(r), []).append(r)

    all_keys, seen = [], set()
    for r in ecc_lines:
        k = _key(r)
        if k not in seen:
            all_keys.append(k)
            seen.add(k)
    for r in s4_lines:
        k = _key(r)
        if k not in seen:
            all_keys.append(k)
            seen.add(k)

    results = []
    for k in all_keys:
        ecc_list = ecc_map.get(k, [])
        s4_list  = s4_map.get(k, [])
        for i in range(max(len(ecc_list), len(s4_list))):
            ecc_r = ecc_list[i] if i < len(ecc_list) else None
            s4_r  = s4_list[i]  if i < len(s4_list)  else None
            if ecc_r and s4_r:
                status = "match"
            else:
                status = "missing_in_s4" if ecc_r else "extra_in_s4"
            results.append({"key": k, "ecc": ecc_r, "s4": s4_r, "status": status})
    return results


# ─── Excel Report Helpers ─────────────────────────────────────────────────────

def _banner(ws, row, ecc_name, s4_name):
    """Write the top branded banner. Returns next row."""
    n = 9

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    _set(ws.cell(row, 1), "MONDELEZ INTERNATIONAL  |  ACCENTURE",
         bg=BANNER_BG, size=13, bold=True, color="FFFFFF", h="center")
    ws.row_dimensions[row].height = 28
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    _set(ws.cell(row, 1), "ZB6 Comparison Report",
         bg=SEC_BG, size=11, bold=True, color="D8C8FF", h="center")
    ws.row_dimensions[row].height = 20
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    label = (f"ECC Baseline: {ecc_name}   <->   S4 New System: {s4_name}"
             f"   |   Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    _set(ws.cell(row, 1), label, bg="F3EEF9", size=9, color="555555", h="center", italic=True)
    ws.row_dimensions[row].height = 16
    row += 1

    return row


def _legend(ws, row):
    ws.row_dimensions[row].height = 18
    ws.cell(row, 1).value = "KEY:"
    ws.cell(row, 1).font = _font(9, bold=True, color="444444")
    items = [
        ("  MATCH — values identical  ",           MATCH_BG,   "1B5E20"),
        ("  DIFFERENT — values changed  ",         "FFEBEE",   "B71C1C"),
        ("  MISSING IN S4 — in ECC only  ",        MISSING_BG, "7B4600"),
        ("  EXTRA IN S4 — not in ECC  ",           EXTRA_BG,   "0D47A1"),
    ]
    for i, (label, bg, fg) in enumerate(items, 2):
        c = ws.cell(row, i)
        _set(c, label, bg=bg, size=9, bold=True, color=fg, h="center")
    return row + 2


def _section_hdr(ws, row, title, n=9):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    _set(ws.cell(row, 1), f"  {title}", bg=SEC_BG, size=10, bold=True, color="FFFFFF", h="left")
    for col in range(2, n + 1):
        ws.cell(row, col).fill = _fill(SEC_BG)
    ws.row_dimensions[row].height = 20
    return row + 1


def _col_hdrs(ws, row, headers, height=28):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h)
        c.fill = _fill(HDR_BG)
        c.font = _font(9, bold=True, color="FFFFFF")
        c.alignment = _align(h="center", wrap=True)
        c.border = _border()
    ws.row_dimensions[row].height = height
    return row + 1


def _status_style(status):
    if status == "match":
        return MATCH_BG, "MATCH", "1B5E20"
    if status == "missing_in_s4":
        return MISSING_BG, "MISSING IN S4", "7B4600"
    if status == "extra_in_s4":
        return EXTRA_BG, "EXTRA IN S4", "0D47A1"
    return "FFEBEE", "DIFFERENT", "B71C1C"


# ─── Report Builder ───────────────────────────────────────────────────────────

def build_report(ecc_path, s4_path, output_path=None):
    ecc_name = os.path.basename(ecc_path)
    s4_name  = os.path.basename(s4_path)

    print(f"  Reading ECC : {ecc_name}")
    ecc_hdr, ecc_lines, ecc_docnum = parse_file(ecc_path)

    print(f"  Reading S4  : {s4_name}")
    s4_hdr,  s4_lines,  s4_docnum  = parse_file(s4_path)

    ecc_label = ecc_docnum or ecc_name
    s4_label  = s4_docnum  or s4_name

    print(f"  ECC header rows : {len(ecc_hdr)}   line rows : {len(ecc_lines)}")
    print(f"  S4  header rows : {len(s4_hdr)}   line rows : {len(s4_lines)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"

    # Column widths
    for i, w in enumerate([22, 30, 30, 18, 18, 14, 14, 14, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1
    row = _banner(ws, row, ecc_name, s4_name)
    row = _legend(ws, row)

    # ── SECTION 1: HEADER DETAILS ─────────────────────────────────────────────
    hdr_results  = compare_headers(ecc_hdr, s4_hdr)
    n_match      = sum(1 for r in hdr_results if r["status"] == "match")
    n_missing    = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
    n_extra      = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")

    row = _section_hdr(ws, row,
        f"SECTION 1 — HEADER DETAILS  (Column J)   "
        f"| ECC: {len(ecc_hdr)} fields   S4: {len(s4_hdr)} fields   "
        f"Match: {n_match}   Missing in S4: {n_missing}   Extra in S4: {n_extra}",
        n=4)

    row = _col_hdrs(ws, row, [
        "FIELD  (TextTypeCode)",
        f"ECC Value\n({ecc_label})",
        f"S4 Value\n({s4_label})",
        "STATUS",
    ], height=24)

    if not hdr_results:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws.cell(row, 1, "No header data found in either file.").font = _font(9, italic=True, color="888888")
        row += 1
    else:
        for item in hdr_results:
            bg, status_text, fg = _status_style(item["status"])
            vals   = [item["field"],
                      item["ecc_value"] if item["ecc_value"] is not None else "—",
                      item["s4_value"]  if item["s4_value"]  is not None else "—",
                      status_text]
            aligns = ["left", "left", "left", "center"]
            bolds  = [False, False, False, True]
            for i, (v, ha, b) in enumerate(zip(vals, aligns, bolds), 1):
                c = ws.cell(row, i, v)
                c.fill      = _fill(bg)
                c.font      = _font(10, bold=b, color=(fg if i == 4 else "212121"))
                c.alignment = _align(h=ha)
                c.border    = _border()
            ws.row_dimensions[row].height = 17
            row += 1

    row += 1

    # ── SECTION 2: LINE ITEM / PRICING DETAILS ────────────────────────────────
    line_results = compare_line_items(ecc_lines, s4_lines)
    n_match   = sum(1 for r in line_results if r["status"] == "match")
    n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
    n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")

    row = _section_hdr(ws, row,
        f"SECTION 2 — LINE ITEM / PRICING DETAILS  (Cols Z, AA, AB, AC, AK, AM, AN)   "
        f"| ECC: {len(ecc_lines)} rows   S4: {len(s4_lines)} rows   "
        f"Match: {n_match}   Missing in S4: {n_missing}   Extra in S4: {n_extra}",
        n=9)

    row = _col_hdrs(ws, row, [
        "Line #",
        "Charge Type",
        f"ECC: MaterialNumber\n({ecc_label})",
        f"S4: MaterialNumber\n({s4_label})",
        f"ECC: Amount\n({ecc_label})",
        f"S4: Amount\n({s4_label})",
        f"ECC: Description\n({ecc_label})",
        f"S4: Description\n({s4_label})",
        "STATUS",
    ], height=30)

    if not line_results:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        ws.cell(row, 1, "No line item data found in either file.").font = _font(9, italic=True, color="888888")
        row += 1
    else:
        for item in line_results:
            bg, status_text, fg = _status_style(item["status"])
            e = item["ecc"] or {}
            s = item["s4"]  or {}
            vals = [
                str(e.get("line_num")      or s.get("line_num")      or "—"),
                str(e.get("charge_type")   or s.get("charge_type")   or "—"),
                e.get("material_num")  or "—",
                s.get("material_num")  or "—",
                e.get("amount")        or "—",
                s.get("amount")        or "—",
                e.get("description")   or "—",
                s.get("description")   or "—",
                status_text,
            ]
            aligns = ["center","center","left","left","center","center","left","left","center"]
            bolds  = [False, True, False, False, False, False, False, False, True]
            for i, (v, ha, b) in enumerate(zip(vals, aligns, bolds), 1):
                c = ws.cell(row, i, v)
                c.fill      = _fill(bg)
                c.font      = _font(10, bold=b, color=(fg if i == 9 else "212121"))
                c.alignment = _align(h=ha)
                c.border    = _border()
            ws.row_dimensions[row].height = 17
            row += 1

    # ── Output ────────────────────────────────────────────────────────────────
    if output_path is None:
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.dirname(os.path.abspath(ecc_path))
        output_path = os.path.join(out_dir, f"SAP_Comparison_{ts}.xlsx")

    wb.save(output_path)
    print(f"\n  Report saved : {output_path}")
    return output_path


# ─── Raw Data Export ─────────────────────────────────────────────────────────

def build_raw_export(ecc_path, s4_path, output_path=None):
    """
    Build a two-sheet Excel with the raw parsed data from each file.
    Sheet names are the document numbers (or filenames as fallback).
    """
    ecc_hdr, ecc_lines, ecc_docnum = parse_file(ecc_path)
    s4_hdr,  s4_lines,  s4_docnum  = parse_file(s4_path)

    ecc_label = ecc_docnum or os.path.basename(ecc_path)
    s4_label  = s4_docnum  or os.path.basename(s4_path)

    wb = Workbook()

    for label, hdr_rows, line_rows in [
        (ecc_label, ecc_hdr, ecc_lines),
        (s4_label,  s4_hdr,  s4_lines),
    ]:
        ws = wb.create_sheet(title=label[:31])  # Excel sheet name max 31 chars

        # ── Header fields section ─────────────────────────────────────────────
        row = 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        _set(ws.cell(row, 1), "HEADER FIELDS", bg=SEC_BG, size=10, bold=True, color="FFFFFF", h="left")
        ws.cell(row, 2).fill = _fill(SEC_BG)
        ws.row_dimensions[row].height = 20
        row += 1

        for col, title in [(1, "TextTypeCode"), (2, "Value")]:
            c = ws.cell(row, col, title)
            c.fill = _fill(HDR_BG)
            c.font = _font(9, bold=True, color="FFFFFF")
            c.alignment = _align(h="center")
            c.border = _border()
        ws.row_dimensions[row].height = 18
        row += 1

        for r in hdr_rows:
            _set(ws.cell(row, 1), r["field"],  h="left")
            _set(ws.cell(row, 2), r["value"],  h="left")
            ws.row_dimensions[row].height = 16
            row += 1

        row += 1

        # ── Line items section ────────────────────────────────────────────────
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _set(ws.cell(row, 1), "LINE ITEMS", bg=SEC_BG, size=10, bold=True, color="FFFFFF", h="left")
        for col in range(2, 9):
            ws.cell(row, col).fill = _fill(SEC_BG)
        ws.row_dimensions[row].height = 20
        row += 1

        line_cols = ["Line #", "Charge Type", "Material Number", "Description",
                     "Amount", "Material Desc", "Product Desc", "Net Weight"]
        for i, title in enumerate(line_cols, 1):
            c = ws.cell(row, i, title)
            c.fill = _fill(HDR_BG)
            c.font = _font(9, bold=True, color="FFFFFF")
            c.alignment = _align(h="center", wrap=True)
            c.border = _border()
        ws.row_dimensions[row].height = 18
        row += 1

        for r in line_rows:
            vals = [
                r.get("line_num")      or "",
                r.get("charge_type")   or "",
                r.get("material_num")  or "",
                r.get("description")   or "",
                r.get("amount")        or "",
                r.get("material_desc") or "",
                r.get("product_desc")  or "",
                r.get("net_weight")    or "",
            ]
            for i, v in enumerate(vals, 1):
                _set(ws.cell(row, i), v, h="left" if i not in (1, 5) else "center")
            ws.row_dimensions[row].height = 16
            row += 1

        # Column widths
        for i, w in enumerate([10, 14, 18, 30, 14, 24, 24, 14], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # Remove default empty sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    if output_path is None:
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.dirname(os.path.abspath(ecc_path))
        output_path = os.path.join(out_dir, f"SAP_RawData_{ts}.xlsx")

    wb.save(output_path)
    return output_path


# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    ecc_path = s4_path = None

    if len(sys.argv) == 3:
        ecc_path = sys.argv[1]
        s4_path  = sys.argv[2]
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)

            messagebox.showinfo(
                "SAP Comparison Tool — Mondelez | Accenture",
                "Step 1 of 2: Select the ECC (old system) Excel file.",
                parent=root,
            )
            ecc_path = filedialog.askopenfilename(
                title="Select ECC (Old System) XML/HTML File",
                filetypes=[("XML/HTML Files", "*.xml *.html"), ("All Files", "*.*")],
                parent=root,
            )
            if not ecc_path:
                messagebox.showwarning("Cancelled", "No ECC file selected.", parent=root)
                root.destroy()
                return

            messagebox.showinfo(
                "SAP Comparison Tool — Mondelez | Accenture",
                "Step 2 of 2: Select the S4 (new system) Excel file.",
                parent=root,
            )
            s4_path = filedialog.askopenfilename(
                title="Select S4 (New System) XML/HTML File",
                filetypes=[("XML/HTML Files", "*.xml *.html"), ("All Files", "*.*")],
                parent=root,
            )
            if not s4_path:
                messagebox.showwarning("Cancelled", "No S4 file selected.", parent=root)
                root.destroy()
                return

            root.destroy()
        except Exception as e:
            print(f"GUI unavailable: {e}")
            print("Usage: python compare_sap.py <ecc_file.xlsx> <s4_file.xlsx>")
            return

    print("\nSAP ECC -> S4 Comparison Tool  |  Mondelez | Accenture")
    print("=" * 55)
    try:
        output = build_report(ecc_path, s4_path)
        try:
            import subprocess
            subprocess.Popen(["start", "", output], shell=True)
        except Exception:
            pass
        print("\nDone.")
    except Exception as err:
        import traceback
        print(f"\nERROR: {err}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
