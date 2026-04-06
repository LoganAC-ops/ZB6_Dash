#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP ECC -> S4 Comparison Tool  |  Streamlit UI
Mondelez International | Accenture
"""

import streamlit as st
import pandas as pd
import tempfile
import os

from compare_sap import parse_file, compare_headers, compare_line_items, build_report, check_dates

# ─── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SAP Comparison Tool",
    page_icon="📊",
    layout="wide",
)

# ─── Brand colors ─────────────────────────────────────────────────────────────

MATCH_BG   = "#E8F5E9"
MATCH_FG   = "#1B5E20"
DIFF_BG    = "#FFEBEE"
DIFF_FG    = "#B71C1C"
MISSING_BG = "#FFF3CD"
MISSING_FG = "#7B4600"
EXTRA_BG   = "#E3F2FD"
EXTRA_FG   = "#0D47A1"

STATUS_COLORS = {
    "match":         (MATCH_BG,   MATCH_FG),
    "missing_in_s4": (MISSING_BG, MISSING_FG),
    "extra_in_s4":   (EXTRA_BG,   EXTRA_FG),
}

STATUS_LABELS = {
    "match":         "MATCH",
    "missing_in_s4": "MISSING IN S4",
    "extra_in_s4":   "EXTRA IN S4",
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def style_row(row):
    status = row.get("Status", "")
    reverse = {v: k for k, v in STATUS_LABELS.items()}
    key = reverse.get(status, "")
    bg, fg = STATUS_COLORS.get(key, ("#FFFFFF", "#000000"))
    return [f"background-color: {bg}; color: {fg}"] * len(row)


def render_legend():
    cols = st.columns(3)
    items = [
        ("MATCH",          MATCH_BG,   MATCH_FG),
        ("MISSING IN S4",  MISSING_BG, MISSING_FG),
        ("EXTRA IN S4",    EXTRA_BG,   EXTRA_FG),
    ]
    for col, (label, bg, fg) in zip(cols, items):
        col.markdown(
            f'<div style="background:{bg};color:{fg};padding:6px 10px;'
            f'border-radius:4px;font-weight:bold;font-size:13px;text-align:center">'
            f'{label}</div>',
            unsafe_allow_html=True,
        )


def save_upload(uploaded_file):
    """Save an UploadedFile to a temp file and return the path."""
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()
    return tmp.name


# ─── Header ───────────────────────────────────────────────────────────────────

st.markdown(
    """
    <div style="background:#1E0A4C;padding:18px 24px;border-radius:6px;margin-bottom:4px">
        <h2 style="color:white;margin:0;font-family:Calibri,sans-serif">
            MONDELEZ INTERNATIONAL &nbsp;|&nbsp; ACCENTURE
        </h2>
        <p style="color:#D8C8FF;margin:4px 0 0;font-size:15px">
            ZB6 Comparison Report
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("")

# ─── File upload ──────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)
with col1:
    ecc_file = st.file_uploader(
        "ECC File (old system)",
        type=["xml", "html"],
        help="XML/HTML export from SAP ECC",
    )
with col2:
    s4_file = st.file_uploader(
        "S4 File (new system)",
        type=["xml", "html"],
        help="XML/HTML export from SAP S4",
    )

st.markdown("")
run = st.button("Run Comparison", type="primary", disabled=not (ecc_file and s4_file))

# ─── Run ──────────────────────────────────────────────────────────────────────

if run and ecc_file and s4_file:

    with st.spinner("Reading and comparing files..."):

        ecc_path = save_upload(ecc_file)
        s4_path  = save_upload(s4_file)

        try:
            ecc_hdr,  ecc_lines = parse_file(ecc_path)
            s4_hdr,   s4_lines  = parse_file(s4_path)

            hdr_results  = compare_headers(ecc_hdr, s4_hdr)
            line_results = compare_line_items(ecc_lines, s4_lines)

            # Build downloadable Excel report
            from datetime import datetime
            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = tempfile.gettempdir()
            out_path = os.path.join(out_dir, f"SAP_Comparison_{ts}.xlsx")
            build_report(ecc_path, s4_path, output_path=out_path)

        finally:
            os.unlink(ecc_path)
            os.unlink(s4_path)

    st.success("Comparison complete.")

    # ── Date format warnings ──────────────────────────────────────────────────
    ecc_date_issues = check_dates(ecc_hdr)
    s4_date_issues  = check_dates(s4_hdr)
    if ecc_date_issues or s4_date_issues:
        with st.expander("⚠️ Date Format Warnings — expected YYYYMMDD (e.g. 20260206)", expanded=True):
            if ecc_date_issues:
                st.markdown(f"**ECC ({ecc_file.name})**")
                for issue in ecc_date_issues:
                    st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")
            if s4_date_issues:
                st.markdown(f"**S4 ({s4_file.name})**")
                for issue in s4_date_issues:
                    st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")

    # ── Legend ────────────────────────────────────────────────────────────────
    render_legend()
    st.markdown("")

    # ── Section 1: Headers ────────────────────────────────────────────────────
    n_match   = sum(1 for r in hdr_results if r["status"] == "match")
    n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
    n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")

    st.markdown(
        f"<div style='background:#2D0E6A;color:white;padding:8px 14px;"
        f"border-radius:4px;font-weight:bold'>"
        f"SECTION 1 — HEADER DETAILS &nbsp;|&nbsp; "
        f"ECC: {len(ecc_hdr)} &nbsp; S4: {len(s4_hdr)} &nbsp; "
        f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        f"</div>",
        unsafe_allow_html=True,
    )

    def make_hdr_row(r):
        return {
            "Field (TextTypeCode)": r["field"],
            f"ECC Value ({ecc_file.name})": r["ecc_value"] if r["ecc_value"] is not None else "—",
            f"S4 Value ({s4_file.name})":   r["s4_value"]  if r["s4_value"]  is not None else "—",
            "Status": STATUS_LABELS[r["status"]],
        }

    def make_line_row(r):
        e, s = r["ecc"] or {}, r["s4"] or {}
        return {
            "Line #":                              str(e.get("line_num") or s.get("line_num") or "—"),
            "Charge Type":                         str(e.get("charge_type") or s.get("charge_type") or "—"),
            f"ECC: Material ({ecc_file.name})":    e.get("material_num")  or "—",
            f"S4: Material ({s4_file.name})":      s.get("material_num")  or "—",
            f"ECC: Amount ({ecc_file.name})":      e.get("amount")        or "—",
            f"S4: Amount ({s4_file.name})":        s.get("amount")        or "—",
            f"ECC: Description ({ecc_file.name})": e.get("description")   or "—",
            f"S4: Description ({s4_file.name})":   s.get("description")   or "—",
            "Status": STATUS_LABELS[r["status"]],
        }

    # ── Section 1: Headers — mismatches only ─────────────────────────────────
    hdr_gaps = [r for r in hdr_results if r["status"] != "match"]

    if hdr_gaps:
        hdr_gap_df = pd.DataFrame([make_hdr_row(r) for r in hdr_gaps])
        st.dataframe(
            hdr_gap_df.style.apply(style_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("All header fields match.")

    st.markdown("")

    # ── Section 2: Line Items — mismatches only ───────────────────────────────
    n_match   = sum(1 for r in line_results if r["status"] == "match")
    n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
    n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")

    st.markdown(
        f"<div style='background:#2D0E6A;color:white;padding:8px 14px;"
        f"border-radius:4px;font-weight:bold'>"
        f"SECTION 2 — LINE ITEM / PRICING DETAILS &nbsp;|&nbsp; "
        f"ECC: {len(ecc_lines)} rows &nbsp; S4: {len(s4_lines)} rows &nbsp; "
        f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        f"</div>",
        unsafe_allow_html=True,
    )

    line_gaps = [r for r in line_results if r["status"] != "match"]

    if line_gaps:
        line_gap_df = pd.DataFrame([make_line_row(r) for r in line_gaps])
        st.dataframe(
            line_gap_df.style.apply(style_row, axis=1),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("All line items match.")

    st.markdown("")

    # ── Combined view (expandable) ────────────────────────────────────────────
    with st.expander("Full Combined View — all rows (matched, missing, extra)"):
        st.markdown("**Header Details**")
        if hdr_results:
            st.dataframe(
                pd.DataFrame([make_hdr_row(r) for r in hdr_results])
                  .style.apply(style_row, axis=1),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No header data found.")

        st.markdown("**Line Item / Pricing Details**")
        if line_results:
            st.dataframe(
                pd.DataFrame([make_line_row(r) for r in line_results])
                  .style.apply(style_row, axis=1),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No line item data found.")

    st.markdown("")

    # ── Download ──────────────────────────────────────────────────────────────
    with open(out_path, "rb") as f:
        st.download_button(
            label="Download Excel Report",
            data=f,
            file_name=os.path.basename(out_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
