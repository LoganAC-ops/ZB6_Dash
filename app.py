#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
E-Invoicing Comparison Report  |  Streamlit UI
Mondelez International | Accenture
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime

from compare_sap import (
    parse_file, compare_headers, compare_line_items, build_report, build_raw_export, check_dates,
    parse_file_cr, compare_cr_lines, build_report_cr, build_raw_export_cr,
    parse_file_pa, compare_pa_lines, build_report_pa, build_raw_export_pa,
    parse_file_idoc, compare_idoc_lines, build_report_idoc, build_raw_export_idoc,
    parse_file_do, compare_do_lines, build_report_do, build_raw_export_do,
)

# ─── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="E-Invoicing Comparison Report",
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

ARG_STATUS_LABELS = {
    "match":         "MATCH",
    "missing_in_s4": "MISSING IN S4",
    "extra_in_s4":   "EXTRA IN S4",
}

CR_STATUS_LABELS = {
    "match":         "MATCH",
    "missing_in_s4": "MISSING IN TESTING",
    "extra_in_s4":   "EXTRA IN TESTING",
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def style_row(row, status_labels):
    status = row.get("Status", "")
    reverse = {v: k for k, v in status_labels.items()}
    key = reverse.get(status, "")
    bg, fg = STATUS_COLORS.get(key, ("#FFFFFF", "#000000"))
    return [f"background-color: {bg}; color: {fg}"] * len(row)


def render_legend(missing_label="MISSING IN S4", extra_label="EXTRA IN S4"):
    cols = st.columns(3)
    items = [
        ("MATCH",        MATCH_BG,   MATCH_FG),
        (missing_label,  MISSING_BG, MISSING_FG),
        (extra_label,    EXTRA_BG,   EXTRA_FG),
    ]
    for col, (label, bg, fg) in zip(cols, items):
        col.markdown(
            f'<div style="background:{bg};color:{fg};padding:6px 10px;'
            f'border-radius:4px;font-weight:bold;font-size:13px;text-align:center">'
            f'{label}</div>',
            unsafe_allow_html=True,
        )


def save_upload(uploaded_file):
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()
    return tmp.name


def section_banner(title):
    st.markdown(
        f"<div style='background:#2D0E6A;color:white;padding:8px 14px;"
        f"border-radius:4px;font-weight:bold'>{title}</div>",
        unsafe_allow_html=True,
    )


# ─── Page Header ──────────────────────────────────────────────────────────────

st.markdown(
    """
    <div style="background:#1E0A4C;padding:18px 24px;border-radius:6px;margin-bottom:4px">
        <h2 style="color:white;margin:0;font-family:Calibri,sans-serif">
            MONDELEZ INTERNATIONAL &nbsp;|&nbsp; ACCENTURE
        </h2>
        <p style="color:#D8C8FF;margin:4px 0 0;font-size:15px">
            E-Invoicing Comparison Report
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("")

# ─── Country Tabs ─────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Argentina (ZB6)",
    "Costa Rica",
    "Panama (ZB6)",
    "IDOC Countries (UY · HN · VE)",
    "Dominican Republic (ZB6)",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ARGENTINA
# ══════════════════════════════════════════════════════════════════════════════

with tab1:

    st.markdown("### Argentina (ZB6)")

    col1, col2 = st.columns(2)
    with col1:
        ecc_file = st.file_uploader(
            "EWP File (old system)",
            type=["xml", "html"],
            help="XML/HTML export from SAP ECC",
            key="ecc_arg",
        )
    with col2:
        s4_file = st.file_uploader(
            "S4 File (new system)",
            type=["xml", "html"],
            help="XML/HTML export from SAP S4",
            key="s4_arg",
        )

    st.markdown("")
    run_arg = st.button(
        "Run Comparison", type="primary",
        disabled=not (ecc_file and s4_file),
        key="run_arg",
    )

    if run_arg and ecc_file and s4_file:

        with st.spinner("Reading and comparing files..."):
            ecc_path = save_upload(ecc_file)
            s4_path  = save_upload(s4_file)

            try:
                ecc_hdr,  ecc_lines, ecc_docnum = parse_file(ecc_path)
                s4_hdr,   s4_lines,  s4_docnum  = parse_file(s4_path)

                ecc_label = ecc_docnum or ecc_file.name
                s4_label  = s4_docnum  or s4_file.name

                hdr_results  = compare_headers(ecc_hdr, s4_hdr)
                line_results = compare_line_items(ecc_lines, s4_lines)

                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"ARG_Comparison_{ts}.xlsx")
                build_report(ecc_path, s4_path, output_path=out_path)

                raw_path = os.path.join(out_dir, f"ARG_RawData_{ts}.xlsx")
                build_raw_export(ecc_path, s4_path, output_path=raw_path)

            finally:
                os.unlink(ecc_path)
                os.unlink(s4_path)

        st.success("Comparison complete — Argentina (ZB6)")

        # Date warnings
        ecc_date_issues = check_dates(ecc_hdr)
        s4_date_issues  = check_dates(s4_hdr)
        if ecc_date_issues or s4_date_issues:
            with st.expander("⚠️ Date Format Warnings — expected YYYYMMDD", expanded=True):
                if ecc_date_issues:
                    st.markdown(f"**EWP ({ecc_file.name})**")
                    for issue in ecc_date_issues:
                        st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")
                if s4_date_issues:
                    st.markdown(f"**S4 ({s4_file.name})**")
                    for issue in s4_date_issues:
                        st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")

        render_legend("MISSING IN S4", "EXTRA IN S4")
        st.markdown("")

        def style_arg(row):
            return style_row(row, ARG_STATUS_LABELS)

        def make_hdr_row(r):
            return {
                "Field (TextTypeCode)":      r["field"],
                f"EWP Value ({ecc_label})":  r["ecc_value"] if r["ecc_value"] is not None else "—",
                f"S4 Value ({s4_label})":    r["s4_value"]  if r["s4_value"]  is not None else "—",
                "Status": ARG_STATUS_LABELS[r["status"]],
            }

        def make_line_row(r):
            e, s = r["ecc"] or {}, r["s4"] or {}
            return {
                "Line #":                              str(e.get("line_num") or s.get("line_num") or "—"),
                "Charge Type":                         str(e.get("charge_type") or s.get("charge_type") or "—"),
                f"EWP: Material ({ecc_label})":        e.get("material_num")  or "—",
                f"S4: Material ({s4_label})":          s.get("material_num")  or "—",
                f"EWP: Amount ({ecc_label})":          e.get("amount")        or "—",
                f"S4: Amount ({s4_label})":            s.get("amount")        or "—",
                f"EWP: Description ({ecc_label})":     e.get("description")   or "—",
                f"S4: Description ({s4_label})":       s.get("description")   or "—",
                "Status": ARG_STATUS_LABELS[r["status"]],
            }

        # Section 1: Headers
        n_match   = sum(1 for r in hdr_results if r["status"] == "match")
        n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")
        section_banner(
            f"ARGENTINA — SECTION 1: HEADER DETAILS &nbsp;|&nbsp; "
            f"EWP: {len(ecc_hdr)} &nbsp; S4: {len(s4_hdr)} &nbsp; "
            f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        )

        hdr_gaps = [r for r in hdr_results if r["status"] != "match"]
        if hdr_gaps:
            st.dataframe(
                pd.DataFrame([make_hdr_row(r) for r in hdr_gaps]).style.apply(style_arg, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All header fields match.")

        st.markdown("")

        # Section 2: Line Items
        n_match   = sum(1 for r in line_results if r["status"] == "match")
        n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")
        section_banner(
            f"ARGENTINA — SECTION 2: LINE ITEM / PRICING DETAILS &nbsp;|&nbsp; "
            f"EWP: {len(ecc_lines)} rows &nbsp; S4: {len(s4_lines)} rows &nbsp; "
            f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        )

        line_gaps = [r for r in line_results if r["status"] != "match"]
        if line_gaps:
            st.dataframe(
                pd.DataFrame([make_line_row(r) for r in line_gaps]).style.apply(style_arg, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All line items match.")

        st.markdown("")

        with st.expander("Full Combined View — all rows (matched, missing, extra)"):
            st.markdown("**Header Details**")
            if hdr_results:
                st.dataframe(
                    pd.DataFrame([make_hdr_row(r) for r in hdr_results]).style.apply(style_arg, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No header data found.")

            st.markdown("**Line Item / Pricing Details**")
            if line_results:
                st.dataframe(
                    pd.DataFrame([make_line_row(r) for r in line_results]).style.apply(style_arg, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No line item data found.")

        st.markdown("")

        dl1, dl2 = st.columns(2)
        with dl1:
            with open(out_path, "rb") as f:
                st.download_button(
                    label="Download Comparison Report",
                    data=f,
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_arg_cmp",
                )
        with dl2:
            with open(raw_path, "rb") as f:
                st.download_button(
                    label="Download Raw Data",
                    data=f,
                    file_name=os.path.basename(raw_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_arg_raw",
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — COSTA RICA
# ══════════════════════════════════════════════════════════════════════════════

with tab2:

    st.markdown("### Costa Rica (?)")

    col1, col2 = st.columns(2)
    with col1:
        prod_file = st.file_uploader(
            "Production File (baseline)",
            type=["xml", "html"],
            help="Production XML export from Costa Rica",
            key="prod_cr",
        )
    with col2:
        test_file = st.file_uploader(
            "Testing File (new system)",
            type=["xml", "html"],
            help="Testing XML export from Costa Rica",
            key="test_cr",
        )

    st.markdown("")
    run_cr = st.button(
        "Run Comparison", type="primary",
        disabled=not (prod_file and test_file),
        key="run_cr",
    )

    if run_cr and prod_file and test_file:

        with st.spinner("Reading and comparing files..."):
            prod_path = save_upload(prod_file)
            test_path = save_upload(test_file)

            try:
                prod_hdr, prod_lines, prod_docnum = parse_file_cr(prod_path)
                test_hdr, test_lines, test_docnum = parse_file_cr(test_path)

                prod_label = prod_docnum or prod_file.name
                test_label = test_docnum or test_file.name

                hdr_results  = compare_headers(prod_hdr, test_hdr)
                line_results = compare_cr_lines(prod_lines, test_lines)

                ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir   = tempfile.gettempdir()
                out_path  = os.path.join(out_dir, f"CR_Comparison_{ts}.xlsx")
                build_report_cr(prod_path, test_path, output_path=out_path)

                raw_path = os.path.join(out_dir, f"CR_RawData_{ts}.xlsx")
                build_raw_export_cr(prod_path, test_path, output_path=raw_path)

            finally:
                os.unlink(prod_path)
                os.unlink(test_path)

        st.success("Comparison complete — Costa Rica")

        render_legend("MISSING IN TESTING", "EXTRA IN TESTING")
        st.markdown("")

        def style_cr(row):
            return style_row(row, CR_STATUS_LABELS)

        def make_hdr_row_cr(r):
            return {
                "Field":                              r["field"],
                f"Production Value ({prod_label})":   r["ecc_value"] if r["ecc_value"] is not None else "—",
                f"Testing Value ({test_label})":       r["s4_value"]  if r["s4_value"]  is not None else "—",
                "Status": CR_STATUS_LABELS[r["status"]],
            }

        def make_line_row_cr(r):
            p, t = r["prod"] or {}, r["test"] or {}
            return {
                "Line #":                                   str(p.get("line_num") or t.get("line_num") or "—"),
                f"Prod: Cod. Interno ({prod_label})":       p.get("codigo_interno")    or "—",
                f"Test: Cod. Interno ({test_label})":       t.get("codigo_interno")    or "—",
                "Detalle":                                  str(p.get("detalle") or t.get("detalle") or "—"),
                f"Prod: PrecioUnitario ({prod_label})":     p.get("precio_unitario")   or "—",
                f"Test: PrecioUnitario ({test_label})":     t.get("precio_unitario")   or "—",
                f"Prod: MontoTotalLinea ({prod_label})":    p.get("monto_total_linea") or "—",
                f"Test: MontoTotalLinea ({test_label})":    t.get("monto_total_linea") or "—",
                "Status": CR_STATUS_LABELS[r["status"]],
            }

        # Section 1: Headers
        n_match   = sum(1 for r in hdr_results if r["status"] == "match")
        n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")
        section_banner(
            f"COSTA RICA — SECTION 1: HEADER DETAILS &nbsp;|&nbsp; "
            f"Production: {len(prod_hdr)} &nbsp; Testing: {len(test_hdr)} &nbsp; "
            f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
        )

        hdr_gaps = [r for r in hdr_results if r["status"] != "match"]
        if hdr_gaps:
            st.dataframe(
                pd.DataFrame([make_hdr_row_cr(r) for r in hdr_gaps]).style.apply(style_cr, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All header fields match.")

        st.markdown("")

        # Section 2: Line Items
        n_match   = sum(1 for r in line_results if r["status"] == "match")
        n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")
        section_banner(
            f"COSTA RICA — SECTION 2: LINE ITEMS &nbsp;|&nbsp; "
            f"Production: {len(prod_lines)} rows &nbsp; Testing: {len(test_lines)} rows &nbsp; "
            f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
        )

        line_gaps = [r for r in line_results if r["status"] != "match"]
        if line_gaps:
            st.dataframe(
                pd.DataFrame([make_line_row_cr(r) for r in line_gaps]).style.apply(style_cr, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All line items match.")

        st.markdown("")

        with st.expander("Full Combined View — all rows (matched, missing, extra)"):
            st.markdown("**Header Details**")
            if hdr_results:
                st.dataframe(
                    pd.DataFrame([make_hdr_row_cr(r) for r in hdr_results]).style.apply(style_cr, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No header data found.")

            st.markdown("**Line Items**")
            if line_results:
                st.dataframe(
                    pd.DataFrame([make_line_row_cr(r) for r in line_results]).style.apply(style_cr, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No line item data found.")

        st.markdown("")

        dl1, dl2 = st.columns(2)
        with dl1:
            with open(out_path, "rb") as f:
                st.download_button(
                    label="Download Comparison Report",
                    data=f,
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cr_cmp",
                )
        with dl2:
            with open(raw_path, "rb") as f:
                st.download_button(
                    label="Download Raw Data",
                    data=f,
                    file_name=os.path.basename(raw_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cr_raw",
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — PANAMA
# ══════════════════════════════════════════════════════════════════════════════

with tab3:

    st.markdown("### Panama (ZB6)")

    col1, col2 = st.columns(2)
    with col1:
        prod_file_pa = st.file_uploader(
            "Production File (baseline)",
            type=["xml", "html"],
            help="Production XML export from Panama",
            key="prod_pa",
        )
    with col2:
        test_file_pa = st.file_uploader(
            "Testing File (new system)",
            type=["xml", "html"],
            help="Testing XML export from Panama",
            key="test_pa",
        )

    st.markdown("")
    run_pa = st.button(
        "Run Comparison", type="primary",
        disabled=not (prod_file_pa and test_file_pa),
        key="run_pa",
    )

    if run_pa and prod_file_pa and test_file_pa:

        with st.spinner("Reading and comparing files..."):
            prod_path = save_upload(prod_file_pa)
            test_path = save_upload(test_file_pa)

            try:
                prod_hdr, prod_lines, prod_docnum = parse_file_pa(prod_path)
                test_hdr, test_lines, test_docnum = parse_file_pa(test_path)

                prod_label = prod_docnum or prod_file_pa.name
                test_label = test_docnum or test_file_pa.name

                hdr_results  = compare_headers(prod_hdr, test_hdr)
                line_results = compare_pa_lines(prod_lines, test_lines)

                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"PA_Comparison_{ts}.xlsx")
                build_report_pa(prod_path, test_path, output_path=out_path)

                raw_path = os.path.join(out_dir, f"PA_RawData_{ts}.xlsx")
                build_raw_export_pa(prod_path, test_path, output_path=raw_path)

            finally:
                os.unlink(prod_path)
                os.unlink(test_path)

        st.success("Comparison complete — Panama (ZB6)")

        render_legend("MISSING IN TESTING", "EXTRA IN TESTING")
        st.markdown("")

        def style_pa(row):
            return style_row(row, CR_STATUS_LABELS)

        def make_hdr_row_pa(r):
            return {
                "Field":                              r["field"],
                f"Production Value ({prod_label})":   r["ecc_value"] if r["ecc_value"] is not None else "—",
                f"Testing Value ({test_label})":       r["s4_value"]  if r["s4_value"]  is not None else "—",
                "Status": CR_STATUS_LABELS[r["status"]],
            }

        def make_line_row_pa(r):
            p, t = r["prod"] or {}, r["test"] or {}
            return {
                "Line #":                                   str(p.get("line_num") or t.get("line_num") or "—"),
                f"Prod: MaterialNumber ({prod_label})":     p.get("material_num")  or "—",
                f"Test: MaterialNumber ({test_label})":     t.get("material_num")  or "—",
                "Material Desc":                            str(p.get("material_desc") or t.get("material_desc") or "—"),
                f"Prod: LineItemAmount ({prod_label})":     p.get("line_amount")   or "—",
                f"Test: LineItemAmount ({test_label})":     t.get("line_amount")   or "—",
                f"Prod: NetPrice ({prod_label})":           p.get("net_price")     or "—",
                f"Test: NetPrice ({test_label})":           t.get("net_price")     or "—",
                "Status": CR_STATUS_LABELS[r["status"]],
            }

        # Section 1: Headers
        n_match   = sum(1 for r in hdr_results if r["status"] == "match")
        n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")
        section_banner(
            f"PANAMA — SECTION 1: HEADER DETAILS &nbsp;|&nbsp; "
            f"Production: {len(prod_hdr)} &nbsp; Testing: {len(test_hdr)} &nbsp; "
            f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
        )

        hdr_gaps = [r for r in hdr_results if r["status"] != "match"]
        if hdr_gaps:
            st.dataframe(
                pd.DataFrame([make_hdr_row_pa(r) for r in hdr_gaps]).style.apply(style_pa, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All header fields match.")

        st.markdown("")

        # Section 2: Line Items
        n_match   = sum(1 for r in line_results if r["status"] == "match")
        n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")
        section_banner(
            f"PANAMA — SECTION 2: LINE ITEMS &nbsp;|&nbsp; "
            f"Production: {len(prod_lines)} rows &nbsp; Testing: {len(test_lines)} rows &nbsp; "
            f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
        )

        line_gaps = [r for r in line_results if r["status"] != "match"]
        if line_gaps:
            st.dataframe(
                pd.DataFrame([make_line_row_pa(r) for r in line_gaps]).style.apply(style_pa, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All line items match.")

        st.markdown("")

        with st.expander("Full Combined View — all rows (matched, missing, extra)"):
            st.markdown("**Header Details**")
            if hdr_results:
                st.dataframe(
                    pd.DataFrame([make_hdr_row_pa(r) for r in hdr_results]).style.apply(style_pa, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No header data found.")

            st.markdown("**Line Items**")
            if line_results:
                st.dataframe(
                    pd.DataFrame([make_line_row_pa(r) for r in line_results]).style.apply(style_pa, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No line item data found.")

        st.markdown("")

        dl1, dl2 = st.columns(2)
        with dl1:
            with open(out_path, "rb") as f:
                st.download_button(
                    label="Download Comparison Report",
                    data=f,
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_pa_cmp",
                )
        with dl2:
            with open(raw_path, "rb") as f:
                st.download_button(
                    label="Download Raw Data",
                    data=f,
                    file_name=os.path.basename(raw_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_pa_raw",
                )


# ══════════════════════════════════════════════════════════════════════════════
# IDOC TAB HELPER — shared rendering logic for UY / HN / VE
# ══════════════════════════════════════════════════════════════════════════════

def _render_idoc_tab(country_name, country_upper, prefix, tab_key,
                     prod_file_obj, test_file_obj, run_btn):
    """Render the comparison results for an IDOC country tab."""

    if not (run_btn and prod_file_obj and test_file_obj):
        return

    with st.spinner("Reading and comparing IDOC files..."):
        prod_path = save_upload(prod_file_obj)
        test_path = save_upload(test_file_obj)

        try:
            prod_hdr, prod_lines, prod_docnum = parse_file_idoc(prod_path)
            test_hdr, test_lines, test_docnum = parse_file_idoc(test_path)

            prod_label = prod_docnum or prod_file_obj.name
            test_label = test_docnum or test_file_obj.name

            hdr_results  = compare_headers(prod_hdr, test_hdr)
            line_results = compare_idoc_lines(prod_lines, test_lines)

            ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir  = tempfile.gettempdir()
            out_path = os.path.join(out_dir, f"{prefix}_Comparison_{ts}.xlsx")
            build_report_idoc(prod_path, test_path, country_name, prefix, output_path=out_path)

            raw_path = os.path.join(out_dir, f"{prefix}_RawData_{ts}.xlsx")
            build_raw_export_idoc(prod_path, test_path, country_name, prefix, output_path=raw_path)

        finally:
            os.unlink(prod_path)
            os.unlink(test_path)

    st.success(f"Comparison complete — {country_name} (IDOC)")

    render_legend("MISSING IN TESTING", "EXTRA IN TESTING")
    st.markdown("")

    def _style(row):
        return style_row(row, CR_STATUS_LABELS)

    def _hdr_row(r):
        return {
            "Field (Technical Name)":             r["field"],
            f"Production Value ({prod_label})":   r["ecc_value"] if r["ecc_value"] is not None else "—",
            f"Testing Value ({test_label})":       r["s4_value"]  if r["s4_value"]  is not None else "—",
            "Status": CR_STATUS_LABELS[r["status"]],
        }

    def _line_row(r):
        p, t = r["prod"] or {}, r["test"] or {}
        return {
            "Line #":                                    str(p.get("line_num")   or t.get("line_num")   or "—"),
            f"Prod: EAN/Material ({prod_label})":        p.get("ean")            or "—",
            f"Test: EAN/Material ({test_label})":        t.get("ean")            or "—",
            "Qty":                                       str(p.get("quantity")   or t.get("quantity")   or "—"),
            f"Prod: Net Amount ({prod_label})":          p.get("net_amount")     or "—",
            f"Test: Net Amount ({test_label})":          t.get("net_amount")     or "—",
            f"Prod: Unit Price ({prod_label})":          p.get("unit_price")     or "—",
            f"Test: Unit Price ({test_label})":          t.get("unit_price")     or "—",
            "Status": CR_STATUS_LABELS[r["status"]],
        }

    # Section 1: Headers
    n_match   = sum(1 for r in hdr_results if r["status"] == "match")
    n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
    n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")
    section_banner(
        f"{country_upper} — SECTION 1: HEADER DETAILS &nbsp;|&nbsp; "
        f"Production: {len(prod_hdr)} &nbsp; Testing: {len(test_hdr)} &nbsp; "
        f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
    )

    hdr_gaps = [r for r in hdr_results if r["status"] != "match"]
    if hdr_gaps:
        st.dataframe(
            pd.DataFrame([_hdr_row(r) for r in hdr_gaps]).style.apply(_style, axis=1),
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("All header fields match.")

    st.markdown("")

    # Section 2: Line Items
    n_match   = sum(1 for r in line_results if r["status"] == "match")
    n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
    n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")
    section_banner(
        f"{country_upper} — SECTION 2: LINE ITEMS &nbsp;|&nbsp; "
        f"Production: {len(prod_lines)} rows &nbsp; Testing: {len(test_lines)} rows &nbsp; "
        f"Match: {n_match} &nbsp; Missing in Testing: {n_missing} &nbsp; Extra in Testing: {n_extra}"
    )

    line_gaps = [r for r in line_results if r["status"] != "match"]
    if line_gaps:
        st.dataframe(
            pd.DataFrame([_line_row(r) for r in line_gaps]).style.apply(_style, axis=1),
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("All line items match.")

    st.markdown("")

    with st.expander("Full Combined View — all rows (matched, missing, extra)"):
        st.markdown("**Header Details**")
        if hdr_results:
            st.dataframe(
                pd.DataFrame([_hdr_row(r) for r in hdr_results]).style.apply(_style, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.info("No header data found.")

        st.markdown("**Line Items**")
        if line_results:
            st.dataframe(
                pd.DataFrame([_line_row(r) for r in line_results]).style.apply(_style, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.info("No line item data found.")

    st.markdown("")

    dl1, dl2 = st.columns(2)
    with dl1:
        with open(out_path, "rb") as f:
            st.download_button(
                label="Download Comparison Report",
                data=f,
                file_name=os.path.basename(out_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{tab_key}_cmp",
            )
    with dl2:
        with open(raw_path, "rb") as f:
            st.download_button(
                label="Download Raw Data",
                data=f,
                file_name=os.path.basename(raw_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{tab_key}_raw",
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — IDOC COUNTRIES  (UY02 · HN02 · VE02)
# ══════════════════════════════════════════════════════════════════════════════

with tab4:

    st.markdown("### IDOC Countries")
    st.caption("Format: IDOC — shared base structure (E1EDK01 / E1EDP01) confirmed identical across all countries")

    st.info(
        "**Supported countries (all use the same IDOC base format):**  \n"
        "🇺🇾 Uruguay (UY02) &nbsp;·&nbsp; "
        "🇭🇳 Honduras (HN02) &nbsp;·&nbsp; "
        "🇻🇪 Venezuela (VE02)"
    )

    _IDOC_COUNTRIES = {
        "Uruguay (UY02)":   ("Uruguay",   "URUGUAY",   "UY"),
        "Honduras (HN02)":  ("Honduras",  "HONDURAS",  "HN"),
        "Venezuela (VE02)": ("Venezuela", "VENEZUELA", "VE"),
    }

    idoc_country = st.selectbox(
        "Select Country",
        list(_IDOC_COUNTRIES.keys()),
        key="idoc_country_select",
    )

    _cname, _cupper, _prefix = _IDOC_COUNTRIES[idoc_country]

    col1, col2 = st.columns(2)
    with col1:
        prod_file_idoc = st.file_uploader(
            "Production File (baseline)",
            type=["html", "htm"],
            help=f"Production IDOC HTML report from {_cname}",
            key=f"prod_idoc_{_prefix}",
        )
    with col2:
        test_file_idoc = st.file_uploader(
            "Testing File (new system)",
            type=["html", "htm"],
            help=f"Testing IDOC HTML report from {_cname}",
            key=f"test_idoc_{_prefix}",
        )

    st.markdown("")
    run_idoc = st.button(
        "Run Comparison", type="primary",
        disabled=not (prod_file_idoc and test_file_idoc),
        key=f"run_idoc_{_prefix}",
    )

    _render_idoc_tab(
        country_name=_cname, country_upper=_cupper, prefix=_prefix,
        tab_key=f"idoc_{_prefix.lower()}",
        prod_file_obj=prod_file_idoc, test_file_obj=test_file_idoc, run_btn=run_idoc,
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — DOMINICAN REPUBLIC
# ══════════════════════════════════════════════════════════════════════════════

with tab5:

    st.markdown("### Dominican Republic (ZB6)")

    col1, col2 = st.columns(2)
    with col1:
        ecc_file_do = st.file_uploader(
            "EWP File (old system)",
            type=["xml", "html"],
            help="XML export from SAP ECC (ZB6)",
            key="ecc_do",
        )
    with col2:
        s4_file_do = st.file_uploader(
            "S4 File (new system)",
            type=["xml", "html"],
            help="XML export from SAP S4 (F2)",
            key="s4_do",
        )

    st.markdown("")
    run_do = st.button(
        "Run Comparison", type="primary",
        disabled=not (ecc_file_do and s4_file_do),
        key="run_do",
    )

    if run_do and ecc_file_do and s4_file_do:

        with st.spinner("Reading and comparing files..."):
            ecc_path = save_upload(ecc_file_do)
            s4_path  = save_upload(s4_file_do)

            try:
                ecc_hdr,  ecc_lines, ecc_docnum = parse_file_do(ecc_path)
                s4_hdr,   s4_lines,  s4_docnum  = parse_file_do(s4_path)

                ecc_label = ecc_docnum or ecc_file_do.name
                s4_label  = s4_docnum  or s4_file_do.name

                hdr_results  = compare_headers(ecc_hdr, s4_hdr)
                line_results = compare_do_lines(ecc_lines, s4_lines)

                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"DO_Comparison_{ts}.xlsx")
                build_report_do(ecc_path, s4_path, output_path=out_path)

                raw_path = os.path.join(out_dir, f"DO_RawData_{ts}.xlsx")
                build_raw_export_do(ecc_path, s4_path, output_path=raw_path)

            finally:
                os.unlink(ecc_path)
                os.unlink(s4_path)

        st.success("Comparison complete — Dominican Republic (ZB6)")

        # Date warnings
        ecc_date_issues = check_dates(ecc_hdr)
        s4_date_issues  = check_dates(s4_hdr)
        if ecc_date_issues or s4_date_issues:
            with st.expander("⚠️ Date Format Warnings — expected YYYYMMDD", expanded=True):
                if ecc_date_issues:
                    st.markdown(f"**EWP ({ecc_file_do.name})**")
                    for issue in ecc_date_issues:
                        st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")
                if s4_date_issues:
                    st.markdown(f"**S4 ({s4_file_do.name})**")
                    for issue in s4_date_issues:
                        st.warning(f"`{issue['field']}` → `{issue['value']}`  — not a valid YYYYMMDD date")

        render_legend("MISSING IN S4", "EXTRA IN S4")
        st.markdown("")

        def style_do(row):
            return style_row(row, ARG_STATUS_LABELS)

        def make_hdr_row_do(r):
            return {
                "Field":                            r["field"],
                f"EWP Value ({ecc_label})":         r["ecc_value"] if r["ecc_value"] is not None else "—",
                f"S4 Value ({s4_label})":           r["s4_value"]  if r["s4_value"]  is not None else "—",
                "Status": ARG_STATUS_LABELS[r["status"]],
            }

        def make_line_row_do(r):
            e, s = r["ecc"] or {}, r["s4"] or {}
            return {
                "Line #":                                str(e.get("line_num")      or s.get("line_num")      or "—"),
                f"EWP: MaterialNumber ({ecc_label})":    e.get("material_num")      or "—",
                f"S4: MaterialNumber ({s4_label})":      s.get("material_num")      or "—",
                "Material Desc":                         str(e.get("material_desc") or s.get("material_desc") or "—"),
                f"EWP: LineItemAmount ({ecc_label})":    e.get("line_amount")       or "—",
                f"S4: LineItemAmount ({s4_label})":      s.get("line_amount")       or "—",
                f"EWP: GrossPrice ({ecc_label})":        e.get("gross_price")       or "—",
                f"S4: GrossPrice ({s4_label})":          s.get("gross_price")       or "—",
                "Status": ARG_STATUS_LABELS[r["status"]],
            }

        # Section 1: Headers
        n_match   = sum(1 for r in hdr_results if r["status"] == "match")
        n_missing = sum(1 for r in hdr_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in hdr_results if r["status"] == "extra_in_s4")
        section_banner(
            f"DOMINICAN REPUBLIC — SECTION 1: HEADER DETAILS &nbsp;|&nbsp; "
            f"EWP: {len(ecc_hdr)} &nbsp; S4: {len(s4_hdr)} &nbsp; "
            f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        )

        hdr_gaps = [r for r in hdr_results if r["status"] != "match"]
        if hdr_gaps:
            st.dataframe(
                pd.DataFrame([make_hdr_row_do(r) for r in hdr_gaps]).style.apply(style_do, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All header fields match.")

        st.markdown("")

        # Section 2: Line Items
        n_match   = sum(1 for r in line_results if r["status"] == "match")
        n_missing = sum(1 for r in line_results if r["status"] == "missing_in_s4")
        n_extra   = sum(1 for r in line_results if r["status"] == "extra_in_s4")
        section_banner(
            f"DOMINICAN REPUBLIC — SECTION 2: LINE ITEM / PRICING DETAILS &nbsp;|&nbsp; "
            f"EWP: {len(ecc_lines)} rows &nbsp; S4: {len(s4_lines)} rows &nbsp; "
            f"Match: {n_match} &nbsp; Missing in S4: {n_missing} &nbsp; Extra in S4: {n_extra}"
        )

        line_gaps = [r for r in line_results if r["status"] != "match"]
        if line_gaps:
            st.dataframe(
                pd.DataFrame([make_line_row_do(r) for r in line_gaps]).style.apply(style_do, axis=1),
                use_container_width=True, hide_index=True,
            )
        else:
            st.success("All line items match.")

        st.markdown("")

        with st.expander("Full Combined View — all rows (matched, missing, extra)"):
            st.markdown("**Header Details**")
            if hdr_results:
                st.dataframe(
                    pd.DataFrame([make_hdr_row_do(r) for r in hdr_results]).style.apply(style_do, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No header data found.")

            st.markdown("**Line Item / Pricing Details**")
            if line_results:
                st.dataframe(
                    pd.DataFrame([make_line_row_do(r) for r in line_results]).style.apply(style_do, axis=1),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("No line item data found.")

        st.markdown("")

        dl1, dl2 = st.columns(2)
        with dl1:
            with open(out_path, "rb") as f:
                st.download_button(
                    label="Download Comparison Report",
                    data=f,
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_do_cmp",
                )
        with dl2:
            with open(raw_path, "rb") as f:
                st.download_button(
                    label="Download Raw Data",
                    data=f,
                    file_name=os.path.basename(raw_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_do_raw",
                )
