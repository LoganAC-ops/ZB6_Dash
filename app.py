#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
E-Invoicing Comparison Report  |  Streamlit UI
Mondelez International | Accenture
"""

import streamlit as st
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

# ─── Tab Styling ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* ZB6 group — first 3 tabs: subtle blue tint */
.stTabs [data-baseweb="tab-list"] button:nth-child(1),
.stTabs [data-baseweb="tab-list"] button:nth-child(2),
.stTabs [data-baseweb="tab-list"] button:nth-child(3) {
    background-color: #E8F0FE;
    border-top: 3px solid #4A90D9;
    border-radius: 6px 6px 0 0;
    font-weight: 600;
}
/* Costa Rica tab — 4th: neutral */
.stTabs [data-baseweb="tab-list"] button:nth-child(4) {
    background-color: #F0F4F0;
    border-top: 3px solid #88A888;
    border-radius: 6px 6px 0 0;
}
/* IDOC tab — 5th: warm amber tint */
.stTabs [data-baseweb="tab-list"] button:nth-child(5) {
    background-color: #FFF8E7;
    border-top: 3px solid #E6A817;
    border-radius: 6px 6px 0 0;
    font-weight: 600;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
}
</style>
""", unsafe_allow_html=True)

# ─── Country Tabs ─────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Argentina (ZB6)",
    "Panama (ZB6)",
    "Dominican Republic (ZB6)",
    "Costa Rica",
    "IDOC Countries (UY · HN · VE)",
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
        with st.spinner("Building comparison report..."):
            ecc_path = save_upload(ecc_file)
            s4_path  = save_upload(s4_file)
            try:
                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"ARG_Comparison_{ts}.xlsx")
                build_report(ecc_path, s4_path, output_path=out_path)
                raw_path = os.path.join(out_dir, f"ARG_RawData_{ts}.xlsx")
                build_raw_export(ecc_path, s4_path, output_path=raw_path)
            finally:
                os.unlink(ecc_path)
                os.unlink(s4_path)
        st.success("Report ready — Argentina (ZB6)")
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

with tab4:

    st.markdown("### Costa Rica")

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

        with st.spinner("Building comparison report..."):
            prod_path = save_upload(prod_file)
            test_path = save_upload(test_file)

            try:
                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_path = os.path.join(tempfile.gettempdir(), f"CR_Comparison_{ts}.xlsx")
                build_report_cr(prod_path, test_path, output_path=out_path)
            finally:
                os.unlink(prod_path)
                os.unlink(test_path)

        st.success("Report ready — Costa Rica")
        with open(out_path, "rb") as f:
            st.download_button(
                label="Download Comparison Report",
                data=f,
                file_name=os.path.basename(out_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_cr_cmp",
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PANAMA
# ══════════════════════════════════════════════════════════════════════════════

with tab2:

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
        with st.spinner("Building comparison report..."):
            prod_path = save_upload(prod_file_pa)
            test_path = save_upload(test_file_pa)
            try:
                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"PA_Comparison_{ts}.xlsx")
                build_report_pa(prod_path, test_path, output_path=out_path)
                raw_path = os.path.join(out_dir, f"PA_RawData_{ts}.xlsx")
                build_raw_export_pa(prod_path, test_path, output_path=raw_path)
            finally:
                os.unlink(prod_path)
                os.unlink(test_path)
        st.success("Report ready — Panama (ZB6)")
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


def _render_idoc_tab(country_name, prefix, tab_key, prod_file_obj, test_file_obj, run_btn):
    if not (run_btn and prod_file_obj and test_file_obj):
        return
    with st.spinner("Building comparison report..."):
        prod_path = save_upload(prod_file_obj)
        test_path = save_upload(test_file_obj)
        try:
            ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir  = tempfile.gettempdir()
            out_path = os.path.join(out_dir, f"{prefix}_Comparison_{ts}.xlsx")
            build_report_idoc(prod_path, test_path, country_name, prefix, output_path=out_path)
            raw_path = os.path.join(out_dir, f"{prefix}_RawData_{ts}.xlsx")
            build_raw_export_idoc(prod_path, test_path, country_name, prefix, output_path=raw_path)
        finally:
            os.unlink(prod_path)
            os.unlink(test_path)
    st.success(f"Report ready — {country_name} (IDOC)")
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
# TAB 5 — IDOC COUNTRIES  (UY02 · HN02 · VE02)
# ══════════════════════════════════════════════════════════════════════════════

with tab5:

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
        country_name=_cname, prefix=_prefix,
        tab_key=f"idoc_{_prefix.lower()}",
        prod_file_obj=prod_file_idoc, test_file_obj=test_file_idoc, run_btn=run_idoc,
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — DOMINICAN REPUBLIC
# ══════════════════════════════════════════════════════════════════════════════

with tab3:

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
        with st.spinner("Building comparison report..."):
            ecc_path = save_upload(ecc_file_do)
            s4_path  = save_upload(s4_file_do)
            try:
                ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_dir  = tempfile.gettempdir()
                out_path = os.path.join(out_dir, f"DO_Comparison_{ts}.xlsx")
                build_report_do(ecc_path, s4_path, output_path=out_path)
                raw_path = os.path.join(out_dir, f"DO_RawData_{ts}.xlsx")
                build_raw_export_do(ecc_path, s4_path, output_path=raw_path)
            finally:
                os.unlink(ecc_path)
                os.unlink(s4_path)
        st.success("Report ready — Dominican Republic (ZB6)")
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
