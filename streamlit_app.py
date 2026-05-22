"""
streamlit_app.py — ABHL & Associates Excel → Tally Prime XML
Web interface. Deployed at: https://abhl-tally-import.streamlit.app
"""

import io
import pandas as pd
import streamlit as st

from converter import convert_workbook


# --------------------------------------------------------------------------
# page config
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="ABHL Excel → Tally XML Converter",
    page_icon="📒",
    layout="centered",
)

# --------------------------------------------------------------------------
# header
# --------------------------------------------------------------------------

st.markdown(
    """
    <div style="text-align:center; padding: 0.5rem 0 1rem 0;">
        <h1 style="margin-bottom:0.2rem;">📒 Excel → Tally Prime XML</h1>
        <p style="color:#666; margin-top:0;">
            ABHL &amp; Associates · Chartered Accountants
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write(
    "Convert your Excel voucher sheet into a Tally Prime-importable XML in "
    "three clicks. Auto-detects the header row no matter where it starts "
    "(A1, C3, anywhere)."
)

# --------------------------------------------------------------------------
# instructions expander
# --------------------------------------------------------------------------

with st.expander("ℹ️  How to use", expanded=False):
    st.markdown(
        """
        **Step 1.** Upload your Excel file (`.xlsx`).  
        **Step 2.** Confirm the sheet name (defaults to `Sheet1`).  
        **Step 3.** Click **Convert**, then download the XML.  
        **Step 4.** In Tally Prime: **Gateway of Tally → Import → Vouchers
        → XML File → browse to the downloaded XML.**

        **Required columns** (case-insensitive, in any position on the sheet):
        `DATE`, `Vch_num`, `Narration`, `VCHTYPE`, `Dr_Ledger_1`, `Dr_amt_1`,
        `Cr_Ledger_1`, `Cr_amt_1`, and optionally up to `Dr_Ledger_4` /
        `Cr_Ledger_4` for multi-leg journals.

        **Voucher types supported:** PAYMENT, RECEIPT, JOURNAL, CONTRA, or
        any custom type already defined in your Tally company.

        **Dates** can be Excel date cells or text in any of:
        DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD, DD-Mon-YYYY.
        """
    )

# --------------------------------------------------------------------------
# upload + options
# --------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Upload your Excel file",
    type=["xlsx", "xlsm"],
    help="Only .xlsx / .xlsm files. Max 200 MB.",
)

col1, col2 = st.columns(2)
with col1:
    sheet_name = st.text_input(
        "Sheet name",
        value="Sheet1",
        help="The sheet holding voucher data. Leave as 'Sheet1' if unsure — "
             "the app will auto-detect if it isn't found.",
    )
with col2:
    st.write("")  # spacer
    st.write("")
    convert_clicked = st.button("🔄 Convert to Tally XML",
                                type="primary", use_container_width=True)

# --------------------------------------------------------------------------
# conversion
# --------------------------------------------------------------------------

if convert_clicked:
    if not uploaded:
        st.error("Please upload an Excel file first.")
        st.stop()

    with st.spinner("Reading workbook and building XML…"):
        try:
            xml_text, log = convert_workbook(uploaded.getvalue(),
                                             sheet_name=sheet_name.strip() or None)
        except ValueError as e:
            st.error(f"❌ {e}")
            st.stop()
        except Exception as e:
            st.error(f"❌ Unexpected error: {e}")
            st.exception(e)
            st.stop()

    ok_count = sum(1 for r in log if r["status"] == "ok")
    skip_count = sum(1 for r in log if r["status"] == "skipped")

    # ------------------ summary ------------------
    st.success(
        f"✅ Converted **{ok_count} vouchers** successfully"
        + (f" · ⚠️ Skipped **{skip_count}** problematic rows" if skip_count else "")
    )

    # ------------------ download button ------------------
    file_stem = uploaded.name.rsplit(".", 1)[0]
    st.download_button(
        label="⬇️  Download Tally XML",
        data=xml_text.encode("utf-8"),
        file_name=f"{file_stem}_tally.xml",
        mime="application/xml",
        type="primary",
        use_container_width=True,
    )

    # ------------------ row-level log ------------------
    if log:
        with st.expander(f"📋 Row-by-row log ({len(log)} rows)", expanded=skip_count > 0):
            df = pd.DataFrame(log)
            df = df.rename(columns={
                "row": "Excel Row",
                "status": "Status",
                "reason": "Reason (if skipped)",
                "voucher_num": "Voucher #",
            })

            def colour(s):
                return ("background-color:#d4edda" if s == "ok"
                        else "background-color:#fff3cd")

            st.dataframe(
                df.style.map(colour, subset=["Status"]),
                use_container_width=True,
                hide_index=True,
            )

    # ------------------ preview ------------------
    with st.expander("🔍 Preview generated XML", expanded=False):
        st.code(xml_text[:3000] + ("\n..." if len(xml_text) > 3000 else ""),
                language="xml")

# --------------------------------------------------------------------------
# footer
# --------------------------------------------------------------------------

st.divider()
st.markdown(
    """
    <div style="text-align:center; color:#888; font-size:0.85rem;">
        Built for ABHL &amp; Associates · Chartered Accountants (FRN: 139200W)<br>
        Akshar Business Park, Sector 25, Vashi, Navi Mumbai
    </div>
    """,
    unsafe_allow_html=True,
)
