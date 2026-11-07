"""Printing Daily Report Fixer – upload the downloaded report, get it back corrected."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from report_fixer import fix_report
from report_fixer.processor import DASH, Sheet

st.set_page_config(page_title="Printing Daily Report Fixer", page_icon="🧾", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1200px; }
    .hero {
        background: radial-gradient(ellipse at center, #2a2a2a 0%, #111111 75%);
        border: 1px solid #B8860B; border-radius: 18px;
        padding: 1.75rem 2.25rem; margin-bottom: 1.5rem;
    }
    .hero h1 {
        margin: 0; font-size: 1.9rem; font-weight: 800;
        background: linear-gradient(180deg, #FDE68A 0%, #F2B705 45%, #B8860B 100%);
        -webkit-background-clip: text; background-clip: text; color: transparent;
    }
    .hero p { margin: .5rem 0 0 0; color: #C9C9C9; font-size: .95rem; line-height: 1.5; }
    div[data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #B8860B; border-radius: 10px; background-color: #1C1C1C;
    }
    .stDownloadButton > button { border-radius: 10px; font-weight: 700; border: 1px solid #B8860B; }
    </style>
    <div class="hero">
      <h1>Printing Daily Report Fixer</h1>
      <p>Upload the Printing Daily Report exactly as you downloaded it. The website corrects
      the data and formatting, and you download a clean report.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("Upload the Printing Daily Report (.xlsx)", type=["xlsx", "xlsm"])

with st.expander("What gets corrected?"):
    st.markdown(
        """
1. **Empty item rows removed**: a row is removed only when it has an item name and nothing else.
2. **Day Consumption** is converted from tons to **Kgs**, and *No of days Consumption* is recalculated
   as Closing Stock ÷ Day Consumption (Kgs).
3. **All RAW Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty and Machine Line.
4. **All TagEnv Aging Stock**: removes Lot Code, Lot No, Supplier Lot No, Width, GSM and Source Type.
   Repeated item names are grouped into one row showing the **average Qty and Holding Days**. In Excel, click the
   **+** at the left edge (or the **1 / 2** buttons at the top-left) to show the individual lots.
5. **Zero values** are shown as **-** on every sheet.
6. **Decimals** are rounded to the nearest whole number.
7. Every sheet gets the same professional layout: coloured headers, borders, alignment, column widths,
   frozen header rows and print settings.
        """
    )

zero_is_empty = st.checkbox(
    "Treat 0 as empty when removing rows",
    value=True,
    help="When ticked, a row whose values are all 0 or blank (only the item name is filled) is removed. "
    "Untick to remove only rows where every other cell is completely blank.",
)

if not uploaded:
    st.info("Waiting for a report to be uploaded.")
    st.stop()


@st.cache_data(show_spinner=False)
def run(data: bytes, zero_is_empty: bool) -> tuple[bytes, list[Sheet]]:
    return fix_report(data, zero_is_empty=zero_is_empty)


try:
    with st.spinner("Correcting the report…"):
        output, sheets = run(uploaded.getvalue(), zero_is_empty)
except Exception as exc:  # noqa: BLE001 – show any parsing problem to the user
    st.error(f"Could not process this file. Is it the Printing Daily Report export?\n\n{exc}")
    st.stop()

out_name = f"{Path(uploaded.name).stem}_Corrected.xlsx"
st.success("Your report is ready.")
st.download_button(
    "⬇️  Download corrected report",
    data=output,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

with st.expander("Changes made", expanded=False):
    for sheet in sheets:
        st.markdown(f"**{sheet.name}**")
        st.markdown("\n".join(f"- {n}" for n in sheet.notes) or "- No changes needed")


def to_frame(table, kinds=("data", "group", "total")) -> pd.DataFrame:
    rows = [r.values for r in table.rows if r.kind in kinds]
    df = pd.DataFrame(rows, columns=table.header)
    # Mixed number / "-" columns are shown as text so Streamlit does not complain.
    return df.astype(object).where(df.notna(), "").astype(str)


st.subheader("Preview")
tabs = st.tabs([s.name for s in sheets])
for tab, sheet in zip(tabs, sheets):
    with tab:
        if sheet.title:
            st.markdown(f"#### {sheet.title}")
        for t_index, table in enumerate(sheet.tables):
            if table.title:
                st.markdown(f"**{table.title}**")

            has_groups = any(r.kind == "group" for r in table.rows)
            if not has_groups:
                st.dataframe(to_frame(table), hide_index=True, use_container_width=True)
                continue

            # Grouped sheet: main view lists every item once; click a row to see its lots.
            search = st.text_input("Search item name", key=f"search-{sheet.name}-{t_index}")
            summary = to_frame(table, kinds=("data", "group"))
            if search:
                summary = summary[summary.iloc[:, 0].str.contains(search, case=False, regex=False)]
            st.caption(
                f"{len(summary)} items. Rows with more than one lot show averages. "
                "Click a row to see its individual lots."
            )
            event = st.dataframe(
                summary,
                hide_index=True,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                key=f"grid-{sheet.name}-{t_index}",
            )
            picked = event.selection.rows if event and event.selection else []
            if picked:
                item = summary.iloc[picked[0], 0]
                lots, inside = [], False
                for r in table.rows:
                    if r.kind in ("data", "group"):
                        inside = r.values[0] == item
                        if inside and r.kind == "data":
                            lots.append(r.values)
                    elif r.kind == "detail" and inside:
                        lots.append(r.values)
                detail = pd.DataFrame(lots, columns=table.header).drop(columns=["No. of Lots"], errors="ignore")
                detail = detail.rename(columns={"Qty (Avg)": "Qty", "Holding Days (Avg)": "Holding Days"})
                detail = detail.astype(object).where(detail.notna(), "").astype(str).replace({"": DASH})
                st.markdown(f"**{item}** – {len(lots)} lot(s)")
                st.dataframe(detail, hide_index=True, use_container_width=True)
