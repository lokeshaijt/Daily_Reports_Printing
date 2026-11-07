# Printing Daily Report Fixer

A simple website for correcting the **Printing Daily Report** export. Upload the `.xlsx` exactly as
downloaded, then download a corrected and properly formatted report. No manual clean-up is needed.

## What it corrects

| # | Rule | Where |
|---|------|-------|
| 1 | Removes a row only when the item name has **no value in any other column**. Rows with at least one value are kept. With *Treat 0 as empty* ticked (the default), rows that contain only zeros are also removed. | all sheets |
| 2 | **Day Consumption** is converted from tons to **Kgs** (× 1000). **No of days Consumption** is recalculated as *Master Reel Closing in Kgs ÷ Day Consumption in Kgs*. The Grand Total gets the total day consumption and its overall days figure. | Daily Report – Master Roll |
| 3 | Removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty and Machine Line. | All RAW Aging Stock |
| 4 | Removes Lot Code, Lot No, Supplier Lot No, Width, GSM and Source Type. Each item name appears once, with **No. of Lots**, **Qty (Avg)** and **Holding Days (Avg)**. The individual lots are grouped underneath and collapsed. Click **+** in the left margin, or the **1 / 2** outline buttons, to expand them. On the website, click a row in the preview to see its lots. | All TagEnv Aging Stock |
| 5, 7 | Zero values are shown as `-`. | all sheets |
| 6 | Decimals are rounded half-up to the nearest whole number. Fraction-style percentage columns (0.15) are first shown as percentages (15) so they do not round to 0. | all sheets |
| 8 | Applies consistent formatting: title and section bands, coloured headers, borders, number format `#,##0`, left-aligned text, right-aligned numbers, column widths sized to the content, frozen header rows and landscape fit-to-width printing. | all sheets |

Small extras: obvious header typos are fixed (*Consumbtion*, *Watage*, *Ouput*, and the pivot's *Row Labels* becomes
*Item Name*). Blank Grand Total cells are filled with the column sum. Stray scratch formulas outside the tables are
dropped. Sheet names and sheet order stay the same as the original.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens at http://localhost:8501.

## Tests

```bash
pip install pytest
python -m pytest -q
```

## Deploy (Streamlit Community Cloud, free)

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **New app**, pick this repository and branch, and set the main file to `app.py`.
3. Click **Deploy** and share the link with the team.

## Code layout

- `app.py`: the web page (upload, download, change log, preview).
- `report_fixer/processor.py`: parses the workbook into tables and applies every correction rule.
- `report_fixer/writer.py`: writes the corrected, formatted workbook.
- `tests/`: tests that run against a synthetic report.
