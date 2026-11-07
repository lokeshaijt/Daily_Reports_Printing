"""Run with:  python -m pytest -q"""

import io

import openpyxl
import pytest

from report_fixer import fix_report

DAILY_HEADER = ["Item Name", "Master Reel Opening in Kgs", "Master Reel Inward in Kgs", "Transfer to Slit in Kgs",
                "Master Reel Closing in Kgs", "Pending PO Quantity", "Day Consumption", "No of days Consumption"]
AGING_HEADER = ["Lot Code", "Item Name", "Lot No", "Supplier lot no", "Category", "Width", "GSM", "Source Type",
                "Qty", "Machine Line", "Holding Days"]


def build_report() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Report"
    ws.append(["PRINTING DAILY REPORT - 05 Nov 2026"])
    ws.append([])
    ws.append(["Master Roll"])
    ws.append(DAILY_HEADER)
    ws.append(["PAPER 70 GSM", 76960, 0, 0, 76960, 0, 5.77, 13338, "=G5*1000"])
    ws.append(["PAPER B", None, 20, None, None, None, None, None])
    ws.append(["PAPER C", 0, 0, 0, 0, 0, None, 0])
    ws.append(["PAPER D", None, None, None, None, None, None, None])
    ws.append(["Grand Total", 76960, 20, 0, 76960, 0])

    raw = wb.create_sheet("All RAW Aging Stock")
    raw.append(AGING_HEADER)
    raw.append([1, "PAPER 70 GSM", "L1", "S1", "RAWPAPER", 1250, 70, "SEG", 1, None, 36])

    tag = wb.create_sheet("All TagEnv Aging Stock")
    tag.append(AGING_HEADER)
    tag.append([1, "ENV A", "L1", None, "FGENV", 0, 0, "PRINT", 10.4, "Line 1", 10])
    tag.append([2, "ENV B", "L2", None, "FGENV", 0, 0, "PRINT", 5, "Line 1", 3])
    tag.append([3, "ENV A", "L3", None, "FGENV", 0, 0, "PRINT", 20, "Line 1", 21])

    ink = wb.create_sheet("Ink Stock")
    ink.append(["Row Labels", "Total Qty", "Holding Days", "Per Day Consumbtion"])
    ink.append(["INK GOLD", 0, 0, 4.4])
    ink.append(["INK METALLIC", 0, 0, 0])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@pytest.fixture(scope="module")
def result():
    data, sheets = fix_report(build_report())
    return openpyxl.load_workbook(io.BytesIO(data)), sheets


def rows(ws):
    return [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]


def test_sheet_order_preserved(result):
    wb, _ = result
    assert wb.sheetnames == ["Daily Report", "All RAW Aging Stock", "All TagEnv Aging Stock", "Ink Stock"]


def test_daily_report(result):
    wb, _ = result
    data = rows(wb["Daily Report"])
    names = [r[0] for r in data]
    # Rule 1: B (one value) kept, C (only zeros) and D (only blanks) removed.
    assert "PAPER B" in names and "PAPER C" not in names and "PAPER D" not in names
    header = data[2]
    assert header[6] == "Day Consumption in Kgs"
    paper70 = data[3]
    # Rule 2: 5.77 t -> 5770 kg; 76960 / 5770 = 13.3 -> 13 days.  Rule 7: zeros -> "-".
    assert paper70 == ("PAPER 70 GSM", 76960, "-", "-", 76960, "-", 5770, 13)
    total = data[-1]
    assert total[0] == "Grand Total" and total[6] == 5770 and total[7] == 13
    # The stray scratch formula outside the table is gone.
    assert wb["Daily Report"].max_column == len(DAILY_HEADER)


def test_raw_aging_columns(result):
    wb, _ = result
    data = rows(wb["All RAW Aging Stock"])
    assert data[0] == ("Item Name", "Width", "Source Type", "Holding Days")
    assert data[1] == ("PAPER 70 GSM", 1250, "SEG", 36)


def test_tagenv_grouping(result):
    wb, _ = result
    ws = wb["All TagEnv Aging Stock"]
    data = rows(ws)
    assert data[0] == ("Item Name", "No. of Lots", "Category", "Qty (Avg)", "Machine Line", "Holding Days (Avg)")
    # ENV A appears twice -> summary with averages (10.4+20)/2=15.2 -> 15, (10+21)/2=15.5 -> 16 (half-up).
    assert data[1] == ("ENV A", 2, "FGENV", 15, "Line 1", 16)
    assert data[2][3] == 10 and data[3][3] == 20
    assert ws.row_dimensions[3].outline_level == 1 and ws.row_dimensions[3].hidden
    assert data[4] == ("ENV B", 1, "FGENV", 5, "Line 1", 3)


def test_ink_stock(result):
    wb, _ = result
    data = rows(wb["Ink Stock"])
    assert data[0] == ("Item Name", "Total Qty", "Holding Days", "Per Day Consumption")
    assert data[1:] == [("INK GOLD", "-", "-", 4)]


def test_zero_kept_when_option_off():
    data, _ = fix_report(build_report(), zero_is_empty=False)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    names = [r[0] for r in rows(wb["Daily Report"])]
    assert "PAPER C" in names and "PAPER D" not in names
