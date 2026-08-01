"""
test_ingest.py
Covers src/ingest.py: reading CSV/Excel files, handling messy real-world
formats (title rows, multiple sheets), and clear errors for unusable files.
"""

import pytest
from src import ingest


def test_reads_clean_csv(sample_csv_path):
    df = ingest.from_any_file(sample_csv_path)
    assert list(df.columns) == ["Date", "Region", "Salesperson", "Product", "Target", "Collection", "Orders", "Cost"]
    assert len(df) == 4


def test_handles_title_row_and_blank_row(messy_excel_path):
    """A company name + blank row before the real header shouldn't produce 'Unnamed' columns."""
    df = ingest.from_any_file(messy_excel_path)
    assert not any(str(c).lower().startswith("unnamed") for c in df.columns)
    assert "Sale Date" in df.columns
    assert len(df) == 2


def test_picks_the_sheet_with_real_data(multisheet_excel_path):
    """When sheet 1 is a cover page, the loader should find the sheet with actual data."""
    df = ingest.from_any_file(multisheet_excel_path)
    assert "Date" in df.columns
    assert "Collection" in df.columns
    assert len(df) == 2


def test_unsupported_extension_raises_clear_error(tmp_path):
    bad_file = tmp_path / "data.txt"
    bad_file.write_text("not a real data file")
    with pytest.raises(ValueError, match="Unsupported file type"):
        ingest.from_any_file(str(bad_file))


def test_old_doc_format_raises_helpful_error(tmp_path):
    fake_doc = tmp_path / "report.doc"
    fake_doc.write_text("legacy binary format placeholder")
    with pytest.raises(ValueError, match="\\.docx"):
        ingest.from_any_file(str(fake_doc))


def test_merged_two_row_header(tmp_path):
    """A top row with a merged label + a sub-header row below it should combine correctly."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Region", "Sales Figures", None, None])
    ws.append([None, None, "Target", "Collection", "Orders"])
    ws.merge_cells("C1:E1")
    ws.append(["01-06-2026", "Mumbai", 50000, 48000, 12])
    path = tmp_path / "merged.xlsx"
    wb.save(path)

    df = ingest.from_any_file(str(path))
    assert "Target" in df.columns
    assert "Collection" in df.columns
    assert "Orders" in df.columns
