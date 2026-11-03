import io
import zipfile

import pytest

import betteroffice_xlsx as bo


def protect_budget(data):
    source = zipfile.ZipFile(io.BytesIO(data))
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            part = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                part = part.replace(b"</sheetData>", b'</sheetData><sheetProtection sheet="1"/>', 1)
                assert b"sheetProtection" in part
            target.writestr(item, part)
    return output.getvalue()


def test_locked_cells_on_a_protected_sheet_refuse_edits(sample_bytes):
    wb = bo.Workbook.open(protect_budget(sample_bytes))
    version = wb.version()
    with pytest.raises(bo.XlsxError, match=r'^B3 is locked on protected sheet "Budget"$'):
        wb.set("Budget", "B3", "1000")
    with pytest.raises(bo.XlsxError, match=r'^B3 is locked on protected sheet "Budget"$'):
        wb.set_many("Budget", [("B3", "1000")])
    with pytest.raises(bo.XlsxError, match=r'^sheet "Budget" is protected$'):
        wb.set_style("Budget", "B3", bold=True)
    assert wb.version() == version
    assert wb["Budget"].formula("D3") == "B3+C3"
    assert wb["Budget"]["D3"] == pytest.approx(157.0)
    assert wb.set("Summary", "A1", "open").applied is True


def test_an_unprotected_sheet_takes_the_same_edit(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.set("Budget", "B3", "1000").applied is True
