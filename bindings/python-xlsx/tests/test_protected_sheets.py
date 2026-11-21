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


def test_protected_sheets_allow_writes_locked_cells_and_keeps_the_protection(sample_bytes):
    wb = bo.Workbook.open(protect_budget(sample_bytes))
    assert wb.set("Budget", "B3", "1000", protected_sheets="allow").applied is True
    assert wb.set_many("Budget", [("C3", "7")], protected_sheets="allow").applied is True
    wb.set_style("Budget", "B3", bold=True, protected_sheets="allow")
    wb.set_number_format("Budget", "B3", "percent", protected_sheets="allow")
    assert wb["Budget"]["D3"] == pytest.approx(1007.0)
    saved = wb.save()
    with zipfile.ZipFile(io.BytesIO(saved)) as package:
        assert b"<sheetProtection" in package.read("xl/worksheets/sheet1.xml")
    reopened = bo.Workbook.open(saved)
    with pytest.raises(bo.XlsxError, match=r'^B3 is locked on protected sheet "Budget"$'):
        reopened.set("Budget", "B3", "1")


def test_protected_sheets_refuse_and_the_default_raise(sample_bytes):
    wb = bo.Workbook.open(protect_budget(sample_bytes))
    version = wb.version()
    for options in [{}, {"protected_sheets": "refuse"}]:
        with pytest.raises(bo.XlsxError, match=r'^B3 is locked on protected sheet "Budget"$'):
            wb.set("Budget", "B3", "1000", **options)
        with pytest.raises(bo.XlsxError, match=r'^B3 is locked on protected sheet "Budget"$'):
            wb.set_many("Budget", [("B3", "1000")], **options)
        with pytest.raises(bo.XlsxError, match=r'^sheet "Budget" is protected$'):
            wb.set_style("Budget", "B3", bold=True, **options)
        with pytest.raises(bo.XlsxError, match=r'^sheet "Budget" is protected$'):
            wb.set_number_format("Budget", "B3", "percent", **options)
    assert wb.version() == version


def test_protected_sheets_allow_is_refused_on_a_collaborative_workbook(sample_bytes):
    wb = bo.Workbook.open_collaborative(protect_budget(sample_bytes), client_id=7)
    message = r"protectedSheets allow requires a standalone agent workbook"
    with pytest.raises(ValueError, match=message):
        wb.set("Budget", "B3", "1000", protected_sheets="allow")
    with pytest.raises(ValueError, match=message):
        wb.set_many("Summary", [("A1", "x")], protected_sheets="allow")
    with pytest.raises(ValueError, match=message):
        wb.set_style("Budget", "B3", bold=True, protected_sheets="allow")
    assert wb["Budget"].formula("D3") == "B3+C3"


def test_protected_sheets_rejects_other_values(sample_bytes):
    wb = bo.Workbook.open(protect_budget(sample_bytes))
    version = wb.version()
    for value in ["Allow", "", "yes", None, 5, True, ["allow"]]:
        with pytest.raises(ValueError, match=r"protectedSheets must be refuse or allow"):
            wb.set("Summary", "A1", "x", protected_sheets=value)
        with pytest.raises(ValueError, match=r"protectedSheets must be refuse or allow"):
            wb.set_many("Summary", [("A1", "x")], protected_sheets=value)
        with pytest.raises(ValueError, match=r"protectedSheets must be refuse or allow"):
            wb.set_style("Summary", "A1", bold=True, protected_sheets=value)
        with pytest.raises(ValueError, match=r"protectedSheets must be refuse or allow"):
            wb.set_number_format("Summary", "A1", "percent", protected_sheets=value)
    assert wb.version() == version
    assert wb.set("Summary", "A1", "x").applied is True
