import io
import zipfile
from xml.sax.saxutils import escape

import pytest

import betteroffice_xlsx as bo


def formula_workbook(sample_bytes, formula):
    worksheet = f"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1"><c r="A1" t="str"><f>{escape(formula)}</f><v/></c></row>
            <row r="2"><c r="D2"><v>2</v></c></row>
            <row r="4"><c r="B4"><v>4</v></c></row>
            <row r="6"><c r="B6"><v>6</v></c></row>
        </sheetData>
    </worksheet>"""
    buffer = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(sample_bytes)) as source,
        zipfile.ZipFile(buffer, "w") as target,
    ):
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                data = worksheet.encode()
            target.writestr(part, data)
    return buffer.getvalue()


@pytest.mark.parametrize(
    "formula",
    ['$B$6&"A"&$B$4&"B"&$D$2', 'B6&"A"&B4&"B"&D2', '$B6&"A"&B$4&"B"&D2'],
)
def test_absolute_and_relative_concatenation(sample_bytes, formula, tmp_path):
    path = tmp_path / "demo.xlsx"
    path.write_bytes(formula_workbook(sample_bytes, formula))
    workbook = bo.Workbook.open_path(path)
    sheet = workbook[0]
    assert sheet.formula("A1") == formula
    assert sheet["A1"] == ""
    workbook.recalculate()
    assert sheet["A1"] == "6A4B2"
    sheet["B6"] = 9
    assert sheet["A1"] == "9A4B2"
    reopened = bo.Workbook.open(workbook.save())
    assert reopened[0].formula("A1") == formula
    assert reopened[0]["A1"] == "9A4B2"
