import io
import zipfile
from pathlib import Path
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
    return replace_worksheet(sample_bytes, worksheet.encode())


def replace_worksheet(sample_bytes, worksheet):
    buffer = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(sample_bytes)) as source,
        zipfile.ZipFile(buffer, "w") as target,
    ):
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                data = worksheet
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


def test_shared_formulas_recalculate_and_survive_save(sample_bytes, tmp_path):
    fixture = (
        Path(__file__).resolve().parents[3]
        / "crates/xlsx-parse/tests/fixtures/shared-formulas.xml"
    )
    path = tmp_path / "shared.xlsx"
    path.write_bytes(replace_worksheet(sample_bytes, fixture.read_bytes()))
    workbook = bo.Workbook.open_path(path)
    sheet = workbook[0]
    formula = '$B$6&"A"&$B$4&"B"&$D$2'
    for address in ("A1", "A2", "A3"):
        assert sheet.formula(address) == formula
    workbook.recalculate()
    assert [sheet[address] for address in ("A1", "A2", "A3")] == ["6A4B2"] * 3
    sheet["B6"] = 9
    assert [sheet[address] for address in ("A1", "A2", "A3")] == ["9A4B2"] * 3
    reopened = bo.Workbook.open(workbook.save())
    for address in ("A1", "A2", "A3"):
        assert reopened[0].formula(address) == formula
        assert reopened[0][address] == "9A4B2"


def test_shared_mixed_references_track_translated_dependencies(sample_bytes):
    worksheet = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1"><c r="A1"><f t="shared" si="0" ref="A1:C1">A2+$D2+A$3+$D$3</f></c><c r="B1"><f t="shared" si="0"/></c><c r="C1"><f t="shared" si="0"/></c></row>
            <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2"><v>4</v></c></row>
            <row r="3"><c r="A3"><v>10</v></c><c r="B3"><v>20</v></c><c r="C3"><v>30</v></c><c r="D3"><v>40</v></c></row>
        </sheetData>
    </worksheet>"""
    workbook = bo.Workbook.open(replace_worksheet(sample_bytes, worksheet))
    sheet = workbook[0]
    addresses = ("A1", "B1", "C1")
    formulas = ("A2+$D2+A$3+$D$3", "B2+$D2+B$3+$D$3", "C2+$D2+C$3+$D$3")
    assert [sheet.formula(address) for address in addresses] == list(formulas)
    workbook.recalculate()
    assert [sheet[address] for address in addresses] == [55, 66, 77]
    sheet["B2"] = 5
    assert [sheet[address] for address in addresses] == [55, 69, 77]
    sheet["D2"] = 8
    assert [sheet[address] for address in addresses] == [59, 73, 81]
    reopened = bo.Workbook.open(workbook.save())[0]
    assert [reopened.formula(address) for address in addresses] == list(formulas)
    assert [reopened[address] for address in addresses] == [59, 73, 81]
