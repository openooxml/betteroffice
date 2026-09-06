import io
import zipfile
from pathlib import Path

import pytest

import betteroffice_xlsx as bo


@pytest.mark.parametrize(
    "columns", ["S:V", "$S:$V", "$S:V", "S:$V", "Budget!S:V", "'Budget'!$S:$V"]
)
def test_whole_column_lookup_recalculates_and_round_trips(sample_bytes, columns):
    fixture = (
        Path(__file__).resolve().parents[3]
        / "crates/xlsx-calc/tests/fixtures/whole-column-vlookup.xml"
    )
    worksheet = fixture.read_bytes().replace(b"S:V", columns.encode())
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
    workbook = bo.Workbook.open(buffer.getvalue())
    sheet = workbook[0]
    formula = f"VLOOKUP(2,{columns},4,FALSE)"
    assert sheet.formula("A1") == formula
    assert sheet["A1"] == "two"
    sheet["V2"] = "updated"
    assert sheet["A1"] == "updated"
    workbook.recalculate()
    assert sheet["A1"] == "updated"
    sheet["B1"] = "=" + formula
    assert sheet.formula("B1") == formula
    assert sheet["B1"] == "updated"
    reopened = bo.Workbook.open(workbook.save())
    assert reopened[0].formula("A1") == formula
    assert reopened[0]["A1"] == "updated"
    reopened[0]["V2"] = "reopened"
    assert reopened[0]["A1"] == "reopened"
    assert reopened[0]["B1"] == "reopened"
