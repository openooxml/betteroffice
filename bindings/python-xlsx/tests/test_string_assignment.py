import pytest

import betteroffice_xlsx as bo


@pytest.mark.parametrize(
    "value", ["1.0", "001", "1e3", " 42 ", "TRUE", "false", "", "hello", "inf", "nan"]
)
def test_string_values_survive_assignment_and_save(sample_bytes, value):
    wb = bo.Workbook.open(sample_bytes)
    sheet = wb["Budget"]
    sheet["H1"] = value

    assert sheet["H1"] == value
    assert isinstance(sheet["H1"], str)
    assert sheet.formula("H1") is None
    assert bo.Workbook.open(wb.save())["Budget"]["H1"] == value

    sheet["H1"] = None
    assert sheet["H1"] is None


@pytest.mark.parametrize("method", ["assignment", "set", "set_many", "propose"])
@pytest.mark.parametrize("number_format", ["automatic", "text"])
@pytest.mark.parametrize("collaborative", [False, True])
def test_numeric_text_recalculates_without_losing_decimal_places(
    sample_bytes, method, number_format, collaborative
):
    opener = bo.Workbook.open_collaborative if collaborative else bo.Workbook.open
    wb = opener(sample_bytes)
    sheet = wb["Budget"]
    wb.set_number_format("Budget", "H1", number_format)
    wb.set_many("Budget", {"H1": 2, "H2": '="TEST"&$H$1'})

    if method == "assignment":
        sheet["H1"] = "1.0"
    elif method == "set":
        wb.set("Budget", "H1", "1.0")
    elif method == "set_many":
        wb.set_many("Budget", {"H1": "1.0"})
    else:
        proposal = wb.propose("test", [("Budget", "H1", "1.0")])
        assert proposal.edits[0].after == "1.0"
        wb.accept_proposal(proposal.id)

    assert sheet["H1"] == "1.0"
    assert sheet["H2"] == "TEST1.0"
    assert not wb.set("Budget", "H1", "1.0").applied

    wb.undo()
    assert sheet["H1"] == 2.0
    assert sheet["H2"] == "TEST2"
    wb.redo()
    assert sheet["H1"] == "1.0"
    assert sheet["H2"] == "TEST1.0"

    reopened = bo.Workbook.open_recalculated(wb.save())["Budget"]
    assert reopened["H1"] == "1.0"
    assert reopened["H2"] == "TEST1.0"
