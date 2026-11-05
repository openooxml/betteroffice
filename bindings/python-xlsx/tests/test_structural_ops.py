import io
import pickle
import re
import typing
import zipfile
from pathlib import Path

import pytest

import betteroffice_xlsx as bo
from betteroffice_xlsx import ops as op_types

FIXTURES = Path(__file__).resolve().parents[3] / "packages" / "xlsx" / "test-fixtures"
ACCEPTED_UNTYPED_TAGS = {"restoreColDimension", "restoreRowDimension"}
INTERNAL_TAGS = {
    "setCharts",
    "setConditionalFormats",
    "setDefinedNames",
    "restoreAutoFilter",
    "restoreColStyles",
    "restoreRowStyles",
    "restoreSheet",
}


def a1_range(a1):
    def point(text):
        letters = "".join(ch for ch in text if ch.isalpha())
        col = 0
        for ch in letters:
            col = col * 26 + ord(ch) - 64
        return {"row": int(text[len(letters):]) - 1, "col": col - 1}

    start, _, end = a1.partition(":")
    return {"start": point(start), "end": point(end or start)}


def sheet_xml(data, index=1):
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return package.read(f"xl/worksheets/sheet{index}.xml").decode()


def charts_workbook():
    return bo.Workbook.open((FIXTURES / "charts.xlsx").read_bytes())


def test_apply_ops_inserts_rows_as_one_undo_step(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    depth = workbook.history().undo_depth
    result = workbook.apply_ops(
        [
            {"type": "insertRows", "sheet": 0, "at": 2, "count": 2},
            {"type": "insertCols", "sheet": 0, "at": 0, "count": 1},
        ]
    )
    assert isinstance(result, bo.Mutation)
    assert result.applied is True
    assert workbook.formula(0, "E5") == "C5+D5"
    assert workbook.value(0, "C5") == 100
    assert workbook.history().undo_depth == depth + 1
    workbook.undo()
    assert workbook.formula(0, "D3") == "B3+C3"
    assert workbook.value(0, "B3") == 100


def test_apply_ops_accepts_any_iterable_and_no_ops(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.apply_ops([]).applied is False
    generated = ({"type": "addSheet", "index": 3, "name": name} for name in ["Extra"])
    assert workbook.apply_ops(generated).applied is True
    assert workbook.apply_ops(({"type": "renameSheet", "sheet": 3, "name": "Notes"},)).applied
    assert workbook.sheet_names == ["Budget", "Summary", "Styled", "Notes"]


def test_apply_ops_reads_integral_floats_as_javascript_numbers(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.apply_ops([{"type": "insertRows", "sheet": 0.0, "at": 2.0, "count": 1.0}])
    assert workbook.value(0, "B4") == 100
    with pytest.raises(ValueError):
        workbook.apply_ops([{"type": "setColWidth", "sheet": 0, "col": 1, "width": float("nan")}])


def test_apply_ops_uses_the_calculation_context(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    workbook.apply_ops(
        [{"type": "setCell", "sheet": 0, "at": {"row": 0, "col": 6}, "cell": {"formula": "TODAY()"}}],
        now_serial=45000.75,
    )
    assert workbook.value(0, "G1") == 45000


def test_apply_ops_accepts_the_untyped_restore_ops(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    workbook.apply_ops([{"type": "setColWidth", "sheet": 0, "col": 3, "width": 30}])
    result = workbook.apply_ops(
        [{"type": "restoreColDimension", "sheet": 0, "col": 3, "width": None, "hidden_width": None}]
    )
    assert result.applied is True


def test_typed_ops_cover_every_public_tag_the_engine_decodes(sample_bytes):
    with pytest.raises(ValueError) as error:
        bo.Workbook.open(sample_bytes).apply_ops([{"type": "bogus"}])
    message = str(error.value)
    assert message.startswith("bad ops: unknown variant `bogus`")
    accepted = set(re.findall(r"`(\w+)`", message.split("expected one of", 1)[1]))
    typed = {
        typing.get_args(typing.get_type_hints(op)["type"])[0]
        for op in typing.get_args(op_types.XlsxOp)
    }
    assert len(typed) == 28
    assert typed.isdisjoint(ACCEPTED_UNTYPED_TAGS | INTERNAL_TAGS)
    assert typed | ACCEPTED_UNTYPED_TAGS | INTERNAL_TAGS == accepted


def test_apply_ops_refuses_the_internal_ops_as_every_binding_does(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    before = workbook.save()
    with pytest.raises(bo.OperationError, match="^restore sheet operations are internal$"):
        workbook.apply_ops([{"type": "setCharts", "sheet": 0, "charts": []}])
    assert workbook.save() == before


@pytest.mark.parametrize(
    "ops,error,message",
    [
        ([{"type": "insertRows", "sheet": 0, "at": "x", "count": 1}], ValueError, "bad ops: "),
        ([{"type": "insertRows", "sheet": 9, "at": 0, "count": 1}], bo.RangeError, "sheet 9 out of range"),
        (
            [{"type": "insertRows", "sheet": 0, "at": 0, "count": 0}],
            bo.OperationError,
            "row operation count must be positive",
        ),
        (
            [{"type": "removeConditionalFormat", "sheet": 0, "id": "cf9"}],
            bo.OperationError,
            "conditional format cf9 not found",
        ),
    ],
)
def test_apply_ops_failures_are_typed_and_change_nothing(sample_bytes, ops, error, message):
    workbook = bo.Workbook.open(sample_bytes)
    version = workbook.version()
    with pytest.raises(error) as raised:
        workbook.apply_ops(ops)
    assert message in str(raised.value)
    assert workbook.version() == version
    assert workbook.can_undo is False


def test_collaborative_workbooks_refuse_structural_ops(sample_bytes):
    workbook = bo.Workbook.open_collaborative(sample_bytes, client_id=7)
    with pytest.raises(bo.CollaborativeStructureError) as raised:
        workbook.apply_ops([{"type": "insertRows", "sheet": 0, "at": 2, "count": 1}])
    assert str(raised.value) == "structural operations are unavailable in collaborative mode"
    with pytest.raises(bo.CollaborativeStructureError):
        workbook.edit_charts([{"sheet": 0, "chart": 0, "action": "set_chart_title", "title": "x"}])
    assert workbook.can_undo is False


def test_collaborative_sheet_limits_raise_the_structure_error(sample_bytes):
    left = bo.Workbook.open_collaborative(sample_bytes, client_id=7)
    right = bo.Workbook.open_collaborative(sample_bytes, client_id=8)
    added = left.sheet_count
    with pytest.raises(bo.CollaborativeStructureError) as raised:
        left.apply_ops([{"type": "addSheet", "index": 0, "name": "First"}])
    assert str(raised.value) == "collaborative sheets can only be added at the end of the workbook"
    left.apply_ops([{"type": "addSheet", "index": added, "name": "Notes"}])
    right.set(0, "Z9", "peer elsewhere")
    right.apply_update(left.diff(right.state_vector()))
    right.set(added, "A1", "peer")
    left.apply_update(right.diff(left.state_vector()))
    with pytest.raises(bo.CollaborativeStructureError) as raised:
        left.undo()
    assert str(raised.value) == "undo would remove a sheet that holds collaborators' edits"
    assert left.sheet_count == added + 1
    assert left.value(added, "A1") == "peer"


def test_new_errors_roll_up_to_xlsx_error():
    assert issubclass(bo.CollaborativeStructureError, bo.XlsxError)
    assert not issubclass(bo.CollaborativeStructureError, ValueError)
    assert issubclass(bo.OperationError, bo.XlsxError)
    assert issubclass(bo.OperationError, ValueError)
    assert bo.OperationError.__module__ == "betteroffice_xlsx"


def test_a_refused_op_is_caught_as_value_error_and_pickles(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    try:
        workbook.apply_ops([{"type": "insertRows", "sheet": 0, "at": 0, "count": 0}])
    except ValueError as error:
        refused = error
    else:
        raise AssertionError("the op was not refused")
    assert type(refused) is bo.OperationError
    restored = pickle.loads(pickle.dumps(refused))
    assert type(restored) is bo.OperationError
    assert str(restored) == str(refused) == "row operation count must be positive"


def test_move_range_rewrites_references(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.move_range(0, "B3:C3", "F3").applied is True
    assert workbook.value(0, "F3") == 100
    assert workbook.value(0, "B3") is None
    assert workbook.formula(0, "D3") == "F3+G3"
    workbook.undo()
    assert workbook.formula(0, "D3") == "B3+C3"


@pytest.mark.parametrize(
    "args,error,message",
    [
        ((0, "B3:", "F3"), bo.RangeError, "bad range: "),
        ((0, "B3:C3", "F"), bo.RangeError, "bad move destination: "),
        ((9, "B3:C3", "F3"), bo.RangeError, "sheet 9 out of range"),
        (("Budget", "B3:C3", "F3"), ValueError, "bad move range args: "),
    ],
)
def test_move_range_argument_failures(sample_bytes, args, error, message):
    with pytest.raises(error) as raised:
        bo.Workbook.open(sample_bytes).move_range(*args)
    assert str(raised.value).startswith(message)


def test_hyperlinks_set_and_remove_with_typescript_field_names(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    workbook.set_hyperlink(
        0, {"range": a1_range("B3"), "externalTarget": "https://example.com/b3", "tooltip": "Line"}
    )
    saved = sheet_xml(workbook.save())
    assert re.search(r'<hyperlink [^>]*ref="B3"[^>]*tooltip="Line"', saved)
    workbook.set_hyperlink(1, {"range": a1_range("A1"), "location": "Budget!A1"})
    assert 'location="Budget!A1"' in sheet_xml(workbook.save(), 2)
    assert workbook.remove_hyperlink(0, a1_range("B3")).applied is True
    assert "<hyperlink " not in sheet_xml(workbook.save())
    with pytest.raises(ValueError) as raised:
        workbook.set_hyperlink(0, {"location": "Budget!A1"})
    assert str(raised.value).startswith("bad ops: missing field `range`")


def test_edit_charts_and_move_chart(sample_bytes):
    workbook = charts_workbook()
    result = workbook.edit_charts(
        [
            {"sheet": 0, "chart": 0, "action": "set_chart_title", "title": "Quarterly revenue"},
            {"sheet": 0, "chart": 1, "action": "set_series", "series": 0, "name": {"text": "Sales"}},
        ]
    )
    assert result.applied is True
    charts = workbook.charts(0)
    assert charts[0]["title"] == "Quarterly revenue"
    assert charts[1]["series"][0]["name"] == "Sales"
    before = charts[0]["anchorRange"]
    assert workbook.move_chart(0, charts[0]["id"], 64, 20).applied is True
    assert workbook.charts(0)[0]["anchorRange"] != before
    with pytest.raises(bo.OperationError):
        workbook.edit_charts([])
    with pytest.raises(bo.OperationError):
        workbook.move_chart(0, charts[2]["id"], 10, 10)
    with pytest.raises(ValueError) as raised:
        workbook.edit_charts([{"sheet": 0, "chart": 0, "action": "retitle"}])
    assert str(raised.value).startswith("bad chart edits: ")
    with pytest.raises(ValueError) as raised:
        workbook.move_chart(0, 0, 1, 1)
    assert str(raised.value).startswith("bad chart move args: ")
    assert bo.Workbook.open(sample_bytes).charts(0) == []


def test_capture_and_apply_format(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    workbook.set_style(2, "A2", bold=True, fill_color="#ffcc00")
    captured = workbook.capture_format(2, "A2")
    assert captured["rows"] == 1 and captured["columns"] == 1
    assert len(captured["formats"]) == 1
    assert workbook.apply_format(2, "C5:D6", captured).applied is True
    formatting = workbook.selection_formatting(2, "C5:D6")
    assert formatting["bold"] is True
    assert formatting["fillColor"].lower() == "#ffcc00"
    with pytest.raises(bo.RangeError) as raised:
        workbook.capture_format(2, "A2:")
    assert str(raised.value).startswith("bad range: ")
    with pytest.raises(ValueError) as raised:
        workbook.apply_format(2, "C5", {"rows": 1})
    assert str(raised.value).startswith("bad format args: ")
    with pytest.raises(bo.OperationError):
        workbook.apply_format(2, "C5", {**captured, "rows": 2})


def test_existing_calls_raise_the_subclass_and_still_catch_as_xlsx_error(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    with pytest.raises(bo.XlsxError) as raised:
        workbook.set(0, "A1", "x" * 40000)
    assert type(raised.value) is bo.OperationError
    assert str(raised.value) == "cell text exceeds Excel's length limit"
    with pytest.raises(bo.OperationError):
        workbook.set_many(0, {"A1": "x" * 40000})
