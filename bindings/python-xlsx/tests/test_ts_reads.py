import io
import zipfile
from pathlib import Path

import pytest

import betteroffice_xlsx as bo

FIXTURES = Path(__file__).resolve().parents[3] / "packages" / "xlsx" / "test-fixtures"


def with_sheet_xml(data, body, dxfs=""):
    output = io.BytesIO()
    worksheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + body
        + "</worksheet>"
    ).encode()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as source,
        zipfile.ZipFile(output, "w") as target,
    ):
        for part in source.infolist():
            contents = source.read(part.filename)
            if part.filename == "xl/worksheets/sheet1.xml":
                contents = worksheet
            elif part.filename == "xl/styles.xml" and dxfs:
                contents = contents.replace(b"</styleSheet>", dxfs.encode() + b"</styleSheet>")
            target.writestr(part, contents)
    return output.getvalue()


def test_cell_returns_editable_input_with_zero_based_coordinates(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.cell("Budget", 0, 0) == {
        "a1": "A1", "input": "Quarterly Budget Report", "isFormula": False,
    }
    assert workbook.cell(0, 2, 3) == {
        "a1": "D3", "input": "=B3+C3", "isFormula": True,
    }
    assert workbook.cell(0, 99, 99) == {
        "a1": "CV100", "input": "", "isFormula": False,
    }


def test_cell_and_range_cells_include_array_flag_only_for_members(sample_bytes):
    data = with_sheet_xml(
        sample_bytes,
        '<sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A2">'
        'ROW(A1:A2)</f><v>1</v></c></row>'
        '<row r="2"><c r="A2"><v>2</v></c></row></sheetData>',
    )
    workbook = bo.Workbook.open(data)
    assert workbook.cell(0, 0, 0) == {
        "a1": "A1", "input": "=ROW(A1:A2)", "isFormula": True, "isArrayFormula": True,
    }
    assert workbook.cell(0, 1, 0)["isArrayFormula"] is True
    assert "isArrayFormula" not in workbook.cell(0, 2, 0)
    assert workbook.range_cells(0, "A1:A3") == [
        [workbook.cell(0, row, 0)] for row in range(3)
    ]


def test_range_cells_returns_a_row_major_matrix(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.range_cells("Budget", "C3:D4") == [
        [workbook.cell(0, row, col) for col in (2, 3)] for row in (2, 3)
    ]
    assert workbook.range_cells(0, "H30:I31") == [
        [
            {"a1": "H30", "input": "", "isFormula": False},
            {"a1": "I30", "input": "", "isFormula": False},
        ],
        [
            {"a1": "H31", "input": "", "isFormula": False},
            {"a1": "I31", "input": "", "isFormula": False},
        ],
    ]


def test_sheet_info_and_sheet_info_for_keep_the_active_tab(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    version = workbook.version()
    saved = workbook.save()
    info = workbook.sheet_info()
    assert set(info) == {
        "sheetIds", "sheetNames", "activeSheet", "contentWidth", "contentHeight",
        "frozenRows", "frozenCols", "initialScrollX", "initialScrollY", "autoFilter",
    }
    assert info["sheetNames"] == workbook.sheet_names
    assert info["sheetIds"] == ["sheet:0", "sheet:1", "sheet:2"]
    assert info["activeSheet"] == 0
    assert info["autoFilter"] is None
    assert info["contentWidth"] > 0
    assert info["contentHeight"] > 0
    assert workbook.sheet_info_for("Summary")["activeSheet"] == 1
    assert workbook.sheet_info_for(0) == info
    assert workbook.active_sheet == 0
    assert workbook.version() == version
    assert workbook.save() == saved


def test_sheet_info_includes_the_requested_sheets_auto_filter(sample_bytes):
    data = with_sheet_xml(
        sample_bytes,
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
        '<autoFilter ref="A1:A3"/>',
    )
    workbook = bo.Workbook.open(data)
    assert workbook.sheet_info()["autoFilter"] == "A1:A3"
    workbook.set_active_sheet(1)
    assert workbook.sheet_info()["autoFilter"] is None
    assert workbook.sheet_info_for("Budget")["autoFilter"] == "A1:A3"
    assert workbook.active_sheet == 1


def test_calculation_status_uses_camel_case_and_omits_false_budget(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.calculation_status() == {"limitedCells": []}
    workbook.set(0, "H1", "=1+1")
    assert workbook.calculation_status() == {"limitedCells": []}


def test_list_defined_names_defaults_pagination_and_scope():
    workbook = bo.Workbook.open((FIXTURES / "defined-names.xlsx").read_bytes())
    all_names = workbook.list_defined_names()
    assert workbook.list_defined_names(None) == all_names
    assert all_names["version"] == workbook.version()
    assert all_names["total"] == 5
    assert all_names["nextOffset"] is None
    first = workbook.list_defined_names({"limit": 2})
    assert [item["name"] for item in first["items"]] == ["Rows", "Header"]
    assert first["items"][0] == {
        "name": "Rows", "hidden": False, "scope": None,
        "refersTo": "Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))",
    }
    assert first["nextOffset"] == 2
    last = workbook.list_defined_names({"offset": first["nextOffset"]})
    assert last["items"] == all_names["items"][2:]
    assert last["nextOffset"] is None
    assert workbook.list_defined_names({"scope": None}) == all_names
    assert workbook.list_defined_names({"scope": 0})["items"] == []
    assert workbook.list_defined_names({"offset": 99})["items"] == []


@pytest.mark.parametrize("names_request", [{"limit": 0}, {"scope": 99}, {"unknown": True}])
def test_list_defined_names_maps_core_errors(sample_bytes, names_request):
    with pytest.raises(bo.XlsxError):
        bo.Workbook.open(sample_bytes).list_defined_names(names_request)


def test_conditional_formats_preserve_wire_shape_and_optional_fields(sample_bytes):
    data = with_sheet_xml(
        sample_bytes,
        '<sheetData/><conditionalFormatting sqref="A1:A3">'
        '<cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0">'
        '<formula>1</formula></cfRule><cfRule type="dataBar" priority="2">'
        '<dataBar/></cfRule></conditionalFormatting>',
        '<dxfs count="1"><dxf><font><b val="0"/><i/><color rgb="FFFF0000"/></font>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/>'
        '</patternFill></fill></dxf></dxfs>',
    )
    workbook = bo.Workbook.open(data)
    assert workbook.conditional_formats("Budget") == [{
        "id": "cf1", "range": "A1:A3", "type": "cellIs", "priority": 1,
        "operator": "greaterThan", "formulas": ["1"], "editable": True, "dxfId": 0,
        "format": {"bold": False, "italic": True, "fontColor": "#ff0000", "fill": "#00ff00"},
    }, {
        "id": "cf2", "range": "A1:A3", "type": "dataBar", "priority": 2,
        "formulas": [], "editable": False, "format": {},
    }]
    assert workbook.conditional_formats(1) == []


def test_search_text_options_and_sheet_metadata(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.search_text("quarterly") == [{
        "sheet": 0, "sheetId": "sheet:0", "sheetName": "Budget",
        "row": 0, "col": 0, "a1": "A1", "text": "Quarterly Budget Report",
    }]
    assert workbook.search_text("quarterly", case_sensitive=True) == []
    assert workbook.search_text("Quarterly", case_sensitive=True) == workbook.search_text("quarterly")
    assert workbook.search_text("Line item", limit=2) == workbook.search_text("Line item")[:2]
    assert len(workbook.search_text("Line item", limit=2)) == 2
    assert workbook.search_text("Line item", limit=0) == []
    assert workbook.search_text("") == []
    workbook.set(1, "H1", "Quarterly summary")
    match = workbook.search_text("Quarterly")[-1]
    assert (match["sheet"], match["sheetId"], match["sheetName"], match["a1"]) == (
        1, "sheet:1", "Summary", "H1",
    )


@pytest.mark.parametrize("limit", [-1, 2**32, 1.5, True])
def test_search_text_rejects_invalid_limits(sample_bytes, limit):
    workbook = bo.Workbook.open(sample_bytes)
    with pytest.raises(bo.RangeError):
        workbook.search_text("Quarterly", limit=limit)


def test_selection_formatting_reports_uniform_fields_and_omits_mixed_ones(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    workbook.set_style(0, "H20:I20", bold=True, font_family="Arial", font_size=14)
    workbook.set_number_format(0, "H20:I20", "0.000")
    formatting = workbook.selection_formatting("Budget", "H20:I20")
    assert formatting["bold"] is True
    assert formatting["fontFamily"] == "Arial"
    assert formatting["fontSize"] == 14
    assert formatting["numberFormat"] == "custom"
    assert formatting["numberFormatPattern"] == "0.000"
    workbook.set_style(0, "I20", bold=False)
    assert "bold" not in workbook.selection_formatting(0, "H20:I20")


def test_merged_range_list_preserves_existing_a1_accessor(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    assert workbook.merged_ranges("Budget", "A1:H30") == ["A1:D1"]
    assert workbook.merged_range_list("Budget", "A1:H30") == [{
        "start": {"row": 0, "col": 0}, "end": {"row": 0, "col": 3},
    }]
    assert workbook.merged_range_list(0, "H20:I21") == []


def test_visible_merged_ranges_accepts_viewport_dictionary(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    viewport = {"x": 0, "y": 0, "width": 800, "height": 600}
    assert workbook.visible_merged_ranges("Budget", viewport) == workbook.merged_range_list(
        0, "A1:H30"
    )
    assert workbook.visible_merged_ranges(1, viewport) == []
    with pytest.raises(bo.RangeError):
        workbook.visible_merged_ranges(0, {**viewport, "width": 0})
    with pytest.raises(ValueError):
        workbook.visible_merged_ranges(0, {"x": 0})


def test_charts_returns_metadata_as_plain_dictionaries(sample_bytes):
    assert bo.Workbook.open(sample_bytes).charts(0) == []
    workbook = bo.Workbook.open((FIXTURES / "charts.xlsx").read_bytes())
    charts = workbook.charts("Charts")
    assert charts == workbook.charts(0)
    assert len(charts) > 0
    chart = charts[0]
    assert chart["index"] == 0
    assert chart["editable"] is True
    assert chart["title"] == "Revenue and costs"
    assert chart["type"] == "barChart"
    assert chart["error"] is None
    assert "anchorRange" in chart
    assert isinstance(chart["series"], list)


@pytest.mark.parametrize(
    "read",
    [
        lambda wb, sheet: wb.cell(sheet, 0, 0),
        lambda wb, sheet: wb.range_cells(sheet, "A1:B2"),
        lambda wb, sheet: wb.sheet_info_for(sheet),
        lambda wb, sheet: wb.conditional_formats(sheet),
        lambda wb, sheet: wb.selection_formatting(sheet, "A1:B2"),
        lambda wb, sheet: wb.merged_range_list(sheet, "A1:B2"),
        lambda wb, sheet: wb.visible_merged_ranges(
            sheet, {"x": 0, "y": 0, "width": 100, "height": 100}
        ),
        lambda wb, sheet: wb.charts(sheet),
    ],
)
def test_new_sheet_reads_use_existing_sheet_errors(sample_bytes, read):
    workbook = bo.Workbook.open(sample_bytes)
    for sheet, error in [(True, TypeError), (99, IndexError), ("Missing", KeyError)]:
        with pytest.raises(error):
            read(workbook, sheet)


def test_new_reads_preserve_workbook_state(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    version, saved, history = workbook.version(), workbook.save(), workbook.history()
    workbook.cell(0, 2, 3)
    workbook.range_cells(0, "A1:D4")
    workbook.sheet_info()
    workbook.sheet_info_for(1)
    workbook.calculation_status()
    workbook.list_defined_names()
    workbook.conditional_formats(0)
    workbook.search_text("Quarterly")
    workbook.selection_formatting(0, "A1:D4")
    workbook.merged_range_list(0, "A1:D4")
    workbook.visible_merged_ranges(0, {"x": 0, "y": 0, "width": 800, "height": 600})
    workbook.charts(0)
    assert workbook.version() == version
    assert workbook.save() == saved
    assert workbook.active_sheet == 0
    assert workbook.history().undo_depth == history.undo_depth


def test_new_cell_and_range_reads_map_range_errors(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    with pytest.raises(bo.RangeError):
        workbook.cell(0, 1_048_576, 0)
    with pytest.raises(bo.RangeError):
        workbook.range_cells(0, "bad")
    with pytest.raises(bo.RangeError):
        workbook.selection_formatting(0, "bad")
    with pytest.raises(bo.RangeError):
        workbook.merged_range_list(0, "bad")
