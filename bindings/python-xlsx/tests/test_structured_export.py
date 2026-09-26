import re

import pytest

from betteroffice_xlsx import (
    ParseError,
    Workbook,
    export_xlsx_markdown,
    export_xlsx_structured,
    render_xlsx_markdown,
)


def cell(content, sheet, a1):
    return next(
        candidate
        for candidate in content["sheets"][sheet]["cells"]
        if candidate["anchor"]["a1"] == a1
    )


def test_live_export_carries_the_version_and_changes_nothing(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    version = workbook.version()
    saved = workbook.save()
    result = workbook.export_structured()
    assert result["ok"] is True
    assert result["version"] == version
    content = result["content"]
    assert content["anchorScope"] == "session"
    assert content["calculation"] == {"policy": "asStored", "freshness": "unverified"}
    assert [sheet["anchor"]["sheet"]["name"] for sheet in content["sheets"]] == [
        "Budget",
        "Summary",
        "Styled",
    ]
    assert cell(content, 0, "D3") == {
        "id": "s0!D3",
        "anchor": {
            "kind": "cell",
            "sheet": {"sheetId": "sheet:0", "index": 0, "name": "Budget"},
            "a1": "D3",
        },
        "value": {"kind": "number", "value": 157.0},
        "formula": "B3+C3",
        "displayText": "157",
        "numberFormat": "General",
        "formulaResult": "unverified",
        "merge": None,
    }

    markdown = workbook.export_markdown(
        scope=[{"sheet": 0, "range": "A1:D4"}], markdown_options={"maxRows": 3}
    )
    assert markdown["ok"] is True
    rendered = markdown["content"]
    assert "## Budget" in rendered["markdown"]
    assert re.findall(r"<!-- xlsx-export:\d+ -->", rendered["markdown"]) == [
        anchor["marker"] for anchor in rendered["anchors"]
    ]
    assert rendered["truncated"] is True
    assert workbook.version() == version
    assert workbook.save() == saved


@pytest.mark.parametrize("collaborative", [False, True])
def test_cell_anchors_are_batch_targets(sample_bytes, collaborative):
    workbook = (
        Workbook.open_collaborative(sample_bytes, client_id=301)
        if collaborative
        else Workbook.open(sample_bytes)
    )
    exported = workbook.export_structured(scope=[{"sheet": 1}])
    anchor = exported["content"]["sheets"][0]["cells"][0]["anchor"]
    target = {
        "sheetId": anchor["sheet"]["sheetId"],
        "range": {"kind": "a1", "a1": anchor["a1"]},
    }
    catalog = workbook.read_cells({"ranges": []})
    assert target["sheetId"] == catalog["sheets"][1]["sheetId"]
    step = {"op": "setCellInputs", "target": target, "inputs": [["edited"]]}
    request = {"expectVersion": exported["version"], "steps": [step]}
    assert workbook.apply_edits(request)["ok"] is True
    cells = workbook.read_cells({"ranges": [target]})["ranges"][0]["cells"]
    assert cells[0][0]["value"] == {"kind": "text", "value": "edited"}
    assert export_xlsx_structured(sample_bytes)["sheets"][1]["anchor"]["sheet"] == {
        "sheetId": "sheet:1",
        "index": 1,
        "name": "Summary",
    }


def test_refusals_are_data_and_malformed_options_raise(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    refused = workbook.export_structured(scope=[{"sheet": 99}])
    assert refused["ok"] is False
    assert refused["version"] == workbook.version()
    assert refused["failure"]["code"] == "invalid-scope"
    assert refused["failure"]["target"] is None
    hidden = workbook.export_structured(max_cells=0)
    assert hidden["failure"]["code"] == "invalid-options"
    with pytest.raises(ValueError):
        workbook.export_structured(max_cells=-1)
    with pytest.raises(ValueError):
        workbook.export_markdown(markdown_options={"rows": 1})


def test_bytes_export_is_deterministic_and_renders_the_same_markdown(sample_bytes):
    first = export_xlsx_structured(sample_bytes, options={"maxCells": 10})
    assert export_xlsx_structured(sample_bytes, options={"maxCells": 10}) == first
    assert first["anchorScope"] == "snapshot"
    assert "version" not in first
    assert first["truncated"] is True
    assert first["diagnostics"][-1]["code"] == "truncated"

    full = export_xlsx_structured(sample_bytes)
    assert render_xlsx_markdown(full, max_rows=5) == export_xlsx_markdown(
        sample_bytes, markdown_options={"maxRows": 5}
    )
    with pytest.raises(ValueError):
        render_xlsx_markdown({**full, "schemaVersion": 2})
    with pytest.raises(ValueError):
        export_xlsx_structured(sample_bytes, options={"maxBytes": 10})
    with pytest.raises(ParseError):
        export_xlsx_structured(b"not a workbook")
