import pytest

from betteroffice_xlsx import StaleProposalError, Workbook


def target(a1, sheet="sheet:0"):
    return {"sheetId": sheet, "range": {"kind": "a1", "a1": a1}}


def set_b3(version, value):
    return {
        "expectVersion": version,
        "steps": [{"op": "setCellInputs", "target": target("B3"), "inputs": [[value]]}],
    }


def test_reads_and_finds_carry_the_version(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    read = workbook.read_cells({"ranges": [target("D3")]})
    assert read["ok"] is True
    assert read["version"] == workbook.version()
    assert read["ranges"][0]["cells"][0][0] == {
        "a1": "D3",
        "value": {"kind": "number", "value": 157.0},
        "formula": "B3+C3",
        "displayText": "157",
    }
    assert [sheet["name"] for sheet in read["sheets"]] == ["Budget", "Summary", "Styled"]
    found = workbook.find_text({"text": "Quarterly"})
    assert found["matches"][0]["cell"]["a1"] == "A1"
    assert found["truncated"] is False


def test_applies_a_batch_as_one_undo_step(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    base = workbook.version()
    depth = workbook.history().undo_depth
    result = workbook.apply_edits(
        {
            **set_b3(base, "1000"),
            "calculation": {"nowSerial": 45000.0},
        }
    )
    assert result["ok"] is True
    assert result["applied"] is True
    assert result["baseVersion"] == base
    assert result["version"] == workbook.version() != base
    assert "D3" in [cell["a1"] for cell in result["calculation"]["changed"]]
    assert workbook.value(0, "D3") == 1057
    assert workbook.history().undo_depth == depth + 1
    workbook.undo()
    assert workbook.value(0, "B3") == 100


def test_refusals_are_data_and_malformed_requests_raise(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    version = workbook.version()
    workbook.set(0, "H1", "typed")
    stale = workbook.apply_edits(set_b3(version, "1"))
    assert stale["ok"] is False
    assert stale["failure"]["code"] == "stale-version"
    assert stale["version"] == workbook.version()
    assert workbook.value(0, "B3") == 100

    guarded = workbook.validate_edits(
        {
            "expectVersion": workbook.version(),
            "steps": [
                {
                    "op": "patchStyle",
                    "target": target("B3"),
                    "patch": {"bold": True},
                    "expect": {"cells": [[{"displayText": "101"}]]},
                }
            ],
        }
    )
    assert guarded["failure"]["code"] == "content-mismatch"
    assert guarded["failure"]["stepIndex"] == 0

    with pytest.raises(ValueError):
        workbook.apply_edits({"steps": []})
    with pytest.raises(ValueError):
        workbook.apply_edits({"expectVersion": version, "steps": [{"op": "insertRows"}]})


def test_history_none_keeps_undo_and_redo(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    workbook.set(0, "H1", "typed")
    workbook.undo()
    before = workbook.history()
    result = workbook.apply_edits(
        {**set_b3(workbook.version(), "5"), "history": "none", "source": "agent"}
    )
    assert result["source"] == "agent"
    after = workbook.history()
    assert (after.undo_depth, after.redo_depth) == (before.undo_depth, before.redo_depth)
    workbook.redo()
    assert workbook.value(0, "H1") == "typed"


def test_stale_proposals_name_their_sheets(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    proposal = workbook.propose("agent", [(1, "B2", "10")])
    workbook.set(1, "B2", "moved")
    with pytest.raises(StaleProposalError) as raised:
        workbook.accept_proposal(proposal.id)
    assert raised.value.cells == ["B2"]
    assert raised.value.targets == [
        {"sheet": 1, "sheetId": "sheet:1", "row": 1, "col": 1, "a1": "B2"}
    ]


def test_oversized_requests_are_refused(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    result = workbook.apply_edits(set_b3(workbook.version(), "x" * (16 * 1024 * 1024)))
    assert result["ok"] is False
    assert result["failure"]["code"] == "limit-exceeded"
    assert workbook.value(0, "B3") == 100
