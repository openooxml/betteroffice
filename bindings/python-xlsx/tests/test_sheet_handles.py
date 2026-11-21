import pytest

import betteroffice_xlsx as bo


@pytest.fixture
def workbook_with_extra_sheet(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    wb.apply_ops([{"type": "addSheet", "index": wb.sheet_count, "name": "Extra"}])
    return bo.Workbook.open(wb.save())


def test_held_sheet_follows_move(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    first = wb.sheet_names[0]
    held = wb[first]
    last = wb.sheet_count - 1
    wb.apply_ops([{"type": "moveSheet", "from": 0, "to": last}])
    held["A1"] = "Forecast"
    saved = bo.Workbook.open(wb.save())
    assert held.name == first
    assert held.index == last
    assert saved.value(first, "A1") == "Forecast"
    assert saved.value(saved.sheet_names[0], "A1") != "Forecast"


def test_held_sheet_follows_add_before_it(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    first = wb.sheet_names[0]
    held = wb[first]
    wb.apply_ops([{"type": "addSheet", "index": 0, "name": "Inserted"}])
    held["B2"] = "kept"
    saved = bo.Workbook.open(wb.save())
    assert saved.value(first, "B2") == "kept"
    assert saved.value("Inserted", "B2") != "kept"


def test_held_sheet_writes_without_structural_ops(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    first = wb.sheet_names[0]
    held = wb[first]
    held["A1"] = "Forecast"
    saved = bo.Workbook.open(wb.save())
    assert saved.value(first, "A1") == "Forecast"
    assert saved.value("Extra", "A1") != "Forecast"


def test_removed_handle_raises_and_remaining_handle_follows(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    removed, survivor = list(wb)[:2]
    name = survivor.name
    wb.apply_ops([{"type": "removeSheet", "index": 0}])

    with pytest.raises(KeyError, match="sheet was removed"):
        removed["A1"]
    with pytest.raises(KeyError, match="sheet was removed"):
        removed["A1"] = "missing"
    with pytest.raises(KeyError, match="sheet was removed"):
        removed.name
    with pytest.raises(KeyError, match="sheet was removed"):
        removed.index
    assert repr(removed) == "Sheet(<removed>)"

    assert survivor.index == 0
    assert survivor.name == name
    survivor["A1"] = "kept"
    assert wb.value(name, "A1") == "kept"


def test_held_sheet_follows_undo_and_redo_of_move(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    first = wb.sheet_names[0]
    held = wb[first]
    wb.apply_ops([{"type": "moveSheet", "from": 0, "to": 1}])
    assert (held.name, held.index) == (first, 1)

    assert wb.undo().applied
    assert (held.name, held.index) == (first, 0)
    assert wb.redo().applied
    assert (held.name, held.index) == (first, 1)


def test_undo_revives_removed_handle(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    held = wb[0]
    name = held.name
    value = held["B3"]
    wb.apply_ops([{"type": "removeSheet", "index": 0}])

    assert wb.undo().applied
    assert (held.name, held.index) == (name, 0)
    assert held["B3"] == value
    held["A1"] = "restored"
    assert wb.value(name, "A1") == "restored"


def test_new_sheet_handle_follows_a_move_in_a_batch(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    wb.apply_ops([{"type": "addSheet", "index": wb.sheet_count, "name": "New"}])
    held = wb["New"]
    wb.apply_ops(
        [
            {"type": "addSheet", "index": 0, "name": "Inserted"},
            {"type": "moveSheet", "from": held.index + 1, "to": 0},
        ]
    )

    assert (held.name, held.index) == ("New", 0)
    held["A1"] = "new value"
    assert wb.value("New", "A1") == "new value"
    assert wb.value("Inserted", "A1") is None


def test_held_sheet_follows_positional_wire_ops(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    held = wb[0]
    name = held.name
    wb.apply_ops([("addSheet", 0, "Inserted"), ("moveSheet", 1, 2), ("removeSheet", 0)])
    assert (held.name, held.index) == (name, 1)


def test_cell_edit_after_undo_replaces_sheet_history(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    first = wb.sheet_names[0]
    held = wb[first]
    value = held["A1"]
    wb.apply_ops([{"type": "moveSheet", "from": 0, "to": 1}])
    assert wb.undo().applied

    held["A1"] = "Forecast"
    assert wb.undo().applied
    assert (held.name, held.index, held["A1"]) == (first, 0, value)
    assert wb.redo().applied
    assert (held.name, held.index, held["A1"]) == (first, 0, "Forecast")


def test_held_collaborative_sheets_follow_concurrent_appends(sample_bytes):
    left = bo.Workbook.open_collaborative(sample_bytes, client_id=101)
    right = bo.Workbook.open_collaborative(sample_bytes, client_id=202)
    end = left.sheet_count
    left.apply_ops([{"type": "addSheet", "index": end, "name": "Left"}])
    right.apply_ops([{"type": "addSheet", "index": end, "name": "Right"}])
    left_sheet = left["Left"]
    right_sheet = right["Right"]
    assert left_sheet.index == right_sheet.index == end

    left.apply_update(right.diff(left.state_vector()))
    right.apply_update(left.diff(right.state_vector()))
    assert left.sheet_names == right.sheet_names
    assert {left_sheet.index, right_sheet.index} == {end, end + 1}
    assert left_sheet.name == "Left"
    assert right_sheet.name == "Right"

    left_sheet["A1"] = "left value"
    right_sheet["A1"] = "right value"
    left.apply_update(right.diff(left.state_vector()))
    assert left.value("Left", "A1") == "left value"
    assert left.value("Right", "A1") == "right value"


def test_sheet_constructor_takes_a_current_index(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    last = wb.sheet_count - 1
    wb.apply_ops([{"type": "moveSheet", "from": 0, "to": last}])
    sheet = bo.Sheet(wb, 0)
    assert sheet.index == 0
    assert sheet.name == wb.sheet_names[0]


def test_cell_writes_keep_one_history_anchor_per_undo_step(workbook_with_extra_sheet):
    wb = workbook_with_extra_sheet
    held = wb[wb.sheet_names[0]]
    for row in range(1, 201):
        held[f"A{row}"] = row
    assert len(wb._sheet_history) == wb.history().undo_depth + 1
    assert all(before is after for before, after in filter(None, wb._sheet_history))


def test_collaborative_handles_read_the_engine_sheet_ids(sample_bytes):
    wb = bo.Workbook.open_collaborative(sample_bytes, client_id=41)
    assert wb._inner.sheet_ids == wb.sheet_info()["sheetIds"]
    held = wb[0]
    held["A1"] = "x"
    assert held.index == 0
    assert wb.value(0, "A1") == "x"
