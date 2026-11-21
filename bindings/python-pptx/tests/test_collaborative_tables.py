import pytest

import betteroffice_pptx as bo
from test_parity_ops import _table_deck


def _replicas(sample_bytes):
    deck, shape_id = _table_deck(sample_bytes)
    first = bo.Presentation.open_collaborative(deck.save(), client_id=83)
    joined = bo.Presentation.open_from_update(first.state_as_update(), client_id=84)
    return first, joined, shape_id


def test_collaborative_constructors_refuse_table_structure_edits(sample_bytes):
    first, joined, shape_id = _replicas(sample_bytes)
    for replica in (first, joined):
        assert replica.is_collaborative
        with pytest.raises(bo.PptxError, match="collaborative mode"):
            replica.insert_table_row(0, shape_id, 0, "after", ["New row"])
        with pytest.raises(bo.PptxError, match="collaborative mode"):
            replica.insert_table_column(0, shape_id, 0, "after", ["New column"])


def test_standalone_open_allows_table_structure_edits(sample_bytes):
    deck, shape_id = _table_deck(sample_bytes)
    assert not deck.is_collaborative
    row = deck.insert_table_row(0, shape_id, 0, "after", ["New row"])
    assert (row["rowCount"], row["columnCount"]) == (3, 2)
    column = deck.insert_table_column(0, shape_id, 0, "after", ["New column"])
    assert (column["rowCount"], column["columnCount"]) == (3, 3)


def test_concurrent_table_cell_text_edits_merge(sample_bytes):
    first, joined, shape_id = _replicas(sample_bytes)
    table = next(shape for shape in first[0].shapes if shape.id == shape_id)
    first_cell, second_cell = table.stories[:2]
    first.insert_text(first_cell.id, 0, "First ")
    joined.insert_text(second_cell.id, 0, "Second ")

    first_update = first.state_as_update()
    joined_update = joined.state_as_update()
    first.apply_update(joined_update)
    joined.apply_update(first_update)

    for replica in (first, joined):
        assert replica.story(first_cell.id).text == "First Cell"
        assert replica.story(second_cell.id).text == "Second Cell"
        shapes = replica.read_content()["slides"][0]["shapes"]
        graphic = next(shape["graphic"] for shape in shapes if shape["id"] == shape_id)
        assert (len(graphic["rows"]), len(graphic["grid"])) == (2, 2)
