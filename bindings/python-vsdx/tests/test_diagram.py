import pytest

from betteroffice_vsdx import Diagram, ParseError, VsdxError

def test_opens_foundation_fixture(foundation_path):
    diagram = Diagram.open_path(foundation_path)

    assert len(diagram) == 1
    page = diagram.pages[0]
    assert (page.id, page.name, page.source_part_path) == (1, "Page-1", "visio/pages/page1.xml")
    shape = page.shapes[0]
    assert (shape.id, shape.name, shape.text) == (1, "Process", " AB\n\t C ")
    assert [(cell.name, cell.formula, cell.value, cell.unit) for cell in shape.cells] == [
        ("FOnly", "Width*2", None, None),
        ("VOnly", None, "5", None),
        ("Both", "Height*2", "2", None),
        ("LineWeight", None, "0.01", None),
    ]


def test_parse_errors_are_vsdx_errors():
    with pytest.raises(ParseError) as error:
        Diagram.open(b"not a VSDX")

    assert isinstance(error.value, VsdxError)
