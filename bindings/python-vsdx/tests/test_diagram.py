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


def test_page_order_and_names_follow_the_catalog(foundation_path):
    import io
    import zipfile

    with zipfile.ZipFile(foundation_path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["visio/pages/page2.xml"] = parts["visio/pages/page1.xml"]
    parts["visio/pages/pages.xml"] = b"""<v:Pages xmlns:v="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><!-- <Page ID='2' Name='Wrong'/> --><v:Page ID='2' Name='First &amp; correct'><v:Rel r:id='second'/></v:Page><v:Page ID='1' Name='Last'><v:Rel r:id='first'/></v:Page></v:Pages>"""
    parts["visio/pages/_rels/pages.xml.rels"] = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id='first' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/><Relationship Id='second' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page2.xml'/></Relationships>"""
    data = io.BytesIO()
    with zipfile.ZipFile(data, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    diagram = Diagram.open(data.getvalue())
    assert [(page.id, page.name) for page in diagram.pages] == [(2, "First & correct"), (1, "Last")]


@pytest.mark.parametrize("buffer_type", [bytes, bytearray, memoryview])
def test_accepts_documented_buffers(foundation_path, buffer_type):
    data = buffer_type(foundation_path.read_bytes())
    diagram = Diagram.open(data)
    assert len(diagram) == 1
    assert diagram.pages[0].name == "Page-1"


def test_snapshots_mutable_buffer_before_parsing(foundation_path):
    data = bytearray(foundation_path.read_bytes())
    diagram = Diagram.open(memoryview(data))
    data[:] = b"x" * len(data)
    assert diagram.pages[0].shapes[0].name == "Process"


@pytest.mark.parametrize("data", [None, 10, "not bytes", [1, 2, 3]])
def test_rejects_non_buffers(data):
    with pytest.raises(TypeError):
        Diagram.open(data)


def test_missing_path_raises_file_not_found(tmp_path):
    path = tmp_path / "missing.vsdx"
    with pytest.raises(FileNotFoundError) as error:
        Diagram.open_path(path)
    assert error.value.filename == str(path)
