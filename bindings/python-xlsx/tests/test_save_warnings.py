import json
import warnings

from betteroffice_xlsx import SaveResult, SaveWarning, Workbook

WARNING = {
    "code": "unwritable-reference",
    "part": "xl/charts/chart1.xml",
    "element": "chart",
    "sheet": "A\x01B",
    "message": "xl/charts/chart1.xml on sheet A\x01B was kept as the file wrote it",
}


class Engine:
    """The engine surface a save goes through, answering with fixed warnings:
    no public path can yet bring a workbook into a state the save warns
    about, so the mapping is checked against the engine's JSON shape."""

    def __init__(self, found):
        self.found = found

    def save(self):
        return b"saved"

    def save_path(self, path):
        pass

    def _save_warnings_json(self):
        return json.dumps(self.found)


def test_an_ordinary_save_carries_none_and_warns_nothing(sample_bytes):
    workbook = Workbook.open(sample_bytes)
    workbook.apply_ops([{"type": "addSheet", "index": 1, "name": "Notes"}])
    result = workbook.save_with_warnings()
    assert result.warnings == []
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        assert workbook.save() == result.bytes


def test_save_with_warnings_maps_the_engine_json():
    result = Workbook(Engine([WARNING, {**WARNING, "element": None, "sheet": None}])).save_with_warnings()
    assert result == SaveResult(
        bytes=b"saved",
        warnings=[
            SaveWarning(**WARNING),
            SaveWarning(code=WARNING["code"], part=WARNING["part"], message=WARNING["message"]),
        ],
    )


def test_plain_save_returns_the_bytes_and_warns_once_under_the_default_filter(tmp_path):
    workbook = Workbook(Engine([WARNING]))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("default")
        assert workbook.save() == b"saved"
        workbook.save_path(tmp_path / "saved.xlsx")
    assert [(entry.category, str(entry.message)) for entry in caught] == [
        (RuntimeWarning, WARNING["message"]),
        (RuntimeWarning, WARNING["message"]),
    ]
    assert {entry.filename for entry in caught} == {__file__}
