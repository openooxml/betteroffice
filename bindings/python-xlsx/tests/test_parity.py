import hashlib
import json
import os
import re
from pathlib import Path

import pytest

from betteroffice_xlsx import Workbook

ROOT = Path(__file__).resolve().parents[3]
PARITY = ROOT / "parity" / "xlsx"
EXPORT_OPTIONS = {
    "scope": "scope",
    "includeHiddenSheets": "include_hidden_sheets",
    "includeHiddenRows": "include_hidden_rows",
    "includeHiddenColumns": "include_hidden_columns",
    "includeDefinedNames": "include_defined_names",
    "includeHiddenNames": "include_hidden_names",
    "maxCells": "max_cells",
    "maxBytes": "max_bytes",
}


def map_strings(value, transform):
    if isinstance(value, str):
        return transform(value)
    if isinstance(value, list):
        return [map_strings(item, transform) for item in value]
    if isinstance(value, dict):
        return {key: map_strings(value[key], transform) for key in sorted(value)}
    return value


def history_state(workbook):
    state = workbook.history()
    return {
        "canUndo": state.can_undo,
        "canRedo": state.can_redo,
        "undoDepth": state.undo_depth,
        "redoDepth": state.redo_depth,
    }


def mutation_result(result):
    return {
        "applied": result.applied,
        "changed": sorted(result.changed),
        "limitedCells": sorted(result.limited),
    }


def export_options(args):
    return {EXPORT_OPTIONS[key]: value for key, value in args.items()}


def saved_bytes(workbook):
    data = workbook.save()
    return {"sha256": hashlib.sha256(data).hexdigest(), "byteLength": len(data)}


def run():
    script = json.loads((PARITY / "script.json").read_text(encoding="utf-8"))
    now_serial = script["calculation"]["nowSerial"]
    workbook = Workbook.open_recalculated(
        (ROOT / script["fixture"]).read_bytes(), now_serial=now_serial
    )
    operations = {
        "sheetNames": lambda args: workbook.sheet_names,
        "version": lambda args: workbook.version(),
        "readCells": lambda args: workbook.read_cells(args),
        "findText": lambda args: workbook.find_text(args),
        "validateEdits": lambda args: workbook.validate_edits(args),
        "applyEdits": lambda args: workbook.apply_edits(args),
        "historyState": lambda args: history_state(workbook),
        "undo": lambda args: mutation_result(workbook.undo(now_serial=now_serial)),
        "redo": lambda args: mutation_result(workbook.redo(now_serial=now_serial)),
        "exportStructured": lambda args: workbook.export_structured(**export_options(args)),
        "exportMarkdown": lambda args: workbook.export_markdown(
            **export_options(args["options"]), markdown_options=args["markdownOptions"]
        ),
        "save": lambda args: saved_bytes(workbook),
    }
    versions = []

    def normalize(value):
        if not re.fullmatch(r"[0-9a-f]{32}-[0-9]+", value):
            return value
        if value not in versions:
            versions.append(value)
        return f"<version:{versions.index(value) + 1}>"

    def resolve(value):
        match = re.fullmatch(r"<version:([1-9][0-9]*)>", value)
        if match is None:
            return value
        index = int(match[1]) - 1
        if index >= len(versions):
            raise ValueError(f"Unknown version reference {value}")
        return versions[index]

    results = []
    for index, step in enumerate(script["steps"]):
        op = step["op"]
        if op not in operations:
            raise AssertionError(f"Unknown parity op {op} at step {index}")
        try:
            result = operations[op](map_strings(step["args"], resolve))
            results.append({"op": op, "result": map_strings(result, normalize)})
        except Exception as error:
            raise AssertionError(f"Parity step {index} ({op}) failed") from error
    return results


def assert_json_equal(expected, actual, path="$"):
    if type(expected) is bool or type(actual) is bool:
        if type(expected) is bool and type(actual) is bool and expected == actual:
            return
    elif type(expected) in (int, float) and type(actual) in (int, float):
        if expected == actual:
            return
    elif isinstance(expected, dict) and isinstance(actual, dict):
        if expected.keys() == actual.keys():
            for key in sorted(expected):
                assert_json_equal(expected[key], actual[key], f"{path}[{json.dumps(key)}]")
            return
    elif isinstance(expected, list) and isinstance(actual, list):
        if len(expected) == len(actual):
            for index, (expected_item, actual_item) in enumerate(zip(expected, actual)):
                assert_json_equal(expected_item, actual_item, f"{path}[{index}]")
            return
    elif type(expected) is type(actual) and expected == actual:
        return
    raise AssertionError(
        f"JSON mismatch at {path}\n"
        f"- expected: {json.dumps(expected, indent=2, ensure_ascii=False)}\n"
        f"+ actual: {json.dumps(actual, indent=2, ensure_ascii=False)}"
    )


@pytest.mark.parametrize(
    "expected,actual,path",
    [
        (True, 1, "$"),
        (1, True, "$"),
        (False, 0.0, "$"),
        (0.0, False, "$"),
        ({"result": [True]}, {"result": [1]}, '$["result"][0]'),
        ({"result": [1, 2]}, {"result": [2, 1]}, '$["result"][0]'),
        ({"result": [1]}, {"result": []}, '$["result"]'),
        ({"result": {"value": 1}}, {"result": {"other": 1}}, '$["result"]'),
    ],
)
def test_json_comparison_reports_mismatch_path(expected, actual, path):
    with pytest.raises(AssertionError) as error:
        assert_json_equal(expected, actual)
    assert f"JSON mismatch at {path}\n" in str(error.value)


def test_json_comparison_accepts_equal_numbers_and_booleans():
    assert_json_equal(
        {"numbers": [1, 2.5], "booleans": [True, False], "empty": None},
        {"empty": None, "booleans": [True, False], "numbers": [1.0, 2.5]},
    )


def test_scripted_api_parity_matches_shared_golden():
    actual = run()
    expected_path = PARITY / "expected.json"
    if os.environ.get("BETTEROFFICE_PARITY_WRITE") == "1":
        expected_path.write_text(
            json.dumps(actual, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )
    expected = json.loads(expected_path.read_text(encoding="utf-8"))
    assert_json_equal(expected, actual)
