import hashlib
import json
import os
import re
from pathlib import Path

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
        "undo": lambda args: workbook.undo(now_serial=now_serial).applied,
        "redo": lambda args: workbook.redo(now_serial=now_serial).applied,
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


def test_scripted_api_parity_matches_shared_golden():
    actual = run()
    expected_path = PARITY / "expected.json"
    if os.environ.get("BETTEROFFICE_PARITY_WRITE") == "1":
        expected_path.write_text(
            json.dumps(actual, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )
    expected = json.loads(expected_path.read_text(encoding="utf-8"))
    assert len(actual) == len(expected), "Parity step count differs"
    for index, (actual_step, expected_step) in enumerate(zip(actual, expected)):
        assert actual_step == expected_step, (
            f"Parity step {index} ({actual_step['op']}) differs\n"
            f"- expected: {json.dumps(expected_step, indent=2, ensure_ascii=False)}\n"
            f"+ actual: {json.dumps(actual_step, indent=2, ensure_ascii=False)}"
        )
