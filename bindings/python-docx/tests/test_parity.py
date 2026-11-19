import hashlib
import io
import json
import zipfile
from pathlib import Path

import pytest

from betteroffice_docx import Document

ROOT = Path(__file__).resolve().parents[3]
PARITY = ROOT / "parity" / "docx"
EXPORT_OPTIONS = {
    "revisionView": "revision_view",
    "stories": "stories",
    "includeFormatting": "include_formatting",
    "maxBlocks": "max_blocks",
    "maxBytes": "max_bytes",
}
CONTROL_OPTIONS = {
    "stories": "stories",
    "maxControls": "max_controls",
    "maxBytes": "max_bytes",
}


def keyword_options(args, names):
    return {names[key]: value for key, value in args.items()}


def run_result(run):
    return {
        "text": run.text,
        "bold": run.bold,
        "italic": run.italic,
        "underline": run.underline,
        "strike": run.strike,
        "fontSize": run.font_size,
        "fontFamily": run.font_family,
        "color": run.color,
        "styleId": run.style_id,
    }


def paragraph_result(paragraph):
    return {
        "id": paragraph.id,
        "style": paragraph.style,
        "alignment": paragraph.alignment,
        "runs": [run_result(run) for run in paragraph.runs],
        "text": paragraph.text,
    }


def cell_result(cell):
    return {
        "paragraphs": [paragraph_result(paragraph) for paragraph in cell.paragraphs],
        "tables": [table_result(table) for table in cell.tables],
        "text": cell.text,
    }


def table_result(table):
    return {
        "rows": [
            {"cells": [cell_result(cell) for cell in row.cells], "text": row.text}
            for row in table.rows
        ],
        "columnWidths": table.column_widths,
        "text": table.text,
    }


def run_highlights_result(paragraph):
    return {"id": paragraph.id, "highlights": [run.highlight for run in paragraph.runs]}


def cell_merges_result(table):
    return {
        "rows": [
            [{"gridSpan": cell.grid_span, "vMerge": cell.v_merge} for cell in row.cells]
            for row in table.rows
        ]
    }


def section_result(section):
    return {
        "id": section.id,
        "pageWidth": section.page_width,
        "pageHeight": section.page_height,
        "orientation": section.orientation,
        "marginTop": section.margin_top,
        "marginBottom": section.margin_bottom,
        "marginLeft": section.margin_left,
        "marginRight": section.margin_right,
        "paragraphs": [paragraph_result(paragraph) for paragraph in section.paragraphs],
        "text": section.text,
    }


def header_footer_result(story):
    return {
        "relId": story.rel_id,
        "kind": story.kind,
        "paragraphs": [paragraph_result(paragraph) for paragraph in story.paragraphs],
        "text": story.text,
    }


def structure_result(document):
    structure = document.structure()
    return {
        "bodyParagraphs": structure.body_paragraphs,
        "bodyTables": structure.body_tables,
        "sections": structure.sections,
        "headers": structure.headers,
        "footers": structure.footers,
        "footnotes": structure.footnotes,
        "endnotes": structure.endnotes,
    }


def receipt_result(edit):
    return {
        "story": edit.story,
        "start": edit.start,
        "end": edit.end,
        "revisionIds": list(edit.revision_ids),
    }


def kept_paragraph_result(paragraph):
    result = {
        "part": paragraph.part,
        "ordinal": paragraph.ordinal,
        "text": paragraph.text,
        "scope": paragraph.scope,
        "dropped": list(paragraph.dropped),
    }
    if paragraph.para_id is not None:
        result["paraId"] = paragraph.para_id
    return result


def dropped_structure_result(structure):
    return {
        "part": structure.part,
        "dropped": list(structure.dropped),
        "paragraphs": [
            {
                "ordinal": paragraph.ordinal,
                "paraId": paragraph.para_id,
                "text": paragraph.text,
                "dropped": list(paragraph.dropped),
            }
            for paragraph in structure.paragraphs
        ],
    }


def report_result(report):
    kept = report.kept_unchanged
    dropped = report.structure_dropped
    return {
        "keptUnchanged": None if kept is None else {
            "structure": kept.structure,
            "paragraphs": [kept_paragraph_result(paragraph) for paragraph in kept.paragraphs],
        },
        "structureDropped": None if dropped is None else [
            dropped_structure_result(structure) for structure in dropped
        ],
        **parts_result(report.bytes),
    }


def png_result(png):
    data = png.bytes
    return {
        "width": int.from_bytes(data[16:20], "big"),
        "height": int.from_bytes(data[20:24], "big"),
        "skippedPictures": png.skipped_images,
    }


def parts_result(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = [
            (info.filename, archive.read(info))
            for info in archive.infolist()
            if not info.is_dir()
        ]
    return {
        "parts": [
            {
                "name": name,
                "sha256": hashlib.sha256(part).hexdigest(),
                "byteLength": len(part),
            }
            for name, part in sorted(parts)
        ]
    }


def cell_paragraph_ids(tables):
    ids = set()
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                ids.update(paragraph.id for paragraph in cell.paragraphs)
    ids.discard(None)
    return ids


def edited_result(document, para_id, except_ids):
    in_cells = cell_paragraph_ids(document.tables())
    excepted = {key.upper() for key in except_ids}
    story = [
        paragraph.text
        for paragraph in document.paragraphs()
        if paragraph.id not in in_cells
        and (paragraph.id or "").upper() not in excepted
        and paragraph.text
    ]
    return {"paragraph": document.paragraph(para_id).text, "story": "\n".join(story)}


def script_sha256():
    return hashlib.sha256((PARITY / "script.json").read_bytes()).hexdigest()


def read_golden(name):
    golden = json.loads((PARITY / name).read_text(encoding="utf-8"))
    recorded = golden.get("scriptSha256") if isinstance(golden, dict) else None
    assert recorded == script_sha256(), f"{name} was written for another script.json"
    return golden["results"]


def includes_target(step, target):
    return "targets" not in step or target in step["targets"]


def run():
    script = json.loads((PARITY / "script.json").read_text(encoding="utf-8"))
    now = script["now"]
    document = Document.open((ROOT / script["fixture"]).read_bytes())
    saved = None

    def reopen(data):
        nonlocal document, saved
        document = Document.open(data)
        saved = None
        return None

    def open_saved(args):
        if saved is None:
            raise AssertionError("openSaved needs a save first")
        return reopen(saved)

    def paragraph(args):
        key = args["index"] if "index" in args else args["id"]
        return paragraph_result(document.paragraph(key))

    def save(args):
        nonlocal saved
        saved = document.save(now=now, update_modified_date=False)
        return parts_result(saved)

    def save_report(args):
        nonlocal saved
        report = document.save_report(now=now, update_modified_date=False)
        saved = report.bytes
        return report_result(report)

    operations = {
        "open": lambda args: reopen((ROOT / args["fixture"]).read_bytes()),
        "openSaved": open_saved,
        "text": lambda args: document.text,
        "structure": lambda args: structure_result(document),
        "paragraphIds": lambda args: document.paragraph_ids,
        "paragraphs": lambda args: [paragraph_result(item) for item in document.paragraphs()],
        "runHighlights": lambda args: [
            run_highlights_result(item) for item in document.paragraphs()
        ],
        "paragraph": paragraph,
        "tables": lambda args: [table_result(table) for table in document.tables()],
        "cellMerges": lambda args: [cell_merges_result(table) for table in document.tables()],
        "sections": lambda args: [section_result(section) for section in document.sections()],
        "headers": lambda args: [header_footer_result(story) for story in document.headers()],
        "footers": lambda args: [header_footer_result(story) for story in document.footers()],
        "warnings": lambda args: document.warnings,
        "listContentControls": lambda args: document.list_content_controls(
            **keyword_options(args, CONTROL_OPTIONS)
        ),
        "findContentControls": lambda args: document.find_content_controls(
            args["query"], **keyword_options(args["options"], CONTROL_OPTIONS)
        ),
        "exportStructured": lambda args: document.export_structured(
            **keyword_options(args, EXPORT_OPTIONS)
        ),
        "exportMarkdown": lambda args: document.export_markdown(
            **keyword_options(args, EXPORT_OPTIONS)
        ),
        "renderPng": lambda args: png_result(document.render_png({"page": args["page"]})),
        "replaceParagraphText": lambda args: receipt_result(
            document.replace_paragraph_text(args["paraId"], args["text"])
        ),
        "replaceParagraphTextAt": lambda args: receipt_result(
            document.replace_paragraph_text_at(args["index"], args["text"])
        ),
        "readEdited": lambda args: edited_result(
            document, args["paraId"], args.get("exceptParaIds", [])
        ),
        "save": save,
        "saveReport": save_report,
    }
    results = []
    for index, step in enumerate(script["steps"]):
        if not includes_target(step, "python"):
            continue
        op = step["op"]
        if op not in operations:
            raise AssertionError(f"Unknown parity op {op} at step {index}")
        try:
            results.append({"op": op, "result": operations[op](step["args"])})
        except Exception as error:
            raise AssertionError(f"Parity step {index} ({op}) failed") from error
    engine = []
    for index, step in enumerate(script["engineSteps"]):
        op = step["op"]
        try:
            engine.append({"op": op, "result": operations[op](step["args"])})
        except Exception as error:
            raise AssertionError(f"Engine step {index} ({op}) failed") from error
    return results, engine


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
        ({"result": [1, 2]}, {"result": [2, 1]}, '$["result"][0]'),
        ({"result": {"value": 1}}, {"result": {"other": 1}}, '$["result"]'),
    ],
)
def test_json_comparison_reports_mismatch_path(expected, actual, path):
    with pytest.raises(AssertionError) as error:
        assert_json_equal(expected, actual)
    assert f"JSON mismatch at {path}\n" in str(error.value)


def test_step_targets_default_to_all_and_filter_explicit_lists():
    assert includes_target({"op": "text"}, "python")
    assert includes_target({"op": "text", "targets": ["python"]}, "python")
    assert not includes_target({"op": "text", "targets": ["ts", "rust"]}, "python")


def test_shared_and_engine_only_scripted_steps_match_their_goldens():
    actual, engine = run()
    script = json.loads((PARITY / "script.json").read_text(encoding="utf-8"))
    expected = read_golden("expected.json")
    assert len(expected) == len(script["steps"]), "Golden step count differs"
    expected = [
        result for step, result in zip(script["steps"], expected)
        if includes_target(step, "python")
    ]
    assert_json_equal(expected, actual)
    assert_json_equal(read_golden("engine-expected.json"), engine)
