import hashlib
import json
import os
import re
from pathlib import Path

import pytest

from betteroffice_pptx import Presentation

ROOT = Path(__file__).resolve().parents[3]
PARITY = ROOT / "parity" / "pptx"
EXPORT_OPTIONS = {
    "includeHiddenSlides": "include_hidden_slides",
    "includeHiddenShapes": "include_hidden_shapes",
    "includeNotes": "include_notes",
    "includeComments": "include_comments",
    "includeFormatting": "include_formatting",
    "maxBlocks": "max_blocks",
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


def slide_result(result):
    return {
        "slideId": result.slide_id,
        "fromIndex": result.from_index,
        "toIndex": result.to_index,
    }


def comment_result(result):
    return {
        "commentId": result.comment_id,
        "slideId": result.slide_id,
        "parentId": result.parent_id,
        "resolved": result.resolved,
    }


def story_result(story):
    return {
        "id": story.id,
        "length": story.length,
        "text": story.text,
        "paragraphs": [
            {
                "id": paragraph.id,
                "alignment": paragraph.alignment,
                "level": paragraph.level,
                "text": paragraph.text,
                "runs": [
                    {
                        "text": run.text,
                        "style": {
                            "bold": run.bold,
                            "italic": run.italic,
                            "fontSizePt": run.font_size,
                            "color": run.color,
                            "fontFamily": run.font_family,
                            "underline": run.underline,
                            "spacingPt": run.spacing_pt,
                            "baselinePct": run.baseline_pct,
                            "kernPt": run.kern_pt,
                            "caps": run.caps,
                        },
                    }
                    for run in paragraph.runs
                ],
            }
            for paragraph in story.paragraphs
        ],
    }


def history_result(presentation, direction):
    result = presentation.history_step(direction)
    return {**result, "changedSlideIds": sorted(result["changedSlideIds"])}


def add_comment(presentation, args):
    comment = args["comment"]
    return comment_result(
        presentation.add_comment(
            args["slideId"],
            comment["text"],
            author=comment["author"],
            initials=comment["initials"],
            created=comment["created"],
            x=comment["xEmu"],
            y=comment["yEmu"],
        )
    )


def export_options(args):
    return {EXPORT_OPTIONS[key]: value for key, value in args.items()}


def saved_bytes(presentation):
    data = presentation.save()
    return {"sha256": hashlib.sha256(data).hexdigest(), "byteLength": len(data)}


def run():
    script = json.loads((PARITY / "script.json").read_text(encoding="utf-8"))
    presentation = Presentation.open_collaborative(
        (ROOT / script["fixture"]).read_bytes(), client_id=script["clientId"]
    )
    presentation.author = "wasm"
    operations = {
        "version": lambda args: presentation.version(),
        "readContent": lambda args: presentation.read_content(args),
        "findText": lambda args: presentation.find_text(args),
        "validateEdits": lambda args: presentation.validate_edits(args),
        "applyEdits": lambda args: presentation.apply_edits(args),
        "exportStructured": lambda args: presentation.export_structured(**export_options(args)),
        "exportMarkdown": lambda args: presentation.export_markdown(**export_options(args)),
        "historyStep": lambda args: history_result(presentation, args["direction"]),
        "duplicateSlide": lambda args: slide_result(
            presentation.duplicate_slide(args["slideId"], args["toIndex"])
        ),
        "bringShapeToFront": lambda args: presentation.bring_shape_to_front(
            args["slideId"], args["shapeId"]
        ),
        "addComment": lambda args: add_comment(presentation, args),
        "setCommentPosition": lambda args: comment_result(
            presentation.set_comment_position(
                args["commentId"], args["position"]["xEmu"], args["position"]["yEmu"]
            )
        ),
        "setStoryParagraphs": lambda args: story_result(
            presentation.set_story_paragraphs(args["storyId"], args["paragraphs"])
        ),
        "readCommandContent": lambda args: presentation.read_command_content(args),
        "save": lambda args: saved_bytes(presentation),
    }
    versions = []

    def normalize(value):
        if not re.fullmatch(r"[0-9a-f]{16}-[0-9]+", value):
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
