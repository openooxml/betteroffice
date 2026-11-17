import json
from pathlib import Path

from betteroffice_xlsx import font_substitute

CASES = Path(__file__).resolve().parents[3] / "crates" / "betteroffice-fonts" / "data" / "resolve-cases.json"


def test_font_substitutes_match_the_typescript_table():
    rows = json.loads(CASES.read_text(encoding="utf-8"))
    cases = [case for case in rows if "family" in case["query"] and "script" not in case["query"]]
    assert len(cases) == sum("script" not in case["query"] for case in rows)
    assert len(cases) > 500
    assert any(case["answer"]["wordFamilySubstitute"] for case in cases)
    for case in cases:
        expected = case["answer"]["wordFamilySubstitute"]
        actual = font_substitute(case["query"]["family"])
        assert actual == (None if expected is None else (expected["remote"], expected["generic"])), case["query"]


def test_font_substitute_is_case_insensitive_and_unknown_is_none():
    assert font_substitute("  CALIBRI ") == ("Carlito", "sans-serif")
    assert font_substitute("Courier New") == ("Cousine", "monospace")
    assert font_substitute("not-a-family") is None
