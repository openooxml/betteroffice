use docx_layout::display_list::build_display_list_json;
use serde_json::{Value, json};

/// A paragraph fragment showing lines `[from_line, to_line)` at column `x`.
fn fragment(from_line: usize, to_line: usize, x: u32) -> Value {
    json!({
        "kind": "paragraph",
        "blockId": "breaks",
        "x": x,
        "y": 20,
        "width": 120,
        "height": 20 * (to_line - from_line),
        "fromLine": from_line,
        "toLine": to_line,
        "pmStart": 0,
        "pmEnd": 14
    })
}

fn page(fragments: Vec<Value>) -> Value {
    json!({
        "number": 1,
        "size": { "w": 300, "h": 400 },
        "margins": { "top": 20, "right": 20, "bottom": 20, "left": 20 },
        "fragments": fragments
    })
}

/// "Ab", two line breaks, "Cd", a tab and "Ef gh" soft-wrapped after "Ef ",
/// laid out as `pages`.
fn display(pages: Vec<Value>) -> Value {
    let line = |head: (usize, usize), tail: (usize, usize)| {
        json!({
            "headRun": head.0, "headChar": head.1, "tailRun": tail.0, "tailChar": tail.1,
            "width": 60, "ascent": 14, "descent": 4, "lineHeight": 20
        })
    };
    let input = json!({
        "contractVersion": 1,
        "measured": [{
            "block": {
                "kind": "paragraph",
                "id": "breaks",
                "runs": [
                    { "kind": "text", "text": "Ab", "pmStart": 1, "pmEnd": 3 },
                    { "kind": "lineBreak", "pmStart": 3, "pmEnd": 4 },
                    { "kind": "lineBreak", "pmStart": 4, "pmEnd": 5 },
                    { "kind": "text", "text": "Cd", "pmStart": 5, "pmEnd": 7 },
                    { "kind": "tab", "pmStart": 7, "pmEnd": 8, "width": 20 },
                    { "kind": "text", "text": "Ef gh", "pmStart": 8, "pmEnd": 13 }
                ],
                "pmStart": 0,
                "pmEnd": 14
            },
            "measure": {
                "kind": "paragraph",
                "totalHeight": 80,
                "lines": [
                    line((0, 0), (1, 0)),
                    line((2, 0), (2, 0)),
                    line((3, 0), (5, 3)),
                    line((5, 3), (5, 5))
                ]
            }
        }],
        "options": {},
        "layout": { "pages": pages }
    });
    serde_json::from_str(&build_display_list_json(&input.to_string()).expect("display list builds"))
        .expect("valid display JSON")
}

/// `(docStart, "before" | "after", separator)` for every marked primitive of `page`.
fn marked(output: &Value, page: usize) -> Vec<(i64, &'static str, String)> {
    let mut out = Vec::new();
    for primitive in output["pages"][page]["primitives"]
        .as_array()
        .expect("primitive array")
    {
        let Some(start) = primitive["docStart"].as_i64() else {
            continue;
        };
        for (key, side) in [("separatorBefore", "before"), ("separatorAfter", "after")] {
            if let Some(separator) = primitive.get(key).and_then(Value::as_str) {
                out.push((start, side, separator.to_owned()));
            }
        }
    }
    out
}

#[test]
fn explicit_breaks_and_tabs_mark_the_next_text_but_soft_wraps_do_not() {
    let output = display(vec![page(vec![fragment(0, 4, 20)])]);
    let texts: Vec<(i64, &str)> = output["pages"][0]["primitives"]
        .as_array()
        .expect("primitive array")
        .iter()
        .filter(|primitive| primitive["kind"] == "text" || primitive["kind"] == "glyphRun")
        .filter_map(|primitive| {
            Some((primitive["docStart"].as_i64()?, primitive["text"].as_str()?))
        })
        .collect();
    assert!(
        texts
            .iter()
            .any(|&(start, text)| start == 11 && text.starts_with('g')),
        "{texts:?}"
    );
    assert_eq!(
        marked(&output, 0),
        vec![
            (5, "before", "\n\n".to_owned()),
            (8, "before", "\t".to_owned())
        ]
    );
}

#[test]
fn breaks_that_end_a_fragment_read_after_its_last_text_on_that_page_only() {
    let output = display(vec![
        page(vec![fragment(0, 2, 20)]),
        page(vec![fragment(2, 4, 20)]),
    ]);
    assert_eq!(marked(&output, 0), vec![(4, "after", "\n\n".to_owned())]);
    assert_eq!(marked(&output, 1), vec![(8, "before", "\t".to_owned())]);
}

#[test]
fn a_paragraph_continued_in_the_next_column_keeps_its_break_between_the_columns() {
    let output = display(vec![page(vec![fragment(0, 2, 20), fragment(2, 4, 160)])]);
    assert_eq!(
        marked(&output, 0),
        vec![
            (4, "after", "\n\n".to_owned()),
            (8, "before", "\t".to_owned())
        ]
    );
}
