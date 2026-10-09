#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, SegmentContent, seed_from_docx};
use serde_json::{Value, json};

fn split_document(preceding: u32, lines: usize, height: u32, widow: bool) -> Vec<u8> {
    let paragraph = |id: u32, height: u32, content: &str| {
        fixture::p(
            &format!("{id:08X}"),
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{}" w:lineRule="exact"/><w:widowControl w:val="{}"/></w:pPr>{content}"#,
                height * 15,
                u8::from(widow),
            ),
        )
    };
    let split = (0..lines)
        .map(|index| {
            let line = fixture::r(&format!("Split {index}"));
            if index == 0 {
                line
            } else {
                format!("<w:r><w:br/></w:r>{line}")
            }
        })
        .collect::<String>();
    let mut body = [
        paragraph(1, 100, &fixture::r("First")),
        paragraph(2, 100, &fixture::r("Second")),
        paragraph(3, preceding, &fixture::r("Preceding")),
        paragraph(4, height, &split),
        paragraph(
            5,
            10,
            &format!(
                r#"{}<w:ins w:id="1" w:author="Ann" w:date="2026-09-29T12:00:00Z">{}</w:ins>"#,
                fixture::r("After"),
                fixture::r(" changed"),
            ),
        ),
    ]
    .concat();
    for id in 6..12 {
        body.push_str(&paragraph(id, 100, &fixture::r("Tail")));
    }
    body.push_str("<w:sectPr/>");
    fixture::with_body(&body)
}

fn prime(bytes: &[u8], font: u32, slices: &[(usize, usize, usize)]) -> (EngineSession, Value) {
    let engine = EngineSession::new(75301);
    engine.layout_document_with_regions_retained("{}").unwrap();
    let seed = docx_edit::EditingDoc::new(75300);
    seed_from_docx(&seed, bytes).unwrap();
    engine
        .doc()
        .apply_host_update_v1(&seed.encode_state_as_update_v1())
        .unwrap();
    engine
        .doc()
        .set_note_separator_state(seed.note_separator_state().unwrap());
    let mut request = fixture::region_request(&engine, bytes, font);
    request["regions"]["sections"] = json!([{
        "sectionId": "main",
        "pageSize": {"w": 200, "h": 120},
        "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
    }]);
    request["notes"]["contents"] = json!([]);
    let output = engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    let layout: Value = serde_json::from_str(&output).unwrap();
    let input: Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    let split_id = &input["measured"][3]["block"]["id"];
    let actual: Vec<_> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .flat_map(|(page, value)| {
            value["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .filter(move |fragment| &fragment["blockId"] == split_id)
                .map(move |fragment| {
                    (
                        page,
                        fragment["fromLine"].as_u64().unwrap() as usize,
                        fragment["toLine"].as_u64().unwrap() as usize,
                    )
                })
        })
        .collect();
    assert_eq!(actual, slices);
    (engine, request)
}

fn assert_resumed_matches_fresh(engine: &EngineSession, request: &Value) {
    let before = engine.stats();
    let output = engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls,
        before.incremental_pagination_calls + 1
    );
    assert!(after.pagination_blocks_placed - before.pagination_blocks_placed <= 4);
    assert!(after.rebuilt_pages < after.retained_pages);
    let fresh = EngineSession::new(75302);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let expected = fresh
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    assert_eq!(output.as_bytes(), expected.as_bytes());
    assert_eq!(engine.retained_layout_json().unwrap(), expected);
}

#[test]
fn an_edit_after_a_split_paragraph_resumes_and_matches_fresh() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    for (preceding, lines, height, widow, slices) in [
        (95, 2, 10, false, vec![(3, 0, 1), (3, 1, 2)]),
        (90, 2, 10, false, vec![(2, 0, 1), (3, 1, 2)]),
        (40, 4, 20, true, vec![(2, 0, 2), (3, 2, 4)]),
    ] {
        let bytes = split_document(preceding, lines, height, widow);
        let (engine, request) = prime(&bytes, font, &slices);
        let mut offset = 0;
        let mut at = None;
        for segment in engine.doc().story_segments("body").unwrap() {
            match segment.content {
                SegmentContent::Text(text) => {
                    if text == "After" {
                        at = Some(offset + 2);
                        break;
                    }
                    offset += text.encode_utf16().count() as u32;
                }
                _ => offset += 1,
            }
        }
        let edit = docx_edit::EditingDoc::new(75303);
        edit.apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        edit.insert_text(
            &EditCtx::local("", ""),
            Position::new("body", at.unwrap()),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
        engine
            .doc()
            .apply_host_update_v1(&edit.encode_state_as_update_v1())
            .unwrap();
        assert_resumed_matches_fresh(&engine, &request);
    }
}
