use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};

#[test]
fn unicode_simple_folding_preserves_matches_and_utf16_offsets() {
    let text = include_str!("fixtures/unicode-search.txt").trim_end();
    let session = DeckSession::open(
        include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx"),
        84091,
    )
    .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let added = session
        .add_text_box(
            &EditCtx::local("test"),
            &slide,
            &ShapeDraft {
                name: "Unicode search".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                text: text.to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let before = session.encode_state_as_update_v1();
    let cases: &[(&str, &[&str])] = &[
        ("i", &["I", "i"]),
        ("I", &["I", "i"]),
        ("ı", &["ı"]),
        ("İ", &["İ"]),
        ("ΐ", &["ΐ", "ΐ"]),
        ("ΐ", &["ΐ", "ΐ"]),
        ("ΰ", &["ΰ", "ΰ"]),
        ("ΰ", &["ΰ", "ΰ"]),
        ("ﬅ", &["ﬅ", "ﬆ"]),
        ("ﬆ", &["ﬅ", "ﬆ"]),
    ];
    for &(query, expected) in cases {
        for case_sensitive in [false, true] {
            let matches = session
                .search_text(query, case_sensitive, None)
                .unwrap()
                .into_iter()
                .filter(|m| m.shape_id == added.shape_id)
                .collect::<Vec<_>>();
            let sensitive = [query];
            let expected = if case_sensitive { &sensitive } else { expected };
            assert_eq!(
                matches.iter().map(|m| m.text.as_str()).collect::<Vec<_>>(),
                expected,
                "query {query:?}, case_sensitive={case_sensitive}"
            );
            for found in matches {
                let start = text[..text.find(&found.text).unwrap()]
                    .encode_utf16()
                    .count() as u32;
                assert_eq!(found.start, start);
                assert_eq!(found.end, start + found.text.encode_utf16().count() as u32);
            }
        }
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}
