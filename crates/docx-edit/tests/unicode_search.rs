use docx_edit::EditingDoc;

#[test]
fn unicode_simple_folding_preserves_matches_and_utf16_offsets() {
    let text = include_str!("fixtures/unicode-search.txt").trim_end();
    let doc = EditingDoc::new(83091);
    doc.create_story("body", text, "Normal", "left").unwrap();
    let before = doc.encode_state_as_update_v1();
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
            let matches = doc.search_text(query, case_sensitive, None).unwrap();
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
    assert_eq!(doc.encode_state_as_update_v1(), before);
}
