use docx_edit::{EditingDoc, StoryInfoKind, seed_from_docx};

#[test]
fn list_stories_names_kinds_containers_parts_and_inherited_uses() {
    let doc = EditingDoc::new(7101);
    seed_from_docx(
        &doc,
        include_bytes!(
            "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
        ),
    )
    .unwrap();
    let listed = doc.list_stories();
    assert_eq!(listed.version, doc.version());
    let summary: Vec<_> = listed
        .stories
        .iter()
        .map(|info| {
            (
                info.story.as_str(),
                info.kind,
                info.parent.as_deref(),
                info.root.as_str(),
            )
        })
        .collect();
    use StoryInfoKind::*;
    assert_eq!(
        summary,
        [
            ("body", Body, None, "body"),
            ("body:sdt0", ContentControl, Some("body"), "body"),
            ("body:t0:r0c0", TableCell, Some("body"), "body"),
            ("body:t0:r0c1", TableCell, Some("body"), "body"),
            ("body:t0:r0c2", TableCell, Some("body"), "body"),
            ("body:t0:r1c0", TableCell, Some("body"), "body"),
            ("body:t0:r1c1", TableCell, Some("body"), "body"),
            ("body:t0:r2c0", TableCell, Some("body"), "body"),
            ("body:t0:r2c1", TableCell, Some("body"), "body"),
            ("body:t1:r0c0", TableCell, Some("body"), "body"),
            (
                "body:t1:r0c0:t0:r0c0",
                TableCell,
                Some("body:t1:r0c0"),
                "body"
            ),
            (
                "body:t1:r0c0:t0:r0c1",
                TableCell,
                Some("body:t1:r0c0"),
                "body"
            ),
            ("body:t1:r0c1", TableCell, Some("body"), "body"),
            ("en:1", Endnote, None, "en:1"),
            ("fn:1", Footnote, None, "fn:1"),
            ("hf:rIdFooter1", Footer, None, "hf:rIdFooter1"),
            ("hf:rIdHeader1", Header, None, "hf:rIdHeader1"),
            ("hf:rIdHeader2", Header, None, "hf:rIdHeader2"),
            ("hf:rIdHeader3", Header, None, "hf:rIdHeader3"),
        ]
    );
    let header = listed
        .stories
        .iter()
        .find(|info| info.story == "hf:rIdHeader2")
        .unwrap();
    assert!(
        header
            .part
            .as_deref()
            .is_some_and(|part| part.ends_with(".xml"))
    );
    let uses = serde_json::to_value(&header.uses).unwrap();
    assert_eq!(
        uses,
        serde_json::json!([
            { "sectionIndex": 0, "variant": "first" },
            { "sectionIndex": 1, "variant": "first" },
            { "sectionIndex": 2, "variant": "first" },
        ])
    );
}
