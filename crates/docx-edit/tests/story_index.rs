use docx_edit::{
    CellLoc, EditCtx, EditingDoc, Position, StoryInfoKind, TableRange, seed_from_docx,
};

#[path = "support/header_footer_alias_fixture.rs"]
mod fixture;

#[test]
fn header_relationship_ids_are_not_nesting_markers() {
    let mut parts =
        ooxml_opc::unzip_parts(&fixture::package(&[("titleHeader", "header1.xml")], &[])).unwrap();
    let (_, header) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/header1.xml")
        .unwrap();
    *header = String::from_utf8(header.clone())
        .unwrap()
        .replace(
            "</w:hdr>",
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="20000002"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>"#,
        )
        .into_bytes();
    let doc = EditingDoc::new(7102);
    seed_from_docx(&doc, &ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let listed = doc.list_stories();
    let header = listed
        .stories
        .iter()
        .find(|info| info.story == "hf:titleHeader")
        .unwrap();
    assert_eq!(header.kind, StoryInfoKind::Header);
    assert_eq!(header.parent, None);
    assert_eq!(header.root, "hf:titleHeader");
    assert_eq!(header.part.as_deref(), Some("word/header1.xml"));
    let cell = listed
        .stories
        .iter()
        .find(|info| info.story == "hf:titleHeader:t0:r0c0")
        .unwrap();
    assert_eq!(cell.kind, StoryInfoKind::TableCell);
    assert_eq!(cell.parent.as_deref(), Some("hf:titleHeader"));
    assert_eq!(cell.root, "hf:titleHeader");
}

#[test]
fn tables_in_stories_the_session_created_list_their_cells() {
    for story in ["aux", "hf", "fn", "hf:rId7:aux"] {
        let doc = EditingDoc::new(7104);
        doc.create_story(story, "", "Normal", "left").unwrap();
        doc.insert_table(&EditCtx::local("", ""), Position::new(story, 0), 1, 1)
            .unwrap();
        let listed = doc.list_stories();
        let created = listed
            .stories
            .iter()
            .find(|info| info.story == story)
            .unwrap();
        assert_eq!(created.kind, StoryInfoKind::Other, "{story}");
        assert_eq!(created.root, story, "{story}");
        let cell = listed
            .stories
            .iter()
            .find(|info| info.story == format!("{story}:t0:r0c0~7104"))
            .unwrap();
        assert_eq!(cell.kind, StoryInfoKind::TableCell, "{story}");
        assert_eq!(cell.parent.as_deref(), Some(story), "{story}");
        assert_eq!(cell.root, story, "{story}");
    }
}

#[test]
fn a_nested_table_moved_by_a_cell_merge_names_the_cell_that_holds_it() {
    let doc = EditingDoc::new(7103);
    doc.create_story("body", "", "Normal", "left").unwrap();
    let ctx = EditCtx::local("", "");
    doc.insert_table(&ctx, Position::new("body", 0), 1, 2)
        .unwrap();
    doc.insert_table(&ctx, Position::new("body:t0:r0c1~7103", 0), 1, 1)
        .unwrap();
    doc.merge_cells(
        &ctx,
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    let listed = doc.list_stories();
    assert!(
        listed
            .stories
            .iter()
            .all(|info| info.story != "body:t0:r0c1~7103")
    );
    let nested = listed
        .stories
        .iter()
        .find(|info| info.story == "body:t0:r0c1~7103:t0:r0c0~7103")
        .unwrap();
    assert_eq!(nested.kind, StoryInfoKind::TableCell);
    assert_eq!(nested.parent.as_deref(), Some("body:t0:r0c0~7103"));
    assert_eq!(nested.root, "body");
}

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
