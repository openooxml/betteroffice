use pptx_edit::{DeckSession, EditCtx, EditError, TableInsertPosition, TextStyle};
use pptx_parse::GraphicFrameData;

const BASE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn source(merged: bool) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(BASE).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let extent = xml.find("<a:ext ").unwrap();
    let end = extent + xml[extent..].find("/>").unwrap() + 2;
    xml.replace_range(extent..end, r#"<a:ext cx="6000000" cy="1200000"/>"#);
    let start = xml.find("<a:tbl>").unwrap();
    let end = start + xml[start..].find("</a:tbl>").unwrap() + 8;
    let mut table = String::from(
        "<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w=\"1000000\"/><a:gridCol w=\"2000000\"/><a:gridCol w=\"3000000\"/></a:tblGrid>",
    );
    for row in 0..3 {
        table.push_str("<a:tr h=\"400000\">");
        for col in 0..3 {
            let attr = if merged && row == 0 && col == 0 {
                " gridSpan=\"2\""
            } else {
                ""
            };
            table.push_str(&format!(r#"<a:tc{attr}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="3000" algn="ctr"/><a:r><a:rPr b="1" lang="en-US" strike="sngStrike"/><a:t>R{row}C{col}</a:t></a:r><a:endParaRPr sz="1800"/></a:p></a:txBody><a:tcPr marL="91440"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:tcPr></a:tc>"#));
        }
        table.push_str("</a:tr>");
    }
    table.push_str("</a:tbl>");
    xml.replace_range(start..end, &table);
    *bytes = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn target(deck: &DeckSession) -> (String, String) {
    let snapshot = deck.snapshot().unwrap();
    (
        snapshot.slides[0].id.clone(),
        snapshot.slides[0].shapes[0].id.clone(),
    )
}

fn table(deck: &DeckSession) -> pptx_parse::Table {
    let snapshot = deck.snapshot().unwrap();
    let Some(GraphicFrameData::Table(table)) = &snapshot.slides[0].shapes[0].graphic else {
        panic!("table")
    };
    table.clone()
}

#[test]
fn delete_column_updates_grid_frame_and_rejects_stale_stories() {
    let deck = DeckSession::open(&source(false), 51).unwrap();
    let (slide, shape) = target(&deck);
    let before = deck.snapshot().unwrap();
    let old = before.slides[0].shapes[0].text_stories[2].id.clone();
    let receipt = deck
        .delete_table_column(&EditCtx::local("test"), &slide, &shape, 1)
        .unwrap();
    assert_eq!((receipt.row_count, receipt.column_count), (3, 2));
    assert_eq!(table(&deck).grid, [1_000_000, 3_000_000]);
    assert_eq!(
        deck.snapshot().unwrap().slides[0].shapes[0].width,
        4_000_000
    );
    assert!(matches!(
        deck.insert_text(
            &EditCtx::local("test"),
            &old,
            0,
            "wrong",
            &TextStyle::default()
        ),
        Err(EditError::StoryNotFound(_))
    ));
    let reopened = DeckSession::open(&deck.save().unwrap(), 52).unwrap();
    assert_eq!(table(&reopened).grid, table(&deck).grid);
    assert_eq!(
        reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[1].plain_text(),
        "R0C2"
    );
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), before);
    assert!(deck.redo());
    assert_eq!(table(&deck).grid.len(), 2);
}

#[test]
fn insert_column_splits_anchor_width_and_copies_cell_formatting() {
    let deck = DeckSession::open(&source(false), 53).unwrap();
    let (slide, shape) = target(&deck);
    let receipt = deck
        .insert_table_column(
            &EditCtx::local("test"),
            &slide,
            &shape,
            1,
            TableInsertPosition::After,
            &["New".into()],
        )
        .unwrap();
    assert_eq!((receipt.row_count, receipt.column_count), (3, 4));
    assert_eq!(receipt.new_cell_refs.len(), 3);
    assert_eq!(
        table(&deck).grid,
        [1_000_000, 1_000_000, 1_000_000, 3_000_000]
    );
    assert_eq!(
        deck.snapshot().unwrap().slides[0].shapes[0].width,
        6_000_000
    );
    let saved = deck.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(p, _)| p == "ppt/slides/slide1.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert_eq!(xml.matches("strike=\"sngStrike\"").count(), 12);
    assert_eq!(xml.matches("val=\"ABCDEF\"").count(), 12);
    let reopened = DeckSession::open(&saved, 54).unwrap();
    assert_eq!(
        reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[2].plain_text(),
        "New"
    );
}

#[test]
fn row_edits_adjust_height_undo_and_replicate() {
    let bytes = source(false);
    let deck = DeckSession::open(&bytes, 55).unwrap();
    let peer = DeckSession::open(&bytes, 56).unwrap();
    let (slide, shape) = target(&deck);
    let before = deck.snapshot().unwrap();
    let added = deck
        .insert_table_row(
            &EditCtx::local("test"),
            &slide,
            &shape,
            2,
            TableInsertPosition::After,
            &["Added".into()],
        )
        .unwrap();
    assert_eq!((added.row_count, added.column_count), (4, 3));
    assert_eq!(added.new_cell_refs.len(), 3);
    assert_eq!(
        deck.snapshot().unwrap().slides[0].shapes[0].height,
        1_600_000
    );
    peer.apply_update_v1(&deck.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(peer.snapshot().unwrap(), deck.snapshot().unwrap());
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), before);
    assert!(deck.redo());
    deck.delete_table_row(&EditCtx::local("test"), &slide, &shape, 3)
        .unwrap();
    assert_eq!(
        deck.snapshot().unwrap().slides[0].shapes[0].height,
        1_200_000
    );
    assert_eq!(table(&deck).rows.len(), 3);
}

#[test]
fn structural_refusals_leave_state_unchanged() {
    let ctx = EditCtx::local("test");
    for op in 0..4 {
        let deck = DeckSession::open(&source(true), 57).unwrap();
        let (slide, shape) = target(&deck);
        let before = deck.encode_state_as_update_v1();
        let result = match op {
            0 => deck.delete_table_column(&ctx, &slide, &shape, 0),
            1 => deck.delete_table_row(&ctx, &slide, &shape, 0),
            2 => deck.insert_table_column(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[]),
            _ => deck.insert_table_row(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[]),
        };
        assert!(matches!(result, Err(EditError::InvalidTable(_))));
        assert_eq!(deck.encode_state_as_update_v1(), before);
    }
    let deck = DeckSession::open(&source(false), 58).unwrap();
    let (slide, shape) = target(&deck);
    for _ in 0..2 {
        deck.delete_table_row(&ctx, &slide, &shape, 0).unwrap();
        deck.delete_table_column(&ctx, &slide, &shape, 0).unwrap();
    }
    let before = deck.encode_state_as_update_v1();
    assert!(deck.delete_table_row(&ctx, &slide, &shape, 0).is_err());
    assert!(deck.delete_table_column(&ctx, &slide, &shape, 0).is_err());
    assert!(
        deck.insert_table_row(&ctx, &slide, &shape, 4, TableInsertPosition::Before, &[])
            .unwrap_err()
            .to_string()
            .contains("row")
    );
    assert_eq!(deck.encode_state_as_update_v1(), before);
}

#[test]
fn every_structural_operation_has_a_lossless_undo_and_redo() {
    for operation in 0..6 {
        let deck = DeckSession::open(&source(false), 61).unwrap();
        let (slide, shape) = target(&deck);
        let before = deck.snapshot().unwrap();
        let ctx = EditCtx::local("test");
        match operation {
            0 => deck.delete_table_row(&ctx, &slide, &shape, 1),
            1 => deck.delete_table_column(&ctx, &slide, &shape, 1),
            2 => deck.insert_table_row(&ctx, &slide, &shape, 0, TableInsertPosition::Before, &[]),
            3 => deck.insert_table_row(&ctx, &slide, &shape, 2, TableInsertPosition::After, &[]),
            4 => {
                deck.insert_table_column(&ctx, &slide, &shape, 0, TableInsertPosition::Before, &[])
            }
            _ => deck.insert_table_column(&ctx, &slide, &shape, 2, TableInsertPosition::After, &[]),
        }
        .unwrap();
        let after = deck.snapshot().unwrap();
        let saved = deck.save().unwrap();
        assert!(deck.undo());
        assert_eq!(deck.snapshot().unwrap(), before);
        assert!(deck.redo());
        assert_eq!(deck.snapshot().unwrap(), after);
        assert_eq!(deck.save().unwrap(), saved);
    }
}

#[test]
fn structural_replay_and_inverse_history_match_on_peers() {
    use serde_json::{Value, json};
    let bytes = source(false);
    let worker = DeckSession::open_replay_baseline(&bytes, Some(71), None).unwrap();
    let identity = worker.peer_identity(Vec::new()).unwrap();
    let peer = DeckSession::open_peer_deck(&bytes, &identity, None).unwrap();
    peer.register_peer_fonts(Vec::new()).unwrap();
    peer.adopt_peer_identity().unwrap();
    let (slide, shape) = target(&worker);
    let ops = [
        json!({"method":"insertTableRow","args":[slide,shape,2,"after",["Row"]]}),
        json!({"method":"insertTableColumn","args":[slide,shape,1,"before",["Column"]]}),
        json!({"method":"deleteTableRow","args":[slide,shape,0]}),
        json!({"method":"deleteTableColumn","args":[slide,shape,0]}),
        json!({"method":"undo","args":[]}),
        json!({"method":"redo","args":[]}),
    ];
    for (index, op) in ops.into_iter().enumerate() {
        let input = json!({"sequence":index+1,"baseVersion":worker.version(),"op":op}).to_string();
        let expected: Value = serde_json::from_str(&worker.replay_json(&input).unwrap()).unwrap();
        let actual: Value = serde_json::from_str(&peer.replay_json(&input).unwrap()).unwrap();
        assert_eq!(actual, expected);
        assert_eq!(peer.snapshot().unwrap(), worker.snapshot().unwrap());
        assert_eq!(peer.save().unwrap(), worker.save().unwrap());
    }
}

#[test]
fn structural_edits_preserve_concurrent_text_in_surviving_cells() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 81).unwrap();
    let second = DeckSession::open(&bytes, 82).unwrap();
    let (slide, shape) = target(&first);
    let story = first.snapshot().unwrap().slides[0].shapes[0].text_stories[8]
        .id
        .clone();
    first
        .delete_table_column(&EditCtx::local("test"), &slide, &shape, 0)
        .unwrap();
    second
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "Peer ",
            &TextStyle::default(),
        )
        .unwrap();
    let a = first.encode_state_as_update_v1();
    let b = second.encode_state_as_update_v1();
    first.apply_update_v1(&b).unwrap();
    second.apply_update_v1(&a).unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    let text = first.snapshot().unwrap().slides[0].shapes[0]
        .text_stories
        .last()
        .unwrap()
        .plain_text();
    assert_eq!(text, "Peer R2C2");
}

#[test]
fn concurrent_structure_conflicts_refuse_without_losing_either_edit() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 83).unwrap();
    let second = DeckSession::open(&bytes, 84).unwrap();
    let (slide, shape) = target(&first);
    first
        .insert_table_row(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["First".into()],
        )
        .unwrap();
    second
        .insert_table_column(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["Second".into()],
        )
        .unwrap();
    let a = first.encode_state_as_update_v1();
    let b = second.encode_state_as_update_v1();
    assert!(matches!(
        first.apply_update_v1(&b),
        Err(EditError::InvalidTable(_))
    ));
    assert!(matches!(
        second.apply_update_v1(&a),
        Err(EditError::InvalidTable(_))
    ));
    assert_eq!(first.encode_state_as_update_v1(), a);
    assert_eq!(second.encode_state_as_update_v1(), b);
}

#[test]
fn all_merge_forms_refuse_structural_edits() {
    for attribute in ["rowSpan=\"2\"", "hMerge=\"1\"", "vMerge=\"1\""] {
        let mut parts = ooxml_opc::unzip_parts(&source(true)).unwrap();
        let (_, xml) = parts
            .iter_mut()
            .find(|(name, _)| name == "ppt/slides/slide1.xml")
            .unwrap();
        *xml = String::from_utf8(xml.clone())
            .unwrap()
            .replace("gridSpan=\"2\"", attribute)
            .into_bytes();
        let deck = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 85).unwrap();
        let (slide, shape) = target(&deck);
        let before = deck.encode_state_as_update_v1();
        assert!(matches!(
            deck.delete_table_row(&EditCtx::local("test"), &slide, &shape, 0),
            Err(EditError::InvalidTable(_))
        ));
        assert!(matches!(
            deck.delete_table_column(&EditCtx::local("test"), &slide, &shape, 0),
            Err(EditError::InvalidTable(_))
        ));
        assert_eq!(deck.encode_state_as_update_v1(), before);
    }
}

#[test]
fn edited_tables_can_be_edited_again_and_duplicated() {
    let deck = DeckSession::open(&source(false), 86).unwrap();
    let (slide, shape) = target(&deck);
    let ctx = EditCtx::local("test");
    deck.insert_table_row(
        &ctx,
        &slide,
        &shape,
        0,
        TableInsertPosition::After,
        &["New".into()],
    )
    .unwrap();
    let current = deck.snapshot().unwrap();
    let story = &current.slides[0].shapes[0].text_stories[3].id;
    deck.insert_text(&ctx, story, 0, "Updated ", &TextStyle::default())
        .unwrap();
    deck.duplicate_slide(&ctx, &slide, 1).unwrap();
    let saved = deck.save().unwrap();
    let reopened = DeckSession::open(&saved, 87).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides.len(), 2);
    for slide in &snapshot.slides {
        assert_eq!(slide.shapes[0].text_stories.len(), 12);
        assert_eq!(slide.shapes[0].text_stories[3].plain_text(), "Updated New");
    }
}

#[test]
fn inserted_cells_have_exact_text_when_the_template_has_multiple_paragraphs() {
    let mut parts = ooxml_opc::unzip_parts(&source(false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    *xml = String::from_utf8(xml.clone())
        .unwrap()
        .replacen(
            "</a:p>",
            "</a:p><a:p><a:r><a:rPr i=\"1\"/><a:t>More</a:t></a:r></a:p>",
            1,
        )
        .into_bytes();
    let deck = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 88).unwrap();
    let (slide, shape) = target(&deck);
    deck.insert_table_row(
        &EditCtx::local("test"),
        &slide,
        &shape,
        0,
        TableInsertPosition::After,
        &["Exact".into()],
    )
    .unwrap();
    let snapshot = deck.snapshot().unwrap();
    assert_eq!(
        snapshot.slides[0].shapes[0].text_stories[3].plain_text(),
        "Exact"
    );
    assert_eq!(
        snapshot.slides[0].shapes[0].text_stories[4].plain_text(),
        ""
    );
    let reopened = DeckSession::open(&deck.save().unwrap(), 89).unwrap();
    assert_eq!(
        reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[3].plain_text(),
        "Exact"
    );
}

#[test]
fn given_text_inherits_formatting_from_an_empty_template_cell() {
    let mut parts = ooxml_opc::unzip_parts(&source(false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    *xml = String::from_utf8(xml.clone())
        .unwrap()
        .replace("R0C0", "")
        .into_bytes();
    let deck = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 90).unwrap();
    let (slide, shape) = target(&deck);
    deck.insert_table_row(
        &EditCtx::local("test"),
        &slide,
        &shape,
        0,
        TableInsertPosition::After,
        &["New".into()],
    )
    .unwrap();
    let reopened = DeckSession::open(&deck.save().unwrap(), 91).unwrap();
    let table = table(&reopened);
    let run = &table.rows[1].cells[0].text.paragraphs[0].runs[0];
    assert_eq!(run.text, "New");
    assert_eq!(run.properties.bold, Some(true));
}

#[test]
fn pending_structural_redo_refuses_remote_structure_until_history_is_cleared() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 92).unwrap();
    let second = DeckSession::open(&bytes, 93).unwrap();
    let (slide, shape) = target(&first);
    first
        .insert_table_row(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    second
        .insert_table_column(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    assert!(first.undo());
    let before = first.encode_state_as_update_v1();
    let remote = second.encode_state_as_update_v1();
    assert!(matches!(
        first.apply_update_v1(&remote),
        Err(EditError::InvalidUpdate(_))
    ));
    assert_eq!(first.encode_state_as_update_v1(), before);
    assert!(first.redo());
    assert_eq!(table(&first).rows.len(), 4);
    assert!(first.undo());
    first.clear_undo_history();
    first.apply_update_v1(&remote).unwrap();
    assert!(!first.redo());
    assert_eq!(table(&first).grid.len(), 4);
}

#[test]
fn downstream_structure_updates_do_not_invalidate_local_structural_undo() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 94).unwrap();
    let second = DeckSession::open(&bytes, 95).unwrap();
    let (slide, shape) = target(&first);
    first
        .insert_table_row(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    second
        .insert_table_column(
            &EditCtx::local("test"),
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &[],
        )
        .unwrap();
    let before = first.encode_state_as_update_v1();
    let remote = second.encode_state_as_update_v1();
    assert!(matches!(
        first.apply_update_v1(&remote),
        Err(EditError::InvalidUpdate(_))
    ));
    assert_eq!(first.encode_state_as_update_v1(), before);
    first.clear_undo_history();
    first.apply_update_v1(&remote).unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    assert!(!first.undo());
}
