use pptx_edit::{DeckSession, EditCtx, EditError, TableEdit, TableInsertPosition, TextStyle};
use pptx_parse::GraphicFrameData;

const BASE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn source(merged: bool) -> Vec<u8> {
    source_sized(merged, 6_000_000, 1_200_000)
}

fn source_sized(merged: bool, cx: i64, cy: i64) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(BASE).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let extent = xml.find("<a:ext ").unwrap();
    let end = extent + xml[extent..].find("/>").unwrap() + 2;
    xml.replace_range(extent..end, &format!(r#"<a:ext cx="{cx}" cy="{cy}"/>"#));
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
fn concurrent_structure_edits_converge_on_every_replica() {
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
    first.apply_update_v1(&b).unwrap();
    second.apply_update_v1(&a).unwrap();
    let third = DeckSession::open(&bytes, 85).unwrap();
    third.apply_update_v1(&b).unwrap();
    third.apply_update_v1(&a).unwrap();
    let discarded = [
        first.take_discarded_table_edits(),
        second.take_discarded_table_edits(),
    ];
    assert_eq!(
        discarded
            .iter()
            .filter(|ids| ids.as_slice() == std::slice::from_ref(&shape))
            .count(),
        1
    );
    assert!(discarded.iter().any(Vec::is_empty));
    assert!(third.take_discarded_table_edits().is_empty());
    let merged = first.encode_state_as_update_v1();
    let reopened = DeckSession::open_from_update(&merged, 86).unwrap();
    let expected = first.snapshot().unwrap();
    for replica in [&second, &third, &reopened] {
        assert_eq!(replica.snapshot().unwrap(), expected);
    }
    let table = table(&first);
    let shape_snapshot = &expected.slides[0].shapes[0];
    assert!(
        (table.rows.len(), table.grid.len()) == (4, 3)
            || (table.rows.len(), table.grid.len()) == (3, 4)
    );
    assert_eq!(
        shape_snapshot.text_stories.len(),
        table.rows.len() * table.grid.len()
    );
    assert!(!first.can_undo());
    assert!(!second.can_undo());
    let saved = first.save().unwrap();
    let attached = DeckSession::open_from_update_with_source(&merged, &bytes, 88).unwrap();
    assert_eq!(saved, attached.save().unwrap());
    let from_file = DeckSession::open(&saved, 87).unwrap();
    let saved_table = {
        let snapshot = from_file.snapshot().unwrap();
        let Some(GraphicFrameData::Table(table)) = &snapshot.slides[0].shapes[0].graphic else {
            panic!("table")
        };
        table.clone()
    };
    assert_eq!(saved_table.grid, table.grid);
    assert_eq!(saved_table.rows.len(), table.rows.len());
    let texts: Vec<_> = from_file.snapshot().unwrap().slides[0].shapes[0]
        .text_stories
        .iter()
        .map(|story| story.plain_text())
        .collect();
    let in_memory: Vec<_> = shape_snapshot
        .text_stories
        .iter()
        .map(|story| story.plain_text())
        .collect();
    assert_eq!(texts, in_memory);
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
fn remote_structure_drops_pending_structural_redo_for_that_table() {
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
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    assert!(!first.redo());
    assert!(matches!(table(&first).grid.len(), 3 | 4));
    assert_eq!(table(&first).rows.len(), 3);
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
}

#[test]
fn downstream_structure_update_applies_and_drops_only_that_tables_structural_undo() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 94).unwrap();
    let second = DeckSession::open(&bytes, 95).unwrap();
    let (slide, shape) = target(&first);
    let story = first.snapshot().unwrap().slides[0].shapes[0].text_stories[8]
        .id
        .clone();
    first
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "Local ",
            &TextStyle::default(),
        )
        .unwrap();
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
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    assert_eq!(table(&first).rows.len(), 4);
    assert_eq!(table(&first).grid.len(), 4);
    assert!(first.undo());
    let snapshot = first.snapshot().unwrap();
    let stories = &snapshot.slides[0].shapes[0].text_stories;
    assert_eq!(stories.len(), 16);
    assert_eq!(stories[15].plain_text(), "R2C2");
    assert_eq!(table(&first).rows.len(), 4);
    assert_eq!(table(&first).grid.len(), 4);
    assert!(!first.undo());
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
}

#[test]
fn duplicated_tables_accept_remote_structure_edits() {
    let bytes = source(false);
    let first = DeckSession::open(&bytes, 96).unwrap();
    let second = DeckSession::open(&bytes, 97).unwrap();
    let (slide, _) = target(&first);
    let copy = first
        .duplicate_slide(&EditCtx::local("test"), &slide, 1)
        .unwrap()
        .slide_id;
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    let copied_shape = second
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find(|slide| slide.id == copy)
        .unwrap()
        .shapes[0]
        .id
        .clone();
    second
        .insert_table_row(
            &EditCtx::local("test"),
            &copy,
            &copied_shape,
            0,
            TableInsertPosition::After,
            &["Peer".into()],
        )
        .unwrap();
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    assert!(first.can_undo());
}

#[test]
fn row_and_column_edits_keep_grown_frame_size_by_delta() {
    let ctx = EditCtx::local("test");
    let deck = DeckSession::open(&source_sized(false, 7_000_000, 2_000_000), 98).unwrap();
    let (slide, shape) = target(&deck);
    let size = |deck: &DeckSession| {
        let snapshot = deck.snapshot().unwrap();
        (
            snapshot.slides[0].shapes[0].width,
            snapshot.slides[0].shapes[0].height,
        )
    };
    deck.insert_table_row(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[])
        .unwrap();
    assert_eq!(size(&deck), (7_000_000, 2_400_000));
    deck.delete_table_row(&ctx, &slide, &shape, 0).unwrap();
    deck.delete_table_row(&ctx, &slide, &shape, 0).unwrap();
    assert_eq!(size(&deck), (7_000_000, 1_600_000));
    deck.insert_table_column(&ctx, &slide, &shape, 1, TableInsertPosition::After, &[])
        .unwrap();
    assert_eq!(size(&deck), (7_000_000, 1_600_000));
    deck.delete_table_column(&ctx, &slide, &shape, 3).unwrap();
    assert_eq!(size(&deck), (4_000_000, 1_600_000));
    let reopened = DeckSession::open(&deck.save().unwrap(), 99).unwrap();
    assert_eq!(size(&reopened), (4_000_000, 1_600_000));
}

#[test]
fn new_cell_text_takes_the_first_template_run_style_whole() {
    let mut parts = ooxml_opc::unzip_parts(&source(false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    *xml = String::from_utf8(xml.clone())
        .unwrap()
        .replacen(
            r#"<a:r><a:rPr b="1" lang="en-US" strike="sngStrike"/><a:t>R0C0</a:t></a:r>"#,
            r#"<a:r><a:rPr b="1" lang="en-US"/><a:t>Q1 </a:t></a:r><a:r><a:rPr b="0" lang="en-US"/><a:t>revenue</a:t></a:r>"#,
            1,
        )
        .into_bytes();
    let deck = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 100).unwrap();
    let (slide, shape) = target(&deck);
    deck.insert_table_row(
        &EditCtx::local("test"),
        &slide,
        &shape,
        0,
        TableInsertPosition::After,
        &["Total".into()],
    )
    .unwrap();
    let reopened = DeckSession::open(&deck.save().unwrap(), 101).unwrap();
    let table = table(&reopened);
    let runs = &table.rows[1].cells[0].text.paragraphs[0].runs;
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text, "Total");
    assert_eq!(runs[0].properties.bold, Some(true));
}

#[test]
fn cells_with_line_breaks_are_refused() {
    let ctx = EditCtx::local("test");
    let deck = DeckSession::open(&source(false), 102).unwrap();
    let (slide, shape) = target(&deck);
    let before = deck.encode_state_as_update_v1();
    for text in ["Two\nlines", "Two\rlines"] {
        let error = deck
            .insert_table_row(
                &ctx,
                &slide,
                &shape,
                0,
                TableInsertPosition::After,
                &[text.into()],
            )
            .unwrap_err();
        assert!(matches!(error, EditError::InvalidTable(_)));
        assert!(error.to_string().contains("line break"));
        assert!(
            deck.insert_table_column(
                &ctx,
                &slide,
                &shape,
                0,
                TableInsertPosition::After,
                &[text.into()]
            )
            .is_err()
        );
    }
    assert_eq!(deck.encode_state_as_update_v1(), before);
}

#[test]
fn inserted_rows_and_columns_get_fresh_unique_powerpoint_ids() {
    const A16: &str = "http://schemas.microsoft.com/office/drawing/2014/main";
    let mut parts = ooxml_opc::unzip_parts(&source(false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut text = String::from_utf8(xml.clone()).unwrap();
    for (i, width) in ["1000000", "2000000", "3000000"].into_iter().enumerate() {
        text = text.replacen(
            &format!(r#"<a:gridCol w="{width}"/>"#),
            &format!(
                r#"<a:gridCol w="{width}"><a:extLst><a:ext uri="{{9D8B030D-6E8A-4147-A177-3AD203B41FA5}}"><a16:colId xmlns:a16="{A16}" val="{}"/></a:ext></a:extLst></a:gridCol>"#,
                20_000 + i
            ),
            1,
        );
    }
    let mut row = 0;
    while let Some(at) = text.find("</a:tc></a:tr>") {
        text.replace_range(
            at..at + "</a:tc></a:tr>".len(),
            &format!(
                r#"</a:tc><a:extLst><a:ext uri="{{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}}"><a16:rowId xmlns:a16="{A16}" val="{}"/></a:ext></a:extLst></a:tr>"#,
                10_000 + row
            ),
        );
        row += 1;
    }
    assert_eq!(row, 3);
    *xml = text.into_bytes();
    let deck = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 103).unwrap();
    let (slide, shape) = target(&deck);
    let ctx = EditCtx::local("test");
    deck.insert_table_row(&ctx, &slide, &shape, 1, TableInsertPosition::After, &[])
        .unwrap();
    deck.insert_table_column(&ctx, &slide, &shape, 0, TableInsertPosition::Before, &[])
        .unwrap();
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
    let ids = |name: &str| -> Vec<String> {
        xml.match_indices(&format!("<a16:{name} "))
            .map(|(at, _)| {
                let rest = &xml[at..];
                let start = rest.find("val=\"").unwrap() + 5;
                let end = start + rest[start..].find('"').unwrap();
                rest[start..end].to_owned()
            })
            .collect()
    };
    let rows = ids("rowId");
    let columns = ids("colId");
    assert_eq!(rows.len(), 4);
    assert_eq!(columns.len(), 4);
    let mut all = rows.clone();
    all.extend(columns.iter().cloned());
    let unique: std::collections::HashSet<_> = all.iter().collect();
    assert_eq!(unique.len(), all.len());
    assert_eq!([&rows[0], &rows[1], &rows[3]], ["10000", "10001", "10002"]);
    assert_eq!(
        [&columns[1], &columns[2], &columns[3]],
        ["20000", "20001", "20002"]
    );
    let reopened = DeckSession::open(&saved, 104).unwrap();
    assert_eq!(table(&reopened).rows.len(), 4);
    assert_eq!(table(&reopened).grid.len(), 4);
}

#[test]
fn collaborative_sessions_refuse_structure_and_drop_its_local_history() {
    let deck = DeckSession::open(&source(false), 105).unwrap();
    let (slide, shape) = target(&deck);
    let ctx = EditCtx::local("test");
    let story = deck.snapshot().unwrap().slides[0].shapes[0].text_stories[8]
        .id
        .clone();
    deck.insert_text(&ctx, &story, 0, "Kept ", &TextStyle::default())
        .unwrap();
    deck.insert_table_row(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[])
        .unwrap();
    assert!(!deck.is_collaborative());
    deck.set_collaborative().unwrap();
    assert!(deck.is_collaborative());
    let before = deck.encode_state_as_update_v1();
    for op in 0..4 {
        let result = match op {
            0 => deck.delete_table_column(&ctx, &slide, &shape, 0),
            1 => deck.delete_table_row(&ctx, &slide, &shape, 0),
            2 => deck.insert_table_column(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[]),
            _ => deck.insert_table_row(&ctx, &slide, &shape, 0, TableInsertPosition::After, &[]),
        };
        assert!(matches!(
            result,
            Err(EditError::CollaborativeStructureOperation)
        ));
    }
    assert_eq!(deck.encode_state_as_update_v1(), before);
    assert!(deck.undo());
    assert_eq!(table(&deck).rows.len(), 4);
    let stories = &deck.snapshot().unwrap().slides[0].shapes[0].text_stories;
    assert_eq!(stories[11].plain_text(), "R2C2");
    assert!(!deck.undo());
}

#[test]
fn joining_a_collaborative_session_drops_pending_structural_redo() {
    let deck = DeckSession::open(&source(false), 106).unwrap();
    let (slide, shape) = target(&deck);
    deck.insert_table_row(
        &EditCtx::local("test"),
        &slide,
        &shape,
        0,
        TableInsertPosition::After,
        &[],
    )
    .unwrap();
    assert!(deck.undo());
    assert!(deck.can_redo());
    deck.set_collaborative().unwrap();
    assert!(!deck.can_redo());
    assert!(!deck.redo());
    assert_eq!(table(&deck).rows.len(), 3);
}

#[test]
fn table_edit_checks_change_nothing_and_refuse_like_the_edit() {
    let bytes = source(false);
    let checked = DeckSession::open(&bytes, 107).unwrap();
    let plain = DeckSession::open(&bytes, 107).unwrap();
    let (slide, shape) = target(&checked);
    let after = TableInsertPosition::After;
    let before = checked.encode_state_as_update_v1();
    for edit in [
        TableEdit::InsertRow,
        TableEdit::DeleteRow,
        TableEdit::InsertColumn,
        TableEdit::DeleteColumn,
    ] {
        checked
            .check_table_edit(edit, &slide, &shape, 0, after, &[])
            .unwrap();
        assert!(matches!(
            checked.check_table_edit(edit, &slide, &shape, 3, after, &[]),
            Err(EditError::InvalidTable(_))
        ));
    }
    assert!(matches!(
        checked.check_table_edit(
            TableEdit::InsertRow,
            &slide,
            &shape,
            0,
            after,
            &["Two\nlines".into()]
        ),
        Err(EditError::InvalidTable(_))
    ));
    assert!(
        checked
            .check_table_edit(TableEdit::InsertRow, &slide, "missing", 0, after, &[])
            .is_err()
    );
    assert_eq!(checked.encode_state_as_update_v1(), before);
    assert!(!checked.can_undo());
    let ctx = EditCtx::local("test");
    checked
        .insert_table_row(&ctx, &slide, &shape, 0, after, &[])
        .unwrap();
    plain
        .insert_table_row(&ctx, &slide, &shape, 0, after, &[])
        .unwrap();
    assert_eq!(checked.snapshot().unwrap(), plain.snapshot().unwrap());
    checked.set_collaborative().unwrap();
    assert!(matches!(
        checked.check_table_edit(TableEdit::InsertRow, &slide, &shape, 0, after, &[]),
        Err(EditError::CollaborativeStructureOperation)
    ));
    let merged = DeckSession::open(&source(true), 108).unwrap();
    let (slide, shape) = target(&merged);
    assert!(matches!(
        merged.check_table_edit(TableEdit::DeleteRow, &slide, &shape, 0, after, &[]),
        Err(EditError::InvalidTable(_))
    ));
}
