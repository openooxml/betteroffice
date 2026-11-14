use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TableInsertPosition, TextStyle};

const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const TABLE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn grouped_session() -> DeckSession {
    let mut parts = ooxml_opc::unzip_parts(DEMO).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(bytes.clone()).unwrap();
    assert!(xml.contains("<p:sp>"));
    *bytes = xml
        .replacen(
            "<p:sp>",
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"900\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp>",
            1,
        )
        .replacen("</p:spTree>", "</p:grpSp></p:spTree>", 1)
        .into_bytes();
    DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 84093).unwrap()
}

fn table_cell(session: &DeckSession, row: usize, column: usize) -> String {
    let snapshot = session.snapshot().unwrap();
    let table = &snapshot.slides[0].shapes[0];
    let base = format!("story:{}:table:{row}:{column}", table.id);
    table
        .text_stories
        .iter()
        .find(|story| story.id == base || story.id.starts_with(&format!("{base}:")))
        .unwrap()
        .id
        .clone()
}

fn session_with_text(text: &str) -> (DeckSession, String) {
    let session = DeckSession::open(
        include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx"),
        84092,
    )
    .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let added = session
        .add_text_box(
            &EditCtx::local("test"),
            &slide,
            &ShapeDraft {
                name: "Anchors".to_owned(),
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
    let story = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == added.shape_id)
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session.add_undo_barrier();
    (session, story)
}

#[test]
fn caret_anchor_follows_undo_and_redo_of_matching_text() {
    let (session, story) = session_with_text("ABC");
    let context = EditCtx::local("test");
    session
        .insert_text(&context, &story, 0, "A", &TextStyle::default())
        .unwrap();
    let after_typed = session.anchor_caret(&story, 1).unwrap();

    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&after_typed), Some(0));
    assert!(session.redo());
    assert_eq!(session.resolve_caret_anchor(&after_typed), Some(1));
}

#[test]
fn caret_anchor_at_story_start_stays_before_later_insertions() {
    let (session, story) = session_with_text("ABC");
    let start = session.anchor_caret(&story, 0).unwrap();
    let end = session.anchor_caret(&story, 3).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "xy",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&start), Some(0));
    assert_eq!(session.resolve_caret_anchor(&end), Some(5));
    assert!(session.anchor_caret(&story, 6).is_err());
}

fn story_home(session: &DeckSession, story: &str) -> (String, String) {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .iter()
                .find(|shape| shape.text_stories.iter().any(|text| text.id == story))
                .map(|shape| (slide.id.clone(), shape.id.clone()))
        })
        .unwrap()
}

#[test]
fn caret_anchor_resolves_to_none_once_its_slide_is_deleted() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    session
        .delete_slide(&EditCtx::local("test"), &slide)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn caret_anchor_resolves_to_none_once_its_shape_is_removed() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, shape) = story_home(&session, &story);
    session
        .remove_shape(&EditCtx::local("test"), &slide, &shape)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn caret_anchor_resolves_again_after_its_slide_deletion_is_undone() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    session
        .delete_slide(&EditCtx::local("test"), &slide)
        .unwrap();
    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}

#[test]
fn caret_anchor_on_a_remaining_slide_survives_another_slide_deletion() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    let other = session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .find(|candidate| candidate.id != slide)
        .unwrap()
        .id
        .clone();
    session
        .delete_slide(&EditCtx::local("test"), &other)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}

#[test]
fn caret_anchor_report_insert_into_story_of_deleted_slide() {
    const MARK: &str = "ANCHORMARK";
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    let events = std::rc::Rc::new(std::cell::Cell::new(0usize));
    let counted = events.clone();
    let _subscription = session.observe_update_v1(move |_| counted.set(counted.get() + 1));
    println!();
    match session.delete_slide(&EditCtx::local("test"), &slide) {
        Ok(_) => println!("ANCHOR-REPORT delete_slide ok"),
        Err(error) => println!("ANCHOR-REPORT delete_slide error: {error}"),
    }
    let after_delete = events.get();
    let resolved = session.resolve_caret_anchor(&anchor);
    println!("ANCHOR-REPORT resolve_caret_anchor after delete: {resolved:?}");
    match resolved {
        Some(index) => match session.insert_text(
            &EditCtx::local("test"),
            &story,
            index,
            MARK,
            &TextStyle::default(),
        ) {
            Ok(receipt) => println!("ANCHOR-REPORT insert_text ok: {receipt:?}"),
            Err(error) => println!("ANCHOR-REPORT insert_text error: {error}"),
        },
        None => println!("ANCHOR-REPORT insert_text skipped: anchor resolved to None"),
    }
    match session.story(&story) {
        Ok(snapshot) => println!(
            "ANCHOR-REPORT dead story contains marker: {} (text {:?})",
            snapshot.plain_text().contains(MARK),
            snapshot.plain_text()
        ),
        Err(error) => println!("ANCHOR-REPORT dead story read error: {error}"),
    }
    match session.save() {
        Ok(bytes) => match ooxml_opc::unzip_parts(&bytes) {
            Ok(parts) => {
                let hits: Vec<&str> = parts
                    .iter()
                    .filter(|(name, _)| name.starts_with("ppt/slides/"))
                    .filter(|(_, data)| String::from_utf8_lossy(data).contains(MARK))
                    .map(|(name, _)| name.as_str())
                    .collect();
                println!(
                    "ANCHOR-REPORT saved slide parts contain marker: {} {hits:?}",
                    !hits.is_empty()
                );
            }
            Err(error) => println!("ANCHOR-REPORT saved bytes unzip error: {error}"),
        },
        Err(error) => println!("ANCHOR-REPORT save error: {error}"),
    }
    println!(
        "ANCHOR-REPORT update events: {} on delete, {} after insert",
        after_delete,
        events.get() - after_delete
    );
}

#[test]
fn caret_anchor_in_a_group_child_survives_another_slide_deletion() {
    let session = grouped_session();
    let snapshot = session.snapshot().unwrap();
    let story = snapshot.slides[0]
        .shapes
        .iter()
        .flat_map(|shape| &shape.children)
        .flat_map(|child| &child.text_stories)
        .find(|text| !text.plain_text().is_empty())
        .unwrap()
        .id
        .clone();
    let anchor = session.anchor_caret(&story, 1).unwrap();
    let other = snapshot.slides[1].id.clone();
    session
        .delete_slide(&EditCtx::local("test"), &other)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(1));
}

#[test]
fn caret_anchor_resolves_to_none_once_its_table_row_is_deleted() {
    let session = DeckSession::open(TABLE, 84094).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let context = EditCtx::local("test");
    session
        .insert_table_row(
            &context,
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["XY".to_owned()],
        )
        .unwrap();
    session.add_undo_barrier();
    let anchor = session
        .anchor_caret(&table_cell(&session, 1, 0), 1)
        .unwrap();
    session
        .delete_table_row(&context, &slide, &shape, 1)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn an_anchor_taken_before_a_row_insert_resolves_to_none() {
    let session = DeckSession::open(TABLE, 84095).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let context = EditCtx::local("test");
    session.add_undo_barrier();
    let anchor = session
        .anchor_caret(&table_cell(&session, 0, 0), 0)
        .unwrap();
    session
        .insert_table_row(
            &context,
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["XY".to_owned()],
        )
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn an_anchor_in_another_row_taken_before_a_row_delete_resolves_to_none() {
    let session = DeckSession::open(TABLE, 84096).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].id.clone();
    let shape = snapshot.slides[0].shapes[0].id.clone();
    let context = EditCtx::local("test");
    session
        .insert_table_row(
            &context,
            &slide,
            &shape,
            0,
            TableInsertPosition::After,
            &["XY".to_owned()],
        )
        .unwrap();
    session.add_undo_barrier();
    let anchor = session
        .anchor_caret(&table_cell(&session, 0, 0), 0)
        .unwrap();
    session
        .delete_table_row(&context, &slide, &shape, 1)
        .unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn caret_anchors_on_a_duplicated_slide_and_its_original_after_the_duplicate_is_deleted() {
    let (session, story) = session_with_text("ABC");
    let (slide, _) = story_home(&session, &story);
    let context = EditCtx::local("test");
    session.duplicate_slide(&context, &slide, 1).unwrap();
    let duplicate = session.snapshot().unwrap().slides[1].clone();
    let copy = duplicate
        .shapes
        .iter()
        .find(|shape| shape.name == "Anchors")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    assert_ne!(copy, story);
    let original = session.anchor_caret(&story, 2).unwrap();
    let copied = session.anchor_caret(&copy, 1).unwrap();
    session.delete_slide(&context, &duplicate.id).unwrap();
    assert_eq!(session.resolve_caret_anchor(&copied), None);
    assert_eq!(session.resolve_caret_anchor(&original), Some(2));
}

#[test]
fn caret_anchor_survives_moving_its_slide() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    session
        .move_slide(&EditCtx::local("test"), &slide, 1)
        .unwrap();
    assert_eq!(session.snapshot().unwrap().slides[1].id, slide);
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}

#[test]
fn caret_anchor_follows_undo_and_redo_of_its_shape_removal() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, shape) = story_home(&session, &story);
    session
        .remove_shape(&EditCtx::local("test"), &slide, &shape)
        .unwrap();
    session.add_undo_barrier();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
    assert!(session.redo());
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
    assert!(session.undo());
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}
