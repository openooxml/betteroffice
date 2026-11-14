use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};

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
    session.delete_slide(&EditCtx::local("test"), &slide).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn caret_anchor_resolves_to_none_once_its_shape_is_removed() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, shape) = story_home(&session, &story);
    session.remove_shape(&EditCtx::local("test"), &slide, &shape).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), None);
}

#[test]
fn caret_anchor_resolves_again_after_its_slide_deletion_is_undone() {
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    session.delete_slide(&EditCtx::local("test"), &slide).unwrap();
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
    session.delete_slide(&EditCtx::local("test"), &other).unwrap();
    assert_eq!(session.resolve_caret_anchor(&anchor), Some(2));
}

#[test]
fn caret_anchor_report_insert_into_story_of_deleted_slide() {
    const MARK: &str = "G31MARK";
    let (session, story) = session_with_text("ABC");
    let anchor = session.anchor_caret(&story, 2).unwrap();
    let (slide, _) = story_home(&session, &story);
    let events = std::rc::Rc::new(std::cell::Cell::new(0usize));
    let counted = events.clone();
    let _subscription = session.observe_update_v1(move |_| counted.set(counted.get() + 1));
    println!();
    match session.delete_slide(&EditCtx::local("test"), &slide) {
        Ok(_) => println!("G31-REPORT delete_slide ok"),
        Err(error) => println!("G31-REPORT delete_slide error: {error}"),
    }
    let after_delete = events.get();
    let resolved = session.resolve_caret_anchor(&anchor);
    println!("G31-REPORT resolve_caret_anchor after delete: {resolved:?}");
    match resolved {
        Some(index) => match session.insert_text(
            &EditCtx::local("test"),
            &story,
            index,
            MARK,
            &TextStyle::default(),
        ) {
            Ok(receipt) => println!("G31-REPORT insert_text ok: {receipt:?}"),
            Err(error) => println!("G31-REPORT insert_text error: {error}"),
        },
        None => println!("G31-REPORT insert_text skipped: anchor resolved to None"),
    }
    match session.story(&story) {
        Ok(snapshot) => println!(
            "G31-REPORT dead story contains marker: {} (text {:?})",
            snapshot.plain_text().contains(MARK),
            snapshot.plain_text()
        ),
        Err(error) => println!("G31-REPORT dead story read error: {error}"),
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
                    "G31-REPORT saved slide parts contain marker: {} {hits:?}",
                    !hits.is_empty()
                );
            }
            Err(error) => println!("G31-REPORT saved bytes unzip error: {error}"),
        },
        Err(error) => println!("G31-REPORT save error: {error}"),
    }
    println!(
        "G31-REPORT update events: {} on delete, {} after insert",
        after_delete,
        events.get() - after_delete
    );
}
