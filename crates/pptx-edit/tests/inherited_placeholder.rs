use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, ShapeRect, TextStyle};

const FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/style-matrix-deck.pptx");

/// The layout transform of `slide2`'s first placeholder, which carries no
/// `a:xfrm` of its own.
const LAYOUT_RECT: ShapeRect = ShapeRect {
    x: 838_200,
    y: 365_125,
    width: 10_515_600,
    height: 1_325_563,
};

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn context() -> EditCtx {
    EditCtx::local("test")
}

#[test]
fn a_placeholder_without_a_transform_reports_the_layout_geometry() {
    let session = DeckSession::open(FIXTURE, 901).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = &snapshot.slides[1].shapes[0];

    assert_eq!((shape.x, shape.y, shape.width, shape.height), (0, 0, 0, 0));
    let inherited = shape
        .inherited
        .expect("the layout placeholder has geometry");
    assert_eq!(
        (inherited.x, inherited.y, inherited.width, inherited.height),
        (
            LAYOUT_RECT.x,
            LAYOUT_RECT.y,
            LAYOUT_RECT.width,
            LAYOUT_RECT.height
        )
    );
    assert!(snapshot.slides[0].shapes[0].inherited.is_none());
}

#[test]
fn the_first_rect_on_an_inherited_placeholder_writes_all_four_values() {
    let session = DeckSession::open(FIXTURE, 902).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[1].id.clone();
    let shape_id = snapshot.slides[1].shapes[0].id.clone();
    let moved = ShapeRect {
        x: LAYOUT_RECT.x + 400_000,
        y: LAYOUT_RECT.y + 200_000,
        ..LAYOUT_RECT
    };

    session
        .set_shape_rect(&context(), &slide_id, &shape_id, moved)
        .unwrap();

    let edited = &session.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (edited.x, edited.y, edited.width, edited.height),
        (moved.x, moved.y, moved.width, moved.height)
    );
    assert!(edited.inherited.is_none());

    let saved = session.save().unwrap();
    let slide = String::from_utf8(parts(&saved)["ppt/slides/slide2.xml"].clone()).unwrap();
    assert!(slide.contains(&format!(
        "<a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/>",
        moved.x, moved.y, moved.width, moved.height
    )));
    let reopened = DeckSession::open(&saved, 903).unwrap();
    let persisted = &reopened.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (
            persisted.x,
            persisted.y,
            persisted.width,
            persisted.height,
            persisted.inherited
        ),
        (moved.x, moved.y, moved.width, moved.height, None)
    );
}

#[test]
fn an_unmoved_placeholder_keeps_inheriting_across_an_edit_to_its_slide() {
    let session = DeckSession::open(FIXTURE, 904).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[1].shapes[1].text_stories[0].id.clone();

    session
        .insert_text(&context(), &story_id, 0, "edited ", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = String::from_utf8(parts(&saved)["ppt/slides/slide2.xml"].clone()).unwrap();
    assert!(slide.contains("edited "));
    assert_eq!(slide.matches("<p:spPr/>").count(), 5);
    let reopened = DeckSession::open(&saved, 905).unwrap();
    let untouched = &reopened.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (untouched.x, untouched.y, untouched.width, untouched.height),
        (0, 0, 0, 0)
    );
    assert!(untouched.inherited.is_some());
}

/// The fixture with the layout title placeholder turned and flipped.
fn turned_fixture() -> Vec<u8> {
    let mut entries = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    for (path, bytes) in &mut entries {
        if path == "ppt/slideLayouts/slideLayout2.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "<a:xfrm><a:off x=\"838200\" y=\"365125\"/>",
                    "<a:xfrm rot=\"1800000\" flipV=\"1\"><a:off x=\"838200\" y=\"365125\"/>",
                )
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&entries).unwrap()
}

#[test]
fn a_first_move_keeps_the_orientation_the_placeholder_inherits() {
    let fixture = turned_fixture();
    let session = DeckSession::open(&fixture, 906).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[1].id.clone();
    let shape_id = snapshot.slides[1].shapes[0].id.clone();
    let inherited = snapshot.slides[1].shapes[0].inherited.unwrap();
    assert_eq!((inherited.rotation_deg, inherited.flip_v), (30.0, true));

    session
        .move_shape(
            &context(),
            &slide_id,
            &shape_id,
            LAYOUT_RECT.x + 400_000,
            LAYOUT_RECT.y,
        )
        .unwrap();

    let moved = &session.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (moved.rotation_deg, moved.flip_v, moved.flip_h),
        (30.0, true, false)
    );
    assert_eq!(
        (moved.width, moved.height),
        (LAYOUT_RECT.width, LAYOUT_RECT.height)
    );
    assert!(moved.inherited.is_none());

    let reopened = DeckSession::open(&session.save().unwrap(), 907).unwrap();
    let persisted = &reopened.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (
            persisted.x,
            persisted.y,
            persisted.width,
            persisted.height,
            persisted.rotation_deg,
            persisted.flip_v
        ),
        (
            LAYOUT_RECT.x + 400_000,
            LAYOUT_RECT.y,
            LAYOUT_RECT.width,
            LAYOUT_RECT.height,
            30.0,
            true
        )
    );
}

#[test]
fn a_position_only_move_leaves_no_half_materialized_transform() {
    let session = DeckSession::open(FIXTURE, 908).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[1].id.clone();
    let shape_id = snapshot.slides[1].shapes[0].id.clone();

    let receipt = session
        .move_shape(
            &context(),
            &slide_id,
            &shape_id,
            LAYOUT_RECT.x + 400_000,
            LAYOUT_RECT.y + 200_000,
        )
        .unwrap();
    assert_eq!(receipt.before, LAYOUT_RECT);

    let moved = &session.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (moved.x, moved.y, moved.width, moved.height),
        (
            LAYOUT_RECT.x + 400_000,
            LAYOUT_RECT.y + 200_000,
            LAYOUT_RECT.width,
            LAYOUT_RECT.height
        )
    );
    assert!(moved.inherited.is_none());
}

#[test]
fn undoing_a_first_geometry_edit_returns_the_placeholder_to_inheriting() {
    let session = DeckSession::open(FIXTURE, 909).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[1].id.clone();
    let shape_id = snapshot.slides[1].shapes[0].id.clone();

    session
        .move_shape(
            &context(),
            &slide_id,
            &shape_id,
            LAYOUT_RECT.x + 400_000,
            LAYOUT_RECT.y,
        )
        .unwrap();
    assert!(session.undo());

    let restored = &session.snapshot().unwrap().slides[1].shapes[0];
    assert_eq!(
        (restored.x, restored.y, restored.width, restored.height),
        (0, 0, 0, 0)
    );
    assert!(restored.inherited.is_some());

    let saved = session.save().unwrap();
    let slide = String::from_utf8(parts(&saved)["ppt/slides/slide2.xml"].clone()).unwrap();
    assert_eq!(slide.matches("<a:xfrm").count(), 1);
}
