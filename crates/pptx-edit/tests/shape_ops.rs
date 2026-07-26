use std::collections::BTreeMap;

use ooxml_drawingml::{ColorValue, ShapeFill, ShapeOutline};
use pptx_edit::{
    DeckSession, EditCtx, EditError, PresetShapeDraft, ShapeRect, ShapeSnapshot, ShapeStroke,
};
use pptx_parse::ShapeNode;

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

#[test]
fn shape_operations_round_trip_through_history_and_updates() {
    let session = DeckSession::open(FIXTURE, 701).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before_add = session.snapshot().unwrap();
    let receipt = session
        .add_shape(
            &context,
            &slide_id,
            &PresetShapeDraft {
                name: "Rounded rectangle".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: ShapeRect {
                    x: 900_000,
                    y: 1_100_000,
                    width: 3_200_000,
                    height: 1_400_000,
                },
                fill: Some("#D9EAF7".to_owned()),
            },
        )
        .unwrap();
    assert_history_round_trip(&session, before_add);

    let before_fill = session.snapshot().unwrap();
    let fill = session
        .set_shape_fill(&context, &slide_id, &receipt.shape_id, Some("#3367D6"))
        .unwrap();
    assert_eq!(fill.after.as_deref(), Some("#3367D6"));
    assert_history_round_trip(&session, before_fill);

    let before_stroke = session.snapshot().unwrap();
    let stroke = session
        .set_shape_stroke(
            &context,
            &slide_id,
            &receipt.shape_id,
            &ShapeStroke {
                color: Some("#EA4335".to_owned()),
                width_pt: Some(3.0),
            },
        )
        .unwrap();
    assert_eq!(stroke.after.unwrap().width_pt, Some(3.0));
    assert_history_round_trip(&session, before_stroke);

    let before_adjust = session.snapshot().unwrap();
    let adjust = session
        .set_shape_adjust(
            &context,
            &slide_id,
            &receipt.shape_id,
            &BTreeMap::from([("adj".to_owned(), 0.8)]),
        )
        .unwrap();
    assert_eq!(adjust.after.get("adj"), Some(&0.5));
    assert_history_round_trip(&session, before_adjust);

    let before_no_fill = session.snapshot().unwrap();
    session
        .set_shape_fill(&context, &slide_id, &receipt.shape_id, None)
        .unwrap();
    assert_history_round_trip(&session, before_no_fill);

    let before_no_line = session.snapshot().unwrap();
    session
        .set_shape_stroke(
            &context,
            &slide_id,
            &receipt.shape_id,
            &ShapeStroke::default(),
        )
        .unwrap();
    assert_history_round_trip(&session, before_no_line);

    let update = session.encode_state_as_update_v1();
    let replica = DeckSession::open_from_update(&update, 702).unwrap();
    assert_eq!(replica.snapshot().unwrap(), session.snapshot().unwrap());
}

#[test]
fn add_shape_rejects_unknown_geometry_and_invalid_colors() {
    let session = DeckSession::open(FIXTURE, 703).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let draft = PresetShapeDraft {
        name: "Unknown".to_owned(),
        geometry: "not-a-preset".to_owned(),
        rect: ShapeRect {
            x: 0,
            y: 0,
            width: 1_000_000,
            height: 1_000_000,
        },
        fill: None,
    };
    assert!(session.add_shape(&context, &slide_id, &draft).is_err());

    let mut invalid_color = draft;
    invalid_color.geometry = "rect".to_owned();
    invalid_color.fill = Some("#xyz".to_owned());
    assert!(
        session
            .add_shape(&context, &slide_id, &invalid_color)
            .is_err()
    );
}

#[test]
fn adjustable_presets_expose_engine_defaults() {
    let session = DeckSession::open(FIXTURE, 704).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let receipt = session
        .add_shape(
            &context,
            &slide_id,
            &PresetShapeDraft {
                name: "Parallelogram".to_owned(),
                geometry: "parallelogram".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: None,
            },
        )
        .unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == receipt.shape_id)
        .unwrap();

    assert_eq!(shape.adjust_values.get("adj"), Some(&0.25));
}

#[test]
fn set_shape_adjust_rejects_unknown_names_and_oversized_maps() {
    let session = DeckSession::open(FIXTURE, 705).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let receipt = session
        .add_shape(
            &context,
            &slide_id,
            &PresetShapeDraft {
                name: "Parallelogram".to_owned(),
                geometry: "parallelogram".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: None,
            },
        )
        .unwrap();
    let before = session.snapshot().unwrap();

    let invalid_name = session
        .set_shape_adjust(
            &context,
            &slide_id,
            &receipt.shape_id,
            &BTreeMap::from([("width".to_owned(), 0.5)]),
        )
        .unwrap_err();
    assert!(matches!(invalid_name, EditError::InvalidAdjustment(_)));

    let out_of_range_name = session
        .set_shape_adjust(
            &context,
            &slide_id,
            &receipt.shape_id,
            &BTreeMap::from([("adj33".to_owned(), 0.5)]),
        )
        .unwrap_err();
    assert!(matches!(out_of_range_name, EditError::InvalidAdjustment(_)));

    let mut oversized = BTreeMap::from([("adj".to_owned(), 0.5)]);
    for index in 1..=32 {
        oversized.insert(format!("adj{index}"), 0.5);
    }
    let too_many = session
        .set_shape_adjust(&context, &slide_id, &receipt.shape_id, &oversized)
        .unwrap_err();
    assert!(matches!(too_many, EditError::InvalidAdjustment(_)));

    let non_finite = session
        .set_shape_adjust(
            &context,
            &slide_id,
            &receipt.shape_id,
            &BTreeMap::from([("adj".to_owned(), f64::NAN)]),
        )
        .unwrap_err();
    assert!(matches!(non_finite, EditError::InvalidAdjustment(_)));
    assert_eq!(session.snapshot().unwrap(), before);
}

#[test]
fn snapshot_resolves_shape_colors_with_the_presentation_theme() {
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    package.themes[0].theme.color_scheme.accent1 = "123456".to_owned();
    package.themes[0].theme.color_scheme.accent2 = "ABCDEF".to_owned();
    let shape = package
        .slides
        .iter_mut()
        .flat_map(|slide| &mut slide.shapes)
        .find_map(|node| match node {
            ShapeNode::Shape(shape) => Some(shape),
            _ => None,
        })
        .unwrap();
    let source_id = shape.base.id;
    shape.fill = Some(ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            theme_color: Some("accent1".to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    });
    shape.outline = Some(ShapeOutline {
        color: Some(ColorValue {
            theme_color: Some("accent2".to_owned()),
            ..ColorValue::default()
        }),
        ..ShapeOutline::default()
    });

    let session = DeckSession::from_package(package, 706).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .find_map(|shape| find_shape(shape, source_id))
        .unwrap();

    assert_eq!(
        shape
            .fill
            .as_ref()
            .and_then(|fill| fill.color.as_ref())
            .and_then(|color| color.theme_color.as_deref()),
        Some("accent1")
    );
    assert_eq!(shape.resolved_fill_color.as_deref(), Some("#123456"));
    assert_eq!(shape.resolved_outline_color.as_deref(), Some("#ABCDEF"));
}

fn find_shape(shape: &ShapeSnapshot, source_id: u32) -> Option<&ShapeSnapshot> {
    if shape.source_id == source_id {
        return Some(shape);
    }
    shape
        .children
        .iter()
        .find_map(|child| find_shape(child, source_id))
}

fn assert_history_round_trip(session: &DeckSession, before: pptx_edit::DeckSnapshot) {
    let after = session.snapshot().unwrap();
    assert_ne!(after, before);
    session.add_undo_barrier();
    assert!(session.undo());
    assert_eq!(session.snapshot().unwrap(), before);
    assert!(session.redo());
    assert_eq!(session.snapshot().unwrap(), after);
    session.add_undo_barrier();
}

/// A chart-only deck has no text stories at all, so the schema roots have to
/// survive the hydrate round trip while they are still empty.
#[test]
fn a_deck_without_any_text_opens_and_accepts_edits() {
    const CHART_DECK: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/chart-deck.pptx");

    let session = DeckSession::open(CHART_DECK, 702).unwrap();
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.slides.len(), 2);
    assert!(
        snapshot.slides[0]
            .shapes
            .iter()
            .all(|shape| shape.text_stories.is_empty())
    );

    let slide_id = snapshot.slides[0].id.clone();
    let receipt = session
        .add_shape(
            &EditCtx::local("test"),
            &slide_id,
            &PresetShapeDraft {
                name: "Note".to_owned(),
                geometry: "rect".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 100_000,
                    width: 900_000,
                    height: 400_000,
                },
                fill: None,
            },
        )
        .unwrap();
    assert!(
        session.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .any(|shape| shape.id == receipt.shape_id)
    );
}
