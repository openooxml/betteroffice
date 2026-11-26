use std::collections::BTreeMap;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use ooxml_drawingml::{ColorValue, ShapeFill, ShapeOutline};
use pptx_edit::{
    DeckSession, EditCtx, EditError, PresetShapeDraft, ProposalEdit, ProposalRequest, ShapeDraft,
    ShapeRect, ShapeSnapshot, ShapeStroke, TextAutofitMode, TextInsets, TextStyle,
};
use pptx_parse::{ShapeNode, TextAutofit};

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
fn set_shape_rect_is_one_update_and_one_undo_step() {
    let session = DeckSession::open(FIXTURE, 707).unwrap();
    let context = EditCtx::local("test");
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let before = ShapeRect {
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
    };
    let after = ShapeRect {
        x: before.x + 120_000,
        y: before.y + 80_000,
        width: before.width + 300_000,
        height: before.height + 200_000,
    };
    let updates = Arc::new(AtomicUsize::new(0));
    let observed = Arc::clone(&updates);
    let _subscription = session
        .observe_update_v1(move |_| {
            observed.fetch_add(1, Ordering::Relaxed);
        })
        .unwrap();

    session.add_undo_barrier();
    let receipt = session
        .set_shape_rect(&context, &slide.id, &shape.id, after)
        .unwrap();
    assert_eq!((receipt.before, receipt.after), (before, after));
    assert_eq!(updates.load(Ordering::Relaxed), 1);
    let resized = &session.snapshot().unwrap().slides[0].shapes[0];
    assert_eq!(
        (resized.x, resized.y, resized.width, resized.height),
        (after.x, after.y, after.width, after.height)
    );

    session.add_undo_barrier();
    assert!(session.undo());
    let restored = &session.snapshot().unwrap().slides[0].shapes[0];
    assert_eq!(
        (restored.x, restored.y, restored.width, restored.height),
        (before.x, before.y, before.width, before.height)
    );
    assert!(!session.can_undo());
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

#[test]
fn comments_coexist_with_chart_and_hyperlink_relationships() {
    const CHART_DECK: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/chart-deck.pptx");

    let session = DeckSession::open(CHART_DECK, 703).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_comment(
            &context,
            &slide_id,
            "Ada Lovelace",
            "AL",
            "Is this series right?",
            "2026-09-01T10:00:00.000",
            914_400,
            457_200,
        )
        .unwrap();

    let saved = session.save().unwrap();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let rels = String::from_utf8(parts["ppt/slides/_rels/slide1.xml.rels"].clone()).unwrap();
    assert!(
        rels.contains("../charts/chart1.xml"),
        "chart rel was dropped"
    );
    assert!(
        rels.contains("https://example.invalid"),
        "external hyperlink rel was dropped"
    );
    assert!(rels.contains("../comments/comment1.xml"));
    assert!(
        rels.contains("Id=\"rId4\""),
        "the comment must take a fresh id rather than reuse one: {rels}"
    );
    assert_eq!(
        parts["ppt/charts/chart1.xml"],
        chart_deck_part("ppt/charts/chart1.xml")
    );

    let reopened = DeckSession::open(&saved, 704).unwrap();
    assert_eq!(reopened.snapshot().unwrap().comments.len(), 1);
    assert_eq!(reopened.package().charts.len(), 2);

    let first = reopened.snapshot().unwrap().slides[0].id.clone();
    reopened.delete_slide(&context, &first).unwrap();
    let pruned: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&reopened.save().unwrap())
        .unwrap()
        .into_iter()
        .collect();
    assert!(!pruned.contains_key("ppt/comments/comment1.xml"));
    assert!(!pruned.contains_key("ppt/charts/chart1.xml"));
    assert!(pruned.contains_key("ppt/charts/chart2.xml"));
    let content_types = String::from_utf8(pruned["[Content_Types].xml"].clone()).unwrap();
    assert!(!content_types.contains("comments/comment1.xml"));
    assert!(!content_types.contains("charts/chart1.xml"));
}

fn chart_deck_part(path: &str) -> Vec<u8> {
    const CHART_DECK: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/chart-deck.pptx");
    ooxml_opc::unzip_parts(CHART_DECK)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| bytes)
        .expect("part is in the fixture")
}

#[test]
fn layered_presets_can_be_inserted() {
    let session = DeckSession::open(FIXTURE, 644).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    for geometry in ["arc", "cube", "leftBrace", "rightBrace", "ribbon2"] {
        let receipt = session
            .add_shape(
                &EditCtx::local("test"),
                &slide_id,
                &PresetShapeDraft {
                    name: geometry.to_owned(),
                    geometry: geometry.to_owned(),
                    rect: ShapeRect {
                        x: 0,
                        y: 0,
                        width: 952_500,
                        height: 952_500,
                    },
                    fill: Some("#DCE9F7".to_owned()),
                },
            )
            .unwrap();
        assert_eq!(
            session.snapshot().unwrap().slides[0]
                .shapes
                .iter()
                .find(|shape| shape.id == receipt.shape_id)
                .unwrap()
                .geometry,
            geometry
        );
    }
}

#[test]
fn text_box_autofit_and_insets_validate_roundtrip_and_sync() {
    let session = DeckSession::open(FIXTURE, 741).unwrap();
    let peer = DeckSession::open(FIXTURE, 742).unwrap();
    let context = EditCtx::local("test");
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();

    assert!(
        session
            .add_text_box(
                &context,
                &slide_id,
                &ShapeDraft {
                    name: "NegativeInset".to_owned(),
                    rect: ShapeRect {
                        x: 0,
                        y: 0,
                        width: 1_000_000,
                        height: 500_000,
                    },
                    text: "Hi".to_owned(),
                    style: TextStyle::default(),
                    autofit: None,
                    text_insets: Some(TextInsets {
                        left: Some(-1),
                        ..TextInsets::default()
                    }),
                },
            )
            .is_err()
    );

    let zero_insets = TextInsets {
        left: Some(0),
        top: Some(0),
        right: Some(0),
        bottom: Some(0),
    };
    let hero = session
        .add_text_box(
            &context,
            &slide_id,
            &ShapeDraft {
                name: "Hero Metric".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 100_000,
                    width: 3_800_000,
                    height: 914_400,
                },
                text: "62%".to_owned(),
                style: TextStyle {
                    font_size_pt: Some(88.0),
                    ..TextStyle::default()
                },
                autofit: Some(TextAutofitMode::Normal),
                text_insets: Some(zero_insets),
            },
        )
        .unwrap();

    let horizontal_insets = TextInsets {
        left: Some(19_050),
        top: None,
        right: Some(19_050),
        bottom: None,
    };
    let fixed = session
        .add_text_box(
            &context,
            &slide_id,
            &ShapeDraft {
                name: "Fixed Label".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 1_200_000,
                    width: 600_000,
                    height: 300_000,
                },
                text: "02".to_owned(),
                style: TextStyle {
                    font_size_pt: Some(12.0),
                    ..TextStyle::default()
                },
                autofit: Some(TextAutofitMode::None),
                text_insets: Some(horizontal_insets),
            },
        )
        .unwrap();

    let snap = session.snapshot().unwrap();
    let hero_snap = snap.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == hero.shape_id)
        .unwrap();
    assert_eq!(hero_snap.autofit, Some(TextAutofitMode::Normal));
    assert_eq!(hero_snap.text_insets, Some(zero_insets));

    let fixed_snap = snap.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id == fixed.shape_id)
        .unwrap();
    assert_eq!(fixed_snap.autofit, Some(TextAutofitMode::None));
    assert_eq!(fixed_snap.text_insets, Some(horizontal_insets));

    let update = session.encode_state_as_update_v1();
    peer.apply_update_v1(&update).unwrap();
    let peer_snap = peer.snapshot().unwrap();
    assert_eq!(
        peer_snap.slides[0]
            .shapes
            .iter()
            .find(|shape| shape.id == hero.shape_id)
            .unwrap(),
        hero_snap
    );
    assert_eq!(
        peer_snap.slides[0]
            .shapes
            .iter()
            .find(|shape| shape.id == fixed.shape_id)
            .unwrap(),
        fixed_snap
    );

    let proposal = session
        .propose(ProposalRequest {
            agent_id: "layout-agent".to_owned(),
            note: Some("Add autofit box".to_owned()),
            edits: vec![ProposalEdit::AddTextBox {
                reference: Some("box".to_owned()),
                slide_id: slide_id.clone(),
                draft: ShapeDraft {
                    name: "Proposed Box".to_owned(),
                    rect: ShapeRect {
                        x: 200_000,
                        y: 1_600_000,
                        width: 2_000_000,
                        height: 400_000,
                    },
                    text: "Proposed".to_owned(),
                    style: TextStyle::default(),
                    autofit: Some(TextAutofitMode::Normal),
                    text_insets: Some(zero_insets),
                },
            }],
        })
        .unwrap();
    let accepted = session.accept_proposal(&proposal.id, false).unwrap();
    assert!(accepted.applied);
    let proposed_snap = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Proposed Box")
        .cloned()
        .unwrap();
    assert_eq!(proposed_snap.autofit, Some(TextAutofitMode::Normal));
    assert_eq!(proposed_snap.text_insets, Some(zero_insets));

    let saved = session.save().unwrap();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let slide_xml = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    assert!(
        slide_xml.contains(
            r#"<a:bodyPr bIns="0" lIns="0" rIns="0" tIns="0"><a:normAutofit/></a:bodyPr>"#
        ),
        "{slide_xml}"
    );
    assert!(
        slide_xml.contains(r#"<a:bodyPr lIns="19050" rIns="19050"><a:noAutofit/></a:bodyPr>"#),
        "{slide_xml}"
    );

    let parsed = pptx_parse::parse_pptx(&saved).unwrap();
    let hero_body = parsed.slides[0]
        .shapes
        .iter()
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.base.name == "Hero Metric" => shape.text.as_ref(),
            _ => None,
        })
        .unwrap();
    assert_eq!(hero_body.inset_left, Some(0));
    assert_eq!(hero_body.inset_top, Some(0));
    assert_eq!(hero_body.inset_right, Some(0));
    assert_eq!(hero_body.inset_bottom, Some(0));
    assert_eq!(
        hero_body.autofit,
        Some(TextAutofit::Normal {
            font_scale: None,
            line_space_reduction: None,
        })
    );

    let fixed_body = parsed.slides[0]
        .shapes
        .iter()
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.base.name == "Fixed Label" => shape.text.as_ref(),
            _ => None,
        })
        .unwrap();
    assert_eq!(fixed_body.inset_left, Some(19_050));
    assert_eq!(fixed_body.inset_top, None);
    assert_eq!(fixed_body.inset_right, Some(19_050));
    assert_eq!(fixed_body.inset_bottom, None);
    assert_eq!(fixed_body.autofit, Some(TextAutofit::None));
}
