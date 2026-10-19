//! Slide background writes: set, replace and clear across save and reopen.

use std::collections::BTreeMap;

use pptx_edit::{
    ColorValue, DeckSession, EditCtx, EditError, EditFailureCode, EditRequest, GradientFill,
    GradientStop, ShapeFill, ShapeKind,
};
use serde_json::json;

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn context() -> EditCtx {
    EditCtx::local("background-test")
}

fn solid(color: &str) -> ShapeFill {
    ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            rgb: Some(color.to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    }
}

fn stop(position: f64, color: &str) -> GradientStop {
    GradientStop {
        position,
        color: ColorValue {
            rgb: Some(color.to_owned()),
            ..ColorValue::default()
        },
    }
}

fn gradient() -> ShapeFill {
    gradient_with(vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")])
}

fn gradient_with(stops: Vec<GradientStop>) -> ShapeFill {
    ShapeFill {
        fill_type: "gradient".to_owned(),
        color: None,
        gradient: Some(GradientFill {
            gradient_type: "linear".to_owned(),
            angle: Some(90.0),
            stops,
        }),
    }
}

fn background_of(session: &DeckSession, index: usize) -> Option<ShapeFill> {
    session.snapshot().unwrap().slides[index].background.clone()
}

fn part_text(bytes: &[u8], path: &str) -> String {
    let parts: BTreeMap<String, Vec<u8>> =
        ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect();
    String::from_utf8(parts[path].clone()).unwrap()
}

fn slide_part(session: &DeckSession, index: usize) -> String {
    session.snapshot().unwrap().slides[index]
        .source_part_path
        .clone()
        .unwrap()
}

#[test]
fn a_solid_background_round_trips() {
    let session = DeckSession::open(FIXTURE, 7).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before = background_of(&session, 0);

    let receipt = session
        .set_slide_background(&context(), &slide_id, Some(&solid("0F1720")))
        .unwrap();
    assert_eq!(receipt.before, before);
    assert_eq!(receipt.after, Some(solid("0F1720")));

    let saved = session.save().unwrap();
    let slide = part_text(&saved, &slide_part(&session, 0));
    let common = slide.find("<p:cSld>").unwrap();
    let bg = slide.find("<p:bg>").unwrap();
    let tree = slide.find("<p:spTree").unwrap();
    assert!(common < bg && bg < tree, "{slide}");

    let reopened = DeckSession::open(&saved, 8).unwrap();
    assert_eq!(background_of(&reopened, 0), Some(solid("0F1720")));
}

#[test]
fn a_gradient_background_round_trips() {
    let session = DeckSession::open(FIXTURE, 9).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();

    session
        .set_slide_background(&context(), &slide_id, Some(&gradient()))
        .unwrap();
    let saved = session.save().unwrap();
    let slide = part_text(&saved, &slide_part(&session, 0));
    assert!(slide.contains("<a:gradFill>"), "{slide}");
    assert_eq!(slide.matches("<a:gs ").count(), 2, "{slide}");

    let reopened = DeckSession::open(&saved, 10).unwrap();
    assert_eq!(background_of(&reopened, 0), Some(gradient()));
}

#[test]
fn clearing_a_declared_background_removes_the_declaration() {
    let first = DeckSession::open(FIXTURE, 11).unwrap();
    let slide_id = first.snapshot().unwrap().slides[0].id.clone();
    first
        .set_slide_background(&context(), &slide_id, Some(&solid("0F1720")))
        .unwrap();
    let with_background = first.save().unwrap();

    let second = DeckSession::open(&with_background, 12).unwrap();
    let receipt = second
        .set_slide_background(&context(), &slide_id, None)
        .unwrap();
    assert_eq!(receipt.before, Some(solid("0F1720")));
    assert_eq!(receipt.after, None);

    let saved = second.save().unwrap();
    assert!(!part_text(&saved, &slide_part(&second, 0)).contains("<p:bg>"),);
    let reopened = DeckSession::open(&saved, 13).unwrap();
    assert_eq!(background_of(&reopened, 0), None);
}

#[test]
fn a_declared_background_survives_an_unrelated_shape_edit() {
    let first = DeckSession::open(FIXTURE, 14).unwrap();
    let slide_id = first.snapshot().unwrap().slides[0].id.clone();
    first
        .set_slide_background(&context(), &slide_id, Some(&solid("0F1720")))
        .unwrap();
    let with_background = first.save().unwrap();

    let second = DeckSession::open(&with_background, 15).unwrap();
    let shape_id = second.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.kind == ShapeKind::Shape)
        .unwrap()
        .id
        .clone();
    second
        .set_shape_fill(&context(), &slide_id, &shape_id, Some("DC2626"))
        .unwrap();
    let saved = second.save().unwrap();
    let slide = part_text(&saved, &slide_part(&second, 0));
    assert!(slide.contains(r#"<a:srgbClr val="0F1720"/>"#), "{slide}");
    assert!(slide.contains(r#"<a:srgbClr val="DC2626"/>"#), "{slide}");
}

#[test]
fn the_batch_step_sets_and_clears_the_background() {
    let session = DeckSession::open(FIXTURE, 16).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": {"type": "solid", "color": {"rgb": "0F1720"}}
        }]
    }))
    .unwrap();
    let application = session.apply_edits(&request).unwrap().unwrap();
    assert!(application.applied);
    assert_eq!(application.changed_slides, vec![slide_id.clone()]);
    assert_eq!(background_of(&session, 0), Some(solid("0F1720")));

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": null
        }]
    }))
    .unwrap();
    let application = session.apply_edits(&request).unwrap().unwrap();
    assert!(application.applied);
    assert_eq!(background_of(&session, 0), None);

    let saved = session.save().unwrap();
    assert_eq!(
        DeckSession::open(&saved, 17)
            .unwrap()
            .snapshot()
            .unwrap()
            .slides[0]
            .background,
        None
    );
}

#[test]
fn invalid_backgrounds_are_refused_before_saving() {
    let session = DeckSession::open(FIXTURE, 18).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before = background_of(&session, 0);
    let cases = [
        ShapeFill {
            fill_type: "solid".to_owned(),
            color: None,
            gradient: None,
        },
        ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(ColorValue {
                auto: Some(true),
                ..ColorValue::default()
            }),
            gradient: None,
        },
        ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(ColorValue {
                theme_color: Some("not-a-scheme-colour".to_owned()),
                ..ColorValue::default()
            }),
            gradient: None,
        },
        solid("nothex"),
        gradient_with(Vec::new()),
        gradient_with(vec![stop(0.0, "0F1720")]),
        gradient_with(vec![stop(0.0, "0F1720"), stop(100_001.0, "0F766E")]),
        gradient_with(vec![
            stop(0.0, "0F1720"),
            GradientStop {
                position: 100_000.0,
                color: ColorValue::default(),
            },
        ]),
    ];
    for fill in cases {
        let error = session
            .set_slide_background(&context(), &slide_id, Some(&fill))
            .unwrap_err();
        assert!(matches!(error, EditError::InvalidFill(_)), "{error}");
        assert_eq!(background_of(&session, 0), before, "{fill:?}");
    }

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": {"type": "solid"}
        }]
    }))
    .unwrap();
    let refusal = session.apply_edits(&request).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::InvalidStep);
    assert_eq!(background_of(&session, 0), before);
}

#[test]
fn invalid_background_transforms_are_refused_without_mutating_state() {
    use ooxml_drawingml::ColorTransform::{self, *};

    let session = DeckSession::open(FIXTURE, 18).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before = background_of(&session, 0);
    let version = session.version();
    let constructors: [fn(f64) -> ColorTransform; 23] = [
        Tint, Shade, Hue, HueMod, HueOff, Sat, SatMod, SatOff, Lum, LumMod, LumOff, Red, RedMod,
        RedOff, Green, GreenMod, GreenOff, Blue, BlueMod, BlueOff, Alpha, AlphaMod, AlphaOff,
    ];
    let mut invalid = Vec::new();
    for constructor in constructors {
        invalid.extend([f64::NAN, f64::INFINITY, f64::NEG_INFINITY].map(constructor));
    }
    let ranged: [fn(f64) -> ColorTransform; 6] = [Tint, Shade, LumMod, LumOff, SatMod, Alpha];
    for constructor in ranged {
        invalid.extend([-0.01, 1.01].map(constructor));
    }
    for transform in invalid {
        let color = ColorValue {
            rgb: Some("112233".to_owned()),
            transforms: vec![transform],
            ..Default::default()
        };
        for fill in [
            ShapeFill {
                fill_type: "solid".to_owned(),
                color: Some(color.clone()),
                gradient: None,
            },
            gradient_with(vec![
                stop(0.0, "FFFFFF"),
                GradientStop {
                    position: 100_000.0,
                    color: color.clone(),
                },
            ]),
        ] {
            assert!(
                matches!(
                    session.set_slide_background(&context(), &slide_id, Some(&fill)),
                    Err(EditError::InvalidFill(_))
                ),
                "{transform:?}"
            );
            assert_eq!(background_of(&session, 0), before);
            assert_eq!(session.version(), version);
        }
    }
}

#[test]
fn the_batch_guard_checks_the_effective_background() {
    let session = DeckSession::open(FIXTURE, 19).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let expected = background_of(&session, 0);

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": {"type": "solid", "color": {"rgb": "0F1720"}},
            "expect": {"background": expected}
        }]
    }))
    .unwrap();
    assert!(session.apply_edits(&request).unwrap().unwrap().applied);

    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": session.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": {"type": "solid", "color": {"rgb": "DC2626"}},
            "expect": {"background": {"type": "solid", "color": {"rgb": "FFFFFF"}}}
        }]
    }))
    .unwrap();
    let refusal = session.apply_edits(&request).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::ContentMismatch);
    assert_eq!(background_of(&session, 0), Some(solid("0F1720")));
}

#[test]
fn a_duplicated_slide_keeps_its_background() {
    let session = DeckSession::open(FIXTURE, 20).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .set_slide_background(&context(), &slide_id, Some(&solid("0F1720")))
        .unwrap();
    session.duplicate_slide(&context(), &slide_id, 1).unwrap();

    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 21).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides[0].background, Some(solid("0F1720")));
    assert_eq!(snapshot.slides[1].background, Some(solid("0F1720")));
}
