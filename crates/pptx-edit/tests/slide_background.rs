//! Slide background writes: set, replace and clear across save and reopen.

use std::collections::BTreeMap;

use pptx_edit::{
    ColorValue, DeckSession, EditCtx, EditError, EditFailureCode, EditRequest, GradientFill,
    GradientStop, ShapeFill, ShapeKind,
};
use serde_json::json;
use yrs::updates::decoder::Decode;
use yrs::{Doc, Map, MapRef, ReadTxn, Transact, Update};

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
    angled(90.0, stops)
}

fn angled(angle: f64, stops: Vec<GradientStop>) -> ShapeFill {
    ShapeFill {
        fill_type: "gradient".to_owned(),
        color: None,
        gradient: Some(GradientFill {
            gradient_type: "linear".to_owned(),
            angle: Some(angle),
            stops,
        }),
    }
}

fn scheme(slot: &str) -> ShapeFill {
    ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            theme_color: Some(slot.to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    }
}

/// `FIXTURE` with a pattern fill, which the model cannot hold, as the first slide's background.
fn with_pattern_background() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let source = String::from_utf8(xml.clone()).unwrap();
    let solid = r#"<a:solidFill><a:srgbClr val="F8F7F2"/></a:solidFill>"#;
    assert!(source.contains(solid));
    *xml = source
        .replace(
            solid,
            r#"<a:pattFill prst="pct20"><a:fgClr><a:srgbClr val="0F1720"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>"#,
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A peer's update that stores `json` as the slide's background.
fn remote_background(session: &DeckSession, client: u64, slide_id: &str, json: &str) -> Vec<u8> {
    let doc = Doc::with_client_id(client);
    doc.transact_mut()
        .apply_update(Update::decode_v1(&session.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let mut txn = doc.transact_mut();
    let slide = txn
        .get_map("pptx:slides")
        .unwrap()
        .get(&txn, slide_id)
        .unwrap()
        .cast::<MapRef>()
        .unwrap();
    slide.insert(&mut txn, "backgroundJson", json);
    txn.encode_update_v1()
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
        angled(-1.0, vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")]),
        angled(360.0, vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")]),
        angled(
            359.999_999_9,
            vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")],
        ),
        angled(400.0, vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")]),
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

#[test]
fn every_scheme_slot_is_a_settable_background_colour() {
    for (client, slot) in [
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "background1",
        "text2",
        "accent6",
        "folHlink",
    ]
    .into_iter()
    .enumerate()
    {
        let session = DeckSession::open(FIXTURE, 40 + client as u64 * 2).unwrap();
        let slide_id = session.snapshot().unwrap().slides[0].id.clone();
        session
            .set_slide_background(&context(), &slide_id, Some(&scheme(slot)))
            .unwrap();
        let reopened = DeckSession::open(&session.save().unwrap(), 41 + client as u64 * 2).unwrap();
        assert_eq!(background_of(&reopened, 0), Some(scheme(slot)), "{slot}");
    }
}

#[test]
fn a_gradient_angle_on_the_last_degree_round_trips() {
    let session = DeckSession::open(FIXTURE, 60).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let fill = angled(359.5, vec![stop(0.0, "0F1720"), stop(100_000.0, "0F766E")]);
    session
        .set_slide_background(&context(), &slide_id, Some(&fill))
        .unwrap();
    let saved = session.save().unwrap();
    assert!(part_text(&saved, &slide_part(&session, 0)).contains(r#"ang="21570000""#));
    assert_eq!(
        background_of(&DeckSession::open(&saved, 61).unwrap(), 0),
        Some(fill)
    );
}

#[test]
fn a_remote_background_the_writer_cannot_serialize_falls_back_to_the_source() {
    let session = DeckSession::open(FIXTURE, 62).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let before = background_of(&session, 0);
    let declared = part_text(&session.save().unwrap(), &slide_part(&session, 0));
    for (client, json) in [
        r#"{"type":"theme","color":{"themeColor":"accent1"}}"#,
        r#"{"type":"solid","color":{"rgb":"nothex"}}"#,
        r#"{"type":"solid","color":{"themeColor":"not-a-slot"}}"#,
        r#"{"type":"gradient","gradient":{"type":"linear","angle":-90,"stops":[{"position":0,"color":{"rgb":"0F1720"}},{"position":100000,"color":{"rgb":"0F766E"}}]}}"#,
    ]
    .into_iter()
    .enumerate()
    {
        let update = remote_background(&session, 70 + client as u64, &slide_id, json);
        session.apply_update_v1(&update).unwrap();
        assert_eq!(background_of(&session, 0), before, "{json}");
        let saved = session.save().unwrap();
        assert_eq!(part_text(&saved, &slide_part(&session, 0)), declared, "{json}");
    }
}

#[test]
fn clearing_a_background_the_model_cannot_hold_removes_the_declaration() {
    let source = with_pattern_background();
    let direct = DeckSession::open(&source, 63).unwrap();
    let slide_id = direct.snapshot().unwrap().slides[0].id.clone();
    assert_eq!(background_of(&direct, 0), None);
    assert!(part_text(&direct.save().unwrap(), &slide_part(&direct, 0)).contains("<a:pattFill"));
    direct
        .set_slide_background(&context(), &slide_id, None)
        .unwrap();
    let saved = direct.save().unwrap();
    assert!(!part_text(&saved, &slide_part(&direct, 0)).contains("<p:bg>"));

    let batched = DeckSession::open(&source, 64).unwrap();
    let request: EditRequest = serde_json::from_value(json!({
        "expectVersion": batched.version(),
        "steps": [{
            "op": "setSlideBackground",
            "target": {"slideId": slide_id},
            "background": null,
            "expect": {"background": null}
        }]
    }))
    .unwrap();
    let application = batched.apply_edits(&request).unwrap().unwrap();
    assert!(application.applied);
    assert_eq!(application.changed_slides, vec![slide_id.clone()]);
    let saved = batched.save().unwrap();
    assert!(!part_text(&saved, &slide_part(&batched, 0)).contains("<p:bg>"));
}
