use pptx_edit::DeckSession;
use pptx_parse::ShapeNode;

const FIXTURE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/outer-shadow.pptx");
const V10: &[u8] = include_bytes!("fixtures/outer-shadow-main-v10.update.bin");
const V17: &[u8] = include_bytes!("fixtures/outer-shadow-main-v17.update.bin");

#[test]
fn main_shadows_survive_source_attachment_and_subsequent_updates() {
    for update in [V10, V17] {
        check_source_attachment(update);
    }
}

fn check_source_attachment(update: &[u8]) {
    let old = DeckSession::open_from_update(update, 334).unwrap();
    let ShapeNode::Shape(shape) = &old.package().slides[0].shapes[0] else {
        panic!()
    };
    assert!(shape.effects.is_none());
    let attached = DeckSession::open_from_update_with_source(update, FIXTURE, 335).unwrap();
    let reopened =
        DeckSession::open_from_update(&attached.encode_state_as_update_v1(), 336).unwrap();
    assert_eq!(attached.snapshot().unwrap(), reopened.snapshot().unwrap());
    for index in 0..2 {
        let ShapeNode::Shape(shape) = &reopened.package().slides[0].shapes[index] else {
            panic!()
        };
        let shadow = shape
            .effects
            .as_ref()
            .unwrap()
            .outer_shadow
            .as_ref()
            .unwrap();
        assert_eq!(shadow.blur_radius, 76_200);
        assert_eq!(shadow.color.as_ref().unwrap().alpha, Some(0.4));
    }
    assert_eq!(
        ooxml_opc::unzip_parts(&attached.save().unwrap()).unwrap(),
        ooxml_opc::unzip_parts(FIXTURE).unwrap(),
    );
}

#[test]
fn scaled_shadows_and_bitmap_effects_survive_main_updates_and_edited_saves() {
    use pptx_edit::EditCtx;
    for (update, source) in [
        (
            include_bytes!("fixtures/outer-shadow-scale-main-v17.update.bin").as_slice(),
            include_bytes!("../../pptx-render/tests/fixtures/outer-shadow-scale.pptx").as_slice(),
        ),
        (
            include_bytes!("fixtures/blip-shadow-main-v17.update.bin").as_slice(),
            include_bytes!("fixtures/blip-shadow.pptx").as_slice(),
        ),
    ] {
        let fresh = DeckSession::open(source, 33410).unwrap();
        let deferred = DeckSession::open_from_update(update, 33411).unwrap();
        let attached = DeckSession::open_from_update_with_source(
            &deferred.encode_state_as_update_v1(),
            source,
            33412,
        )
        .unwrap();
        assert_eq!(attached.package(), fresh.package());
        assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
        let saved = attached.save().unwrap();
        assert_eq!(
            ooxml_opc::unzip_parts(&saved).unwrap(),
            ooxml_opc::unzip_parts(source).unwrap()
        );
        let current = attached.encode_state_as_update_v1();
        let reopened = DeckSession::open_from_update(&current, 33413).unwrap();
        assert_eq!(
            serde_json::to_vec(reopened.package()).unwrap(),
            serde_json::to_vec(fresh.package()).unwrap()
        );
        assert_eq!(reopened.encode_state_as_update_v1(), current);
        let reattached =
            DeckSession::open_from_update_with_source(&current, source, 33414).unwrap();
        assert_eq!(reattached.encode_state_as_update_v1(), current);
        let before = attached.snapshot().unwrap();
        let slide = &before.slides[0];
        attached
            .move_shape(
                &EditCtx::local("test"),
                &slide.id,
                &slide.shapes[0].id,
                952500,
                1905000,
            )
            .unwrap();
        let saved = attached.save().unwrap();
        let reopened = DeckSession::open(&saved, 33415).unwrap();
        let after = reopened.snapshot().unwrap();
        assert_eq!(
            (after.slides[0].shapes[0].x, after.slides[0].shapes[0].y),
            (952500, 1905000)
        );
        for (expected, actual) in fresh.package().slides[0]
            .shapes
            .iter()
            .zip(&reopened.package().slides[0].shapes)
        {
            match (expected, actual) {
                (ShapeNode::Shape(expected), ShapeNode::Shape(actual)) => {
                    assert_eq!(actual.effects, expected.effects)
                }
                (ShapeNode::Picture(expected), ShapeNode::Picture(actual)) => {
                    assert_eq!(actual.effects, expected.effects);
                    assert_eq!(actual.shape_effects, expected.shape_effects);
                }
                _ => panic!("unexpected fixture shape"),
            }
        }
    }
}
