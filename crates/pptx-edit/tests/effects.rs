use pptx_edit::DeckSession;
use pptx_parse::ShapeNode;

const FIXTURE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/outer-shadow.pptx");
const V10: &[u8] = include_bytes!("fixtures/outer-shadow-main-v10.update.bin");

#[test]
fn main_shadows_survive_source_attachment_and_subsequent_updates() {
    let old = DeckSession::open_from_update(V10, 334).unwrap();
    let ShapeNode::Shape(shape) = &old.package().slides[0].shapes[0] else {
        panic!()
    };
    assert!(shape.effects.is_none());
    let attached = DeckSession::open_from_update_with_source(V10, FIXTURE, 335).unwrap();
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
