use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, ShapeSnapshot, ShapeStroke};
use yrs::{Map, MapRef, ReadTxn, Transact};

const SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2-connectors.pptx");
const V2_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2-connectors.update.bin");
const V2_MOVED_UPDATE: &[u8] =
    include_bytes!("fixtures/deck-schema-v2-connectors-moved.update.bin");
const SLIDE: &str = "ppt/slides/slide1.xml";

fn context() -> EditCtx {
    EditCtx::local("connectors")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(parts(bytes).remove(SLIDE).unwrap()).unwrap()
}

fn element<'a>(xml: &'a str, open: &str, close: &str) -> &'a str {
    let start = xml.find(open).unwrap_or_else(|| panic!("missing {open}"));
    let end = start + xml[start..].find(close).unwrap() + close.len();
    &xml[start..end]
}

fn connector(xml: &str) -> &str {
    element(xml, "<p:cxnSp>", "</p:cxnSp>")
}

fn shape(xml: &str, id: u32) -> &str {
    element(
        xml,
        &format!("<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\""),
        "</p:sp>",
    )
}

fn shape_named(session: &DeckSession, name: &str) -> (String, ShapeSnapshot) {
    let slide = session.snapshot().unwrap().slides.remove(0);
    let shape = slide.shapes.into_iter().find(|shape| shape.name == name);
    (
        slide.id,
        shape.unwrap_or_else(|| panic!("no shape named {name}")),
    )
}

#[test]
fn a_fresh_deck_models_the_connector_and_saves_byte_for_byte() {
    let session = DeckSession::open(SOURCE, 20).unwrap();
    let names: Vec<String> = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .map(|shape| shape.name.clone())
        .collect();
    assert_eq!(names, ["Before", "Straight Arrow Connector 2", "After"]);
    assert_eq!(parts(&session.save().unwrap()), parts(SOURCE));
}

#[test]
fn an_edit_after_a_connector_reaches_the_shape_after_it() {
    let session = DeckSession::open(SOURCE, 21).unwrap();
    let (slide_id, after) = shape_named(&session, "After");
    session
        .move_shape(&context(), &slide_id, &after.id, 952_500, 1_047_750)
        .unwrap();

    let source = slide_xml(SOURCE);
    let saved = slide_xml(&session.save().unwrap());
    assert!(shape(&saved, 4).contains(r#"<a:off x="952500" y="1047750"/>"#));
    assert_eq!(connector(&saved), connector(&source));
    assert_eq!(shape(&saved, 2), shape(&source, 2));
}

#[test]
fn a_connector_edit_keeps_its_joins() {
    let session = DeckSession::open(SOURCE, 22).unwrap();
    let (slide_id, connector_shape) = shape_named(&session, "Straight Arrow Connector 2");
    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &connector_shape.id,
            &ShapeStroke {
                color: Some("#0000FF".to_owned()),
                width_pt: Some(3.0),
            },
        )
        .unwrap();

    let saved = slide_xml(&session.save().unwrap());
    let written = connector(&saved);
    assert!(written.starts_with(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Straight Arrow Connector 2"/><p:cNvCxnSpPr><a:stCxn id="2" idx="1"/><a:endCxn id="4" idx="2"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>"#
    ));
    assert!(written.contains(r#"<a:srgbClr val="0000FF"/>"#));
    assert_eq!(shape(&saved, 4), shape(&slide_xml(SOURCE), 4));
}

#[test]
fn an_untouched_pre_connector_update_saves_its_source_byte_for_byte() {
    let session = DeckSession::open_from_update_with_source(V2_UPDATE, SOURCE, 23).unwrap();
    assert_eq!(slide_xml(&session.save().unwrap()), slide_xml(SOURCE));
    assert_eq!(parts(&session.save().unwrap()), parts(SOURCE));
}

#[test]
fn a_pre_connector_update_edits_the_shape_after_the_connector_not_the_connector() {
    let session = DeckSession::open_from_update_with_source(V2_MOVED_UPDATE, SOURCE, 24).unwrap();

    let source = slide_xml(SOURCE);
    let saved = slide_xml(&session.save().unwrap());
    assert_eq!(connector(&saved), connector(&source));
    assert_eq!(shape(&saved, 2), shape(&source, 2));
    assert!(
        saved.contains(r#"<p:cNvPr id="4" name="After"/>"#),
        "{saved}"
    );
    assert!(shape(&saved, 4).contains(r#"<a:off x="952500" y="1047750"/>"#));
    assert_eq!(saved.matches("<p:sp>").count(), 2);
    assert_eq!(saved.matches("<p:cxnSp>").count(), 1);
}

#[test]
fn a_pre_connector_update_keeps_numbering_shapes_without_the_connector() {
    let session = DeckSession::open_from_update_with_source(V2_MOVED_UPDATE, SOURCE, 25).unwrap();
    let (slide_id, after) = shape_named(&session, "After");
    assert_eq!(after.id, format!("{slide_id}:shape:1"));
    assert_eq!(session.snapshot().unwrap().slides[0].shapes.len(), 2);

    session
        .set_shape_stroke(
            &context(),
            &slide_id,
            &after.id,
            &ShapeStroke {
                color: Some("#00FF00".to_owned()),
                width_pt: Some(1.0),
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    assert!(shape(&saved, 4).contains(r#"<a:srgbClr val="00FF00"/>"#));
    assert_eq!(connector(&saved), connector(&slide_xml(SOURCE)));
}

#[test]
fn legacy_nested_and_trailing_connectors_survive_save_and_child_edits() {
    let source = include_bytes!("fixtures/deck-schema-v2-nested-connectors.pptx");
    let update = include_bytes!("fixtures/deck-schema-v2-nested-connectors.update.bin");
    let session = DeckSession::open_from_update_with_source(update, source, 26).unwrap();
    assert_eq!(parts(&session.save().unwrap()), parts(source));
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let after = &slide.shapes[0].children[1];
    assert_eq!(after.source_id, 4);
    assert_eq!(after.id, "slide:0:256:shape:0.1");
    {
        let mut txn = session.yrs_doc().transact_mut();
        let shapes = txn.get_map("pptx:shapes").unwrap();
        let shape = shapes
            .get(&txn, &after.id)
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        shape.insert(&mut txn, "x", 952_500.0);
        shape.insert(&mut txn, "y", 1_047_750.0);
    }
    let saved = session.save().unwrap();
    let xml = slide_xml(&saved);
    assert_eq!(connector(&xml), connector(&slide_xml(source)));
    assert!(shape(&xml, 4).contains(r#"<a:off x="952500" y="1047750"/>"#));
    assert_eq!(xml.matches("<p:cxnSp>").count(), 3);
    let fresh = DeckSession::open(&saved, 27).unwrap();
    let snapshot = fresh.snapshot().unwrap();
    assert_eq!(snapshot.slides[0].shapes.len(), 2);
    assert_eq!(snapshot.slides[0].shapes[0].children.len(), 4);
    assert_eq!(snapshot.slides[0].shapes[0].children[3].children.len(), 1);
}
