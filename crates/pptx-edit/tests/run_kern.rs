use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/run-spacing.pptx");
const STORY: &str = "story:slide:0:256:shape:0:0";

fn kerned_deck() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    for (path, bytes) in &mut parts {
        if path == "ppt/slides/slide1.xml" {
            let xml = String::from_utf8(bytes.clone()).unwrap();
            let kerned = xml.replacen(
                r#"<a:rPr sz="3200" spc="600">"#,
                r#"<a:rPr sz="3200" spc="600" kern="0">"#,
                1,
            );
            assert_ne!(kerned, xml);
            *bytes = kerned.into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    String::from_utf8(xml.clone()).unwrap()
}

#[test]
fn a_runs_own_kern_threshold_reaches_the_snapshot_and_survives_an_edit() {
    let session = DeckSession::open(&kerned_deck(), 32508).unwrap();
    let kern = |session: &DeckSession| {
        session.story(STORY).unwrap().paragraphs[0].runs[0]
            .style
            .kern_pt
    };
    assert_eq!(kern(&session), Some(0.0));

    let context = EditCtx::local("test");
    let end = session.story(STORY).unwrap().length - 1;
    session
        .format_text(
            &context,
            STORY,
            0,
            end,
            &TextStylePatch {
                bold: Some(true),
                ..Default::default()
            },
        )
        .unwrap();
    let saved = session.save().unwrap();
    assert!(slide_xml(&saved).contains(r#"kern="0""#));
    let reopened = DeckSession::open(&saved, 32509).unwrap();
    assert_eq!(kern(&reopened), Some(0.0));
}

#[test]
fn an_inserted_runs_kern_threshold_must_be_a_size() {
    let session = DeckSession::open(DECK, 32510).unwrap();
    let context = EditCtx::local("test");
    let before = session.encode_state_as_update_v1();
    for kern in [f64::NAN, f64::INFINITY, -1.0, 4000.01] {
        assert!(
            session
                .insert_text(
                    &context,
                    STORY,
                    0,
                    "X",
                    &TextStyle {
                        kern_pt: Some(kern),
                        ..Default::default()
                    }
                )
                .is_err(),
            "{kern}"
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}
