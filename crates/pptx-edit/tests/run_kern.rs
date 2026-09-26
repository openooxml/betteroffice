use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};
use yrs::types::Attrs;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, Text, TextRef, Transact, Update};

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

fn run_kern(session: &DeckSession) -> Option<f64> {
    session.story(STORY).unwrap().paragraphs[0].runs[0]
        .style
        .kern_pt
}

fn story(txn: &yrs::TransactionMut<'_>) -> TextRef {
    txn.get_map("pptx:stories")
        .unwrap()
        .get(txn, STORY)
        .unwrap()
        .cast::<TextRef>()
        .unwrap()
}

fn strip_kern(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(object) => {
            object.remove("kernPt");
            object.values_mut().for_each(strip_kern);
        }
        serde_json::Value::Array(array) => array.iter_mut().for_each(strip_kern),
        _ => {}
    }
}

/// The deck as a session stored before runs carried `kern`.
fn stored_without_kern(deck: &[u8]) -> Vec<u8> {
    let update = DeckSession::open(deck, 32513)
        .unwrap()
        .encode_state_as_update_v1();
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let mut txn = doc.transact_mut();
    let meta = txn.get_map("pptx:meta").unwrap();
    assert!(meta.remove(&mut txn, "kernSeeded").is_some());
    let Some(Out::Any(Any::Buffer(bytes))) = meta.get(&txn, "packageJson") else {
        panic!("expected a packageJson buffer");
    };
    let mut package: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
    strip_kern(&mut package);
    meta.insert(
        &mut txn,
        "packageJson",
        Any::Buffer(std::sync::Arc::from(serde_json::to_vec(&package).unwrap())),
    );
    let story = story(&txn);
    let length = story.len(&txn);
    story.format(
        &mut txn,
        0,
        length,
        Attrs::from([("kern".into(), Any::Null)]),
    );
    drop(txn);
    doc.transact()
        .encode_state_as_update_v1(&Default::default())
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
    assert_eq!(run_kern(&session), Some(0.0));

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
    assert_eq!(run_kern(&reopened), Some(0.0));
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

#[test]
fn a_remote_kern_outside_the_schema_range_is_rejected() {
    let session = DeckSession::open(DECK, 32511).unwrap();
    let before = session.encode_state_as_update_v1();
    let remote = |client: u64, kern: Any| {
        let doc = Doc::with_client_id(client);
        doc.transact_mut()
            .apply_update(Update::decode_v1(&before).unwrap())
            .unwrap();
        let mut txn = doc.transact_mut();
        story(&txn).format(&mut txn, 0, 3, Attrs::from([("kern".into(), kern)]));
        txn.encode_update_v1()
    };
    for (client, kern) in [
        Any::Number(-1.0),
        Any::Number(4_000.01),
        Any::String("12".into()),
    ]
    .into_iter()
    .enumerate()
    {
        let update = remote(32520 + client as u64, kern.clone());
        assert!(session.apply_update_v1(&update).is_err(), "{kern:?}");
        assert_eq!(session.encode_state_as_update_v1(), before);
    }
    session
        .apply_update_v1(&remote(32530, Any::Number(f64::NAN)))
        .unwrap();
    assert_eq!(run_kern(&session), None, "yrs drops a NaN format");
    session
        .apply_update_v1(&remote(32531, Any::Number(12.0)))
        .unwrap();
    assert_eq!(run_kern(&session), Some(12.0));
}

#[test]
fn a_session_stored_before_kern_recovers_it_from_its_source() {
    let deck = kerned_deck();
    let stored = stored_without_kern(&deck);
    let detached = DeckSession::open_from_update(&stored, 32514).unwrap();
    assert_eq!(run_kern(&detached), None);

    let reattached = DeckSession::open_from_update_with_source(&stored, &deck, 32515).unwrap();
    assert_eq!(run_kern(&reattached), Some(0.0));
    assert_eq!(slide_xml(&reattached.save().unwrap()), slide_xml(&deck));

    let plain = stored_without_kern(DECK);
    let reattached = DeckSession::open_from_update_with_source(&plain, DECK, 32516).unwrap();
    assert_eq!(run_kern(&reattached), None);
    assert_eq!(slide_xml(&reattached.save().unwrap()), slide_xml(DECK));
    assert_eq!(
        reattached.encode_state_as_update_v1(),
        DeckSession::open_from_update(&plain, 32517)
            .unwrap()
            .encode_state_as_update_v1(),
        "a source with no kern leaves the stored session untouched"
    );
}
