use pptx_edit::test_counters::{self, AttachWork};
use pptx_edit::{DeckSession, EditCtx};
use yrs::updates::decoder::Decode;
use yrs::{Doc, Map, ReadTxn, StateVector, Transact, Update};

const DEFAULTS_SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2.1-defaults.pptx");
const DEFAULTS_SLIDE: &str = "ppt/slides/slide1.xml";
const META: &str = "pptx:meta";
const STORY_COUNTS: [usize; 3] = [40, 160, 640];
const DELETIONS_PER_STORY: u32 = 6;

fn hydrated(update: &[u8]) -> Doc {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    doc
}

fn restamped(update: &[u8], version: f64) -> Vec<u8> {
    let doc = hydrated(update);
    let meta = doc.transact().get_map(META).unwrap();
    meta.insert(&mut doc.transact_mut(), "schemaVersion", version);
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

fn without_keys(value: &mut serde_json::Value, keys: &[&str]) {
    match value {
        serde_json::Value::Object(map) => {
            for key in keys {
                map.remove(*key);
            }
            for value in map.values_mut() {
                without_keys(value, keys);
            }
        }
        serde_json::Value::Array(values) => values
            .iter_mut()
            .for_each(|value| without_keys(value, keys)),
        _ => {}
    }
}

fn legacy_seed(source: &[u8], client_id: u64) -> Vec<u8> {
    let mut value = serde_json::to_value(pptx_parse::parse_pptx(source).unwrap()).unwrap();
    without_keys(&mut value, &["caps", "colorMap", "colorMapOverride"]);
    let package = serde_json::from_value(value).unwrap();
    let legacy = DeckSession::from_package_with_source(package, source, client_id).unwrap();
    restamped(&restamped(&legacy.encode_state_as_update_v1(), 2.0), 2.1)
}

fn with_copies(count: usize) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(DEFAULTS_SOURCE).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == DEFAULTS_SLIDE)
        .unwrap();
    let slide = String::from_utf8(bytes.clone()).unwrap();
    let start = slide
        .find(r#"<p:sp><p:nvSpPr><p:cNvPr id="10" name="Direct all caps"/>"#)
        .unwrap();
    let end = start + slide[start..].find("</p:sp>").unwrap() + "</p:sp>".len();
    let template = &slide[start..end];
    let copies: String = (0..count)
        .map(|index| {
            template.replace(
                r#"id="10" name="Direct all caps""#,
                &format!(r#"id="{}" name="Copy {index}""#, 1000 + index),
            )
        })
        .collect();
    *bytes = slide
        .replacen("</p:spTree>", &format!("{copies}</p:spTree>"), 1)
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn attach_work(count: usize) -> (usize, AttachWork) {
    let source = with_copies(count);
    let doc = hydrated(&legacy_seed(&source, 44500));
    let fresh = DeckSession::open(&source, 44501).unwrap();
    let stories: Vec<_> = fresh.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .flat_map(|shape| shape.text_stories.iter().map(|story| story.id.clone()))
        .collect();
    {
        let mut txn = doc.transact_mut();
        let map = txn.get_map("pptx:stories").unwrap();
        for story in &stories {
            let text = map
                .get(&txn, story)
                .unwrap()
                .cast::<yrs::TextRef>()
                .unwrap();
            for index in (0..DELETIONS_PER_STORY).rev() {
                yrs::Text::remove_range(&text, &mut txn, index * 2, 1);
            }
        }
    }
    for story in &stories {
        for index in (0..DELETIONS_PER_STORY).rev() {
            fresh
                .delete_text(&EditCtx::local("fixture"), story, index * 2, index * 2 + 1)
                .unwrap();
        }
    }
    let update = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    test_counters::take();
    let attached = DeckSession::open_from_update_with_source(&update, &source, 44502).unwrap();
    let work = test_counters::take();
    for story in &stories {
        assert_eq!(attached.story(story).unwrap(), fresh.story(story).unwrap());
    }
    (stories.len(), work)
}

#[test]
fn migrating_and_attaching_reads_each_story_once_and_builds_snapshots_once() {
    let runs: Vec<_> = STORY_COUNTS.into_iter().map(attach_work).collect();
    println!("{runs:?}");
    let once = |work: &AttachWork| {
        (
            work.visible_builds,
            work.replay_seeds,
            work.seeded_snapshots,
            work.live_snapshots,
        )
    };
    let base = |copies: usize, reads: usize| reads.checked_sub(copies);
    let (_, first) = &runs[0];
    for (copies, (_, work)) in STORY_COUNTS.into_iter().zip(&runs) {
        assert_eq!(once(work), once(first), "{runs:?}");
        assert!(base(copies, work.story_reads).is_some(), "{runs:?}");
        assert_eq!(
            base(copies, work.story_reads),
            base(STORY_COUNTS[0], first.story_reads),
            "{runs:?}"
        );
        assert_eq!(
            base(copies, work.replayed_reads),
            base(STORY_COUNTS[0], first.replayed_reads),
            "{runs:?}"
        );
    }
}
