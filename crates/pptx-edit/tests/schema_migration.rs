//! `deck-schema-v1.update.bin` was produced by release 4bdccdd: it opens
//! `betteroffice-demo.pptx`, adds a text box and edits its story, then persists
//! `encode_state_as_update_v1()`.

use std::collections::BTreeMap;

use pptx_edit::{
    CommentFlavor, DeckSession, DeckSnapshot, EditCtx, EditError, ShapeSnapshot, TextCaps,
    TextStyle,
};
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, MapRef, Out, ReadTxn, StateVector, Transact, Update};

const V2_STYLE_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2.update.bin");
const V1_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v1.update.bin");
const V2_SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2-connectors.pptx");
const V2_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2-connectors.update.bin");
const V2_MOVED_UPDATE: &[u8] =
    include_bytes!("fixtures/deck-schema-v2-connectors-moved.update.bin");
const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const NUMBERED_FIXTURE: &[u8] =
    include_bytes!("../../pptx-parse/tests/fixtures/slide-number-fields.pptx");
const V2_HIDDEN_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2-hidden.update.bin");
const V2_1_DEFAULTS_SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2.1-defaults.pptx");
const V2_1_DEFAULTS_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2.1-defaults.update.bin");
const V2_1_EDITS_SOURCE: &[u8] = include_bytes!("fixtures/deck-schema-v2.1-edits.pptx");
const V2_1_EDITS_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2.1-edits.update.bin");
const SHAPES: &str = "pptx:shapes";
const V2_STORY_ID: &str = "story:shape:4343:0:0";
const V2_HIDDEN_SHAPE_IDS: [&str; 4] = [
    "slide:0:256:shape:0",
    "slide:0:256:shape:8",
    "slide:0:256:shape:8.13",
    "slide:1:257:shape:16",
];
const CUSTOM_V2_UPDATE: &[u8] = include_bytes!("fixtures/deck-custom-schema-v2.update.bin");
const V2_COMMENTS_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v2-comments.update.bin");
const COMMENTS_SOURCE: &[u8] = include_bytes!("fixtures/modern-comments.pptx");
const META: &str = "pptx:meta";
const SHAPE_ID: &str = "shape:4242:0";
const STORY_ID: &str = "story:shape:4242:0:0";

#[test]
fn a_fresh_deck_persists_its_list_styles() {
    let source = include_bytes!("../../pptx-render/tests/fixtures/list-style-bullets.pptx");
    let fresh = DeckSession::open(source, 29413).unwrap();
    let update = fresh.encode_state_as_update_v1();
    assert_eq!(stamped_version(&update), Some(2.2));
    let json = package_json(&update);
    assert!(json.contains("\"listStyle\""));
    assert!(json.contains("\"defaultListStyle\""));
    assert!(json.contains("\"bulletFont\""));
    let reopened = DeckSession::open_from_update(&update, 29414).unwrap();
    assert_eq!(reopened.encode_state_as_update_v1(), update);
    assert_eq!(reopened.snapshot().unwrap(), fresh.snapshot().unwrap());
}

#[test]
fn the_released_v2_comment_fixture_contains_no_comment_model() {
    let update = include_bytes!("fixtures/deck-schema-v2-comments.update.bin");
    assert_eq!(stamped_version(update), Some(2.0));
    assert!(!package_json(update).contains("\"comments\""));
    assert!(!package_json(update).contains("\"commentAuthors\""));
    assert!(
        hydrated(update)
            .transact()
            .get_map("pptx:comments")
            .is_none()
    );
}

#[test]
fn current_main_v2_custom_snapshot_migrates_once_without_losing_shapes() {
    assert_eq!(stamped_version(CUSTOM_V2_UPDATE), Some(2.0));
    let left = DeckSession::open_from_update(CUSTOM_V2_UPDATE, 909).unwrap();
    let right = DeckSession::open_from_update(CUSTOM_V2_UPDATE, 910).unwrap();
    let migrated = left.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    assert_eq!(package_json(&migrated), package_json(CUSTOM_V2_UPDATE));
    let snapshot = left.snapshot().unwrap();
    assert_eq!(
        (snapshot.width_emu, snapshot.height_emu),
        (4_572_000, 2_571_750)
    );
    assert_eq!(snapshot.slides.len(), 2);
    assert_eq!(snapshot.slides[0].shapes.len(), 5);
    assert_eq!(snapshot.slides[0].shapes[0].name, "Mixed paths");
    assert_eq!(snapshot.slides[0].shapes[0].geometry, "custom");
    let reopened = DeckSession::open_from_update(&migrated, 911).unwrap();
    assert_eq!(reopened.snapshot().unwrap(), snapshot);
    assert_eq!(
        StateVector::decode_v1(&reopened.encode_state_vector_v1()).unwrap(),
        StateVector::decode_v1(&left.encode_state_vector_v1()).unwrap()
    );
    right.apply_update_v1(&migrated).unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(
        package_json(&left.encode_state_as_update_v1()),
        package_json(&right.encode_state_as_update_v1())
    );
}

#[test]
fn current_main_generated_v2_snapshot_preserves_content_and_default_serialization() {
    assert_eq!(stamped_version(V2_STYLE_UPDATE), Some(2.0));
    let legacy_json = package_json(V2_STYLE_UPDATE);
    assert!(!legacy_json.contains("formatScheme"));
    let session = DeckSession::open_from_update(V2_STYLE_UPDATE, 909).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story = snapshot.slides[0]
        .shapes
        .iter()
        .find_map(|s| s.text_stories.first())
        .unwrap();
    assert_eq!(
        session.story(&story.id).unwrap().plain_text(),
        "persisted-v2 Styled"
    );
    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    let migrated_package: pptx_parse::PptxPackage =
        serde_json::from_str(&package_json(&migrated)).unwrap();
    assert_eq!(
        migrated_package,
        serde_json::from_str(&legacy_json).unwrap()
    );
    let reopened = DeckSession::open_from_update(&migrated, 910).unwrap();
    assert_eq!(reopened.snapshot().unwrap(), snapshot);
    assert_eq!(reopened.encode_state_as_update_v1().len(), migrated.len());

    let styled = DeckSession::open(
        include_bytes!("../../pptx-parse/tests/fixtures/style-matrix-deck.pptx"),
        911,
    )
    .unwrap();
    let update = styled.encode_state_as_update_v1();
    let restored = DeckSession::open_from_update(&update, 912).unwrap();
    assert_eq!(
        restored.package().themes[0].format_scheme,
        styled.package().themes[0].format_scheme
    );
    assert_eq!(
        restored.package().slides[0].shapes,
        styled.package().slides[0].shapes
    );
    assert_eq!(
        package_json(&restored.encode_state_as_update_v1()),
        package_json(&update)
    );
}

#[test]
fn a_fresh_current_snapshot_preserves_numbering_and_theme_formatting() {
    let mut package = pptx_parse::parse_pptx(include_bytes!(
        "../../pptx-parse/tests/fixtures/style-matrix-deck.pptx"
    ))
    .unwrap();
    package.presentation.first_slide_num = 10;
    let session = DeckSession::from_package(package, 9330).unwrap();
    let update = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&update), Some(2.2));
    let json = package_json(&update);
    assert!(json.contains("\"firstSlideNum\":10"));
    assert!(json.contains("formatScheme"));
    assert!(json.contains("fontColor"));
    let restored = DeckSession::open_from_update(&update, 9331).unwrap();
    assert_eq!(restored.package().presentation.first_slide_num, 10);
    assert_eq!(restored.package().themes, session.package().themes);
    assert_eq!(restored.package().slides, session.package().slides);
    assert_eq!(restored.encode_state_as_update_v1(), update);
}

#[test]
fn released_v1_snapshot_migrates_and_round_trips_as_v2_1() {
    assert_eq!(stamped_version(V1_UPDATE), Some(1.0));

    let session = DeckSession::open_from_update(V1_UPDATE, 901).unwrap();
    assert_v1_content(&session);

    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    assert!(
        package_json(&migrated).contains("\"charts\""),
        "the migrated package must carry the v2 chart field"
    );
    assert!(
        !package_json(&migrated).contains("firstSlideNum"),
        "default numbering must preserve the v2 package representation"
    );

    let reopened = DeckSession::open_from_update(&migrated, 902).unwrap();
    assert_v1_content(&reopened);
    assert_eq!(
        snapshot_shape_ids(&session.snapshot().unwrap()),
        snapshot_shape_ids(&reopened.snapshot().unwrap())
    );
    assert_eq!(
        reopened.encode_state_as_update_v1().len(),
        migrated.len(),
        "reopening a migrated snapshot must not migrate again"
    );
}

#[test]
fn a_migrated_session_still_edits() {
    let session = DeckSession::open_from_update(V1_UPDATE, 903).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            STORY_ID,
            0,
            "re-",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
    let reopened =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 904).unwrap();
    assert_eq!(
        reopened.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
}

#[test]
fn two_clients_migrating_the_same_v1_snapshot_converge() {
    let left = DeckSession::open_from_update(V1_UPDATE, 907).unwrap();
    let right = DeckSession::open_from_update(V1_UPDATE, 908).unwrap();

    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(
        stamped_version(&left.encode_state_as_update_v1()),
        Some(2.2)
    );
    assert_eq!(
        package_json(&left.encode_state_as_update_v1()),
        package_json(&right.encode_state_as_update_v1())
    );
}

#[test]
fn a_fresh_seed_persists_the_connector_filter_at_the_current_schema() {
    let session = DeckSession::open(V2_SOURCE, 909).unwrap();
    let update = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&update), Some(2.2));
    assert!(package_json(&update).contains("\"shapeElements\":\"withConnectors\""));
    let reopened = DeckSession::open_from_update_with_source(&update, V2_SOURCE, 910).unwrap();
    assert!(reopened.package().models_connectors());
    assert_eq!(reopened.snapshot().unwrap(), session.snapshot().unwrap());
    assert_eq!(reopened.save().unwrap(), session.save().unwrap());
}

#[test]
fn a_v2_snapshot_migrates_without_changing_its_package_or_shape_ids() {
    assert_eq!(stamped_version(V2_UPDATE), Some(2.0));
    let session = DeckSession::open_from_update(V2_UPDATE, 911).unwrap();
    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    assert_eq!(package_json(&migrated), package_json(V2_UPDATE));
    assert!(!session.package().models_connectors());
    let snapshot = session.snapshot().unwrap();
    assert_eq!(
        snapshot_shape_ids(&snapshot),
        ["slide:0:256:shape:0", "slide:0:256:shape:1"]
    );
    assert_eq!(snapshot.slides[0].shapes[1].source_id, 4);
    let reopened = DeckSession::open_from_update_with_source(&migrated, V2_SOURCE, 912).unwrap();
    assert_eq!(reopened.snapshot().unwrap(), snapshot);
    assert_eq!(reopened.encode_state_as_update_v1(), migrated);
    let cloned =
        DeckSession::from_package_with_source(session.package().clone(), V2_SOURCE, 913).unwrap();
    let reattached = DeckSession::open_from_update_with_source(
        &cloned.encode_state_as_update_v1(),
        V2_SOURCE,
        914,
    )
    .unwrap();
    assert_eq!(reattached.save().unwrap(), reopened.save().unwrap());
}

#[test]
fn v2_migration_converges_and_accepts_an_existing_peer_edit() {
    let left = DeckSession::open_from_update_with_source(V2_UPDATE, V2_SOURCE, 915).unwrap();
    let right = DeckSession::open_from_update_with_source(V2_UPDATE, V2_SOURCE, 916).unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(V2_MOVED_UPDATE).unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(left.snapshot().unwrap().slides[0].shapes[1].x, 952_500);
    assert_eq!(
        stamped_version(&left.encode_state_as_update_v1()),
        Some(2.2)
    );
    assert_eq!(left.save().unwrap(), right.save().unwrap());
    assert!(!left.package().models_connectors());
}

#[test]
fn unmigratable_schema_versions_stay_rejected() {
    for version in [
        0.0, 1.5, 2.05, 2.5, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0,
        16.0, 17.0, 18.0, 19.0, 20.0, 21.0, 22.0,
    ] {
        assert!(
            matches!(
                DeckSession::open_from_update(&restamped(V1_UPDATE, Some(version)), 905),
                Err(EditError::InvalidState(message))
                    if message == "unsupported deck schema version"
            ),
            "schema version {version} must be rejected"
        );
    }
    assert!(matches!(
        DeckSession::open_from_update(&restamped(V1_UPDATE, None), 906),
        Err(EditError::InvalidState(message))
            if message == "unsupported deck schema version"
    ));
}

#[test]
fn default_numbering_omits_the_default_at_the_current_schema() {
    let session = DeckSession::open(FIXTURE, 913).unwrap();
    let update = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&update), Some(2.2));
    assert!(!package_json(&update).contains("firstSlideNum"));
    let restored = DeckSession::open_from_update(&update, 914).unwrap();
    assert_eq!(restored.package().presentation.first_slide_num, 1);
    assert_eq!(
        restored.encode_state_vector_v1(),
        session.encode_state_vector_v1()
    );
    assert_eq!(
        package_json(&restored.encode_state_as_update_v1()),
        package_json(&update)
    );
}

#[test]
fn slide_number_offsets_survive_migration_from_released_snapshots() {
    let session = DeckSession::open(NUMBERED_FIXTURE, 910).unwrap();
    let update = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&update), Some(2.2));
    assert!(package_json(&update).contains("\"firstSlideNum\":10"));
    for version in [1.0, 2.0] {
        let restored =
            DeckSession::open_from_update(&restamped(&update, Some(version)), 911).unwrap();
        assert_eq!(restored.package().presentation.first_slide_num, 10);
        assert_eq!(restored.snapshot().unwrap(), session.snapshot().unwrap());
        let migrated = restored.encode_state_as_update_v1();
        assert_eq!(stamped_version(&migrated), Some(2.2));
        assert_eq!(package_json(&migrated), package_json(&update));
        let reopened = DeckSession::open_from_update(&migrated, 912).unwrap();
        assert_eq!(
            reopened.encode_state_vector_v1(),
            restored.encode_state_vector_v1()
        );
    }
}

fn assert_v1_content(session: &DeckSession) {
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.width_emu, 12_192_000);
    assert_eq!(snapshot.height_emu, 6_858_000);
    assert_eq!(session.package().presentation.first_slide_num, 1);
    assert_eq!(snapshot.slides.len(), 3);
    assert_eq!(snapshot.slides[0].id, "slide:0:256");
    assert!(
        snapshot_shape_ids(&snapshot)
            .iter()
            .any(|id| id == SHAPE_ID)
    );
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "edited persisted on v1"
    );
    assert!(session.package().charts.is_empty());
    assert!(
        snapshot.comments.is_empty(),
        "a deck migrated from v1 carries no comments"
    );
    assert_eq!(snapshot.comment_flavor, CommentFlavor::Legacy);
}

fn snapshot_shape_ids(snapshot: &DeckSnapshot) -> Vec<String> {
    snapshot
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .map(|shape| shape.id.clone())
        .collect()
}

fn hydrated(update: &[u8]) -> Doc {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    doc
}

fn meta(doc: &Doc) -> MapRef {
    doc.transact().get_map(META).unwrap()
}

fn stamped_version(update: &[u8]) -> Option<f64> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "schemaVersion") {
        Some(Out::Any(Any::Number(value))) => Some(value),
        _ => None,
    }
}

fn package_json(update: &[u8]) -> String {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "packageJson") {
        Some(Out::Any(Any::Buffer(bytes))) => String::from_utf8(bytes.to_vec()).unwrap(),
        _ => panic!("missing packageJson"),
    }
}

fn restamped(update: &[u8], version: Option<f64>) -> Vec<u8> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    {
        let mut txn = doc.transact_mut();
        match version {
            Some(version) => {
                meta.insert(&mut txn, "schemaVersion", version);
            }
            None => {
                meta.remove(&mut txn, "schemaVersion");
            }
        }
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

#[test]
fn released_v2_snapshot_migrates_with_hidden_flags_backfilled() {
    assert_eq!(stamped_version(V2_HIDDEN_UPDATE), Some(2.0));
    assert!(hidden_keys(V2_HIDDEN_UPDATE).is_empty());

    let session = DeckSession::open_from_update(V2_HIDDEN_UPDATE, 911).unwrap();
    assert_v2_content(&session);

    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    assert_eq!(hidden_keys(&migrated), V2_HIDDEN_SHAPE_IDS);

    let reopened = DeckSession::open_from_update(&migrated, 912).unwrap();
    assert_v2_content(&reopened);
    assert_eq!(reopened.snapshot().unwrap(), session.snapshot().unwrap());
    assert_eq!(
        reopened.encode_state_as_update_v1().len(),
        migrated.len(),
        "reopening a migrated snapshot must not migrate again"
    );
}

#[test]
fn two_clients_migrating_the_same_v2_snapshot_converge() {
    let left = DeckSession::open_from_update(V2_HIDDEN_UPDATE, 913).unwrap();
    let right = DeckSession::open_from_update(V2_HIDDEN_UPDATE, 914).unwrap();

    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_v2_content(&left);
    let merged = left.encode_state_as_update_v1();
    assert_eq!(stamped_version(&merged), Some(2.2));
    assert_eq!(hidden_keys(&merged), V2_HIDDEN_SHAPE_IDS);
}

fn assert_v2_content(session: &DeckSession) {
    let snapshot = session.snapshot().unwrap();
    assert_eq!(
        snapshot
            .slides
            .iter()
            .map(|slide| slide.id.as_str())
            .collect::<Vec<_>>(),
        ["slide:2:258", "slide:0:256", "slide:1:257"]
    );
    assert_eq!(hidden_shape_ids(&snapshot), V2_HIDDEN_SHAPE_IDS);
    let ids = snapshot_shape_ids(&snapshot);
    assert!(ids.iter().any(|id| id == "shape:4343:0"));
    assert!(!ids.iter().any(|id| id == "slide:1:257:shape:4"));
    assert_eq!(
        session.story(V2_STORY_ID).unwrap().plain_text(),
        "edited persisted on v2"
    );
}

fn hidden_shape_ids(snapshot: &DeckSnapshot) -> Vec<String> {
    fn collect(shapes: &[ShapeSnapshot], ids: &mut Vec<String>) {
        for shape in shapes {
            if shape.hidden {
                ids.push(shape.id.clone());
            }
            collect(&shape.children, ids);
        }
    }
    let mut ids = Vec::new();
    for slide in &snapshot.slides {
        collect(&slide.shapes, &mut ids);
    }
    ids
}

fn hidden_keys(update: &[u8]) -> Vec<String> {
    let doc = hydrated(update);
    let txn = doc.transact();
    let shapes = txn.get_map(SHAPES).unwrap();
    let mut ids: Vec<String> = shapes
        .iter(&txn)
        .filter_map(|(id, value)| value.cast::<MapRef>().ok().map(|shape| (id, shape)))
        .filter(|(_, shape)| shape.get(&txn, "hidden").is_some())
        .map(|(id, _)| id.to_owned())
        .collect();
    ids.sort();
    ids
}

#[test]
fn future_full_and_differential_updates_are_rejected_atomically() {
    let session = DeckSession::open_from_update(V2_HIDDEN_UPDATE, 9420).unwrap();
    let original = session.encode_state_as_update_v1();
    let future = restamped(&original, Some(3.0));
    let future_doc = hydrated(&future);
    let base = hydrated(&original);
    let diff = future_doc
        .transact()
        .encode_state_as_update_v1(&base.transact().state_vector());
    for update in [&future, &diff] {
        assert!(matches!(
            session.apply_update_v1(update),
            Err(EditError::InvalidState(message)) if message == "unsupported deck schema version"
        ));
        assert_eq!(session.encode_state_as_update_v1(), original);
    }
}

#[test]
fn v2_snapshots_migrate_once_and_import_source_comments() {
    let deferred = DeckSession::open_from_update(V2_COMMENTS_UPDATE, 9610).unwrap();
    assert!(deferred.comments().unwrap().is_empty());
    assert!(deferred.package().comments.is_empty());
    let migrated = deferred.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(2.2));
    let reopened = DeckSession::open_from_update(&migrated, 9611).unwrap();
    assert_eq!(reopened.snapshot().unwrap(), deferred.snapshot().unwrap());
    assert_eq!(reopened.encode_state_as_update_v1(), migrated);

    let attached =
        DeckSession::open_from_update_with_source(V2_COMMENTS_UPDATE, COMMENTS_SOURCE, 9612)
            .unwrap();
    let snapshot = attached.snapshot().unwrap();
    assert_eq!(snapshot.comments.len(), 5);
    assert_eq!(snapshot.comment_flavor, CommentFlavor::Modern);
    assert_eq!(snapshot.slides.len(), 3);
    let imported = attached.encode_state_as_update_v1();
    assert_eq!(stamped_version(&imported), Some(2.2));
    assert!(package_json(&imported).contains("\"comments\""));
    let later =
        DeckSession::open_from_update_with_source(&migrated, COMMENTS_SOURCE, 9613).unwrap();
    assert_eq!(later.snapshot().unwrap(), snapshot);
    let reattached =
        DeckSession::open_from_update_with_source(&imported, COMMENTS_SOURCE, 9614).unwrap();
    assert_eq!(reattached.snapshot().unwrap(), snapshot);
    assert_eq!(
        StateVector::decode_v1(&reattached.encode_state_vector_v1()).unwrap(),
        StateVector::decode_v1(&attached.encode_state_vector_v1()).unwrap()
    );
}

/// Strips every chart text property and stamps `version`, standing in for a
/// package stored before schema 2.1 read `c:txPr`.
fn without_chart_text(update: &[u8], version: f64) -> Vec<u8> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    let mut package: serde_json::Value = serde_json::from_str(&package_json(update)).unwrap();
    for chart in package["charts"].as_array_mut().unwrap() {
        strip_chart_text(&mut chart["chart"]);
    }
    {
        let mut txn = doc.transact_mut();
        meta.insert(
            &mut txn,
            "packageJson",
            Any::Buffer(serde_json::to_vec(&package).unwrap().into()),
        );
        meta.insert(&mut txn, "schemaVersion", version);
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

fn strip_chart_text(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(map) => {
            map.remove("text");
            map.remove("titleText");
            for value in map.values_mut() {
                strip_chart_text(value);
            }
        }
        serde_json::Value::Array(values) => values.iter_mut().for_each(strip_chart_text),
        _ => {}
    }
}

#[test]
fn a_pre_2_1_chart_carries_its_stored_text_and_recovers_it_from_a_source() {
    let source = include_bytes!("../../pptx-render/tests/fixtures/chart-text-properties.pptx");
    let fresh = DeckSession::open(source, 34700).unwrap();
    let current = fresh.encode_state_as_update_v1();
    assert_eq!(stamped_version(&current), Some(2.2));
    assert!(package_json(&current).contains("\"spacingPt\":6.0"));

    let stored = without_chart_text(&current, 2.0);
    assert!(!package_json(&stored).contains("spacingPt"));
    let migrated = DeckSession::open_from_update(&stored, 34701).unwrap();
    let carried = migrated.encode_state_as_update_v1();
    assert_eq!(stamped_version(&carried), Some(2.2));
    assert!(!package_json(&carried).contains("spacingPt"));

    let attached = DeckSession::open_from_update_with_source(&carried, source, 34702).unwrap();
    let imported = attached.encode_state_as_update_v1();
    assert_eq!(package_json(&imported), package_json(&current));
    let reattached = DeckSession::open_from_update_with_source(&imported, source, 34703).unwrap();
    assert_eq!(reattached.encode_state_as_update_v1(), imported);
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

fn without_outline_gradients(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(map) => {
            if let Some(serde_json::Value::Object(outline)) = map.get_mut("outline") {
                outline.remove("gradient");
            }
            for value in map.values_mut() {
                without_outline_gradients(value);
            }
        }
        serde_json::Value::Array(values) => values.iter_mut().for_each(without_outline_gradients),
        _ => {}
    }
}

fn pending_flag(update: &[u8], key: &str) -> Option<bool> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), key) {
        Some(Out::Any(Any::Bool(value))) => Some(value),
        _ => None,
    }
}

/// Seeds `source` through a package stripped of what a released 2.0 writer never
/// stored, then stamps the update 2.0.
fn stored_2_0(source: &[u8], client_id: u64, strip: impl Fn(&mut serde_json::Value)) -> Vec<u8> {
    let mut value = serde_json::to_value(pptx_parse::parse_pptx(source).unwrap()).unwrap();
    strip(&mut value);
    let package = serde_json::from_value(value).unwrap();
    let legacy = DeckSession::from_package_with_source(package, source, client_id).unwrap();
    restamped(&legacy.encode_state_as_update_v1(), Some(2.0))
}

fn assert_source_recovers(
    source: &[u8],
    client_id: u64,
    pending: &[&str],
    strip: impl Fn(&mut serde_json::Value),
) {
    let stored = stored_2_0(source, client_id, strip);
    assert_eq!(stamped_version(&stored), Some(2.0));
    let fresh = DeckSession::open(source, client_id + 1).unwrap();
    let migrated = DeckSession::open_from_update(&stored, client_id + 2).unwrap();
    let carried = migrated.encode_state_as_update_v1();
    assert_eq!(stamped_version(&carried), Some(2.2));
    assert_ne!(
        serde_json::to_value(migrated.package()).unwrap(),
        serde_json::to_value(fresh.package()).unwrap()
    );
    for key in pending {
        assert_eq!(pending_flag(&carried, key), Some(true), "{key}");
    }
    let attached =
        DeckSession::open_from_update_with_source(&carried, source, client_id + 3).unwrap();
    assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
    let imported = attached.encode_state_as_update_v1();
    for key in pending {
        assert_ne!(pending_flag(&imported, key), Some(true), "{key}");
    }
    let reopened = DeckSession::open_from_update(&imported, client_id + 4).unwrap();
    assert_eq!(
        serde_json::to_value(reopened.package()).unwrap(),
        serde_json::to_value(fresh.package()).unwrap()
    );
    assert_eq!(reopened.snapshot().unwrap(), fresh.snapshot().unwrap());
    let reattached =
        DeckSession::open_from_update_with_source(&imported, source, client_id + 5).unwrap();
    assert_eq!(reattached.encode_state_as_update_v1(), imported);
}

#[test]
fn a_2_0_snapshot_recovers_run_baselines() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/text-baseline-script.pptx"),
        41000,
        &["baselinesPendingSource"],
        |value| without_keys(value, &["baselinePct"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_character_spacing() {
    assert_source_recovers(
        include_bytes!("fixtures/run-spacing-shadow.pptx"),
        41100,
        &["spacingPendingSource"],
        |value| without_keys(value, &["spacingPt"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_outline_gradients() {
    assert_source_recovers(
        include_bytes!("../../pptx-parse/tests/fixtures/gradient-outline.pptx"),
        41200,
        &["outlineGradientsPendingSource"],
        without_outline_gradients,
    );
}

#[test]
fn a_2_0_snapshot_recovers_bitmap_effects() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/blip-effects.pptx"),
        41300,
        &[],
        |value| without_keys(value, &["effects"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_shape_shadows() {
    assert_source_recovers(
        include_bytes!("fixtures/blip-shadow.pptx"),
        41400,
        &[],
        |value| without_keys(value, &["effects", "shapeEffects"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_picture_fills() {
    assert_source_recovers(
        include_bytes!("../../pptx-parse/tests/fixtures/picture-fill.pptx"),
        41500,
        &[],
        |value| without_keys(value, &["pictureFill"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_list_styles() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/list-style-bullets.pptx"),
        41600,
        &[],
        |value| without_keys(value, &["listStyle", "defaultListStyle"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_explicit_overflow() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/text-overflow.pptx"),
        41700,
        &[],
        |value| without_keys(value, &["verticalOverflow", "horizontalOverflow"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_chart_fills_and_axis_lines() {
    assert_source_recovers(
        include_bytes!("fixtures/chart-text-overflow.pptx"),
        41800,
        &[],
        |value| {
            if let Some(charts) = value.get_mut("charts") {
                without_keys(charts, &["fill", "line", "text", "titleText"]);
            }
        },
    );
}

#[test]
fn a_2_0_snapshot_recovers_ole_picture_previews() {
    assert_source_recovers(
        include_bytes!("fixtures/metafile-tracking.pptx"),
        41900,
        &["olePicturesPendingSource"],
        |value| without_keys(value, &["picture"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_numbering_restarts() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/autonumber-bullets.pptx"),
        42000,
        &[],
        |value| without_keys(value, &["restart"]),
    );
}

#[test]
fn a_2_0_snapshot_recovers_line_spacing() {
    assert_source_recovers(
        include_bytes!("../../pptx-render/tests/fixtures/line-spacing.pptx"),
        42100,
        &[],
        |value| without_keys(value, &["lineSpacing", "compatLineSpacing"]),
    );
}

#[test]
fn a_2_1_deck_with_integer_media_arrays_migrates_to_base64() {
    let fresh = DeckSession::open(FIXTURE, 4171).unwrap();
    let update = fresh.encode_state_as_update_v1();
    let json = package_json(&update);
    assert!(
        json.contains("\"bytes\":\""),
        "2.2 must write base64 strings"
    );

    let mut value: serde_json::Value = serde_json::from_str(&json).unwrap();
    let media = value["media"]
        .as_array_mut()
        .expect("fixture carries media");
    assert!(!media.is_empty());
    for part in media.iter_mut() {
        part["bytes"] = serde_json::json!([7, 6, 5]);
    }
    let doc = hydrated(&update);
    let meta = meta(&doc);
    {
        let mut txn = doc.transact_mut();
        meta.insert(
            &mut txn,
            "packageJson",
            Any::Buffer(serde_json::to_vec(&value).unwrap().into()),
        );
        meta.insert(&mut txn, "schemaVersion", 2.1);
    }
    let legacy = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());

    let migrated = DeckSession::open_from_update(&legacy, 4172).unwrap();
    let migrated_json = package_json(&migrated.encode_state_as_update_v1());
    assert_eq!(stamped_version(&legacy), Some(2.1));
    assert!(migrated_json.contains("\"bytes\":\"BwYF\""));
    assert!(!migrated_json.contains("[7,6,5]"));
    assert!(migrated_json.contains("betteroffice-mark.png"));
}

#[test]
fn a_released_update_reopened_with_its_source_saves_an_unedited_deck_byte_identically() {
    assert_eq!(stamped_version(V2_1_DEFAULTS_UPDATE), Some(2.1));
    let fresh = DeckSession::open(V2_1_DEFAULTS_SOURCE, 42200).unwrap();
    for (client_id, update) in [
        (42210, V2_1_DEFAULTS_UPDATE.to_vec()),
        (42220, restamped(V2_1_DEFAULTS_UPDATE, Some(2.0))),
    ] {
        let migrated = DeckSession::open_from_update(&update, client_id).unwrap();
        assert_eq!(
            adjust_values(&migrated.snapshot().unwrap()),
            adjust_values(&fresh.snapshot().unwrap())
        );
        let attached = DeckSession::open_from_update_with_source(
            &migrated.encode_state_as_update_v1(),
            V2_1_DEFAULTS_SOURCE,
            client_id + 1,
        )
        .unwrap();
        assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(attached.save().unwrap(), V2_1_DEFAULTS_SOURCE);
        let carried = attached.encode_state_as_update_v1();
        let sourceless = DeckSession::open_from_update(&carried, client_id + 2).unwrap();
        assert_eq!(sourceless.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(
            sourceless.package().masters[0].color_map,
            fresh.package().masters[0].color_map
        );
        let reattached = DeckSession::open_from_update_with_source(
            &carried,
            V2_1_DEFAULTS_SOURCE,
            client_id + 3,
        )
        .unwrap();
        assert_eq!(reattached.save().unwrap(), V2_1_DEFAULTS_SOURCE);
    }
}

#[test]
fn edits_made_after_migrating_a_released_update_survive_the_source_import() {
    let migrated = DeckSession::open_from_update(V2_1_DEFAULTS_UPDATE, 42230).unwrap();
    let context = EditCtx::local("fixture");
    let snapshot = migrated.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = |name: &str| {
        slide
            .shapes
            .iter()
            .find(|shape| shape.name == name)
            .unwrap()
    };
    migrated
        .set_shape_adjust(
            &context,
            &slide.id,
            &shape("Default star").id,
            &BTreeMap::from([("adj".to_owned(), 0.45)]),
        )
        .unwrap();
    migrated
        .insert_text(
            &context,
            &shape("Direct all caps").text_stories[0].id,
            0,
            "QZ",
            &TextStyle {
                font_size_pt: Some(32.0),
                color: Some("#101828".to_owned()),
                caps: Some(TextCaps::Small),
                ..TextStyle::default()
            },
        )
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        V2_1_DEFAULTS_SOURCE,
        42231,
    )
    .unwrap();
    let saved = attached.save().unwrap();
    let parts: BTreeMap<_, _> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let slide = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    let shape_xml = |name: &str| {
        let start = slide.find(&format!("name=\"{name}\"")).unwrap();
        &slide[start..start + slide[start..].find("</p:sp>").unwrap()]
    };
    assert!(shape_xml("Default star").contains(r#"<a:gd fmla="val 45000" name="adj"/>"#));
    let caps = shape_xml("Direct all caps");
    assert!(caps.contains(concat!(
        r#"<a:rPr cap="small" lang="en-US" sz="3200"><a:solidFill><a:srgbClr val="101828"/>"#,
        r#"</a:solidFill></a:rPr><a:t>QZ</a:t>"#
    )));
    assert!(caps.contains(concat!(
        r#"<a:rPr cap="all" lang="en-US" sz="3200"><a:solidFill><a:schemeClr val="tx1"/>"#,
        r#"</a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Mixed Case Title</a:t>"#
    )));
    assert!(!shape_xml("Default trapezoid").contains("<a:gd "));
}

fn adjust_values(snapshot: &DeckSnapshot) -> Vec<BTreeMap<String, f64>> {
    fn collect(shapes: &[ShapeSnapshot], values: &mut Vec<BTreeMap<String, f64>>) {
        for shape in shapes {
            values.push(shape.adjust_values.clone());
            collect(&shape.children, values);
        }
    }
    let mut values = Vec::new();
    for slide in &snapshot.slides {
        collect(&slide.shapes, &mut values);
    }
    values
}

const DEFAULTS_SLIDE: &str = "ppt/slides/slide1.xml";
const STORY_COUNTS: [usize; 3] = [40, 160, 640];
const DELETIONS_PER_STORY: u32 = 6;
const FRAGMENTED_PARAGRAPHS: usize = 40;
const FRAGMENTED_RUNS: usize = 4;
const FRAGMENTED_REPEATS: usize = 12;
const DEFAULTS_LAYOUT: &str = "ppt/slideLayouts/slideLayout1.xml";
const DEFAULTS_MASTER: &str = "ppt/slideMasters/slideMaster1.xml";
const INVERTED_MAP: &str = r#"bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2""#;
const IDENTITY_MAP: &str = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2""#;

type RunLook = (String, Option<TextCaps>, Option<String>);

#[test]
fn a_repeated_character_inserted_before_recovery_stays_the_inserted_text() {
    let migrated = DeckSession::open_from_update(V2_1_DEFAULTS_UPDATE, 42500).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    migrated
        .insert_text(&EditCtx::local("fixture"), &story, 1, "i", &small_caps())
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        V2_1_DEFAULTS_SOURCE,
        42501,
    )
    .unwrap();
    assert_eq!(
        pending_flag(&attached.encode_state_as_update_v1(), "capsPendingSource"),
        None
    );
    let expected = [
        look("M", Some(TextCaps::All), "#FFFFFF"),
        look("i", Some(TextCaps::Small), "#101828"),
        look("ixed Case Title", Some(TextCaps::All), "#FFFFFF"),
    ];
    assert_eq!(run_looks(&attached, &story), expected);
    let saved = DeckSession::open(&attached.save().unwrap(), 42502).unwrap();
    assert_eq!(run_looks(&saved, &story), expected);
}

#[test]
fn a_long_story_with_separated_edits_recovers_its_source_runs() {
    let text = "Mixed Case Title ".repeat(180);
    let text = text.trim_end();
    let source = defaults_variant(&[(
        DEFAULTS_SLIDE,
        "<a:t>Mixed Case Title</a:t>",
        &format!("<a:t>{text}</a:t>"),
    )]);
    let migrated = DeckSession::open_from_update(&legacy_seed(&source, 42600), 42601).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    let context = EditCtx::local("fixture");
    let length = text.len() as u32;
    migrated.delete_text(&context, &story, 0, 1).unwrap();
    migrated
        .insert_text(&context, &story, 0, "X", &TextStyle::default())
        .unwrap();
    migrated
        .delete_text(&context, &story, length - 1, length)
        .unwrap();
    migrated
        .insert_text(&context, &story, length - 1, "Y", &TextStyle::default())
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        &source,
        42602,
    )
    .unwrap();
    assert_eq!(
        run_looks(&attached, &story),
        [
            ("X".to_owned(), None, None),
            look(&text[1..text.len() - 1], Some(TextCaps::All), "#FFFFFF"),
            ("Y".to_owned(), None, None),
        ]
    );
    attached.save().unwrap();
}

#[test]
fn a_released_update_recovers_run_colours_through_every_colour_map() {
    let mapping = format!(
        "<a:overrideClrMapping {INVERTED_MAP} accent1=\"accent1\" accent2=\"accent2\" \
         accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" \
         hlink=\"hlink\" folHlink=\"folHlink\"/>"
    );
    let slide_override = format!("</p:cSld><p:clrMapOvr>{mapping}</p:clrMapOvr>");
    for (client_id, source) in [
        (
            42700,
            defaults_variant(&[(
                DEFAULTS_SLIDE,
                r#"<a:schemeClr val="tx1"/>"#,
                r#"<a:schemeClr val="bg1"/>"#,
            )]),
        ),
        (
            42710,
            defaults_variant(&[
                (DEFAULTS_MASTER, INVERTED_MAP, IDENTITY_MAP),
                (DEFAULTS_SLIDE, "</p:cSld>", &slide_override),
            ]),
        ),
        (
            42720,
            defaults_variant(&[
                (DEFAULTS_MASTER, INVERTED_MAP, IDENTITY_MAP),
                (DEFAULTS_LAYOUT, "<a:masterClrMapping/>", &mapping),
            ]),
        ),
    ] {
        let fresh = DeckSession::open(&source, client_id).unwrap();
        let migrated =
            DeckSession::open_from_update(&legacy_seed(&source, client_id + 1), client_id + 2)
                .unwrap();
        assert_ne!(migrated.snapshot().unwrap(), fresh.snapshot().unwrap());
        let attached = DeckSession::open_from_update_with_source(
            &migrated.encode_state_as_update_v1(),
            &source,
            client_id + 3,
        )
        .unwrap();
        assert_eq!(attached.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(attached.save().unwrap(), source);
        let sourceless =
            DeckSession::open_from_update(&attached.encode_state_as_update_v1(), client_id + 4)
                .unwrap();
        assert_eq!(sourceless.snapshot().unwrap(), fresh.snapshot().unwrap());
        let (package, expected) = (sourceless.package(), fresh.package());
        assert_eq!(package.masters[0].color_map, expected.masters[0].color_map);
        assert_eq!(
            package.layouts[0].color_map_override,
            expected.layouts[0].color_map_override
        );
        assert_eq!(
            package.slides[0].color_map_override,
            expected.slides[0].color_map_override
        );
    }
}

#[test]
fn peers_recovering_the_same_released_update_converge() {
    let fresh = DeckSession::open(V2_1_DEFAULTS_SOURCE, 42800).unwrap();
    let left = DeckSession::open_from_update_with_source(
        V2_1_DEFAULTS_UPDATE,
        V2_1_DEFAULTS_SOURCE,
        42801,
    )
    .unwrap();
    let right = DeckSession::open_from_update_with_source(
        V2_1_DEFAULTS_UPDATE,
        V2_1_DEFAULTS_SOURCE,
        42802,
    )
    .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    for session in [&left, &right] {
        assert_eq!(session.snapshot().unwrap(), fresh.snapshot().unwrap());
        assert_eq!(session.save().unwrap(), V2_1_DEFAULTS_SOURCE);
    }
}

#[test]
fn recovery_concurrent_with_a_peer_edit_converges_and_keeps_the_inserted_text() {
    let editor = DeckSession::open_from_update(V2_1_DEFAULTS_UPDATE, 42900).unwrap();
    let story = defaults_story(&editor, "Direct all caps");
    editor
        .insert_text(&EditCtx::local("fixture"), &story, 1, "i", &small_caps())
        .unwrap();
    let recovering = DeckSession::open_from_update_with_source(
        V2_1_DEFAULTS_UPDATE,
        V2_1_DEFAULTS_SOURCE,
        42901,
    )
    .unwrap();
    editor
        .apply_update_v1(&recovering.encode_state_as_update_v1())
        .unwrap();
    recovering
        .apply_update_v1(&editor.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(editor.snapshot().unwrap(), recovering.snapshot().unwrap());
    let looks = run_looks(&recovering, &story);
    assert_eq!(
        looks.iter().map(|look| look.0.as_str()).collect::<String>(),
        "Miixed Case Title"
    );
    assert_eq!(looks[0], look("M", Some(TextCaps::All), "#FFFFFF"));
    assert_eq!(
        (looks[1].0.as_str(), looks[1].1),
        ("i", Some(TextCaps::Small))
    );
    let saved = DeckSession::open(&recovering.save().unwrap(), 42902).unwrap();
    assert_eq!(run_looks(&saved, &story), looks);
}

#[test]
fn a_surviving_character_between_identical_source_characters_keeps_its_own_run() {
    let source = defaults_variant(&[(
        DEFAULTS_SLIDE,
        "<a:t>Mixed Case Title</a:t></a:r>",
        concat!(
            r#"<a:t>a</a:t></a:r><a:r><a:rPr lang="en-US" sz="3200" cap="small">"#,
            r#"<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>"#,
            r#"<a:latin typeface="Arial"/></a:rPr><a:t>a</a:t></a:r>"#
        ),
    )]);
    let migrated = DeckSession::open_from_update(&legacy_seed(&source, 43100), 43101).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    migrated
        .delete_text(&EditCtx::local("fixture"), &story, 0, 1)
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        &source,
        43102,
    )
    .unwrap();
    let expected = [look("a", Some(TextCaps::Small), "#FFFFFF")];
    assert_eq!(run_looks(&attached, &story), expected);
    let carried = attached.encode_state_as_update_v1();
    assert!(!meta_has(&carried, "capsPendingSource"));
    let saved = DeckSession::open(&attached.save().unwrap(), 43103).unwrap();
    assert_eq!(run_looks(&saved, &story), expected);
}

#[test]
fn an_ambiguous_character_keeps_only_its_story_pending() {
    let source = defaults_variant(&[(
        DEFAULTS_SLIDE,
        "<a:t>Mixed Case Title</a:t></a:r>",
        concat!(
            r#"<a:t>a</a:t></a:r><a:r><a:rPr lang="en-US" sz="3200" cap="small">"#,
            r#"<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>"#,
            r#"<a:latin typeface="Arial"/></a:rPr><a:t>a</a:t></a:r>"#
        ),
    )]);
    let seeded = legacy_seed(&source, 43600);
    let doc = hydrated(&seeded);
    let mut package: serde_json::Value = serde_json::from_str(&package_json(&seeded)).unwrap();
    without_keys(&mut package, &["fontSizePt"]);
    let stored_meta = meta(&doc);
    stored_meta.insert(
        &mut doc.transact_mut(),
        "packageJson",
        Any::Buffer(serde_json::to_vec(&package).unwrap().into()),
    );
    let stored = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let migrated = DeckSession::open_from_update(&stored, 43601).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    migrated
        .delete_text(&EditCtx::local("fixture"), &story, 0, 1)
        .unwrap();
    let mut update = migrated.encode_state_as_update_v1();
    for client_id in [43602, 43603] {
        let attached =
            DeckSession::open_from_update_with_source(&update, &source, client_id).unwrap();
        update = attached.encode_state_as_update_v1();
        let pending: BTreeMap<String, Vec<(u64, u32)>> =
            serde_json::from_str(&meta_string(&update, "capsPendingSource")).unwrap();
        assert_eq!(pending.keys().collect::<Vec<_>>(), [&story]);
        assert_eq!(pending[&story].len(), 1);
        assert_eq!(run_looks(&attached, &story)[0].1, None);
        assert_eq!(
            run_looks(&attached, &defaults_story(&attached, "Direct small caps"))[0].1,
            Some(TextCaps::Small)
        );
    }
}

#[test]
fn retyped_source_text_keeps_the_style_it_was_typed_with() {
    let migrated = DeckSession::open_from_update(V2_1_DEFAULTS_UPDATE, 43200).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    let context = EditCtx::local("fixture");
    migrated.delete_text(&context, &story, 0, 1).unwrap();
    let typed = TextStyle {
        font_size_pt: Some(32.0),
        color: Some("#101828".to_owned()),
        ..TextStyle::default()
    };
    migrated
        .insert_text(&context, &story, 0, "M", &typed)
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        V2_1_DEFAULTS_SOURCE,
        43201,
    )
    .unwrap();
    let expected = [
        look("M", None, "#101828"),
        look("ixed Case Title", Some(TextCaps::All), "#FFFFFF"),
    ];
    assert_eq!(run_looks(&attached, &story), expected);
    let saved = DeckSession::open(&attached.save().unwrap(), 43202).unwrap();
    assert_eq!(run_looks(&saved, &story), expected);
}

#[test]
fn undo_restored_source_text_is_left_as_restored() {
    let migrated = DeckSession::open_from_update(V2_1_DEFAULTS_UPDATE, 43300).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    migrated
        .delete_text(&EditCtx::local("fixture"), &story, 1, 2)
        .unwrap();
    assert!(migrated.undo());
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        V2_1_DEFAULTS_SOURCE,
        43301,
    )
    .unwrap();
    assert_eq!(
        run_looks(&attached, &story),
        [
            look("M", Some(TextCaps::All), "#FFFFFF"),
            look("i", None, "#101828"),
            look("xed Case Title", Some(TextCaps::All), "#FFFFFF"),
        ]
    );
}

#[test]
fn a_preset_value_a_collaborator_set_before_migration_stays() {
    let fresh = DeckSession::open(V2_1_DEFAULTS_SOURCE, 43400).unwrap();
    let shape = |name: &str| {
        fresh.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name == name)
            .unwrap()
            .id
            .clone()
    };
    let (trapezoid, star) = (shape("Default trapezoid"), shape("Default star"));
    let doc = hydrated(V2_1_DEFAULTS_UPDATE);
    {
        let mut txn = doc.transact_mut();
        let shapes = txn.get_map(SHAPES).unwrap();
        let shape = shapes
            .get(&txn, &trapezoid)
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        shape.insert(&mut txn, "adjustValuesJson", r#"{"adj":0.2}"#);
    }
    let edited = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let migrated = DeckSession::open_from_update(&edited, 43401).unwrap();
    let adjust = |session: &DeckSession, id: &str| {
        session.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .find(|shape| shape.id == id)
            .unwrap()
            .adjust_values
            .clone()
    };
    assert_eq!(
        adjust(&migrated, &trapezoid),
        BTreeMap::from([("adj".to_owned(), 0.2)])
    );
    assert_eq!(adjust(&migrated, &star), adjust(&fresh, &star));
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        V2_1_DEFAULTS_SOURCE,
        43402,
    )
    .unwrap();
    let parts: BTreeMap<_, _> = ooxml_opc::unzip_parts(&attached.save().unwrap())
        .unwrap()
        .into_iter()
        .collect();
    let slide = String::from_utf8(parts[DEFAULTS_SLIDE].clone()).unwrap();
    let start = slide.find(r#"name="Default trapezoid""#).unwrap();
    let trapezoid_xml = &slide[start..start + slide[start..].find("</p:sp>").unwrap()];
    assert!(trapezoid_xml.contains(r#"<a:gd fmla="val 20000" name="adj"/>"#));
}

#[test]
fn a_fragmented_story_recovers_in_about_the_time_a_fresh_open_takes() {
    let mut paragraphs = String::new();
    for paragraph in 0..FRAGMENTED_PARAGRAPHS {
        paragraphs.push_str("<a:p><a:pPr/>");
        for run in 0..FRAGMENTED_RUNS {
            let (caps, size) = if (paragraph + run) % 2 == 0 {
                ("all", 3200)
            } else {
                ("small", 3300)
            };
            paragraphs.push_str(&format!(
                "<a:r><a:rPr lang=\"en-US\" sz=\"{size}\" cap=\"{caps}\"><a:solidFill>\
                 <a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr>\
                 <a:t>{}</a:t></a:r>",
                format!("run {run} of {paragraph} ").repeat(FRAGMENTED_REPEATS)
            ));
        }
        paragraphs.push_str("</a:p>");
    }
    let slide = defaults_part(DEFAULTS_SLIDE);
    let start = slide.find("<a:p><a:pPr/><a:r>").unwrap();
    let end = start + slide[start..].find("</a:p>").unwrap() + "</a:p>".len();
    let source = defaults_variant(&[(DEFAULTS_SLIDE, &slide[start..end], &paragraphs)]);
    let stored = legacy_seed(&source, 43500);
    let started = std::time::Instant::now();
    let fresh = DeckSession::open(&source, 43501).unwrap();
    let open_time = started.elapsed();
    let migrated = DeckSession::open_from_update(&stored, 43502).unwrap();
    let story = defaults_story(&migrated, "Direct all caps");
    let context = EditCtx::local("fixture");
    let mut starts = Vec::new();
    let mut offset = 0;
    for paragraph in fresh.story(&story).unwrap().paragraphs {
        starts.push(offset);
        offset += paragraph
            .runs
            .iter()
            .map(|run| run.text.len() as u32)
            .sum::<u32>()
            + 1;
    }
    let typed = TextStyle {
        font_size_pt: Some(32.0),
        ..TextStyle::default()
    };
    for session in [&migrated, &fresh] {
        for index in starts
            .iter()
            .rev()
            .flat_map(|start| [start + 60, start + 30, start + 3])
        {
            session
                .delete_text(&context, &story, index, index + 2)
                .unwrap();
            session
                .insert_text(&context, &story, index, "X", &typed)
                .unwrap();
        }
    }
    let update = migrated.encode_state_as_update_v1();
    let started = std::time::Instant::now();
    let attached = DeckSession::open_from_update_with_source(&update, &source, 43503).unwrap();
    let attach_time = started.elapsed();
    eprintln!("TIMING attach {attach_time:?} open {open_time:?}");
    assert_eq!(
        attached.story(&story).unwrap(),
        fresh.story(&story).unwrap()
    );
    assert!(
        attach_time < open_time * 8,
        "attaching took {attach_time:?}, a fresh open {open_time:?}"
    );
}
fn small_caps() -> TextStyle {
    TextStyle {
        font_size_pt: Some(32.0),
        color: Some("#101828".to_owned()),
        caps: Some(TextCaps::Small),
        ..TextStyle::default()
    }
}

fn look(text: &str, caps: Option<TextCaps>, color: &str) -> RunLook {
    (text.to_owned(), caps, Some(color.to_owned()))
}

fn defaults_story(session: &DeckSession, shape: &str) -> String {
    session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|candidate| candidate.name == shape)
        .unwrap()
        .text_stories[0]
        .id
        .clone()
}

fn run_looks(session: &DeckSession, story: &str) -> Vec<RunLook> {
    session.story(story).unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.clone(), run.style.caps, run.style.color.clone()))
        .collect()
}

fn defaults_part(part: &str) -> String {
    let parts = ooxml_opc::unzip_parts(V2_1_DEFAULTS_SOURCE).unwrap();
    let (_, bytes) = parts.into_iter().find(|(name, _)| name == part).unwrap();
    String::from_utf8(bytes).unwrap()
}

/// The defaults deck with each `(part, from, to)` replacement applied once.
fn defaults_variant(edits: &[(&str, &str, &str)]) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(V2_1_DEFAULTS_SOURCE).unwrap();
    for (part, from, to) in edits {
        let (_, bytes) = parts.iter_mut().find(|(name, _)| name == part).unwrap();
        let xml = String::from_utf8(bytes.clone()).unwrap();
        assert!(xml.contains(from), "{part} lacks {from}");
        *bytes = xml.replacen(from, to, 1).into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Seeds `source` the way a release before schema 2.2 stored it: no caps, and
/// run colours resolved without the slide colour map.
fn legacy_seed(source: &[u8], client_id: u64) -> Vec<u8> {
    let stored = stored_2_0(source, client_id, |value| {
        without_keys(value, &["caps", "colorMap", "colorMapOverride"])
    });
    restamped(&stored, Some(2.1))
}

fn meta_has(update: &[u8], key: &str) -> bool {
    let doc = hydrated(update);
    meta(&doc).contains_key(&doc.transact(), key)
}

fn meta_string(update: &[u8], key: &str) -> String {
    let doc = hydrated(update);
    match meta(&doc).get(&doc.transact(), key) {
        Some(Out::Any(Any::String(value))) => value.to_string(),
        other => panic!("{key} is {other:?}"),
    }
}

#[test]
fn migrating_and_attaching_grows_linearly_with_the_number_of_stories() {
    let slide = defaults_part(DEFAULTS_SLIDE);
    let start = slide
        .find(r#"<p:sp><p:nvSpPr><p:cNvPr id="10" name="Direct all caps"/>"#)
        .unwrap();
    let end = start + slide[start..].find("</p:sp>").unwrap() + "</p:sp>".len();
    let template = &slide[start..end];
    let mut timings = Vec::new();
    for count in STORY_COUNTS {
        let copies: String = (0..count)
            .map(|index| {
                template.replace(
                    r#"id="10" name="Direct all caps""#,
                    &format!(r#"id="{}" name="Copy {index}""#, 1000 + index),
                )
            })
            .collect();
        let source = defaults_variant(&[(
            DEFAULTS_SLIDE,
            "</p:spTree>",
            &format!("{copies}</p:spTree>"),
        )]);
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
                let context = EditCtx::local("fixture");
                fresh
                    .delete_text(&context, story, index * 2, index * 2 + 1)
                    .unwrap();
            }
        }
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let started = std::time::Instant::now();
        let attached = DeckSession::open_from_update_with_source(&update, &source, 44502).unwrap();
        timings.push(started.elapsed());
        for story in &stories {
            assert_eq!(attached.story(story).unwrap(), fresh.story(story).unwrap());
        }
    }
    for step in timings.windows(2) {
        assert!(
            step[1] < step[0] * 8,
            "{STORY_COUNTS:?} stories took {timings:?}"
        );
    }
}

#[test]
fn baseline_and_spacing_recover_text_restored_by_undo() {
    let baselines: &[u8] =
        include_bytes!("../../pptx-render/tests/fixtures/text-baseline-script.pptx");
    let spacing: &[u8] = include_bytes!("fixtures/run-spacing-shadow.pptx");
    for (client_id, source, key) in [
        (44700, baselines, "baselinePct"),
        (44800, spacing, "spacingPt"),
    ] {
        let fresh = DeckSession::open(source, client_id).unwrap();
        let (story, offset) = first_offset(&fresh.snapshot().unwrap(), |style| {
            style.baseline_pct.is_some() || style.spacing_pt.is_some()
        });
        let stored = stored_2_0(source, client_id + 1, |value| without_keys(value, &[key]));
        let migrated = DeckSession::open_from_update(&stored, client_id + 2).unwrap();
        migrated
            .delete_text(&EditCtx::local("fixture"), &story, offset, offset + 1)
            .unwrap();
        assert!(migrated.undo());
        let attached = DeckSession::open_from_update_with_source(
            &migrated.encode_state_as_update_v1(),
            source,
            client_id + 3,
        )
        .unwrap();
        assert_eq!(
            attached.story(&story).unwrap(),
            fresh.story(&story).unwrap(),
            "{key}"
        );
        assert_eq!(attached.save().unwrap(), fresh.save().unwrap(), "{key}");
    }
}

/// The first story offset whose run style matches.
fn first_offset(snapshot: &DeckSnapshot, matches: impl Fn(&TextStyle) -> bool) -> (String, u32) {
    fn find(
        shapes: &[ShapeSnapshot],
        matches: &impl Fn(&TextStyle) -> bool,
    ) -> Option<(String, u32)> {
        for shape in shapes {
            for story in &shape.text_stories {
                let mut offset = 0;
                for paragraph in &story.paragraphs {
                    for run in &paragraph.runs {
                        if matches(&run.style) {
                            return Some((story.id.clone(), offset));
                        }
                        offset += run.text.encode_utf16().count() as u32;
                    }
                    offset += 1;
                }
            }
            if let Some(found) = find(&shape.children, matches) {
                return Some(found);
            }
        }
        None
    }
    snapshot
        .slides
        .iter()
        .find_map(|slide| find(&slide.shapes, &matches))
        .unwrap()
}

#[test]
fn a_released_update_edited_across_runs_and_surrogates_recovers_by_its_clocks() {
    let fresh = DeckSession::open(V2_1_EDITS_SOURCE, 44900).unwrap();
    let story = defaults_story(&fresh, "Direct all caps");
    let context = EditCtx::local("fixture");
    let typed = TextStyle {
        font_size_pt: Some(32.0),
        color: Some("#123456".to_owned()),
        ..TextStyle::default()
    };
    fresh.delete_text(&context, &story, 1, 3).unwrap();
    fresh.delete_text(&context, &story, 9, 11).unwrap();
    fresh
        .insert_text(&context, &story, 0, "NEW", &typed)
        .unwrap();
    fresh.delete_text(&context, &story, 27, 29).unwrap();
    fresh
        .insert_text(&context, &story, 19, "\u{1D402}", &typed)
        .unwrap();
    let attached =
        DeckSession::open_from_update_with_source(V2_1_EDITS_UPDATE, V2_1_EDITS_SOURCE, 44901)
            .unwrap();
    assert_eq!(
        attached.story(&story).unwrap(),
        fresh.story(&story).unwrap()
    );
    assert_eq!(attached.save().unwrap(), fresh.save().unwrap());
    let carried = attached.encode_state_as_update_v1();
    for key in ["capsPendingSource", "colorsPendingSource"] {
        assert!(!meta_has(&carried, key), "{key}");
    }
}

#[test]
fn baseline_and_spacing_recover_a_long_paragraph_restored_by_undo() {
    let (source, story) = long_formatted_paragraph();
    for (client_id, replayable) in [(45000, true), (45100, false)] {
        let migrated = DeckSession::open_from_update(
            &long_paragraph_seed(&source, client_id, replayable),
            client_id + 1,
        )
        .unwrap();
        migrated
            .delete_text(&EditCtx::local("fixture"), &story, 0, LONG_PARAGRAPH)
            .unwrap();
        assert!(migrated.undo());
        let attached = DeckSession::open_from_update_with_source(
            &migrated.encode_state_as_update_v1(),
            &source,
            client_id + 2,
        )
        .unwrap();
        let fresh = DeckSession::open(&source, client_id + 3).unwrap();
        assert_eq!(
            attached.story(&story).unwrap(),
            fresh.story(&story).unwrap()
        );
        assert_eq!(attached.save().unwrap(), source);
    }
}

#[test]
fn a_save_refuses_to_drop_baseline_or_spacing_it_could_not_recover() {
    let (source, story) = long_formatted_paragraph();
    let migrated =
        DeckSession::open_from_update(&long_paragraph_seed(&source, 45200, true), 45201).unwrap();
    let context = EditCtx::local("fixture");
    migrated
        .delete_text(&context, &story, 0, LONG_PARAGRAPH)
        .unwrap();
    migrated
        .insert_text(
            &context,
            &story,
            0,
            &"Z".repeat(2100),
            &TextStyle::default(),
        )
        .unwrap();
    let attached = DeckSession::open_from_update_with_source(
        &migrated.encode_state_as_update_v1(),
        &source,
        45202,
    )
    .unwrap();
    let carried = attached.encode_state_as_update_v1();
    meta_string(&carried, "baselinesPendingSource");
    match attached.save() {
        Err(EditError::Write(message)) => assert!(message.contains("could not be recovered")),
        other => panic!("{other:?}"),
    }
    attached.delete_text(&context, &story, 0, 2100).unwrap();
    attached.save().unwrap();
}

const LONG_PARAGRAPH: u32 = 2000;

/// The defaults deck with a 2,000-character raised, tracked run, and its story.
fn long_formatted_paragraph() -> (Vec<u8>, String) {
    let text: String = "Mixed Case Title "
        .repeat(120)
        .chars()
        .take(LONG_PARAGRAPH as usize)
        .collect();
    let source = defaults_variant(&[
        (
            DEFAULTS_SLIDE,
            r#"<a:rPr lang="en-US" sz="3200" cap="all">"#,
            r#"<a:rPr lang="en-US" sz="3200" cap="all" spc="300" baseline="30000">"#,
        ),
        (
            DEFAULTS_SLIDE,
            "<a:t>Mixed Case Title</a:t>",
            &format!("<a:t>{text}</a:t>"),
        ),
    ]);
    let story = defaults_story(
        &DeckSession::open(&source, 44999).unwrap(),
        "Direct all caps",
    );
    (source, story)
}

/// A 2.0 seed of `source` without baselines or spacing; unless `replayable`,
/// its stored package no longer reproduces the seed.
fn long_paragraph_seed(source: &[u8], client_id: u64, replayable: bool) -> Vec<u8> {
    let seeded = stored_2_0(source, client_id, |value| {
        without_keys(value, &["baselinePct", "spacingPt"])
    });
    if replayable {
        return seeded;
    }
    let doc = hydrated(&seeded);
    let mut package: serde_json::Value = serde_json::from_str(&package_json(&seeded)).unwrap();
    without_keys(&mut package, &["fontSizePt"]);
    let stored_meta = meta(&doc);
    stored_meta.insert(
        &mut doc.transact_mut(),
        "packageJson",
        Any::Buffer(serde_json::to_vec(&package).unwrap().into()),
    );
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}
