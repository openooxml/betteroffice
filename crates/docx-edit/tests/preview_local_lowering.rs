#[allow(dead_code)]
#[path = "support/preview_fixture.rs"]
mod fixture;

use RevisionPreview::{Accepted, Rejected};
use docx_edit::bridge::{RenderEnv, RevisionPreview, yrs_doc_to_mapped_layout_blocks};
use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, StoryRange, seed_from_docx};

fn seeded(bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(75200);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let primer = RenderEnv {
        show_hidden_text: true,
        ..RenderEnv::default()
    };
    engine.lower_story_json("body", &primer).unwrap();
    engine
}

fn oracle(engine: &EngineSession, env: &RenderEnv) {
    let actual = engine.lower_story_json("body", env).unwrap();
    let (blocks, _) = yrs_doc_to_mapped_layout_blocks(engine.doc(), "body", env).unwrap();
    assert_eq!(actual, serde_json::to_string(&blocks).unwrap());
}

fn decision_stream(bytes: &[u8], patches: bool, fields: bool) {
    for seed in 0..40 {
        let engine = seeded(bytes);
        if fields {
            fixture::stamp_fields(&engine);
        }
        let ids = fixture::ids(&engine);
        assert!(!ids.is_empty());
        let mut env = RenderEnv::default();
        let mut random = fixture::Random::new(seed);
        oracle(&engine, &env);
        for _ in 0..24 {
            fixture::decide(&mut env, &ids, &mut random);
            let before = engine.stats();
            oracle(&engine, &env);
            if patches {
                assert_eq!(
                    engine.stats().lower_preview_patches,
                    before.lower_preview_patches + 1
                );
                assert_eq!(engine.stats().lower_cache_misses, before.lower_cache_misses);
            }
        }
    }
}

#[test]
fn preview_local_plain_paragraph_decision_streams() {
    decision_stream(&fixture::plain(), true, false);
}

#[test]
fn preview_local_nested_tables_and_content_controls() {
    decision_stream(&fixture::nested(), true, false);
}

#[test]
fn preview_local_numbered_paragraphs_and_tracked_breaks() {
    decision_stream(&fixture::breaks(), false, false);
    let engine = seeded(&fixture::breaks());
    let env = RenderEnv::default();
    oracle(&engine, &env);
    let before = engine.stats();
    oracle(&engine, &env.with_revision_preview("1", Rejected));
    assert_eq!(
        engine.stats().lower_preview_patches,
        before.lower_preview_patches + 1
    );
}

#[test]
fn preview_local_drawings_and_images() {
    decision_stream(&fixture::drawings(), true, false);
}

#[test]
fn preview_local_fields_inside_revisions() {
    decision_stream(&fixture::fields(), true, true);
    decision_stream(&fixture::hidden_fields(), false, true);
    let engine = seeded(&fixture::hidden_fields());
    fixture::stamp_fields(&engine);
    oracle(&engine, &RenderEnv::default());
    assert!(
        !engine
            .lower_story_json("body", &RenderEnv::default())
            .unwrap()
            .contains("Cached result")
    );
    let before = engine.stats();
    oracle(
        &engine,
        &RenderEnv::default().with_revision_preview("field-1", Rejected),
    );
    assert!(
        engine
            .lower_story_json(
                "body",
                &RenderEnv::default().with_revision_preview("field-1", Rejected)
            )
            .unwrap()
            .contains("Cached result")
    );
    assert_eq!(
        engine.stats().lower_preview_fallbacks,
        before.lower_preview_fallbacks + 1
    );
}

#[test]
fn preview_local_sequence_fields_patch_only_units_without_them() {
    decision_stream(&fixture::sequence(), false, true);
    for (id, patches) in [("1", 1), ("3", 0), ("field-2", 0)] {
        let engine = seeded(&fixture::sequence());
        fixture::stamp_fields(&engine);
        let mut env = RenderEnv::default();
        oracle(&engine, &env);
        for decision in [Accepted, Rejected] {
            env.revision_preview.insert(id.to_owned(), decision);
            let before = engine.stats();
            oracle(&engine, &env);
            assert_eq!(
                engine.stats().lower_preview_patches,
                before.lower_preview_patches + patches,
                "{id}"
            );
            assert_eq!(
                engine.stats().lower_preview_fallbacks,
                before.lower_preview_fallbacks + 1 - patches,
                "{id}"
            );
        }
    }
}

#[test]
fn preview_local_random_edit_and_decision_streams() {
    for seed in 0..40 {
        let engine = seeded(&fixture::plain());
        let mut env = RenderEnv::default();
        let mut random = fixture::Random::new(seed);
        oracle(&engine, &env);
        for _ in 0..24 {
            let suggested = random.next().is_multiple_of(2);
            let plain = EditCtx::local("", "");
            let suggesting = EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting();
            let ctx = if suggested { &suggesting } else { &plain };
            if random.next().is_multiple_of(2) {
                engine
                    .doc()
                    .insert_text(ctx, Position::new("body", 1), "x", FormatPolicy::Plain)
                    .unwrap();
            } else {
                engine
                    .doc()
                    .delete_range(ctx, StoryRange::new("body", 1, 2))
                    .unwrap();
            }
            oracle(&engine, &env);
            let ids = fixture::ids(&engine);
            fixture::decide(&mut env, &ids, &mut random);
            let before = engine.stats();
            oracle(&engine, &env);
            assert_eq!(
                engine.stats().lower_preview_patches,
                before.lower_preview_patches + 1
            );
        }
    }
}

#[test]
fn preview_local_tracked_change_corpus() {
    for (name, bytes) in fixture::corpus() {
        let engine = seeded(bytes);
        let ids = fixture::ids(&engine);
        assert!(!ids.is_empty(), "{name}");
        oracle(&engine, &RenderEnv::default());
        for id in &ids {
            for decision in [Accepted, Rejected] {
                oracle(
                    &engine,
                    &RenderEnv::default().with_revision_preview(id, decision),
                );
            }
        }
        for seed in 0..40 {
            let mut env = RenderEnv::default();
            let mut random = fixture::Random::new(seed);
            for _ in 0..12 {
                fixture::decide(&mut env, &ids, &mut random);
                oracle(&engine, &env);
            }
        }
    }
}
