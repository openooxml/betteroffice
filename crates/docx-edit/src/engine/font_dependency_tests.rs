use super::*;
use serde_json::{Value, json};

struct Pair {
    shared: EngineSession,
    oracle: EngineSession,
    request: Value,
    at: u32,
}

impl Pair {
    fn new(bytes: &[u8], font: u32) -> Self {
        let seed = EditingDoc::new(75400);
        crate::seed_from_docx(&seed, bytes).unwrap();
        let state = seed.encode_state_as_update_v1();
        let shared = EngineSession::new(75401);
        let oracle = EngineSession::new(75401);
        for engine in [&shared, &oracle] {
            engine.doc().apply_host_update_v1(&state).unwrap();
            engine
                .doc()
                .set_note_separator_state(seed.note_separator_state().unwrap());
            engine.set_local_lowering(true);
        }
        oracle.font_dependency_work.borrow_mut().oracle = true;
        let request = lowering_pages::region_request(&shared, bytes, font);
        let at = shared
            .doc()
            .paragraph_index("body")
            .unwrap()
            .para_at(0)
            .unwrap()
            .node_start;
        Self {
            shared,
            oracle,
            request,
            at,
        }
    }

    fn reset_work(&self) {
        *self.shared.font_dependency_work.borrow_mut() = FontDependencyWork::default();
        *self.oracle.font_dependency_work.borrow_mut() = FontDependencyWork {
            oracle: true,
            ..Default::default()
        };
    }

    fn compare(&self) {
        let snapshot = |engine: &EngineSession| {
            let pagination = engine.pagination.borrow();
            let render = engine.render.borrow();
            (
                serde_json::to_vec(render.stories["body"].blocks.as_ref()).unwrap(),
                serde_json::to_vec(pagination.input.as_ref().unwrap()).unwrap(),
                serde_json::to_vec(pagination.layout.as_ref().unwrap()).unwrap(),
                pagination.block_fingerprints.clone(),
                engine.with_display_list(Clone::clone).unwrap(),
            )
        };
        assert_eq!(snapshot(&self.shared), snapshot(&self.oracle));
        assert_eq!(
            self.shared.font_dependency_work.borrow().reuse_sets,
            self.oracle.font_dependency_work.borrow().reuse_sets,
        );
        assert_eq!(
            self.shared.stats().resident_measure_calls,
            self.oracle.stats().resident_measure_calls,
        );
        assert_eq!(
            self.shared.stats().resident_reused_blocks,
            self.oracle.stats().resident_reused_blocks,
        );
    }

    fn layout(&self, trigger: RelayoutTrigger) {
        self.reset_work();
        let request = self.request.to_string();
        let extras = json!({"fontChains": self.request["measurement"]["fontChains"]}).to_string();
        let frame = |engine: &EngineSession| {
            engine
                .layout_regions_for_trigger(&request, None, trigger)
                .unwrap();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.build_display_list_frame(&extras, epoch).unwrap()
        };
        assert_eq!(frame(&self.shared), frame(&self.oracle));
        self.compare();
    }

    fn edit(&self, insert: bool) {
        self.reset_work();
        let frame = |engine: &EngineSession| {
            let range = crate::StoryRange::new("body", self.at, self.at + u32::from(!insert));
            engine
                .edit_resident_text(range, insert.then_some("x"), true)
                .unwrap();
            let epoch = engine.display.borrow().binary_frame_epoch;
            engine.apply_and_layout("body", epoch).unwrap()
        };
        assert_eq!(frame(&self.shared), frame(&self.oracle));
        self.compare();
    }

    fn assert_identity_saving(&self) {
        let shared = self.shared.font_dependency_work.borrow();
        let oracle = self.oracle.font_dependency_work.borrow();
        assert_eq!(shared.validations, 0);
        assert!(shared.identity_skips > 0);
        assert_eq!(shared.identity_skips, oracle.validations);
        assert_eq!(oracle.identity_skips, 0);
        assert!(!shared.reuse_sets.is_empty());
    }

    fn change_chain(&mut self, key: &str, chain: Value) {
        let chains = self.request["measurement"]["fontChains"]
            .as_object_mut()
            .unwrap();
        assert!(chains.contains_key(key));
        chains.insert(key.to_owned(), chain);
    }
}

fn edit_stream(pair: &Pair, trigger: RelayoutTrigger) {
    pair.layout(trigger);
    pair.layout(trigger);
    pair.assert_identity_saving();
    for insert in [true, true, false, true, false, false] {
        pair.edit(insert);
        pair.layout(RelayoutTrigger::Interactive);
        pair.assert_identity_saving();
    }
}

fn font_chain_pair(font: u32) -> Pair {
    let bytes = lowering_pages::with_body(&format!(
        "{}{}",
        lowering_pages::p(
            "00000001",
            r#"<w:r><w:rPr><w:rFonts w:ascii="Requested" w:hAnsi="Requested"/></w:rPr><w:t>Requested text</w:t></w:r>"#,
        ),
        lowering_pages::p(
            "00000002",
            r#"<w:r><w:rPr><w:rFonts w:ascii="Stable" w:hAnsi="Stable"/></w:rPr><w:t>Stable text</w:t></w:r>"#,
        ),
    ));
    let mut pair = Pair::new(&bytes, font);
    pair.request["notes"]["contents"] = json!([]);
    pair
}

#[test]
fn resident_typing_keeps_font_identity_for_the_next_pass() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let pair = font_chain_pair(font);
    pair.layout(RelayoutTrigger::Interactive);
    let before = pair.shared.stats();
    pair.edit(true);
    let after = pair.shared.stats();
    assert_eq!(
        after.resident_measure_calls - before.resident_measure_calls,
        1
    );
    assert_eq!(
        after.resident_reused_blocks - before.resident_reused_blocks,
        1
    );
    assert!(
        pair.shared
            .font_dependency_work
            .borrow()
            .reuse_sets
            .is_empty()
    );
    let chains: BTreeMap<String, Vec<u32>> =
        serde_json::from_value(pair.request["measurement"]["fontChains"].clone()).unwrap();
    for engine in [&pair.shared, &pair.oracle] {
        assert_eq!(
            engine.pagination.borrow().measured_font_chains.as_ref(),
            Some(&chains),
        );
    }
    pair.layout(RelayoutTrigger::Interactive);
    pair.assert_identity_saving();
}

#[test]
fn reentrant_font_layout_matches_per_block_oracle() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let other = docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap();
    let pair = font_chain_pair(font);
    pair.layout(RelayoutTrigger::Interactive);
    let mut nested_request = pair.request.clone();
    nested_request["measurement"]["fontChains"]["requested|0|0"] = json!([other]);
    let nested_chains: BTreeMap<String, Vec<u32>> =
        serde_json::from_value(nested_request["measurement"]["fontChains"].clone()).unwrap();
    pair.reset_work();
    let frame = |engine: &EngineSession| {
        engine
            .edit_resident_text(
                crate::StoryRange::new("body", pair.at, pair.at),
                Some("x"),
                true,
            )
            .unwrap();
        let epoch = engine.display.borrow().binary_frame_epoch;
        let mut ticks = 0_u32;
        let mut nested = false;
        let (frame, profile) = engine
            .apply_and_layout_profiled("body", epoch, &mut || {
                ticks += 1;
                if ticks == 2 {
                    engine
                        .layout_regions_for_trigger(
                            &nested_request.to_string(),
                            None,
                            RelayoutTrigger::Interactive,
                        )
                        .unwrap();
                    nested = true;
                }
                f64::from(ticks)
            })
            .unwrap();
        assert!(nested);
        assert_eq!(profile.lower_ms, 1.0);
        assert_eq!(profile.measure_ms, 0.0);
        let pagination = engine.pagination.borrow();
        assert_eq!(
            pagination.measured_font_chains.as_ref(),
            Some(&nested_chains),
        );
        assert!(
            pagination
                .measured_font_dependencies
                .iter()
                .all(|dependencies| dependencies.matches(FontChains::BTree(&nested_chains)))
        );
        frame
    };
    assert_eq!(frame(&pair.shared), frame(&pair.oracle));
    pair.compare();
    pair.layout(RelayoutTrigger::Interactive);
    let work = pair.shared.font_dependency_work.borrow();
    assert_eq!(work.identity_skips, 0);
    assert!(work.validations > 0);
    assert!(!work.reuse_sets[0].contains(&0));
    assert!(work.reuse_sets[0].contains(&1));
    drop(work);
    pair.layout(RelayoutTrigger::Interactive);
    pair.assert_identity_saving();
}

#[test]
fn failed_font_pass_then_resident_edit_matches_per_block_oracle() {
    for trigger in [RelayoutTrigger::Interactive, RelayoutTrigger::Bulk] {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
        let other = docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap();
        let mut pair = font_chain_pair(font);
        pair.layout(trigger);
        pair.at = pair
            .shared
            .doc()
            .paragraph_index("body")
            .unwrap()
            .para_at(1)
            .unwrap()
            .node_start;
        let original_chains: BTreeMap<String, Vec<u32>> =
            serde_json::from_value(pair.request["measurement"]["fontChains"].clone()).unwrap();
        let mut failed_request = pair.request.clone();
        failed_request["measurement"]["fontChains"]["requested|0|0"] = json!([other]);
        failed_request["regions"]["sections"][0]["properties"]["headerReferences"] =
            json!([{"type": "default", "rId": "missing-font-dependency-test"}]);
        let failed_chains: BTreeMap<String, Vec<u32>> =
            serde_json::from_value(failed_request["measurement"]["fontChains"].clone()).unwrap();
        pair.reset_work();
        for engine in [&pair.shared, &pair.oracle] {
            let mut prepared = engine
                .prepare_region_layout(&failed_request.to_string(), None, trigger)
                .unwrap();
            assert!(prepared.measure(usize::MAX).unwrap());
            {
                let pagination = engine.pagination.borrow();
                assert!(pagination.measured_font_chains.is_none());
                assert!(
                    pagination.measured_font_dependencies[0]
                        .matches(FontChains::BTree(&failed_chains))
                );
                assert!(
                    !pagination.measured_font_dependencies[0]
                        .matches(FontChains::BTree(&original_chains))
                );
            }
            assert_eq!(engine.font_dependency_work.borrow().reuse_sets.len(), 1);
            let error = engine.finish_region_layout(prepared).err().unwrap();
            assert!(error.contains("hf:missing-font-dependency-test"), "{error}");
            assert!(engine.pagination.borrow().measured_font_chains.is_none());
        }
        pair.edit(true);
        assert_eq!(pair.shared.font_dependency_work.borrow().identity_skips, 0);
        if trigger == RelayoutTrigger::Interactive {
            assert!(
                !pair
                    .shared
                    .font_dependency_work
                    .borrow()
                    .reuse_sets
                    .is_empty()
            );
        }
        for reused in &pair.shared.font_dependency_work.borrow().reuse_sets {
            assert!(!reused.contains(&0));
        }
        pair.layout(RelayoutTrigger::Interactive);
        pair.assert_identity_saving();
    }
}

#[test]
fn shared_font_dependencies_match_per_block_oracle_on_synthetic_edits() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for bytes in [
        preview_fixture::plain(),
        preview_fixture::nested(),
        preview_fixture::breaks(),
        preview_fixture::drawings(),
        lowering_pages::unrevised_docx(),
    ] {
        for trigger in [RelayoutTrigger::Interactive, RelayoutTrigger::Bulk] {
            edit_stream(&Pair::new(&bytes, font), trigger);
        }
    }
}

#[test]
fn shared_font_dependencies_match_per_block_oracle_on_corpus_edits() {
    for (name, bytes) in preview_fixture::corpus() {
        for trigger in [RelayoutTrigger::Interactive, RelayoutTrigger::Bulk] {
            let fonts = docx_layout::MeasureFonts::default();
            let _scope = fonts.enter();
            let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
            let mut pair = Pair::new(bytes, font);
            edit_stream(&pair, trigger);
            let key = pair.request["measurement"]["fontChains"]
                .as_object()
                .unwrap()
                .keys()
                .next()
                .unwrap()
                .clone();
            font_changes(&mut pair, trigger, &key, font, false);
            assert!(pair.shared.pagination.borrow().input.is_some(), "{name}");
        }
    }
}

fn font_changes(
    pair: &mut Pair,
    trigger: RelayoutTrigger,
    key: &str,
    font: u32,
    empty_chain: bool,
) {
    pair.layout(trigger);
    pair.layout(trigger);
    pair.assert_identity_saving();
    pair.edit(true);
    pair.edit(false);
    pair.layout(trigger);
    let appended = docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap();
    pair.layout(trigger);
    pair.assert_identity_saving();

    pair.request["measurement"]["fontChains"]["unused|0|0"] = json!([appended]);
    pair.layout(trigger);
    assert!(pair.shared.font_dependency_work.borrow().validations > 0);
    assert_eq!(pair.shared.font_dependency_work.borrow().identity_skips, 0);
    pair.layout(trigger);
    pair.assert_identity_saving();

    let mut chains = vec![json!([font, appended]), json!([appended])];
    if empty_chain {
        chains.push(json!([]));
    }
    chains.push(json!([font]));
    for chain in chains {
        pair.change_chain(key, chain);
        pair.layout(trigger);
        assert!(pair.shared.font_dependency_work.borrow().validations > 0);
        assert_eq!(pair.shared.font_dependency_work.borrow().identity_skips, 0);
        if key == "requested|0|0" {
            let work = pair.shared.font_dependency_work.borrow();
            assert!(!work.reuse_sets[0].contains(&0));
            assert!(work.reuse_sets[0].contains(&1));
        }
        pair.edit(true);
        pair.edit(false);
        pair.layout(RelayoutTrigger::Interactive);
        pair.assert_identity_saving();
        pair.layout(trigger);
    }

    docx_layout::clear_measure_fonts();
    assert_eq!(
        docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
        font,
    );
    pair.layout(trigger);
    assert_eq!(pair.shared.font_dependency_work.borrow().identity_skips, 0);
    assert_eq!(pair.shared.font_dependency_work.borrow().validations, 0);
    assert!(
        pair.shared
            .font_dependency_work
            .borrow()
            .reuse_sets
            .is_empty()
    );
    pair.layout(trigger);
    pair.assert_identity_saving();
}

#[test]
fn shared_font_dependencies_match_per_block_oracle_after_font_changes() {
    for trigger in [RelayoutTrigger::Interactive, RelayoutTrigger::Bulk] {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
        let bytes = lowering_pages::with_body(&format!(
            "{}{}",
            lowering_pages::p(
                "00000001",
                r#"<w:r><w:rPr><w:rFonts w:ascii="Requested" w:hAnsi="Requested"/></w:rPr><w:t>Requested text</w:t></w:r>"#,
            ),
            lowering_pages::p(
                "00000002",
                r#"<w:r><w:rPr><w:rFonts w:ascii="Stable" w:hAnsi="Stable"/></w:rPr><w:t>Stable text</w:t></w:r>"#,
            ),
        ));
        let mut pair = Pair::new(&bytes, font);
        font_changes(&mut pair, trigger, "requested|0|0", font, true);
    }
}

#[test]
fn unchanged_font_identity_does_not_validate_unknown_records() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let pair = Pair::new(&preview_fixture::plain(), font);
    pair.layout(RelayoutTrigger::Interactive);
    for engine in [&pair.shared, &pair.oracle] {
        let mut pagination = engine.pagination.borrow_mut();
        let mut dependencies = pagination.measured_font_dependencies.clone();
        dependencies[0] = FontChainDependencies::unknown();
        let chains = pagination.measured_font_chains.clone();
        pagination.set_font_dependencies(dependencies, chains.as_ref());
    }
    pair.layout(RelayoutTrigger::Interactive);
    pair.assert_identity_saving();
    assert!(!pair.shared.font_dependency_work.borrow().reuse_sets[0].contains(&0));
    pair.layout(RelayoutTrigger::Interactive);
    assert!(pair.shared.font_dependency_work.borrow().reuse_sets[0].contains(&0));
}

#[test]
fn appended_referenced_fonts_keep_the_availability_guard() {
    for trigger in [RelayoutTrigger::Interactive, RelayoutTrigger::Bulk] {
        let fonts = docx_layout::MeasureFonts::default();
        let _scope = fonts.enter();
        let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
        let mut pair = Pair::new(&preview_fixture::plain(), font);
        for chain in pair.request["measurement"]["fontChains"]
            .as_object_mut()
            .unwrap()
            .values_mut()
        {
            *chain = json!([font + 1]);
        }
        pair.layout(trigger);
        pair.layout(trigger);
        pair.assert_identity_saving();
        assert_eq!(
            docx_layout::register_measure_font_bytes(lowering_pages::OTHER_FONT).unwrap(),
            font + 1,
        );
        pair.layout(trigger);
        assert_eq!(pair.shared.font_dependency_work.borrow().identity_skips, 0);
        assert!(
            pair.shared
                .font_dependency_work
                .borrow()
                .reuse_sets
                .is_empty()
        );
        pair.layout(trigger);
        pair.assert_identity_saving();
        pair.edit(true);
        pair.edit(false);
    }
}
