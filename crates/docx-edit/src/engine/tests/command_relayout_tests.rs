use super::*;

/// The worker's text-command loop over one engine: each command edits, lays
/// the document out again and builds the frame the host applies.
struct CommandHost<'a> {
    engine: &'a EngineSession,
    request: String,
    extras: String,
    retained: HashMap<u64, crate::frame_delta::TestFramePage>,
}

impl<'a> CommandHost<'a> {
    fn open(engine: &'a EngineSession, font: u32) -> Self {
        engine.set_local_lowering(true);
        let request = small_page_request(font);
        engine
            .layout_document_with_regions_retained(&request)
            .unwrap();
        let extras = json!({"fontChains": {"calibri|0|0": [font]}}).to_string();
        let mut retained = HashMap::new();
        let frame = engine.build_display_list_frame(&extras, 0).unwrap();
        crate::frame_delta::apply_placeholder_test_frame(&frame, &mut retained);
        Self {
            engine,
            request,
            extras,
            retained,
        }
    }

    fn insert(&mut self, offset: u32, text: &str) {
        self.engine
            .edit_resident_text(
                crate::StoryRange::new("body", offset, offset),
                Some(text),
                true,
            )
            .unwrap();
        self.relayout(Some("body"));
    }

    fn delete(&mut self, start: u32, end: u32) {
        self.engine
            .edit_resident_text(crate::StoryRange::new("body", start, end), None, true)
            .unwrap();
        self.relayout(Some("body"));
    }

    fn relayout(&mut self, story: Option<&str>) {
        self.engine
            .layout_document_with_regions_after_edit(&self.request, story)
            .unwrap();
        let epoch = self.engine.display.borrow().binary_frame_epoch;
        let frame = self
            .engine
            .build_display_list_frame(&self.extras, epoch)
            .unwrap();
        crate::frame_delta::apply_placeholder_test_frame(&frame, &mut self.retained);
    }

    fn assert_matches_cold(&mut self, label: &str) {
        let epoch = self.engine.display.borrow().binary_frame_epoch;
        let frame = self
            .engine
            .build_display_list_frame(&self.extras, epoch)
            .unwrap();
        assert_returned_frame_matches_cold(self.engine, &self.request, &frame, &mut self.retained);
        let cold = EngineSession::new(9712);
        cold.doc()
            .apply_update_v1(&self.engine.doc().encode_state_as_update_v1())
            .unwrap();
        cold.layout_regions_for_trigger(&self.request, None, RelayoutTrigger::Open)
            .unwrap();
        assert_eq!(
            self.engine.retained_headers_footers_json().unwrap(),
            cold.retained_headers_footers_json().unwrap(),
            "{label}: headers and footers"
        );
    }
}

fn paragraph_end(engine: &EngineSession, index: usize) -> u32 {
    let paragraph = engine.doc().paragraphs("body").unwrap()[index]
        .para_id
        .clone();
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_span(&paragraph)
        .unwrap()
        .1
}

/// The work one keystroke does that would scale with the document.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct KeystrokeWork {
    paragraphs_flushed: usize,
    full_lowerings: u64,
    transcoded: usize,
    resolved_lines: usize,
    measured_blocks: u64,
    paginated_pages: usize,
    full_paginations: u64,
    rebuilt_display_pages: u64,
    full_display_builds: u64,
}

fn keystroke_work(engine: &EngineSession, keystroke: impl FnOnce()) -> KeystrokeWork {
    crate::bridge::PARAGRAPHS_FLUSHED.with(|flushed| flushed.set(0));
    docx_layout::work_counters::take();
    let before = engine.stats();
    keystroke();
    let after = engine.stats();
    let layout = docx_layout::work_counters::take();
    KeystrokeWork {
        paragraphs_flushed: crate::bridge::PARAGRAPHS_FLUSHED.with(Cell::get),
        full_lowerings: after.lower_cache_misses - before.lower_cache_misses,
        transcoded: layout.transcoded,
        resolved_lines: layout.resolved_lines,
        measured_blocks: after.resident_measure_calls - before.resident_measure_calls,
        paginated_pages: after.rebuilt_pages,
        full_paginations: (after.pagination_calls - before.pagination_calls)
            - (after.incremental_pagination_calls - before.incremental_pagination_calls),
        rebuilt_display_pages: after.rebuilt_display_pages - before.rebuilt_display_pages,
        full_display_builds: (after.display_builds - before.display_builds)
            - (after.incremental_display_builds - before.incremental_display_builds),
    }
}

/// Types into, then deletes from, the fifth paragraph of a document of
/// `paragraphs` paragraphs, returning each keystroke's work.
fn typing_work(font: u32, paragraphs: usize) -> Vec<KeystrokeWork> {
    let engine = paragraphs_engine(9711, paragraphs);
    let mut host = CommandHost::open(&engine, font);
    let end = paragraph_end(&engine, 4);
    let mut work: Vec<_> = (0..4)
        .map(|key| keystroke_work(&engine, || host.insert(end + key, "x")))
        .collect();
    work.extend(
        (0..2).map(|key| keystroke_work(&engine, || host.delete(end + 3 - key, end + 4 - key))),
    );
    work
}

#[test]
fn command_keystroke_work_does_not_grow_with_the_document() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let small = typing_work(font, 10);
    assert_eq!(small, typing_work(font, 1000));
    for work in &small {
        assert_eq!(work.full_lowerings, 0, "{small:?}");
        assert_eq!(work.full_paginations, 0, "{small:?}");
        assert_eq!(work.full_display_builds, 0, "{small:?}");
        assert_eq!(work.paragraphs_flushed, 1, "{small:?}");
        assert_eq!(work.measured_blocks, 1, "{small:?}");
    }
}

#[test]
fn command_edits_lay_out_as_a_full_recompute_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let engine = paragraphs_engine(9713, 40);
    let twin = paragraphs_engine(9713, 40);
    let mut host = CommandHost::open(&engine, font);
    twin.set_local_lowering(true);
    twin.layout_document_with_regions_retained(&host.request)
        .unwrap();
    let mut undo = engine.doc().undo_manager();
    let mut twin_undo = twin.doc().undo_manager();
    let ctx = crate::EditCtx::local("", "");
    let twin_insert = |offset: u32, text: &str| {
        twin.doc()
            .insert_text(
                &ctx,
                crate::Position::new("body", offset),
                text,
                crate::FormatPolicy::Inherit,
            )
            .unwrap();
    };
    let twin_delete = |start: u32, end: u32| {
        twin.doc()
            .delete_range(&ctx, crate::StoryRange::new("body", start, end))
            .unwrap();
    };
    let check = |host: &mut CommandHost, label: &str| {
        twin.layout_document_with_regions_retained(&host.request)
            .unwrap();
        assert_eq!(
            host.engine.doc().encode_state_as_update_v1(),
            twin.doc().encode_state_as_update_v1(),
            "{label}: document"
        );
        host.assert_matches_cold(label);
    };

    let end = paragraph_end(&engine, 5);
    host.insert(end - 4, "x");
    twin_insert(end - 4, "x");
    check(&mut host, "typing");

    host.insert(end - 2, "pasted, several words long");
    twin_insert(end - 2, "pasted, several words long");
    check(&mut host, "paste");

    host.delete(end - 3, end - 2);
    twin_delete(end - 3, end - 2);
    check(&mut host, "backspace");

    host.delete(end, end + 5);
    twin_delete(end, end + 5);
    check(&mut host, "forward delete");

    let end = paragraph_end(&engine, 7);
    engine
        .doc()
        .split_paragraph(&ctx, crate::Position::new("body", end - 6), None)
        .unwrap();
    twin.doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    host.relayout(None);
    check(&mut host, "enter");
    host.insert(end - 5, "y");
    twin_insert(end - 5, "y");
    check(&mut host, "typing after enter");

    undo.add_undo_barrier();
    twin_undo.add_undo_barrier();
    let end = paragraph_end(&engine, 2);
    let long = "A long pasted passage that wraps over several lines and moves the page breaks. ";
    host.insert(end, &long.repeat(6));
    twin_insert(end, &long.repeat(6));
    check(&mut host, "page-break-moving paste");

    assert!(undo.undo());
    assert!(twin_undo.undo());
    host.relayout(None);
    check(&mut host, "undo");

    let end = paragraph_end(&engine, 2);
    host.insert(end, "z");
    twin_insert(end, "z");
    check(&mut host, "typing after undo");

    let start = paragraph_end(&engine, 0) - 3;
    host.delete(start, start + 3);
    twin_delete(start, start + 3);
    check(&mut host, "delete to the paragraph end");
}
