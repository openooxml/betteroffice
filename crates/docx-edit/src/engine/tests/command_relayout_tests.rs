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
        Self::open_with(engine, small_page_request(font))
    }

    /// [`Self::open`] with the region request `request`, whose fonts the
    /// frames use too.
    fn open_with(engine: &'a EngineSession, request: String) -> Self {
        engine.set_local_lowering(true);
        engine
            .layout_document_with_regions_retained(&request)
            .unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&request).unwrap();
        let extras = json!({"fontChains": parsed["measurement"]["fontChains"]}).to_string();
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

    /// A plain text command as the worker routes it: the resident edit and
    /// relayout where the engine takes them, else the generic edit and the full
    /// relayout. Returns whether it was resident.
    fn command(&mut self, story: &str, start: u32, end: u32, text: Option<&str>) -> bool {
        let index = self.engine.doc().paragraph_index(story).unwrap();
        let paragraph = |offset| index.para_at(offset).unwrap().para_id.to_string();
        let resident =
            self.engine
                .takes_resident_text_edit(story, &paragraph(start), &paragraph(end), text);
        if resident {
            self.engine
                .edit_resident_text(crate::StoryRange::new(story, start, end), text, true)
                .unwrap();
            self.relayout(Some(story));
            return true;
        }
        let ctx = crate::EditCtx::local("", "");
        match text {
            Some(text) => self.engine.doc().insert_text(
                &ctx,
                crate::Position::new(story, start),
                text,
                crate::FormatPolicy::Inherit,
            ),
            None => self
                .engine
                .doc()
                .delete_range(&ctx, crate::StoryRange::new(story, start, end)),
        }
        .unwrap();
        self.relayout(None);
        false
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

/// The body's segments and every root as JSON, whose maps compare whatever
/// order their entries were integrated in.
fn document_content(engine: &EngineSession) -> (Vec<crate::StorySegment>, BTreeMap<String, Any>) {
    use yrs::types::ToJson;
    use yrs::{ReadTxn, Transact};
    let roots = {
        let txn = engine.doc().doc.transact();
        txn.root_refs()
            .map(|(name, root)| (name.to_owned(), root.to_json(&txn)))
            .collect()
    };
    (engine.doc().story_segments("body").unwrap(), roots)
}

#[test]
fn document_content_tells_documents_apart() {
    let ctx = crate::EditCtx::local("", "");
    let engine = paragraphs_engine(9715, 3);
    assert_eq!(
        document_content(&engine),
        document_content(&paragraphs_engine(9715, 3))
    );
    let typed = paragraphs_engine(9715, 3);
    typed
        .doc()
        .insert_text(
            &ctx,
            crate::Position::new("body", 2),
            "x",
            crate::FormatPolicy::Inherit,
        )
        .unwrap();
    let bold = paragraphs_engine(9715, 3);
    bold.doc()
        .toggle_format(
            &ctx,
            crate::StoryRange::new("body", 1, 3),
            crate::SimpleFormat::Bold,
        )
        .unwrap();
    let renumbered = paragraphs_engine(9716, 3);
    for (label, other) in [
        ("an extra character", &typed),
        ("a run attribute", &bold),
        ("paragraph ids", &renumbered),
    ] {
        assert_ne!(
            document_content(&engine),
            document_content(other),
            "{label}"
        );
    }
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
            document_content(host.engine),
            document_content(&twin),
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
        .split_paragraph(&ctx, crate::Position::new("body", end - 6), None)
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

const INLINE_SHAPE: &str = r#"<w:r><w:drawing><wp:inline><wp:extent cx="457200" cy="228600"/><wp:docPr id="3" name="shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;

const ANCHORED_SHAPE: &str = r#"<w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="457200" cy="228600"/><wp:wrapNone/><wp:docPr id="4" name="anchored"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;

/// Six short paragraphs ahead of pages holding paragraphs split across pages
/// and the `objects` named: "table", "image" (inline), "shape" (inline) or
/// "anchored" (a floating shape).
fn objects_engine(client_id: u64, objects: &[&str]) -> EngineSession {
    use super::lowering_fixture::{Package, image, para, run};
    let cell = |content: String| {
        format!(r#"<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{content}</w:tc>"#)
    };
    let words = (0..120)
        .map(|word| format!("word{word}"))
        .collect::<Vec<_>>()
        .join(" ");
    let mut body: String = (0..6)
        .map(|index| {
            para(
                &format!("{:08X}", 0x100 + index),
                &run(&format!(
                    "Paragraph {index}: the quick brown fox jumps over the lazy dog."
                )),
            )
        })
        .collect();
    for round in 0..3 {
        let id = |offset: u32| format!("{:08X}", 0x200 + round * 0x10 + offset);
        body += &para(&id(0), &run(&words));
        if objects.contains(&"table") {
            body += &format!(
                r#"<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>"#,
                cell(para(&id(1), &run("Left cell"))),
                cell(para(&id(2), &run("A right cell long enough to wrap"))),
            );
        }
        if objects.contains(&"image") {
            body += &para(
                &id(3),
                &format!(
                    "{}{}{}",
                    run("Image "),
                    image("rIdImage", "picture"),
                    run(" then text")
                ),
            );
        }
        if objects.contains(&"shape") {
            body += &para(
                &id(4),
                &format!("{}{INLINE_SHAPE}{}", run("Shape "), run(" then text")),
            );
        }
        if objects.contains(&"anchored") && round == 1 {
            body += &para(&id(5), &format!("{}{ANCHORED_SHAPE}", run("Anchor ")));
        }
        body += &para(&id(6), &run("A closing paragraph."));
    }
    let bytes = Package::new(&body)
        .rel("rIdImage", "image", "media/image1.png")
        .bytes();
    let engine = EngineSession::new(client_id);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    engine
}

#[test]
fn command_edits_ahead_of_tables_images_and_split_paragraphs_lay_out_as_a_full_recompute_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for anchored in [false, true] {
        let objects: &[&str] = if anchored {
            &["table", "image", "shape", "anchored"]
        } else {
            &["table", "image", "shape"]
        };
        let engine = objects_engine(9714, objects);
        let mut host = CommandHost::open(&engine, font);
        let mut undo = engine.doc().undo_manager();
        let ctx = crate::EditCtx::local("", "");
        let typed = |host: &mut CommandHost, label: &str, work: KeystrokeWork| {
            let label = format!("{label}, anchored {anchored}");
            if !anchored {
                assert_eq!(work.full_paginations, 0, "{label}: {work:?}");
            }
            host.assert_matches_cold(&label);
        };

        let end = paragraph_end(&engine, 1);
        let work = keystroke_work(&engine, || host.insert(end - 3, "x"));
        typed(&mut host, "typing", work);
        let work = keystroke_work(&engine, || host.delete(end - 3, end - 2));
        typed(&mut host, "backspace", work);
        let work = keystroke_work(&engine, || host.delete(end - 6, end - 5));
        typed(&mut host, "forward delete", work);

        undo.add_undo_barrier();
        let end = paragraph_end(&engine, 2);
        let long =
            "A long pasted passage that wraps over several lines and moves the page breaks. ";
        host.insert(end, &long.repeat(4));
        host.assert_matches_cold(&format!("page-break-moving paste, anchored {anchored}"));
        assert!(undo.undo());
        host.relayout(None);
        host.assert_matches_cold(&format!("undo, anchored {anchored}"));

        let end = paragraph_end(&engine, 3);
        engine
            .doc()
            .split_paragraph(&ctx, crate::Position::new("body", end - 6), None)
            .unwrap();
        host.relayout(None);
        host.assert_matches_cold(&format!("enter, anchored {anchored}"));
        let work = keystroke_work(&engine, || host.insert(end - 2, "y"));
        typed(&mut host, "typing after enter", work);
    }
}

/// `edited` as the first body paragraph, then `between`, then filler that runs
/// over several small pages.
fn breadth_engine(client_id: u64, edited: &str, between: &str) -> EngineSession {
    use super::lowering_fixture::{Package, para, run};
    let filler: String = (0..30)
        .map(|index| {
            para(
                &format!("{:08X}", 0x300 + index),
                &run(&format!(
                    "Filler paragraph {index} with enough words to wrap over two lines."
                )),
            )
        })
        .collect();
    let body = format!("{}{between}{filler}", para("10000001", edited));
    let engine = EngineSession::new(client_id);
    crate::seed_from_docx(engine.doc(), &Package::new(&body).bytes()).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    engine
}

/// The start of the body paragraph keyed `id`.
fn paragraph_start(engine: &EngineSession, id: &str) -> u32 {
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_span(id)
        .unwrap()
        .0
}

#[test]
fn command_edits_around_rich_content_lay_out_as_a_full_recompute_does() {
    use super::lowering_fixture::run;
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let cell = |id: &str, text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{}</w:tc>"#,
            super::lowering_fixture::para(id, &run(text))
        )
    };
    let table = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>"#,
        cell("10000011", "Left cell"),
        cell("10000012", "Right cell"),
    );
    let tracked = format!(
        r#"{}<w:ins w:id="1" w:author="Ann" w:date="2026-01-02T00:00:00Z">{}</w:ins><w:del w:id="2" w:author="Bob" w:date="2026-01-03T00:00:00Z"><w:r><w:delText xml:space="preserve">removed </w:delText></w:r></w:del>{}"#,
        run("Kept "),
        run("inserted "),
        run("tail")
    );
    let control = format!(
        r#"{}<w:sdt><w:sdtPr><w:tag w:val="name"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}"#,
        run("Before "),
        run("Ada Lovelace"),
        run(" after")
    );
    let field_and_link = format!(
        r#"{}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>{}<w:hyperlink w:anchor="target"><w:r><w:t>a link</w:t></w:r></w:hyperlink>{}"#,
        run("Page "),
        run(" and "),
        run(" end")
    );
    let cases: [(&str, String, &str, &str); 4] = [
        (
            "table cell",
            run("Before the table"),
            table.as_str(),
            "body:t0:r0c0",
        ),
        ("tracked changes", tracked, "", "body"),
        ("inline control", control, "", "body"),
        ("field and hyperlink", field_and_link, "", "body"),
    ];
    for (label, edited, between, story) in cases {
        let engine = breadth_engine(9720, &edited, between);
        let mut host = CommandHost::open(&engine, font);
        assert!(host.command(story, 2, 2, Some("x")), "{label}");
        host.assert_matches_cold(&format!("{label}: typing"));
        assert!(host.command(story, 2, 3, None), "{label}");
        host.assert_matches_cold(&format!("{label}: delete"));
    }

    let engine = breadth_engine(9721, &run("First paragraph text"), &table);
    let mut host = CommandHost::open(&engine, font);
    let second = paragraph_start(&engine, "00000300");
    assert!(!host.command("body", 3, second + 3, None));
    host.assert_matches_cold("delete across a table");
    let second = paragraph_start(&engine, "00000301");
    assert!(!host.command("body", 3, second + 3, None));
    host.assert_matches_cold("delete across paragraphs");
}

#[test]
fn command_edits_with_a_page_numbered_header_lay_out_as_a_full_recompute_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let body = format!(
        r#"{}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader2"/><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#,
        (0..24)
            .map(|index| lowering_pages::p(
                &format!("{:08X}", 0x500 + index),
                &lowering_pages::r(&format!(
                    "Body paragraph {index} under a header that numbers its page."
                ))
            ))
            .collect::<String>()
    );
    let bytes = lowering_pages::with_body(&body);
    let engine = EngineSession::new(9722);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    let mut host = CommandHost::open_with(&engine, request);
    host.command("body", 2, 2, Some("x"));
    host.assert_matches_cold("typing under a page-numbered header");
    let long = "A long pasted passage that adds a page and renumbers the pages after it. ";
    host.command("body", 2, 2, Some(&long.repeat(12)));
    host.assert_matches_cold("a paste that adds pages");
    assert!(!host.command("hf:rIdHeader2", 1, 1, Some("x")));
    host.assert_matches_cold("a header edit");
    host.command("body", 2, 2, Some("y"));
    host.assert_matches_cold("typing after a header edit");
}

#[test]
fn retained_pages_of_rich_runs_lay_out_as_a_full_recompute_does() {
    use super::lowering_fixture::{NS, Package, para, run};
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let filler = |from: u32, count: u32| -> String {
        (from..from + count)
            .map(|index| {
                para(
                    &format!("{:08X}", 0x600 + index),
                    &run(&format!(
                        "Filler paragraph {index} with enough words to wrap over two lines."
                    )),
                )
            })
            .collect()
    };
    let rich = [
        format!(
            r#"{}<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>{}<w:fldSimple w:instr=" SEQ Figure "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
            run("Page "),
            run(" figure ")
        ),
        format!(
            r#"{}<w:hyperlink w:anchor="target"><w:r><w:t>a link</w:t></w:r></w:hyperlink>{}<w:commentRangeStart w:id="0"/>{}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>"#,
            run("Before "),
            run(" and "),
            run("commented words")
        ),
        format!(
            r#"{}<w:ins w:id="1" w:author="Ann" w:date="2026-01-02T00:00:00Z">{}</w:ins><w:del w:id="2" w:author="Bob" w:date="2026-01-03T00:00:00Z"><w:r><w:delText xml:space="preserve">removed </w:delText></w:r></w:del>{}"#,
            run("Kept "),
            run("inserted "),
            run("tail")
        ),
        format!(
            r#"{}<w:sdt><w:sdtPr><w:tag w:val="name"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}"#,
            run("Before "),
            run("Ada Lovelace"),
            run(" after")
        ),
        format!(
            r#"<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="3600"/></w:tabs></w:pPr>{}<w:r><w:tab/></w:r>{}"#,
            run("Chapter"),
            run("7")
        ),
    ];
    let body = format!(
        "{}{}{}{}",
        filler(0, 24),
        rich.iter()
            .enumerate()
            .map(|(index, content)| para(&format!("{:08X}", 0x700 + index), content))
            .collect::<String>(),
        filler(24, 4),
        rich.iter()
            .enumerate()
            .map(|(index, content)| para(&format!("{:08X}", 0x710 + index), content))
            .collect::<String>(),
    );
    let comments = format!(
        r#"<w:comments {NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let bytes = Package::new(&body)
        .part(
            "comments.xml",
            "rIdComments",
            "comments",
            "comments",
            &comments,
        )
        .bytes();
    let engine = EngineSession::new(9723);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    let mut host = CommandHost::open(&engine, font);
    let pages = engine
        .pagination
        .borrow()
        .layout
        .as_ref()
        .unwrap()
        .pages
        .len();
    let start = paragraph_start(&engine, "00000601");
    for (label, edit) in [
        ("positions moving up", (start + 2, start + 2, Some("xyz"))),
        ("positions moving down", (start + 2, start + 7, None)),
        ("positions moving down again", (start + 1, start + 3, None)),
        (
            "positions moving up again",
            (start + 1, start + 1, Some("x")),
        ),
    ] {
        let (from, to, text) = edit;
        let work = keystroke_work(&engine, || {
            assert!(host.command("body", from, to, text), "{label}");
        });
        assert_eq!(work.full_paginations, 0, "{label}: {work:?}");
        assert!(work.paginated_pages <= 2, "{label}: {work:?} of {pages}");
        host.assert_matches_cold(label);
    }
}
