use super::*;

/// The worker's text-command loop over one engine: each command edits, lays
/// the document out again and builds the frame the host applies.
struct CommandHost<'a> {
    engine: &'a EngineSession,
    request: String,
    extras: String,
    retained: HashMap<u64, crate::frame_delta::TestFramePage>,
    /// The client id and package `engine` was seeded from, which cold
    /// sessions seed from too.
    seed: Option<(u64, Vec<u8>)>,
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
            seed: None,
        }
    }

    /// Checks against cold sessions seeded from `bytes` under `client_id`, the
    /// package and id `engine` was seeded from.
    fn seeded(mut self, client_id: u64, bytes: Vec<u8>) -> Self {
        self.seed = Some((client_id, bytes));
        self
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
        let seed = self
            .seed
            .as_ref()
            .map(|(client_id, bytes)| (*client_id, bytes.as_slice()));
        assert_returned_frame_matches_seeded_cold(
            self.engine,
            &self.request,
            &frame,
            &mut self.retained,
            seed,
        );
        let cold = cold_session(self.engine, 9712, seed);
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
    seeded_engine(client_id, &objects_package(objects))
}

/// A session seeded from `bytes`, as a host opens it.
fn seeded_engine(client_id: u64, bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(client_id);
    crate::seed_from_docx(engine.doc(), bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    engine
}

/// The package of [`objects_engine`].
fn objects_package(objects: &[&str]) -> Vec<u8> {
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
    Package::new(&body)
        .rel("rIdImage", "image", "media/image1.png")
        .bytes()
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

/// [`small_page_request`] as the worker sends it for a document whose only
/// section is its final one: the final section listed twice.
/// How many pages the retained layout holds.
fn page_count(engine: &EngineSession) -> usize {
    engine
        .pagination
        .borrow()
        .layout
        .as_ref()
        .map_or(0, |layout| layout.pages.len())
}

/// Runs `edit` and checks it against a cold full pass. Returns whether the
/// edit stayed off the full region pass, asserting that it did when the page
/// count held: a page count change falls back by design, as the page number
/// fields of headers and footers may change width.
fn stays_resident<'a>(
    host: &mut CommandHost<'a>,
    label: &str,
    edit: &mut dyn FnMut(&mut CommandHost<'a>),
) -> bool {
    let pages = page_count(host.engine);
    REGION_FULL_PASSES.with(|count| count.set(0));
    edit(host);
    let full = REGION_FULL_PASSES.with(Cell::get);
    if page_count(host.engine) == pages {
        assert_eq!(full, 0, "{label}: the edit ran the full region pass");
    }
    host.assert_matches_cold(label);
    full == 0
}

fn repeated_final_section_request(font: u32) -> String {
    let mut request: serde_json::Value = serde_json::from_str(&small_page_request(font)).unwrap();
    let section = request["regions"]["sections"][0].clone();
    request["regions"]["sections"] = json!([section.clone(), section]);
    request.to_string()
}

#[test]
fn keystrokes_with_a_repeated_final_section_stay_resident() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for (label, engine, seed, paragraph) in [
        ("paragraphs", paragraphs_engine(9724, 40), None, 5),
        (
            "tables",
            objects_engine(9725, &["table"]),
            Some((9725, objects_package(&["table"]))),
            8,
        ),
    ] {
        let mut host = CommandHost::open_with(&engine, repeated_final_section_request(font));
        if let Some((client_id, bytes)) = seed {
            host = host.seeded(client_id, bytes);
        }
        let end = paragraph_end(&engine, paragraph);
        for key in 0..4 {
            let label = format!("{label}: key {key}");
            REGION_FULL_PASSES.with(|count| count.set(0));
            let work = keystroke_work(&engine, || {
                assert!(
                    host.command("body", end + key, end + key, Some("x")),
                    "{label}"
                );
            });
            assert_eq!(REGION_FULL_PASSES.with(Cell::get), 0, "{label}: {work:?}");
            assert_eq!(work.full_lowerings, 0, "{label}: {work:?}");
            assert_eq!(work.full_paginations, 0, "{label}: {work:?}");
            host.assert_matches_cold(&label);
        }
        let work = keystroke_work(&engine, || {
            assert!(
                host.command("body", end + 3, end + 4, None),
                "{label}: backspace"
            );
        });
        assert_eq!(work.full_lowerings, 0, "{label}: backspace {work:?}");
        assert_eq!(work.full_paginations, 0, "{label}: backspace {work:?}");
        host.assert_matches_cold(&format!("{label}: backspace"));
    }
}

/// Filler paragraphs over several small pages around a table and a bookmarked
/// paragraph. The local patch declines edits in either, so typing there
/// re-lowers the whole body.
fn relowering_package() -> Vec<u8> {
    use super::lowering_fixture::{Package, para, run};
    let filler = |from: u32, count: u32| -> String {
        (from..from + count)
            .map(|index| {
                para(
                    &format!("{:08X}", 0x800 + index),
                    &run(&format!(
                        "Filler paragraph {index} with enough words to wrap over two lines."
                    )),
                )
            })
            .collect()
    };
    let cell = |id: &str, text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{}</w:tc>"#,
            para(id, &run(text))
        )
    };
    let body = format!(
        r#"{}<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>{}{}{}"#,
        filler(0, 8),
        cell("00000901", "Left cell"),
        cell("00000902", "A right cell long enough to wrap"),
        filler(8, 8),
        para(
            "00000903",
            &format!(
                r#"<w:bookmarkStart w:id="0" w:name="mark"/>{}<w:bookmarkEnd w:id="0"/>"#,
                run("A bookmarked paragraph")
            )
        ),
        filler(16, 8),
    );
    Package::new(&body).bytes()
}

#[test]
fn typed_edits_that_re_lower_the_body_keep_the_full_pass() {
    use yrs::{ReadTxn, Transact};
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let bytes = relowering_package();
    let engine = seeded_engine(9726, &bytes);
    let mut host =
        CommandHost::open_with(&engine, repeated_final_section_request(font)).seeded(9726, bytes);
    let cell = {
        let txn = engine.doc().yrs_doc().transact();
        let stories = txn.get_map(crate::STORIES).unwrap();
        let mut cells: Vec<String> = crate::live_stories(&stories, &txn)
            .map(|(id, _)| id.to_string())
            .filter(|id| id.starts_with("body:t"))
            .collect();
        cells.sort();
        cells.remove(0)
    };
    let bookmarked = paragraph_start(&engine, "00000903") + 3;
    for (label, story, at) in [
        ("table cell", cell.as_str(), 4),
        ("bookmarked paragraph", "body", bookmarked),
    ] {
        for key in 0..3 {
            let label = format!("{label}: key {key}");
            REGION_FULL_PASSES.with(|count| count.set(0));
            let work = keystroke_work(&engine, || {
                assert!(
                    host.command(story, at + key, at + key, Some("x")),
                    "{label}"
                );
            });
            assert!(work.full_lowerings > 0, "{label}: {work:?}");
            assert_eq!(
                REGION_FULL_PASSES.with(Cell::get),
                1,
                "{label}: the keystroke left the full pass, {work:?}"
            );
            host.assert_matches_cold(&label);
        }
    }
}

#[test]
fn paragraph_breaks_stay_on_the_region_fast_path() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for (label, engine, request, seed, paragraph) in [
        (
            "paragraphs",
            paragraphs_engine(9727, 40),
            small_page_request(font),
            None,
            5,
        ),
        (
            "paragraphs, repeated final section",
            paragraphs_engine(9728, 40),
            repeated_final_section_request(font),
            None,
            5,
        ),
        (
            "after a table",
            objects_engine(9729, &["table"]),
            repeated_final_section_request(font),
            Some((9729, objects_package(&["table"]))),
            8,
        ),
    ] {
        let mut host = CommandHost::open_with(&engine, request);
        if let Some((client_id, bytes)) = seed {
            host = host.seeded(client_id, bytes);
        }
        let mut resident_splits = 0;
        let end = paragraph_end(&engine, paragraph);
        for (name, at) in [("middle", end - 6), ("end", end), ("start", end - 6 - 20)] {
            let at = at.max(paragraph_end(&engine, paragraph - 1) + 1);
            let split = stays_resident(
                &mut host,
                &format!("{label}: enter at the {name}"),
                &mut |host| {
                    engine
                        .split_resident_paragraph(crate::Position::new("body", at))
                        .unwrap();
                    host.relayout(Some("body"));
                },
            );
            resident_splits += usize::from(split);
            let typed = format!("{label}: typing after enter at the {name}");
            let work = keystroke_work(&engine, || {
                assert!(host.command("body", at + 1, at + 1, Some("y")), "{typed}");
            });
            assert_eq!(work.full_lowerings, 0, "{typed}: {work:?}");
            host.assert_matches_cold(&typed);
        }
        assert!(resident_splits > 0, "{label}: no split stayed resident");
    }
}

/// Filler over several small pages around a numbered list and a table,
/// optionally with a section break after the list.
fn structured_engine(client_id: u64, section_break: bool) -> (EngineSession, Vec<u8>) {
    use super::lowering_fixture::{NS, Package, para, run};
    let filler = |from: u32, count: u32| -> String {
        (from..from + count)
            .map(|index| {
                para(
                    &format!("{:08X}", 0xA00 + index),
                    &run(&format!(
                        "Filler paragraph {index} with enough words to wrap over two lines."
                    )),
                )
            })
            .collect()
    };
    let item = |index: u32| {
        format!(
            r#"<w:p w14:paraId="{:08X}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>{}</w:p>"#,
            0xB00 + index,
            run(&format!("List item {index} with a few words"))
        )
    };
    let break_after_list = if section_break {
        r#"<w:p w14:paraId="00000C01"><w:pPr><w:sectPr><w:pgSz w:w="4320" w:h="2880"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:pPr></w:p>"#
    } else {
        ""
    };
    let body = format!(
        r#"{}{}{}{}{}<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{}</w:tc><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>{}"#,
        filler(0, 6),
        item(0),
        item(1),
        item(2),
        break_after_list,
        para("00000D01", &run("Left cell text")),
        para("00000D02", &run("A right cell long enough to wrap")),
        filler(6, 10),
    );
    let numbering = format!(
        r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let bytes = Package::new(&body).numbering(&numbering).bytes();
    let engine = EngineSession::new(client_id);
    crate::seed_from_docx(engine.doc(), &bytes).unwrap();
    engine.set_relayout_trigger(RelayoutTrigger::Open);
    (engine, bytes)
}

#[test]
fn paragraph_breaks_in_lists_cells_and_before_section_breaks_lay_out_as_a_full_recompute_does() {
    use yrs::{ReadTxn, Transact};
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for section_break in [false, true] {
        let (engine, bytes) = structured_engine(9730 + u64::from(section_break), section_break);
        let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
        let client_id = 9730 + u64::from(section_break);
        let mut host = CommandHost::open_with(&engine, request).seeded(client_id, bytes);
        let cell = {
            let txn = engine.doc().yrs_doc().transact();
            let stories = txn.get_map(crate::STORIES).unwrap();
            let mut cells: Vec<String> = crate::live_stories(&stories, &txn)
                .map(|(id, _)| id.to_string())
                .filter(|id| id.starts_with("body:t"))
                .collect();
            cells.sort();
            cells.remove(0)
        };
        let span = |id: &str| {
            engine
                .doc()
                .paragraph_index("body")
                .unwrap()
                .para_span(id)
                .unwrap()
        };
        let mut splits = vec![
            ("a list item", "body"),
            ("the end of the last list item", "body"),
            ("a table cell", cell.as_str()),
        ];
        if section_break {
            splits.push(("the section mark paragraph", "body"));
        }
        for (label, story) in splits {
            let at = match label {
                "a list item" => span("00000B01").0 + 4,
                "the end of the last list item" => span("00000B02").1,
                "a table cell" => 4,
                _ => span("00000C01").0,
            };
            let label = format!("{label}, section break {section_break}");
            let mut split = |host: &mut CommandHost<'_>| {
                engine
                    .split_resident_paragraph(crate::Position::new(story, at))
                    .unwrap();
                host.relayout(Some(story));
            };
            if section_break || story != "body" {
                split(&mut host);
                host.assert_matches_cold(&label);
            } else {
                stays_resident(&mut host, &label, &mut split);
            }
            host.command(story, at + 1, at + 1, Some("z"));
            host.assert_matches_cold(&format!("{label}: typing after it"));
        }
    }
}

#[test]
fn undone_repeated_and_remote_paragraph_breaks_lay_out_as_a_full_recompute_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let engine = paragraphs_engine(9733, 40);
    let mut host = CommandHost::open_with(&engine, repeated_final_section_request(font));
    let mut undo = engine.doc().undo_manager();
    let split = |at: u32| {
        engine
            .split_resident_paragraph(crate::Position::new("body", at))
            .unwrap();
    };

    undo.add_undo_barrier();
    let at = paragraph_end(&engine, 6) - 8;
    split(at);
    host.relayout(Some("body"));
    host.assert_matches_cold("enter");
    assert!(undo.undo());
    host.relayout(None);
    host.assert_matches_cold("undo of enter");
    assert!(undo.redo());
    host.relayout(None);
    host.assert_matches_cold("redo of enter");

    undo.add_undo_barrier();
    let at = paragraph_end(&engine, 9) - 5;
    split(at);
    host.relayout(Some("body"));
    host.command("body", at + 1, at + 1, Some("q"));
    host.command("body", at + 2, at + 2, Some("r"));
    host.assert_matches_cold("enter and typing");
    assert!(undo.undo());
    host.relayout(None);
    host.assert_matches_cold("undo of enter and typing");
    assert!(undo.redo());
    host.relayout(None);
    host.assert_matches_cold("redo of enter and typing");

    for key in 0..5 {
        let at = paragraph_end(&engine, 12) - 3;
        stays_resident(&mut host, &format!("enter {key} of five"), &mut |host| {
            split(at);
            host.relayout(Some("body"));
        });
    }
    let at = paragraph_end(&engine, 20) - 3;
    for _ in 0..5 {
        split(at);
    }
    host.relayout(Some("body"));
    host.assert_matches_cold("five splits laid out once");

    let remote = EngineSession::new(9734);
    remote
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    let at = paragraph_end(&remote, 25) - 4;
    remote
        .doc()
        .split_paragraph(
            &crate::EditCtx::local("", ""),
            crate::Position::new("body", at),
            None,
        )
        .unwrap();
    engine
        .doc()
        .apply_update_v1(&remote.doc().encode_state_as_update_v1())
        .unwrap();
    host.relayout(None);
    host.assert_matches_cold("a remote paragraph break");
    let at = paragraph_end(&remote, 30) - 4;
    remote
        .doc()
        .split_paragraph(
            &crate::EditCtx::local("", ""),
            crate::Position::new("body", at),
            None,
        )
        .unwrap();
    let at = paragraph_end(&engine, 3) - 2;
    split(at);
    engine
        .doc()
        .apply_update_v1(&remote.doc().encode_state_as_update_v1())
        .unwrap();
    host.relayout(Some("body"));
    host.assert_matches_cold("a local paragraph break beside a remote update");
}

/// Filler paragraphs keyed from `base`, each long enough to wrap.
fn filler_paragraphs(base: u32, count: u32) -> String {
    use super::lowering_fixture::{para, run};
    (0..count)
        .map(|index| {
            para(
                &format!("{:08X}", base + index),
                &run(&format!(
                    "Filler paragraph {index} with enough words to wrap over two lines."
                )),
            )
        })
        .collect()
}

#[test]
fn enter_before_a_body_section_break_lays_out_as_a_full_recompute_does() {
    use super::lowering_fixture::{Package, para, run};
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let body = format!(
        r#"{}{}{}"#,
        para(
            "00001101",
            &format!(
                r#"<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>{}"#,
                run("Section 1")
            )
        ),
        para("00001102", &run("Section 2")),
        filler_paragraphs(0x1200, 12),
    );
    let bytes = Package::new(&body).bytes();
    let engine = seeded_engine(9742, &bytes);
    let mut host = CommandHost::open_with(&engine, small_page_request(font)).seeded(9742, bytes);
    let end = engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_span("00001101")
        .unwrap()
        .1;
    engine
        .split_resident_paragraph(crate::Position::new("body", end))
        .unwrap();
    host.relayout(Some("body"));
    host.assert_matches_cold("enter at the end of the paragraph holding a section break");
    host.command("body", end + 1, end + 1, Some("x"));
    host.assert_matches_cold("typing into the paragraph holding the section break");
}

#[test]
fn edits_beside_unequal_columns_lay_out_as_a_full_recompute_does() {
    use super::lowering_fixture::Package;
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let body = format!(
        r#"{}<w:sectPr><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/><w:cols w:num="2" w:equalWidth="0"><w:col w:w="3600" w:space="300"/><w:col w:w="1860"/></w:cols></w:sectPr>"#,
        filler_paragraphs(0x1300, 30),
    );
    let bytes = Package::new(&body).bytes();
    let engine = seeded_engine(9743, &bytes);
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    let mut host = CommandHost::open_with(&engine, request).seeded(9743, bytes);
    for paragraph in [4, 9] {
        let end = paragraph_end(&engine, paragraph);
        host.command("body", end - 3, end - 3, Some("x"));
        host.assert_matches_cold(&format!("typing in paragraph {paragraph}"));
        engine
            .split_resident_paragraph(crate::Position::new("body", end - 8))
            .unwrap();
        host.relayout(Some("body"));
        host.assert_matches_cold(&format!("enter in paragraph {paragraph}"));
    }
}

#[test]
fn edits_beside_a_shape_placed_by_page_side_lay_out_as_a_full_recompute_does() {
    use super::lowering_fixture::{Package, para, run};
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    // Offset from the outside margin into the body, so it has no zone until
    // the page side places it.
    let shape = INSIDE_SHAPE.replace(
        r#"<wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH>"#,
        r#"<wp:positionH relativeFrom="outsideMargin"><wp:posOffset>-3175000</wp:posOffset></wp:positionH>"#,
    );
    assert_ne!(shape, INSIDE_SHAPE);
    let body = format!(
        r#"{}{}<w:sectPr><w:pgSz w:w="5760" w:h="5760"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360" w:header="200" w:footer="200" w:gutter="0"/></w:sectPr>"#,
        para("00001401", &format!("{}{shape}", run("Anchor "))),
        filler_paragraphs(0x1410, 12),
    );
    let bytes = Package::new(&body).bytes();
    let engine = seeded_engine(9744, &bytes);
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    let mut host = CommandHost::open_with(&engine, request).seeded(9744, bytes);
    let end = paragraph_end(&engine, 1);
    engine
        .split_resident_paragraph(crate::Position::new("body", end - 20))
        .unwrap();
    host.relayout(Some("body"));
    host.assert_matches_cold("enter beside the shape");
    host.command("body", end - 25, end - 25, Some("x"));
    host.assert_matches_cold("typing beside the shape");
}
