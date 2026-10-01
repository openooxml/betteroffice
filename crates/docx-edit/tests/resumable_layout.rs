//! A region layout measured a few blocks at a time lays out exactly as one pass, on the
//! repository's DOCX corpus and on generated documents, and leaves the same retained state.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::path::PathBuf;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, SegmentContent, seed_from_docx};
use fixture::{FONT, p, r, region_request, with_body, with_body_and_note};

const CORPUS: &[&str] = &[
    "crates/betteroffice-docx/tests/corpus/fixtures/betteroffice-demo.docx",
    "crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx",
    "crates/docx-edit/tests/fixtures/page-fragments/pages.docx",
    "crates/docx-edit/tests/fixtures/structured-export/principal.docx",
    "crates/docx-edit/tests/fixtures/footnote-anchor.docx",
    "crates/docx-edit/tests/fixtures/suppressed-list-markers.docx",
    "crates/ooxml-redact/tests/fixtures/redaction-integrity.docx",
    "packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx",
];

fn seeded(bytes: &[u8], font: u32) -> (EngineSession, String) {
    let engine = EngineSession::new(71);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = region_request(&engine, bytes, font).to_string();
    (engine, request)
}

fn stepped(engine: &EngineSession, request: &str, blocks: usize) -> String {
    let mut progress = engine.begin_region_layout(request).unwrap();
    while progress.layout_json.is_none() {
        let before = progress.measured_blocks;
        progress = engine.resume_region_layout(blocks).unwrap();
        assert!(progress.layout_json.is_some() || progress.measured_blocks == before + blocks);
    }
    progress.layout_json.unwrap()
}

/// Types into the middle of the body's first run of text.
fn type_into(engine: &EngineSession) {
    let mut offset = 0;
    let mut at = None;
    for segment in engine.doc().story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => {
                let units = text.encode_utf16().count() as u32;
                if units > 1 {
                    at = Some(offset + units / 2);
                    break;
                }
                offset += units;
            }
            _ => offset += 1,
        }
    }
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", at.expect("the body has text")),
            "typed ",
            FormatPolicy::Inherit,
        )
        .unwrap();
}

fn assert_stepped_equals_whole(name: &str, bytes: &[u8], edit_after: bool) {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let (whole, request) = seeded(bytes, font);
    let expected = whole
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    for blocks in [1, 5, 64] {
        let (engine, _) = seeded(bytes, font);
        assert_eq!(
            stepped(&engine, &request, blocks),
            expected,
            "{name}: {blocks} blocks a step"
        );
    }
    let (engine, _) = seeded(bytes, font);
    engine
        .layout_document_with_regions_prefix_retained_json(&request, 1)
        .unwrap();
    assert_eq!(
        stepped(&engine, &request, 3),
        expected,
        "{name}: completing a prefix"
    );
    if !edit_after {
        return;
    }
    // The retained state matches too: the next pass after an edit reuses it alike.
    type_into(&whole);
    type_into(&engine);
    let after_edit = whole
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap(),
        after_edit,
        "{name}: the pass after an edit"
    );
}

#[test]
fn the_corpus_lays_out_alike_in_steps() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in CORPUS {
        assert_stepped_equals_whole(path, &std::fs::read(root.join(path)).unwrap(), true);
    }
}

/// A small deterministic generator, so each document is reproducible from its seed.
struct Rng(u64);

impl Rng {
    fn next(&mut self, below: u64) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0 % below
    }
}

fn section(columns: u64, last: bool) -> String {
    let properties = format!(
        r#"<w:cols w:num="{columns}" w:space="360"/><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#
    );
    if last {
        format!("<w:sectPr>{properties}</w:sectPr>")
    } else {
        p(
            "",
            &format!("<w:pPr><w:sectPr>{properties}</w:sectPr></w:pPr>"),
        )
    }
}

/// A generated document.
fn generated(seed: u64) -> Vec<u8> {
    let mut rng = Rng(seed * 2_654_435_761 + 1);
    let mut body = String::new();
    let text = |rng: &mut Rng, words: u64| {
        (0..words)
            .map(|index| format!("w{}{}", index, rng.next(1000)))
            .collect::<Vec<_>>()
            .join(" ")
    };
    for index in 0..40 + rng.next(80) {
        let id = format!("{:08X}", seed * 1000 + index);
        match rng.next(12) {
            0 => body.push_str(&p(
                &id,
                &format!("<w:pPr><w:keepNext/></w:pPr>{}", r(&text(&mut rng, 3))),
            )),
            1 => body.push_str(&p(&id, r#"<w:r><w:br w:type="page"/></w:r>"#)),
            2 => {
                body.push_str(&section(1 + rng.next(3), false));
            }
            3 => body.push_str(&p(
                &id,
                &format!(
                    r#"{}<w:r><w:footnoteReference w:id="1"/></w:r>{}"#,
                    r(&text(&mut rng, 8)),
                    r(&text(&mut rng, 8))
                ),
            )),
            4 | 5 => {
                let floating = rng.next(3) == 0;
                let rows = 1 + rng.next(12);
                let mut table = String::from("<w:tbl><w:tblPr>");
                if floating {
                    table.push_str(r#"<w:tblpPr w:leftFromText="120" w:rightFromText="120" w:vertAnchor="text" w:horzAnchor="text" w:tblpY="60"/><w:tblW w:w="1800" w:type="dxa"/>"#);
                }
                table.push_str(r#"</w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>"#);
                for row in 0..rows {
                    let header = if row == 0 && !floating {
                        "<w:trPr><w:tblHeader/></w:trPr>"
                    } else {
                        ""
                    };
                    let (left, right) = (1 + rng.next(10), 1 + rng.next(4));
                    table.push_str(&format!(
                        "<w:tr>{header}<w:tc>{}</w:tc><w:tc>{}</w:tc></w:tr>",
                        p("", &r(&text(&mut rng, left))),
                        p("", &r(&text(&mut rng, right)))
                    ));
                }
                table.push_str("</w:tbl>");
                body.push_str(&table);
            }
            6 => body.push_str(&p(&id, "")),
            _ => {
                let words = 5 + rng.next(60);
                body.push_str(&p(&id, &r(&text(&mut rng, words))));
            }
        }
    }
    body.push_str(&section(1, true));
    let note = p(
        "",
        &r("A note long enough to take a line or two of its own."),
    );
    with_body_and_note(&body, &note)
}

#[test]
fn generated_documents_lay_out_alike_in_steps() {
    for seed in 1..=24 {
        let bytes = generated(seed);
        assert_stepped_equals_whole(&format!("generated document {seed}"), &bytes, true);
    }
}

const INSIDE_SHAPE: &str = r#"<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="66675" distR="123825" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Inside shape"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;

fn snapshot_body(body: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&with_body(body)).unwrap();
    for (name, content) in &mut parts {
        let root = match name.as_str() {
            "word/footnotes.xml" => "footnotes",
            "word/endnotes.xml" => "endnotes",
            _ => continue,
        };
        *content = format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn floating_table(anchor: &str, index: usize) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblpPr w:leftFromText="120" w:rightFromText="120" {anchor} w:horzAnchor="text" w:tblpY="60"/><w:tblW w:w="1800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid>{rows}</w:tbl>"#,
        rows = format!("<w:tr><w:tc><w:p><w:r><w:t>Float {index}</w:t></w:r></w:p></w:tc></w:tr>")
            .repeat(4)
    )
}

fn progressive_body(sections: bool, floats: bool) -> Vec<u8> {
    let mut body = String::new();
    for index in 0..140 {
        if sections && index > 0 && index % 35 == 0 {
            let boundary = if index == 70 {
                section(1, false).replace(r#"w:w="7200" w:h="5760""#, r#"w:w="8640" w:h="6480""#)
            } else {
                section(1, false)
            };
            body.push_str(&boundary);
        }
        if floats && index % 17 == 8 {
            body.push_str(&floating_table(r#"w:vertAnchor="text""#, index));
        }
        let properties = if sections && index % 11 < 2 {
            "<w:pPr><w:keepNext/></w:pPr>"
        } else {
            ""
        };
        body.push_str(&p(
            &format!("{:08X}", index + 1),
            &format!(
                "{properties}{}",
                r(&format!(
                    "Paragraph {index} {}",
                    "Measured text fills each page before later blocks arrive. ".repeat(3)
                ))
            ),
        ));
    }
    body.push_str(&section(1, true));
    snapshot_body(&body)
}

fn last_covered_position(kernel: &serde_json::Value) -> Option<u32> {
    kernel["measured"]
        .as_array()
        .unwrap()
        .iter()
        .rev()
        .find_map(|measured| measured["block"]["pmEnd"].as_f64())
        .map(|end| end as u32)
}

#[test]
fn snapshots_preserve_pages_progress_and_final_retained_state() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for (name, bytes) in [
        ("paragraphs", progressive_body(false, false)),
        ("keep groups and sections", progressive_body(true, false)),
        ("text-relative floats", progressive_body(false, true)),
    ] {
        let (whole, request) = seeded(&bytes, font);
        let expected = whole
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let expected_kernel = whole.retained_kernel_inputs_json().unwrap();
        let full: serde_json::Value = serde_json::from_str(&expected).unwrap();
        let full_pages = full["layout"]["pages"].as_array().unwrap();
        let kernel: serde_json::Value = serde_json::from_str(&expected_kernel).unwrap();
        assert!(full_pages.len() > 2, "{name}");
        for blocks in [1, 5, 64] {
            let (engine, _) = seeded(&bytes, font);
            assert!(engine.region_layout_snapshot_json().unwrap().is_none());
            let mut progress = engine.begin_region_layout(&request).unwrap();
            assert!(progress.covered_position.is_none());
            assert!(
                serde_json::to_value(&progress)
                    .unwrap()
                    .get("coveredPosition")
                    .is_none()
            );
            assert!(engine.region_layout_snapshot_json().unwrap().is_none());
            let mut snapshots = 0;
            let mut frame_epoch = 0;
            let mut coverage = None;
            while progress.layout_json.is_none() {
                let before = progress.measured_blocks;
                progress = engine.resume_region_layout(blocks).unwrap();
                assert_eq!(
                    progress.measured_blocks,
                    (before + blocks).min(progress.body_blocks)
                );
                assert!(progress.covered_position >= coverage, "{name}: coverage");
                coverage = progress.covered_position;
                if progress.layout_json.is_some() {
                    break;
                }
                let before_snapshot = serde_json::to_string(&progress).unwrap();
                let snapshot = engine.region_layout_snapshot_json().unwrap();
                assert_eq!(
                    snapshot.is_some(),
                    coverage.is_some(),
                    "{name}: safe prefix"
                );
                if let Some(json) = snapshot {
                    snapshots += 1;
                    let snapshot: serde_json::Value = serde_json::from_str(&json).unwrap();
                    assert_eq!(snapshot["provisional"], true);
                    assert_eq!(snapshot["layout"]["partial"], true);
                    let pages = snapshot["layout"]["pages"].as_array().unwrap();
                    assert!(!pages.is_empty());
                    assert!(pages.len() <= full_pages.len());
                    assert_eq!(
                        pages[..pages.len() - 1],
                        full_pages[..pages.len() - 1],
                        "{name}: {blocks} blocks a step, {} measured",
                        progress.measured_blocks
                    );
                    let prefix_kernel: serde_json::Value =
                        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap())
                            .unwrap();
                    assert_eq!(last_covered_position(&prefix_kernel), coverage);
                    assert!(
                        prefix_kernel["measured"].as_array().unwrap().len()
                            <= progress.measured_blocks
                    );
                    let frame = engine.build_display_list_frame("{}", frame_epoch).unwrap();
                    let next_epoch = u64::from_le_bytes(frame[32..40].try_into().unwrap());
                    assert!(next_epoch > frame_epoch);
                    frame_epoch = next_epoch;
                }
                progress = engine.resume_region_layout(0).unwrap();
                assert_eq!(serde_json::to_string(&progress).unwrap(), before_snapshot);
            }
            assert!(snapshots >= 2, "{name}: {blocks} blocks a step");
            assert_eq!(progress.covered_position, last_covered_position(&kernel));
            assert_eq!(
                progress.layout_json.unwrap(),
                expected,
                "{name}: final layout"
            );
            assert_eq!(
                engine.retained_kernel_inputs_json().unwrap(),
                expected_kernel,
                "{name}: final kernel inputs"
            );
            assert!(engine.region_layout_snapshot_json().unwrap().is_none());
            engine.build_display_list_frame("{}", frame_epoch).unwrap();
        }
    }
}

#[test]
fn snapshots_wait_for_keep_groups_and_section_marks() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = snapshot_body(&format!(
        "{}{}{}{}{}{}{}",
        p(
            "00000001",
            &format!("<w:pPr><w:keepNext/></w:pPr>{}", r("First"))
        ),
        p(
            "00000002",
            &format!("<w:pPr><w:keepNext/></w:pPr>{}", r("Second"))
        ),
        p("00000003", &r("Follower")),
        p("00000004", &r("Before the section mark")),
        section(1, false),
        p("00000005", &r("Next section")),
        section(1, true),
    ));
    let (engine, request) = seeded(&bytes, font);
    engine.begin_region_layout(&request).unwrap();
    for _ in 0..2 {
        let progress = engine.resume_region_layout(1).unwrap();
        assert_eq!(progress.covered_position, None);
        assert!(engine.region_layout_snapshot_json().unwrap().is_none());
    }
    let progress = engine.resume_region_layout(1).unwrap();
    assert!(progress.covered_position.is_some());
    assert!(engine.region_layout_snapshot_json().unwrap().is_some());
    let prefix = engine.retained_kernel_inputs_json().unwrap();
    for _ in 0..2 {
        let next = engine.resume_region_layout(1).unwrap();
        assert_eq!(next.covered_position, progress.covered_position);
        assert!(engine.region_layout_snapshot_json().unwrap().is_some());
        assert_eq!(engine.retained_kernel_inputs_json().unwrap(), prefix);
    }
    let next = engine.resume_region_layout(1).unwrap();
    assert!(next.covered_position > progress.covered_position);
    assert!(engine.region_layout_snapshot_json().unwrap().is_some());
    let prefix: serde_json::Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    assert_eq!(
        prefix["measured"].as_array().unwrap().last().unwrap()["block"]["kind"],
        "sectionBreak"
    );
    assert_eq!(next.covered_position, last_covered_position(&prefix));
}

fn assert_snapshots_refused(bytes: &[u8], font: u32, references_only: bool) {
    let (whole, request) = seeded(bytes, font);
    let request = if references_only {
        let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
        request["notes"]["contents"] = serde_json::json!([]);
        request.to_string()
    } else {
        request
    };
    let expected = whole
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let expected_kernel = whole.retained_kernel_inputs_json().unwrap();
    let (engine, _) = seeded(bytes, font);
    let mut progress = engine.begin_region_layout(&request).unwrap();
    assert!(progress.layout_json.is_none());
    let mut refused_prefixes = 0;
    while progress.layout_json.is_none() {
        assert!(engine.region_layout_snapshot_json().unwrap().is_none());
        refused_prefixes += usize::from(progress.measured_blocks > 0);
        let before = serde_json::to_string(&progress).unwrap();
        progress = engine.resume_region_layout(0).unwrap();
        assert_eq!(serde_json::to_string(&progress).unwrap(), before);
        progress = engine.resume_region_layout(1).unwrap();
    }
    assert!(refused_prefixes > 0);
    assert_eq!(progress.layout_json.unwrap(), expected);
    assert_eq!(
        engine.retained_kernel_inputs_json().unwrap(),
        expected_kernel
    );
}

#[test]
fn coupled_floats_refuse_snapshots_and_finish_alike() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/floating-table-wrap/right-4820-margin.docx"),
    )
    .unwrap();
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    for (name, content) in &mut parts {
        if name == "word/document.xml" {
            *content = String::from_utf8(std::mem::take(content))
                .unwrap()
                .replace(r#"w:vertAnchor="text""#, r#"w:vertAnchor="margin""#)
                .into_bytes();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    assert_snapshots_refused(&bytes, font, false);
    let body = format!(
        "{}{}{}",
        p("00000001", &r("Before the inside shape")),
        p("00000002", INSIDE_SHAPE),
        p(
            "00000003",
            &r(&"Body text wraps around the shape. ".repeat(20))
        ),
    );
    assert_snapshots_refused(&snapshot_body(&body), font, false);
}

#[test]
fn notes_and_balanced_columns_refuse_snapshots() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let body = format!("{}{}", p("", &r("Body text")).repeat(3), section(1, true));
    assert_snapshots_refused(&with_body(&body), font, false);
    let body = format!(
        "{}{}{}",
        p(
            "00000001",
            r#"<w:r><w:t>Note anchor</w:t><w:footnoteReference w:id="1"/></w:r>"#
        ),
        p("00000002", &r("After the note")),
        section(1, true),
    );
    let bytes = with_body_and_note(&body, &p("", &r("Note content")));
    assert_snapshots_refused(&bytes, font, false);
    assert_snapshots_refused(&bytes, font, true);
    let body = format!("{}{}", p("", &r("Columns")).repeat(20), section(2, true));
    assert_snapshots_refused(&snapshot_body(&body), font, false);
}

#[test]
fn a_change_between_steps_abandons_the_pass() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = generated(3);

    let (engine, request) = seeded(&bytes, font);
    assert!(
        engine
            .begin_region_layout(&request)
            .unwrap()
            .layout_json
            .is_none()
    );
    type_into(&engine);
    let snapshot_error = engine.region_layout_snapshot_json().unwrap_err();
    assert_eq!(
        engine.region_layout_snapshot_json().unwrap_err(),
        snapshot_error
    );
    assert_eq!(engine.resume_region_layout(1).unwrap_err(), snapshot_error);
    assert!(
        engine.resume_region_layout(1).is_err(),
        "nothing is left to resume"
    );

    let (engine, request) = seeded(&bytes, font);
    assert!(
        engine
            .begin_region_layout(&request)
            .unwrap()
            .layout_json
            .is_none()
    );
    engine
        .layout_document_with_regions_prefix_retained_json(&request, 1)
        .unwrap();
    assert!(
        engine.resume_region_layout(1).is_err(),
        "another pass replaces it"
    );

    let (engine, request) = seeded(&bytes, font);
    assert!(
        engine
            .begin_region_layout(&request)
            .unwrap()
            .layout_json
            .is_none()
    );
    engine
        .layout_document_json(
            r#"{"measured": [], "options": {"pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}}"#,
        )
        .unwrap();
    assert!(
        engine.resume_region_layout(1).is_err(),
        "a layout without regions replaces it"
    );

    let (engine, request) = seeded(&bytes, font);
    assert!(
        engine
            .begin_region_layout(&request)
            .unwrap()
            .layout_json
            .is_none()
    );
    docx_layout::register_measure_font(fixture::OTHER_FONT).unwrap();
    let snapshot_error = engine.region_layout_snapshot_json().unwrap_err();
    assert_eq!(engine.resume_region_layout(1).unwrap_err(), snapshot_error);
}

/// An abandoned pass that lowered the body with another environment leaves
/// the next resident edit as it would be without it.
#[test]
fn an_abandoned_pass_leaves_resident_edits_alone() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let body = format!(
        "{}{}<w:sectPr/>",
        p("00000001", &r("Visible text before the hidden run.")),
        p(
            "00000002",
            r#"<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">hidden words that take up a line or more of their own when shown</w:t></w:r>"#
        )
    );
    let bytes = with_body(&body);
    let edited = |abandon: bool| {
        let (engine, request) = seeded(&bytes, font);
        // One section and no note contents, so edits take the resident path.
        let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
        request["notes"]["contents"] = serde_json::json!([]);
        let last = request["regions"]["sections"]
            .as_array()
            .unwrap()
            .last()
            .cloned();
        request["regions"]["sections"] = serde_json::json!([last]);
        let request = request.to_string();
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        engine.build_display_list_frame("{}", 0).unwrap();
        if abandon {
            let mut shown: serde_json::Value = serde_json::from_str(&request).unwrap();
            shown["renderEnv"]["showHiddenText"] = true.into();
            shown["measurement"]["defaults"]["fontSize"] = 12.into();
            let progress = engine.begin_region_layout(&shown.to_string()).unwrap();
            assert!(progress.layout_json.is_none());
        }
        type_into(&engine);
        engine.apply_and_layout("body", 1).unwrap();
        engine.retained_kernel_inputs_json().unwrap()
    };
    assert_eq!(edited(true), edited(false));
}
