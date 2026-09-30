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

/// A generated document, and whether it has a section with columns.
fn generated(seed: u64) -> (Vec<u8>, bool) {
    let mut rng = Rng(seed * 2_654_435_761 + 1);
    let mut body = String::new();
    let mut columns = false;
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
                let count = 1 + rng.next(3);
                columns |= count > 1;
                body.push_str(&section(count, false));
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
    (with_body_and_note(&body, &note), columns)
}

#[test]
fn generated_documents_lay_out_alike_in_steps() {
    for seed in 1..=24 {
        let (bytes, columns) = generated(seed);
        // Relayout after an edit resumes placement inside a section; with columns
        // that can panic in terminal column balancing on main, independently of
        // stepping, so those documents check the stepped pass alone.
        assert_stepped_equals_whole(&format!("generated document {seed}"), &bytes, !columns);
    }
}

#[test]
fn a_change_between_steps_abandons_the_pass() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let (bytes, _) = generated(3);

    let (engine, request) = seeded(&bytes, font);
    assert!(
        engine
            .begin_region_layout(&request)
            .unwrap()
            .layout_json
            .is_none()
    );
    type_into(&engine);
    assert!(engine.resume_region_layout(1).is_err());
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
    assert!(
        engine.resume_region_layout(1).is_err(),
        "new fonts abandon it"
    );
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
