//! Relaying out after an edit resumes placement inside a later section and
//! balances that section's columns exactly as a full layout does.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, SegmentContent, seed_from_docx};
use fixture::{FONT, p, r, region_request, with_body_and_note};

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

/// Paragraphs, page breaks and sections of one to three columns.
fn generated(seed: u64) -> Vec<u8> {
    let mut rng = Rng(seed * 2_654_435_761 + 1);
    let mut body = String::new();
    for index in 0..40 + rng.next(80) {
        let id = format!("{:08X}", seed * 1000 + index);
        match rng.next(8) {
            0 => body.push_str(&p(&id, r#"<w:r><w:br w:type="page"/></w:r>"#)),
            1 => body.push_str(&section(1 + rng.next(3), false)),
            2 => body.push_str(&p(&id, "")),
            _ => {
                let words = 5 + rng.next(60);
                let text = (0..words)
                    .map(|index| format!("w{}{}", index, rng.next(1000)))
                    .collect::<Vec<_>>()
                    .join(" ");
                body.push_str(&p(&id, &r(&text)));
            }
        }
    }
    body.push_str(&section(2, true));
    with_body_and_note(&body, &p("", &r("A note.")))
}

/// Types into the middle of the body's last run of text.
fn type_into(engine: &EngineSession) {
    let mut offset = 0;
    let mut at = None;
    for segment in engine.doc().story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => {
                let units = text.encode_utf16().count() as u32;
                if units > 1 {
                    at = Some(offset + units / 2);
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

#[test]
fn an_edit_in_a_later_section_relays_out_like_a_full_layout() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for seed in 1..=24 {
        let bytes = generated(seed);
        let engine = EngineSession::new(72);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let request = region_request(&engine, &bytes, font).to_string();
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        type_into(&engine);
        let incremental = engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();

        let fresh = EngineSession::new(73);
        fresh
            .doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        assert_eq!(
            fresh
                .layout_document_with_regions_retained_json(&request)
                .unwrap(),
            incremental,
            "generated document {seed}"
        );
    }
}
