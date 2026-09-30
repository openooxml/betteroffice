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

/// A section of `columns` columns with a `top` twip top margin, starting as `start` says
/// (`nextPage` when empty).
fn section(columns: u64, top: u64, start: &str, last: bool) -> String {
    let start = if start.is_empty() {
        String::new()
    } else {
        format!(r#"<w:type w:val="{start}"/>"#)
    };
    let properties = format!(
        r#"{start}<w:cols w:num="{columns}" w:space="360"/><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="{top}" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#
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

/// Paragraphs, page breaks and sections of one to three columns, each starting on a new
/// page, continuously or in the next column. The last section has one column and a
/// paragraph, as incremental pagination needs.
fn generated(seed: u64) -> Vec<u8> {
    let mut rng = Rng(seed * 2_654_435_761 + 1);
    let mut body = String::new();
    for index in 0..40 + rng.next(80) {
        let id = format!("{:08X}", seed * 1000 + index);
        match rng.next(9) {
            0 => body.push_str(&p(&id, r#"<w:r><w:br w:type="page"/></w:r>"#)),
            1 => {
                let start = ["", "continuous", "nextColumn"][rng.next(3) as usize];
                let top = [720, 1080][rng.next(2) as usize];
                body.push_str(&section(1 + rng.next(3), top, start, false));
            }
            2 => body.push_str(&p(&id, "")),
            kind => {
                let words = 5 + rng.next(60);
                let text = (0..words)
                    .map(|index| format!("w{}{}", index, rng.next(1000)))
                    .collect::<Vec<_>>()
                    .join(" ");
                let properties = if kind == 3 {
                    "<w:pPr><w:pageBreakBefore/></w:pPr>"
                } else {
                    ""
                };
                body.push_str(&p(&id, &format!("{properties}{}", r(&text))));
            }
        }
    }
    body.push_str(&p("FFFFFFFF", ""));
    body.push_str(&section(1, 720, "continuous", true));
    with_body_and_note(&body, &p("", &r("A note.")))
}

/// Types `text` into the middle of the body's run of text that `pick` chooses from their
/// count.
fn type_into(engine: &EngineSession, text: &str, pick: impl Fn(usize) -> usize) {
    let mut offset = 0;
    let mut runs = Vec::new();
    for segment in engine.doc().story_segments("body").unwrap() {
        match segment.content {
            SegmentContent::Text(text) => {
                let units = text.encode_utf16().count() as u32;
                if units > 1 {
                    runs.push(offset + units / 2);
                }
                offset += units;
            }
            _ => offset += 1,
        }
    }
    let at = runs.get(pick(runs.len())).copied();
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", at.expect("the body has text")),
            text,
            FormatPolicy::Inherit,
        )
        .unwrap();
}

/// Lays `bytes` out, types `text` into the run of text `pick` chooses and lays it out
/// again, which must repaginate incrementally and match a full layout of the edited
/// document.
fn assert_edit_relays_out_like_a_full_layout(
    bytes: &[u8],
    (text, pick): (&str, impl Fn(usize) -> usize),
    font: u32,
    label: &str,
) {
    let engine = EngineSession::new(72);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = region_request(&engine, bytes, font).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    type_into(&engine, text, pick);
    let before = engine.stats().incremental_pagination_calls;
    let incremental = engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    assert_eq!(
        engine.stats().incremental_pagination_calls,
        before + 1,
        "{label} repaginates incrementally"
    );

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
        "{label}"
    );
}

#[test]
fn an_edit_anywhere_relays_out_like_a_full_layout() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for seed in 1..=24 {
        assert_edit_relays_out_like_a_full_layout(
            &generated(seed),
            ("typed ", |count| (seed as usize * 7919) % count),
            font,
            &format!("generated document {seed}"),
        );
    }
}

#[test]
fn a_section_balanced_on_the_page_before_is_not_rebalanced_after_its_page_break() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let body = [
        p("00000001", &r("Before the columns.")),
        section(1, 720, "", false),
        p(
            "00000002",
            &format!("<w:pPr><w:pageBreakBefore/></w:pPr>{}", r("First.")),
        ),
        p("00000003", &r("Second.")),
        section(2, 720, "continuous", false),
        p("00000004", &r("After the columns.")),
        section(1, 720, "continuous", true),
    ]
    .concat();
    assert_edit_relays_out_like_a_full_layout(
        &with_body_and_note(&body, &p("", &r("A note."))),
        ("typed ", |count| count - 1),
        font,
        "a page break opening a balanced section",
    );
}

#[test]
fn an_edit_past_a_page_break_rebalances_its_section() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let mut body = vec![
        p("00000001", &r("Before the columns.")),
        section(1, 720, "", false),
    ];
    for index in 2..6 {
        body.push(p(&format!("{index:08X}"), &r("A short paragraph.")));
    }
    body.push(p(
        "00000006",
        &format!(
            "<w:pPr><w:pageBreakBefore/></w:pPr>{}",
            r("After the break.")
        ),
    ));
    body.push(p("00000007", &r("The last paragraph of the columns.")));
    body.push(section(2, 720, "continuous", false));
    body.push(p("00000008", ""));
    body.push(section(1, 720, "continuous", true));
    assert_edit_relays_out_like_a_full_layout(
        &with_body_and_note(&body.concat(), &p("", &r("A note."))),
        (&"grown ".repeat(40), |count| count - 1),
        font,
        "an edit after a page break in a balanced section",
    );
}
