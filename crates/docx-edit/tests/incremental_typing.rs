//! Sustained typing on a many-page body must stay incremental: each keystroke
//! re-places only a local block window and rebuilds only the damaged pages,
//! including after a reflow that changes the page count (which collapses the
//! block-aligned pagination checkpoints).

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, seed_from_docx};

const SENTENCE: &str = "The issuer shall maintain and make available to the competent authority a complete and accurate record of every crypto-asset service provided during the reporting period, including the identity of the client, the nature of the instruction and the timestamp at which it was executed.";

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

fn build_docx(paragraph_count: usize) -> Vec<u8> {
    let mut body = String::new();
    for i in 1..=paragraph_count {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t xml:space="preserve">Section {i}. {SENTENCE}</w:t></w:r></w:p>"#
        ));
    }
    body.push_str(
        r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#,
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap()
}

fn bootstrap() -> EngineSession {
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), &build_docx(400)).unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "properties": {
                "pageWidth": 12240, "pageHeight": 15840,
                "marginTop": 1440, "marginRight": 1440,
                "marginBottom": 1440, "marginLeft": 1440
            }
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font_id]},
            "defaults": {"fontSize": 11, "fontFamily": "Calibri"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let extras = serde_json::json!({"fontChains": {"calibri|0|0": [font_id]}}).to_string();
    engine.build_display_list_frame(&extras, 0).unwrap();
    engine
}

fn insert(engine: &EngineSession, pos: u32, text: &str) {
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", pos),
            text,
            FormatPolicy::Inherit,
        )
        .unwrap();
}

#[test]
fn sustained_typing_stays_incremental_across_a_page_count_change() {
    let engine = bootstrap();
    let mut epoch = 1_u64;
    let mut pos = 25_u32;
    let key = |engine: &EngineSession, pos: &mut u32, text: &str, epoch: &mut u64| {
        insert(engine, *pos, text);
        *pos += text.len() as u32;
        let frame = engine.apply_and_layout("body", *epoch).unwrap();
        *epoch += 1;
        frame
    };

    let initial_pages = engine.stats().retained_pages;
    assert!(initial_pages >= 30, "fixture must paginate to many pages");

    // steady state: each keystroke is incremental and rebuilds a local window
    for _ in 0..5 {
        let before = engine.stats();
        let frame = key(&engine, &mut pos, "x", &mut epoch);
        let stats = engine.stats();
        assert_eq!(stats.pagination_calls - before.pagination_calls, 1);
        assert_eq!(
            stats.incremental_pagination_calls - before.incremental_pagination_calls,
            1,
            "typing must take the incremental pagination path"
        );
        assert!(
            stats.rebuilt_pages <= 3,
            "a keystroke must not rebuild the whole document ({} pages)",
            stats.rebuilt_pages
        );
        assert!(
            frame.len() < 4_000_000,
            "a keystroke frame must stay page-sized ({} bytes)",
            frame.len()
        );
    }

    // grow the caret paragraph until the document gains a page: this reflow
    // legitimately rebuilds everything once and collapses the block-aligned
    // checkpoints (page starts land mid-paragraph afterwards)
    while engine.stats().retained_pages == initial_pages {
        key(&engine, &mut pos, "grow the paragraph ", &mut epoch);
    }

    // typing after the reflow must return to page-local work even though the
    // checkpoint set is sparse: the damaged-page-span diff bounds the rebuild
    for _ in 0..5 {
        let before = engine.stats();
        let frame = key(&engine, &mut pos, "x", &mut epoch);
        let stats = engine.stats();
        assert_eq!(
            stats.incremental_pagination_calls - before.incremental_pagination_calls,
            1,
            "post-reflow typing must stay on the incremental pagination path"
        );
        assert!(
            stats.rebuilt_pages <= 3,
            "post-reflow keystroke must not rebuild the whole document ({} pages)",
            stats.rebuilt_pages
        );
        assert!(
            frame.len() < 4_000_000,
            "post-reflow keystroke frame must stay page-sized ({} bytes)",
            frame.len()
        );
    }
}
