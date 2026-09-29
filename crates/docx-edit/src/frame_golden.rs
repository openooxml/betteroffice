//! Whole-document frames stay byte-identical when the display list's
//! in-memory shape changes: the wire is what hosts and workers exchange.

use crate::engine::EngineSession;

const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// Documents with tables, comments, revisions, fields, notes and content controls.
const GOLDENS: [(&str, usize, u64); 4] = [
    (
        "betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx",
        72_911,
        2_100_429_671_868_973_425,
    ),
    (
        "betteroffice-docx/tests/corpus/fixtures/betteroffice-demo.docx",
        35_682,
        18_288_183_993_733_024_545,
    ),
    (
        "docx-edit/tests/fixtures/structured-export/principal.docx",
        147_114,
        17_254_889_028_504_211_937,
    ),
    (
        "docx-edit/tests/fixtures/footnote-anchor.docx",
        8_817,
        16_631_091_176_605_866_196,
    ),
];

fn fnv1a(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}

fn frame(docx: &[u8]) -> Vec<u8> {
    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(LIBERATION).unwrap();
    let request = serde_json::json!({
        "bodyStory": "body",
        "regions": { "sections": [{ "sectionId": "main", "properties": {} }] },
        "measurement": {
            "fontChains": { "liberation sans|0|0": [font_id] },
            "defaults": { "fontSize": 11, "fontFamily": "Liberation Sans" },
            "authoritativeShaping": true
        },
        "renderEnv": {}
    })
    .to_string();
    let extras =
        serde_json::json!({ "fontChains": { "liberation sans|0|0": [font_id] } }).to_string();
    let engine = EngineSession::new(7);
    crate::seed::seed_from_docx(engine.doc(), docx).unwrap();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let frame = engine.build_display_list_frame(&extras, 0).unwrap();
    docx_layout::clear_measure_fonts();
    frame
}

#[test]
fn whole_document_frames_match_their_goldens() {
    let crates = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("..");
    let mut actual = Vec::new();
    for (path, _, _) in GOLDENS {
        let bytes = frame(&std::fs::read(crates.join(path)).unwrap());
        actual.push((path, bytes.len(), fnv1a(&bytes)));
    }
    let expected: Vec<_> = GOLDENS.to_vec();
    assert_eq!(actual, expected);
}
