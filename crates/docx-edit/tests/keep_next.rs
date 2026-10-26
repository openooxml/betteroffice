use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn paragraph(label: &str, lines: usize, properties: &str, height: u32) -> String {
    let runs = (1..=lines)
        .map(|line| format!("<w:t>{label}{line}</w:t>"))
        .collect::<Vec<_>>()
        .join("<w:br/>");
    format!(
        concat!(
            "<w:p><w:pPr>{properties}",
            "<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"{height}\" ",
            "w:lineRule=\"exact\"/></w:pPr><w:r>{runs}</w:r></w:p>"
        ),
        properties = properties,
        height = height,
        runs = runs,
    )
}

fn page_lines(body: &str) -> Vec<Vec<String>> {
    let document = format!(
        concat!(
            "<w:document xmlns:w=\"{W}\"><w:body>{body}",
            "<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"3840\"/>",
            "<w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\"/>",
            "</w:sectPr></w:body></w:document>"
        ),
        W = W,
        body = body,
    );
    let content_types = concat!(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
        "<Default Extension=\"rels\" ContentType=\"application/",
        "vnd.openxmlformats-package.relationships+xml\"/>",
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/",
        "vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"
    );
    let relationships = concat!(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
        "<Relationship Id=\"doc\" Type=\"http://schemas.openxmlformats.org/",
        "officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>",
        "</Relationships>"
    );
    let bytes = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            content_types.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), relationships.as_bytes().to_vec()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76700);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {
            "sections": [{"properties": package.document.final_section_properties}],
            "settings": package.settings
        },
        "measurement": {
            "fontChains": {"arial|0|0": [font]},
            "defaults": {"fontFamily": "Arial", "fontSize": 12}
        }
    });
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| fragment["resolvedLines"].as_array().into_iter().flatten())
                .map(|line| {
                    line["segments"]
                        .as_array()
                        .into_iter()
                        .flatten()
                        .filter_map(|segment| segment["text"].as_str())
                        .collect::<String>()
                })
                .filter(|line| !line.is_empty())
                .collect()
        })
        .collect()
}

fn head_fixture(filler_height: u32, properties: &str) -> String {
    [
        paragraph("F", 1, "", filler_height),
        paragraph("P", 4, &format!("<w:keepNext/>{properties}"), 240),
        paragraph("H", 1, "<w:keepNext/>", 240),
        paragraph("B", 1, "", 240),
    ]
    .concat()
}

#[test]
fn keepnext_head_splits() {
    assert_eq!(
        page_lines(&head_fixture(1800, "")),
        [vec!["F1", "P1", "P2"], vec!["P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_head_with_keep_lines_moves_whole() {
    assert_eq!(
        page_lines(&head_fixture(1800, "<w:keepLines/>")),
        [vec!["F1"], vec!["P1", "P2", "P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_head_without_widow_control_leaves_one_line() {
    assert_eq!(
        page_lines(&head_fixture(2040, "<w:widowControl w:val=\"0\"/>")),
        [vec!["F1", "P1"], vec!["P2", "P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_three_member_chain_splits_its_middle() {
    let body = [
        paragraph("F", 1, "", 1560),
        paragraph("H", 1, "<w:keepNext/>", 240),
        paragraph("M", 4, "<w:keepNext/>", 240),
        paragraph("T", 1, "<w:keepNext/>", 240),
        paragraph("B", 1, "", 240),
    ]
    .concat();
    assert_eq!(
        page_lines(&body),
        [vec!["F1", "H1", "M1", "M2"], vec!["M3", "M4", "T1", "B1"]]
    );
}

#[test]
fn keepnext_reserves_the_follower_after_a_head_that_would_fit_whole() {
    assert_eq!(
        page_lines(&head_fixture(1440, "")),
        [vec!["F1", "P1", "P2"], vec!["P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_without_widow_control_carries_only_its_last_line() {
    assert_eq!(
        page_lines(&head_fixture(1440, "<w:widowControl w:val=\"0\"/>")),
        [vec!["F1", "P1", "P2", "P3"], vec!["P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_tail_stays_with_all_single_line_members() {
    assert_eq!(
        page_lines(&head_fixture(1200, "")),
        [vec!["F1", "P1", "P2"], vec!["P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_does_not_orphan_a_single_opening_line() {
    assert_eq!(
        page_lines(&head_fixture(2040, "")),
        [vec!["F1"], vec!["P1", "P2", "P3", "P4", "H1", "B1"]]
    );
}

#[test]
fn keepnext_chain_taller_than_a_page_keeps_ordinary_paragraph_breaks() {
    let body = [
        paragraph("F", 1, "", 600),
        paragraph("A", 4, "<w:keepNext/>", 240),
        paragraph("M", 4, "<w:keepNext/>", 240),
        paragraph("T", 4, "<w:keepNext/>", 240),
        paragraph("B", 1, "", 240),
    ]
    .concat();
    assert_eq!(
        page_lines(&body),
        page_lines(&body.replace("<w:keepNext/>", ""))
    );
    assert_eq!(
        page_lines(&body),
        [
            vec!["F1", "A1", "A2", "A3", "A4", "M1", "M2"],
            vec!["M3", "M4", "T1", "T2", "T3", "T4", "B1"]
        ]
    );
}
