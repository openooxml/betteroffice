use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn section(kind: &str, start: u64, landscape: bool, columns: u64) -> String {
    let (width, height) = if landscape {
        (16838, 11906)
    } else {
        (11906, 16838)
    };
    format!(
        r#"<w:sectPr><w:type w:val="{kind}"/><w:pgSz w:w="{width}" w:h="{height}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:pgNumType w:start="{start}"/><w:cols w:num="{columns}" w:space="720"/></w:sectPr>"#
    )
}

fn paragraph(text: &str, properties: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn document(setting: &str, kind: &str, restart: u64, geometry: &str) -> Vec<u8> {
    let landscape = geometry == "landscape";
    let first_columns = if matches!(geometry, "first-column" | "last-column") {
        2
    } else {
        1
    };
    let columns = if geometry == "two-columns" {
        2
    } else {
        first_columns
    };
    let mut body = String::new();
    if geometry == "last-column" {
        body.push_str(&paragraph("PREFIX", ""));
        body.push_str(r#"<w:p><w:r><w:br w:type="column"/></w:r></w:p>"#);
    }
    body.push_str(&paragraph(
        "FIRST",
        &section("nextPage", 1, false, first_columns),
    ));
    body.push_str(&paragraph(
        "SECOND",
        &section(kind, restart, landscape, columns),
    ));
    body.push_str(&paragraph("THIRD", ""));
    body.push_str(&section("nextPage", 1, landscape, columns));
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#)),
        ("word/settings.xml", format!(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{setting}</w:settings>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn render(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut sections: Vec<_> = package
        .document
        .sections
        .unwrap_or_default()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    sections.push(json!({"properties": package.document.final_section_properties}));
    let engine = EngineSession::new(75200);
    seed_from_docx(engine.doc(), bytes).unwrap();
    serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": sections, "settings": package.settings},
        "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 12}}
    }).to_string()).unwrap()).unwrap()
}

fn page_for(output: &Value, text: &str) -> u64 {
    let block = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["runs"]
                .as_array()
                .is_some_and(|runs| runs.iter().any(|run| run["text"] == text))
        })
        .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .find(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .any(|fragment| fragment["blockId"] == block["block"]["id"])
        })
        .unwrap()["number"]
        .as_u64()
        .unwrap()
}

#[test]
fn imported_facing_page_setting_controls_resident_section_restarts() {
    for (setting, expected_second, expected_third) in [
        ("", 2, 3),
        (r#"<w:evenAndOddHeaders w:val="0"/>"#, 2, 3),
        ("<w:evenAndOddHeaders/>", 3, 5),
    ] {
        let output = render(&document(setting, "nextPage", 1, "same"));
        assert_eq!(page_for(&output, "SECOND"), expected_second, "{setting}");
        assert_eq!(page_for(&output, "THIRD"), expected_third, "{setting}");
        assert_eq!(
            output["layout"]["pages"].as_array().unwrap().len(),
            expected_third as usize
        );
    }
}

#[test]
fn imported_promoted_and_column_section_restarts_match_word() {
    for (kind, geometry) in [
        ("continuous", "landscape"),
        ("nextColumn", "landscape"),
        ("nextColumn", "two-columns"),
        ("nextColumn", "same"),
        ("nextColumn", "first-column"),
        ("nextColumn", "last-column"),
    ] {
        for restart in [1, 2] {
            let (second, third) = match geometry {
                "first-column" => (1, 3),
                "last-column" => (2, 3),
                _ if restart == 1 => (3, 5),
                _ => (2, 3),
            };
            let output = render(&document("<w:evenAndOddHeaders/>", kind, restart, geometry));
            let case = format!("{kind}, {geometry}, restart={restart}");
            assert_eq!(page_for(&output, "SECOND"), second, "{case}");
            assert_eq!(page_for(&output, "THIRD"), third, "{case}");
            assert_eq!(
                output["layout"]["pages"].as_array().unwrap().len(),
                third as usize,
                "{case}"
            );
        }
    }
}
