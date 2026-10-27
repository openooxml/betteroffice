use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn document(flag: &str, mode: Option<u8>, snap: &str) -> Vec<u8> {
    let paragraph = format!(r#"<w:p><w:pPr>{snap}</w:pPr><w:r><w:t>Cell</w:t></w:r></w:p>"#);
    let row = format!(
        r#"<w:tr><w:trPr><w:trHeight w:val="375" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{paragraph}</w:tc></w:tr>"#
    );
    let rows = row.repeat(3);
    let mode = mode.map_or_else(String::new, |mode| {
        format!(
            r#"<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>"#
        )
    });
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#.to_owned()),
        ("word/styles.xml", r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="30"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>"#.to_owned()),
        ("word/settings.xml", format!(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>{flag}{mode}</w:compat></w:settings>"#)),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paragraph}<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>{rows}</w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:docGrid w:type="linesAndChars" w:linePitch="323"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn render(flag: &str, mode: Option<u8>, snap: &str) -> Value {
    let bytes = document(flag, mode, snap);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76403);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {
                        "sections": [{"properties": package.document.final_section_properties}],
                        "settings": package.settings
                    },
                    "measurement": {
                        "fontChains": {"arial|0|0": [font]},
                        "defaults": {"fontFamily": "Arial", "fontSize": 15}
                    }
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn close(actual: &Value, expected: f64) {
    let actual = actual.as_f64().unwrap();
    assert!((actual - expected).abs() < 0.01, "{actual} != {expected}");
}

fn assert_rows(output: &Value, height: f64) {
    let table = &output["measured"][1]["measure"];
    for row in table["rows"].as_array().unwrap() {
        close(&row["height"], height);
    }
    close(&table["totalHeight"], 3.0 * height);
}

#[test]
fn table_rows_ignore_section_line_grid_without_adjust_line_height_in_table() {
    for mode in [None, Some(12), Some(14), Some(15)] {
        for flag in [
            "",
            r#"<w:adjustLineHeightInTable w:val="false"/>"#,
            r#"<w:adjustLineHeightInTable w:val="0"/>"#,
        ] {
            let output = render(flag, mode, "");
            assert_rows(&output, 25.0);
            close(
                &output["measured"][0]["measure"]["totalHeight"],
                646.0 / 15.0,
            );
        }
    }
}

#[test]
fn adjust_line_height_in_table_preserves_cell_grid_and_paragraph_opt_outs() {
    for mode in [None, Some(12), Some(14), Some(15)] {
        for flag in [
            "<w:adjustLineHeightInTable/>",
            r#"<w:adjustLineHeightInTable w:val="true"/>"#,
        ] {
            let output = render(flag, mode, "");
            assert_rows(&output, 646.0 / 15.0);
            close(
                &output["measured"][0]["measure"]["totalHeight"],
                646.0 / 15.0,
            );
            let output = render(flag, mode, r#"<w:snapToGrid w:val="false"/>"#);
            assert_rows(&output, 25.0);
            close(
                &output["measured"][0]["measure"]["totalHeight"],
                18.3984375 * 15.0 / 12.0,
            );
        }
    }
}
