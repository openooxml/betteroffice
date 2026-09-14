use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn paragraph(text: &str, properties: &str) -> String {
    let run = if text.is_empty() {
        String::new()
    } else {
        format!("<w:r><w:t>{text}</w:t></w:r>")
    };
    format!("<w:p><w:pPr>{properties}</w:pPr>{run}</w:p>")
}

fn marker(text: &str, table: bool, different: bool, boundary: bool) -> String {
    let style = if different {
        r#"<w:pStyle w:val="Different"/>"#
    } else {
        ""
    };
    let line = if boundary && text == "FIRST" {
        r#" w:line="2160" w:lineRule="exact""#
    } else {
        ""
    };
    let content = paragraph(
        text,
        &format!(r#"{style}<w:spacing w:before="0" w:after="0"{line}/>"#),
    );
    if !table {
        return content;
    }
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>"#
    )
}

fn document(context: &str, kind: &str, different: bool) -> Vec<u8> {
    let table = matches!(context, "tables" | "floating-tables");
    let boundary = context == "boundary";
    let default_after = if matches!(kind, "defaults" | "zero" | "keep-next" | "contextual") {
        240
    } else {
        0
    };
    let mut properties = match kind {
        "style" => r#"<w:pStyle w:val="Spaced"/>"#.to_owned(),
        "direct" => r#"<w:spacing w:after="240"/>"#.to_owned(),
        "zero" => r#"<w:spacing w:after="0"/>"#.to_owned(),
        "keep-next" => "<w:keepNext/>".to_owned(),
        "contextual" => "<w:contextualSpacing/>".to_owned(),
        _ => String::new(),
    };
    if context == "page-start" {
        properties.push_str("<w:pageBreakBefore/>");
    }
    if boundary && kind == "direct" {
        properties.push_str("<w:keepNext/>");
    }
    let mut body = marker("FIRST", table, false, boundary);
    body.push_str(&paragraph("", &properties));
    let mut last = marker("LAST", table, different, boundary);
    if context == "floating-tables" {
        last = last.replacen(
            "<w:tblPr>",
            r#"<w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="page" w:tblpY="2000"/>"#,
            1,
        );
    }
    body.push_str(&last);
    let (height, margin) = if boundary { (4320, 720) } else { (15840, 1440) };
    body.push_str(&format!(r#"<w:sectPr><w:pgSz w:w="12240" w:h="{height}"/><w:pgMar w:top="{margin}" w:right="1440" w:bottom="{margin}" w:left="1440"/></w:sectPr>"#));
    let styles = format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="{default_after}" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Different"><w:name w:val="Different"/><w:basedOn w:val="Normal"/></w:style></w:styles>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#)),
        ("word/styles.xml", styles),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn render(context: &str, kind: &str, different: bool) -> Value {
    let bytes = document(context, kind, different);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76400);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    }).to_string()).unwrap()).unwrap()
}

fn last_fragment(output: &Value) -> (u64, f64) {
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .rev()
        .find_map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .last()
                .map(|fragment| {
                    (
                        page["number"].as_u64().unwrap(),
                        fragment["y"].as_f64().unwrap(),
                    )
                })
        })
        .unwrap()
}

#[test]
fn imported_empty_paragraph_after_spacing_matches_word() {
    for context in ["text", "tables", "page-start"] {
        let (base_page, base_y) = last_fragment(&render(context, "none", false));
        for kind in [
            "defaults",
            "style",
            "direct",
            "zero",
            "keep-next",
            "contextual",
        ] {
            let (page, y) = last_fragment(&render(context, kind, false));
            let extra = if matches!(kind, "zero" | "contextual") {
                0.0
            } else {
                16.0
            };
            assert_eq!(page, base_page, "{context}, {kind}");
            assert!(
                (y - base_y - extra).abs() < 0.01,
                "{context}, {kind}: {}",
                y - base_y
            );
        }
    }
}

#[test]
fn contextual_empty_paragraph_after_spacing_requires_matching_next_style() {
    for context in ["text", "tables"] {
        let (_, base_y) = last_fragment(&render(context, "none", false));
        let (_, y) = last_fragment(&render(context, "contextual", true));
        assert!(
            (y - base_y - 16.0).abs() < 0.01,
            "{context}: {}",
            y - base_y
        );
    }
}

#[test]
fn empty_keep_next_spacing_fits_once_at_the_page_boundary() {
    for kind in ["keep-next", "direct"] {
        let output = render("boundary", kind, false);
        assert_eq!(
            output["layout"]["pages"].as_array().unwrap().len(),
            1,
            "{kind}"
        );
        let (_, y) = last_fragment(&output);
        assert!((y - 224.0).abs() < 0.01, "{kind}: {y}");
    }
}

#[test]
fn floating_table_does_not_collapse_contextual_separator_spacing() {
    let output = render("floating-tables", "contextual", false);
    let separator = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["kind"] == "paragraph"
                && measured["block"]["runs"]
                    .as_array()
                    .is_some_and(Vec::is_empty)
        })
        .unwrap();
    assert_eq!(separator["block"]["attrs"]["spacing"]["after"], 16.0);
}
