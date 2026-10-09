use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn hf_para(text: &str, style: Option<&str>, spacing: &str, contextual: bool) -> String {
    let style_xml = style
        .map(|name| format!(r#"<w:pStyle w:val="{name}"/>"#))
        .unwrap_or_default();
    let ctx = if contextual {
        "<w:contextualSpacing/>"
    } else {
        ""
    };
    format!(
        r#"<w:p><w:pPr>{style_xml}{spacing}{ctx}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn styles_xml() -> String {
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Different"><w:name w:val="Different"/><w:basedOn w:val="Normal"/></w:style></w:styles>"#.to_owned()
}

fn styles_xml_custom_default() -> String {
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="BodyDefault"><w:name w:val="BodyDefault"/></w:style><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Different"><w:name w:val="Different"/><w:basedOn w:val="Normal"/></w:style></w:styles>"#.to_owned()
}

fn document(header: Option<(String, String)>, footer: Option<(String, String)>) -> Vec<u8> {
    document_with_styles(header, footer, styles_xml())
}

fn document_with_styles(
    header: Option<(String, String)>,
    footer: Option<(String, String)>,
    styles: String,
) -> Vec<u8> {
    let body = r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>BODY</w:t></w:r></w:p>"#;
    let (header_ref, header_part, header_types, header_rel) = match &header {
        Some((first, second)) => (
            r#"<w:headerReference w:type="default" r:id="rIdHeader"/>"#.to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{first}{second}</w:hdr>"#
            ),
            r#"<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#.to_owned(),
            r#"<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>"#.to_owned(),
        ),
        None => (String::new(), String::new(), String::new(), String::new()),
    };
    let (footer_ref, footer_part, footer_types, footer_rel) = match &footer {
        Some((first, second)) => (
            r#"<w:footerReference w:type="default" r:id="rIdFooter"/>"#.to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{first}{second}</w:ftr>"#
            ),
            r#"<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>"#.to_owned(),
            r#"<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>"#.to_owned(),
        ),
        None => (String::new(), String::new(), String::new(), String::new()),
    };
    let sect = format!(
        r#"<w:sectPr>{header_ref}{footer_ref}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360"/><w:cols w:space="720"/></w:sectPr>"#
    );
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>{header_types}{footer_types}</Types>"#
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>{header_rel}{footer_rel}</Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}{sect}</w:body></w:document>"#
            ),
        ),
        ("word/styles.xml".to_owned(), styles),
    ];
    if header.is_some() {
        parts.push(("word/header1.xml".to_owned(), header_part));
    }
    if footer.is_some() {
        parts.push(("word/footer1.xml".to_owned(), footer_part));
    }
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn layout(bytes: &[u8]) -> (String, Value) {
    layout_with_default(bytes, None)
}

fn layout_with_default(bytes: &[u8], default_style: Option<&str>) -> (String, Value) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76401);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let render_env = match default_style {
        Some(style) => json!({"defaultParagraphStyleId": style}),
        None => json!({}),
    };
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": render_env,
                "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
                "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
            })
            .to_string(),
        )
        .unwrap();
    let value: Value = serde_json::from_str(&output).unwrap();
    (output, value)
}

fn block_style_id(output: &Value, kind: &str, index: usize) -> Option<String> {
    variant(output, kind)["measured"][index]["block"]["attrs"]["styleId"]
        .as_str()
        .map(str::to_owned)
}

fn block_effective_style_id(output: &Value, kind: &str, index: usize) -> Option<String> {
    variant(output, kind)["measured"][index]["block"]["attrs"]["effectiveStyleId"]
        .as_str()
        .map(str::to_owned)
}

fn header_story_paragraph_styles(bytes: &[u8]) -> Vec<Option<String>> {
    let engine = EngineSession::new(76403);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
        .doc()
        .paragraphs("hf:rIdHeader")
        .unwrap()
        .into_iter()
        .map(|paragraph| {
            paragraph
                .properties
                .get("pStyle")
                .and_then(|style| match style {
                    yrs::Any::String(style) => Some(style.to_string()),
                    _ => None,
                })
        })
        .collect()
}

fn variant<'a>(output: &'a Value, kind: &str) -> &'a Value {
    output["headersFooters"]["variants"]
        .as_array()
        .unwrap()
        .iter()
        .find(|variant| variant["kind"] == kind)
        .unwrap()
}

fn first_after(output: &Value, kind: &str) -> f64 {
    variant(output, kind)["measured"][0]["block"]["attrs"]["spacing"]["after"]
        .as_f64()
        .unwrap_or(0.0)
}

fn flow_height(output: &Value, kind: &str) -> f64 {
    variant(output, kind)["flowHeight"].as_f64().unwrap()
}

fn last_after(output: &Value, kind: &str) -> f64 {
    variant(output, kind)["measured"][1]["block"]["attrs"]["spacing"]["after"]
        .as_f64()
        .unwrap_or(0.0)
}

fn header_layout_and_baselines(bytes: &[u8]) -> (Value, f64, f64) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(76402);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string();
    let output_str = engine.layout_document_with_regions_json(&request).unwrap();
    let output: Value = serde_json::from_str(&output_str).unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output_str).unwrap()).unwrap();
    let primitives = display["pages"][0]["header"]["primitives"]
        .as_array()
        .unwrap();
    let baseline_for = |suffix: &str| {
        primitives
            .iter()
            .find(|primitive| {
                primitive["blockKey"]
                    .as_str()
                    .is_some_and(|key| key.ends_with(suffix))
            })
            .unwrap()["baselineY"]
            .as_f64()
            .unwrap()
    };
    let first = baseline_for(":p0");
    let second = baseline_for(":p1");
    (output, first, second)
}

const EXPLICIT: &str = r#"<w:spacing w:after="240" w:line="240" w:lineRule="exact"/>"#;

#[test]
fn header_same_style_contextual_collapses_gap_and_body_margin() {
    let collapsed = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
        None,
    );
    let retained = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, false),
            hf_para("SECOND", None, EXPLICIT, false),
        )),
        None,
    );
    let (collapsed_out, collapsed_first, collapsed_second) =
        header_layout_and_baselines(&collapsed);
    let (retained_out, retained_first, retained_second) = header_layout_and_baselines(&retained);
    assert_eq!(first_after(&collapsed_out, "header"), 0.0);
    assert!((first_after(&retained_out, "header") - 16.0).abs() < 0.01);
    assert!((last_after(&collapsed_out, "header") - 16.0).abs() < 0.01);
    let collapsed_flow = flow_height(&collapsed_out, "header");
    let retained_flow = flow_height(&retained_out, "header");
    assert!((retained_flow - collapsed_flow - 16.0).abs() < 0.01);
    let collapsed_top = collapsed_out["options"]["margins"]["top"].as_f64().unwrap();
    let retained_top = retained_out["options"]["margins"]["top"].as_f64().unwrap();
    assert!((retained_top - collapsed_top - 16.0).abs() < 0.01);
    assert!((collapsed_top - (24.0 + collapsed_flow)).abs() < 0.01);
    assert!((collapsed_second - collapsed_first - 16.0).abs() < 0.5);
    assert!((retained_second - retained_first - 32.0).abs() < 0.5);
    let collapsed_visual = variant(&collapsed_out, "header")["visualBottom"]
        .as_f64()
        .unwrap();
    assert!((collapsed_visual - collapsed_flow).abs() < 0.01);
}

#[test]
fn header_different_style_retains_gap_as_control() {
    let collapsed = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
        None,
    );
    let different = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", Some("Different"), EXPLICIT, true),
        )),
        None,
    );
    let (_, collapsed_out) = layout(&collapsed);
    let (_, different_out) = layout(&different);
    assert!((first_after(&different_out, "header") - 16.0).abs() < 0.01);
    let collapsed_flow = flow_height(&collapsed_out, "header");
    let different_flow = flow_height(&different_out, "header");
    assert!((different_flow - collapsed_flow - 16.0).abs() < 0.01);
}

#[test]
fn header_default_spacing_collapses_like_explicit() {
    let explicit = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
        None,
    );
    let default = document(
        Some((
            hf_para("FIRST", None, "", true),
            hf_para("SECOND", None, "", true),
        )),
        None,
    );
    let (_, explicit_out) = layout(&explicit);
    let (_, default_out) = layout(&default);
    assert_eq!(first_after(&default_out, "header"), 0.0);
    let explicit_flow = flow_height(&explicit_out, "header");
    let default_flow = flow_height(&default_out, "header");
    assert!((default_flow - explicit_flow).abs() < 0.01);
}

#[test]
fn footer_same_style_contextual_collapses_gap_and_bottom_margin() {
    let collapsed = document(
        None,
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
    );
    let retained = document(
        None,
        Some((
            hf_para("FIRST", None, EXPLICIT, false),
            hf_para("SECOND", None, EXPLICIT, false),
        )),
    );
    let (_, collapsed_out) = layout(&collapsed);
    let (_, retained_out) = layout(&retained);
    assert_eq!(first_after(&collapsed_out, "footer"), 0.0);
    assert!((first_after(&retained_out, "footer") - 16.0).abs() < 0.01);
    let collapsed_flow = flow_height(&collapsed_out, "footer");
    let retained_flow = flow_height(&retained_out, "footer");
    assert!((retained_flow - collapsed_flow - 16.0).abs() < 0.01);
    let collapsed_bottom = collapsed_out["options"]["margins"]["bottom"]
        .as_f64()
        .unwrap();
    let retained_bottom = retained_out["options"]["margins"]["bottom"]
        .as_f64()
        .unwrap();
    assert!((retained_bottom - collapsed_bottom - 16.0).abs() < 0.01);
}

#[test]
fn header_implicit_and_explicit_default_collapse_both_orders() {
    let implicit_first = document(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", Some("Normal"), EXPLICIT, true),
        )),
        None,
    );
    let explicit_first = document(
        Some((
            hf_para("FIRST", Some("Normal"), EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
        None,
    );
    let (_, implicit_first_out) = layout_with_default(&implicit_first, Some("Normal"));
    let (_, explicit_first_out) = layout_with_default(&explicit_first, Some("Normal"));
    assert_eq!(first_after(&implicit_first_out, "header"), 0.0);
    assert_eq!(first_after(&explicit_first_out, "header"), 0.0);
    assert_eq!(block_style_id(&implicit_first_out, "header", 0), None);
    assert_eq!(
        block_style_id(&implicit_first_out, "header", 1).as_deref(),
        Some("Normal")
    );
    assert_eq!(
        block_effective_style_id(&implicit_first_out, "header", 0).as_deref(),
        Some("Normal")
    );
    assert_eq!(
        block_style_id(&explicit_first_out, "header", 0).as_deref(),
        Some("Normal")
    );
    assert_eq!(block_style_id(&explicit_first_out, "header", 1), None);
    assert_eq!(
        block_effective_style_id(&explicit_first_out, "header", 1).as_deref(),
        Some("Normal")
    );
    let (_, implicit_first_legacy) = layout(&implicit_first);
    assert!((first_after(&implicit_first_legacy, "header") - 16.0).abs() < 0.01);
    assert_eq!(
        header_story_paragraph_styles(&implicit_first),
        vec![None, Some("Normal".to_owned())]
    );
}

#[test]
fn header_custom_default_keeps_explicit_normal_distinct() {
    let implicit_normal = document_with_styles(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", Some("Normal"), EXPLICIT, true),
        )),
        None,
        styles_xml_custom_default(),
    );
    let implicit_body_default = document_with_styles(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", Some("BodyDefault"), EXPLICIT, true),
        )),
        None,
        styles_xml_custom_default(),
    );
    let explicit_body_default_first = document_with_styles(
        Some((
            hf_para("FIRST", Some("BodyDefault"), EXPLICIT, true),
            hf_para("SECOND", None, EXPLICIT, true),
        )),
        None,
        styles_xml_custom_default(),
    );
    let implicit_different = document_with_styles(
        Some((
            hf_para("FIRST", None, EXPLICIT, true),
            hf_para("SECOND", Some("Different"), EXPLICIT, true),
        )),
        None,
        styles_xml_custom_default(),
    );
    let (_, normal_out) = layout_with_default(&implicit_normal, Some("BodyDefault"));
    let (_, default_out) = layout_with_default(&implicit_body_default, Some("BodyDefault"));
    let (_, reversed_out) = layout_with_default(&explicit_body_default_first, Some("BodyDefault"));
    let (_, different_out) = layout_with_default(&implicit_different, Some("BodyDefault"));
    assert!((first_after(&normal_out, "header") - 16.0).abs() < 0.01);
    assert_eq!(first_after(&default_out, "header"), 0.0);
    assert_eq!(first_after(&reversed_out, "header"), 0.0);
    assert!((first_after(&different_out, "header") - 16.0).abs() < 0.01);
    assert_eq!(block_style_id(&normal_out, "header", 0), None);
    assert_eq!(
        block_style_id(&normal_out, "header", 1).as_deref(),
        Some("Normal")
    );
    assert_eq!(
        block_effective_style_id(&normal_out, "header", 0).as_deref(),
        Some("BodyDefault")
    );
}
