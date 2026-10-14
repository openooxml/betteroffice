use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSerif-Italic.ttf");
const SINGLE_10PT: f64 = 2355.0 / 2048.0 * 10.0 * 96.0 / 72.0;

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn run(text: &str, size: u32, properties: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="{size}"/>{properties}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

fn paragraph(runs: &str, spacing: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" {spacing}/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="20"/></w:rPr></w:pPr>{runs}</w:p>"#
    )
}

fn mixed(kind: &str, size: u32, extra: &str, spacing: &str) -> String {
    let properties = format!(r#"<w:vertAlign w:val="{kind}"/>{extra}"#);
    paragraph(
        &format!(
            "{}{}{}",
            run("A", 20, ""),
            run("X", size, &properties),
            run(" End", 20, "")
        ),
        spacing,
    )
}

fn layout(body: &str) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(74850);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "options": {"pageSize":{"w":400,"h":228},
                        "margins":{"top":48,"bottom":48,"left":48,"right":48}},
                    "measurement": {"fontChains": {
                        "times new roman|0|0":[font], "times new roman|0|1":[font]},
                        "defaults":{"fontFamily":"Times New Roman","fontSize":10}}
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn near(actual: &Value, expected: f64) {
    assert!(
        (actual.as_f64().unwrap() - expected).abs() < 0.002,
        "expected {expected}, got {actual}"
    );
}

fn line(output: &Value) -> &Value {
    &output["measured"][0]["measure"]["lines"][0]
}

#[test]
fn repeated_script_prefixes_preserve_word_line_pitch_and_fit_the_page() {
    for kind in ["superscript", "subscript"] {
        let body = mixed(kind, 20, "", r#"w:line="240" w:lineRule="auto""#).repeat(8);
        let output = layout(&body);
        assert_eq!(
            output["layout"]["pages"].as_array().unwrap().len(),
            1,
            "{kind}"
        );
        let measured = output["measured"].as_array().unwrap();
        assert_eq!(measured.len(), 8);
        for paragraph in measured {
            near(&paragraph["measure"]["lines"][0]["lineHeight"], SINGLE_10PT);
        }
        let fragments = output["layout"]["pages"][0]["fragments"]
            .as_array()
            .unwrap();
        near(&fragments[7]["y"], 48.0 + SINGLE_10PT * 7.0);
    }
}

#[test]
fn larger_scripted_run_reserves_its_original_size_with_a_smaller_paragraph_mark() {
    for kind in ["superscript", "subscript"] {
        let output = layout(&mixed(kind, 40, "", r#"w:line="240" w:lineRule="auto""#));
        near(&line(&output)["lineHeight"], SINGLE_10PT * 2.0);
        let standalone = layout(&paragraph(
            &run("X", 40, &format!(r#"<w:vertAlign w:val="{kind}"/>"#)),
            r#"w:line="240" w:lineRule="auto""#,
        ));
        near(&line(&standalone)["lineHeight"], SINGLE_10PT * 2.0);
    }
}

#[test]
fn scripted_fields_reserve_the_original_run_size() {
    for kind in ["superscript", "subscript"] {
        let field = format!(
            r#"<w:fldSimple w:instr="PAGE" w:fldLock="true">{}</w:fldSimple>"#,
            run("1", 40, &format!(r#"<w:vertAlign w:val="{kind}"/>"#))
        );
        let output = layout(&paragraph(
            &format!("{}{field}", run("A", 20, "")),
            r#"w:line="240" w:lineRule="auto""#,
        ));
        assert!(
            output["measured"][0]["block"]["runs"]
                .as_array()
                .unwrap()
                .iter()
                .any(|run| run["kind"] == "field")
        );
        near(&line(&output)["lineHeight"], SINGLE_10PT * 2.0);
    }
}

#[test]
fn explicit_line_spacing_still_controls_scripted_lines() {
    for kind in ["superscript", "subscript"] {
        for (spacing, expected) in [
            (r#"w:line="240" w:lineRule="exact""#, 16.0),
            (r#"w:line="600" w:lineRule="atLeast""#, 40.0),
            (r#"w:line="240" w:lineRule="atLeast""#, SINGLE_10PT * 2.0),
            (r#"w:line="480" w:lineRule="auto""#, SINGLE_10PT * 4.0),
        ] {
            near(
                &line(&layout(&mixed(kind, 40, "", spacing)))["lineHeight"],
                expected,
            );
        }
    }
}

#[test]
fn script_paint_size_width_and_explicit_position_remain_independent_of_line_reservation() {
    for (kind, script_shift) in [("superscript", -4.0), ("subscript", 2.0)] {
        let spacing = r#"w:line="240" w:lineRule="auto""#;
        let ordinary = layout(&mixed(kind, 20, "", spacing));
        let positioned = layout(&mixed(kind, 20, r#"<w:position w:val="6"/>"#, spacing));
        assert_eq!(
            ordinary["measured"][0]["measure"],
            positioned["measured"][0]["measure"]
        );
        let display = |output: &Value| -> Value {
            serde_json::from_str(
                &docx_layout::display_list::build_display_list_json(&output.to_string()).unwrap(),
            )
            .unwrap()
        };
        let plain_display = display(&ordinary);
        let shifted_display = display(&positioned);
        let text = |value: &Value, label: &str| -> Value {
            value["pages"][0]["primitives"]
                .as_array()
                .unwrap()
                .iter()
                .find(|p| p["kind"] == "text" && p["text"] == label)
                .unwrap()
                .clone()
        };
        let normal = text(&plain_display, "A");
        let script = text(&plain_display, "X");
        let shifted = text(&shifted_display, "X");
        assert!(script["font"].as_str().unwrap().contains("10px"));
        near(
            &script["baselineY"],
            normal["baselineY"].as_f64().unwrap() + script_shift,
        );
        near(
            &shifted["baselineY"],
            script["baselineY"].as_f64().unwrap() - 4.0,
        );
        assert_eq!(script["width"], shifted["width"]);
        assert_eq!(script["font"], shifted["font"]);
        let reduced = layout(&paragraph(&run("X", 15, ""), spacing));
        near(&script["width"], line(&reduced)["width"].as_f64().unwrap());
    }
}
