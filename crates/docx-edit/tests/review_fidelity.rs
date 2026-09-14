use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const HORIZONTAL_RULE: &str = r##"<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:rect style="width:0pt;height:1.5pt" o:hr="t" o:hrstd="t" o:hralign="center" fillcolor="#A0A0A0" stroked="f"/></w:pict>"##;
const SHAPE: &str = r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="66675" distR="123825" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Inside shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;

fn document(body: &str, styles: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/styles.xml", format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{styles}</w:styles>"#)),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn paragraph(text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn layout(body: &str, columns: u32) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(74201);
    seed_from_docx(engine.doc(), &document(body, "")).unwrap();
    serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory": "body", "renderEnv": {},
        "options": {"pageSize": {"w":816,"h":1056},
            "margins":{"top":96,"bottom":96,"left":96,"right":96},
            "columns":{"count":columns,"gap":48}},
        "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}
    }).to_string()).unwrap()).unwrap()
}

#[test]
fn vml_horizontal_rule_uses_paragraph_width_without_changing_flow() {
    for columns in [1, 2] {
        let body = format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r>{HORIZONTAL_RULE}</w:r></w:p>{}"#,
            paragraph("After")
        );
        let output = layout(&body, columns);
        let blank = layout(
            &format!(
                r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>{}"#,
                paragraph("After")
            ),
            columns,
        );
        let fragments = &output["layout"]["pages"][0]["fragments"];
        assert_eq!(
            fragments[1]["y"],
            blank["layout"]["pages"][0]["fragments"][1]["y"]
        );
        assert_eq!(
            output["measured"][0]["block"]["attrs"]["horizontalRules"][0]["height"],
            2.0
        );
        let display: Value = serde_json::from_str(
            &docx_layout::display_list::build_display_list_json(&output.to_string()).unwrap(),
        )
        .unwrap();
        let lines: Vec<_> = display["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|primitive| primitive["kind"] == "line")
            .collect();
        assert_eq!(lines.len(), 4);
        assert_eq!(
            lines[0]["x2"].as_f64().unwrap() - lines[0]["x1"].as_f64().unwrap(),
            fragments[0]["width"].as_f64().unwrap()
        );
        assert_eq!(lines[0]["docStart"], 1);
        assert_eq!(lines[0]["docEnd"], 2);
        let display: docx_layout::display_list::DisplayList =
            serde_json::from_value(display.clone()).unwrap();
        for position in [1, 2] {
            let caret = docx_layout::hit::caret_rect(&display, position).unwrap();
            assert_eq!(caret.page_index, 0);
            assert!(caret.y < fragments[1]["y"].as_f64().unwrap());
        }
        let y = lines[0]["y1"].as_f64().unwrap();
        assert!(matches!(
            docx_layout::hit::hit_test(&display, 0, 200.0, y),
            Some(1 | 2)
        ));
    }
}

#[test]
fn vml_horizontal_rule_respects_inline_baselines_and_exact_spacing() {
    for mixed in [false, true] {
        for exact in [false, true] {
            for height in [1.5, 24.0] {
                let rule = HORIZONTAL_RULE.replace("height:1.5pt", &format!("height:{height}pt"));
                let before = if mixed { "<w:t>Before </w:t>" } else { "" };
                let after = if mixed { "<w:t> After</w:t>" } else { "" };
                let line_rule = if exact { "exact" } else { "auto" };
                let body = format!(
                    r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="{line_rule}"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr>{before}{rule}{after}</w:r></w:p>{}"#,
                    paragraph("After")
                );
                let output = layout(&body, 1);
                let plain = layout(&body.replace(&rule, ""), 1);
                let measure = &output["measured"][0]["measure"];
                let line = &measure["lines"][0];
                let plain_line = &plain["measured"][0]["measure"]["lines"][0];
                assert_eq!(measure["lines"].as_array().unwrap().len(), 1);
                let actual = line["lineHeight"].as_f64().unwrap();
                if exact || height == 1.5 {
                    assert!((actual - plain_line["lineHeight"].as_f64().unwrap()).abs() < 0.01);
                } else if mixed {
                    assert_eq!(line["ascent"], 33.0);
                    assert!(actual > 33.0);
                } else {
                    assert_eq!(actual, 33.0);
                    assert_eq!(line["ascent"], 33.0);
                    assert_eq!(line["descent"], 0.0);
                }
                if exact || mixed && height == 1.5 {
                    assert!(
                        (line["ascent"].as_f64().unwrap() - plain_line["ascent"].as_f64().unwrap())
                            .abs()
                            < 0.01
                    );
                }
                let fragments = &output["layout"]["pages"][0]["fragments"];
                assert!(
                    (fragments[1]["y"].as_f64().unwrap()
                        - fragments[0]["y"].as_f64().unwrap()
                        - actual)
                        .abs()
                        < 1e-5
                );
                let display: Value = serde_json::from_str(
                    &docx_layout::display_list::build_display_list_json(&output.to_string())
                        .unwrap(),
                )
                .unwrap();
                let lines: Vec<_> = display["pages"][0]["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|primitive| primitive["kind"] == "line")
                    .collect();
                assert_eq!(lines.len(), 4);
                let y = fragments[0]["y"].as_f64().unwrap();
                let baseline = y
                    + line["ascent"].as_f64().unwrap()
                    + ((actual
                        - line["ascent"].as_f64().unwrap()
                        - line["descent"].as_f64().unwrap())
                        / 2.0)
                        .max(0.0);
                assert!(
                    (lines[0]["y1"].as_f64().unwrap() - (baseline - height * 4.0 / 3.0)).abs()
                        < 0.01
                );
                if mixed {
                    assert!(lines[0]["x1"].as_f64().unwrap() > 110.0);
                    assert_eq!(lines[0]["x1"], lines[0]["x2"]);
                }
            }
        }
    }
}

#[test]
fn vml_horizontal_rule_preserves_run_font_and_follows_hard_breaks() {
    let output = layout(
        &format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>Before</w:t><w:br/></w:r><w:r><w:rPr><w:sz w:val="72"/></w:rPr>{HORIZONTAL_RULE}</w:r></w:p>"#
        ),
        1,
    );
    let measured = &output["measured"][0];
    assert_eq!(measured["block"]["runs"][2]["fontSize"], 36.0);
    assert_eq!(measured["measure"]["lines"].as_array().unwrap().len(), 2);
    let line = &measured["measure"]["lines"][1];
    assert!(line["lineHeight"].as_f64().unwrap() > 48.0);
    assert_eq!(line["ascent"], line["lineHeight"]);
    let display: Value = serde_json::from_str(
        &docx_layout::display_list::build_display_list_json(&output.to_string()).unwrap(),
    )
    .unwrap();
    let lines: Vec<_> = display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| primitive["kind"] == "line")
        .collect();
    assert_eq!(lines.len(), 4);
    assert_eq!(
        lines[0]["x2"].as_f64().unwrap() - lines[0]["x1"].as_f64().unwrap(),
        624.0
    );
}

#[test]
fn break_only_paragraphs_match_word_page_and_column_flow() {
    for (kinds, columns, page_count, x, y) in [
        (vec!["column"], 2, 1, 432.0, 160.0),
        (vec!["page"], 1, 2, 96.0, 96.0),
        (vec!["column", "column"], 2, 2, 96.0, 160.0),
        (vec!["page", "column"], 2, 2, 432.0, 160.0),
        (vec!["column", "page"], 2, 2, 96.0, 96.0),
    ] {
        let breaks = kinds
            .iter()
            .map(|kind| format!(r#"<w:br w:type="{kind}"/>"#))
            .collect::<String>();
        let body = format!(
            r#"{}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="960" w:lineRule="exact"/></w:pPr><w:r>{breaks}</w:r></w:p>{}"#,
            paragraph("Before"),
            paragraph("After")
        );
        let output = layout(&body, columns);
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), page_count, "{kinds:?}");
        let after = pages.last().unwrap()["fragments"]
            .as_array()
            .unwrap()
            .last()
            .unwrap();
        assert_eq!(after["y"].as_f64().unwrap(), y, "{kinds:?}");
        assert_eq!(after["x"].as_f64().unwrap(), x, "{kinds:?}");
    }
}

#[test]
fn inside_shape_wrap_matches_actual_page_and_asymmetric_distances() {
    for (prefix, expected_page) in [
        (String::new(), 1),
        (
            format!(
                r#"{}<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#,
                paragraph("First page")
            ),
            2,
        ),
        (paragraph("Automatic pagination").repeat(30), 2),
    ] {
        let body = format!(
            "{prefix}<w:p>{SHAPE}<w:r><w:t>{}</w:t></w:r></w:p>",
            "Body text wraps around the shape. ".repeat(20)
        );
        let output = layout(&body, 1);
        let pages = output["layout"]["pages"].as_array().unwrap();
        let page = &pages[expected_page - 1];
        let fragments = page["fragments"].as_array().unwrap();
        let shape = fragments
            .iter()
            .find(|fragment| fragment["kind"] == "shape")
            .unwrap();
        let text = fragments
            .iter()
            .find(|fragment| {
                fragment["kind"] == "paragraph" && fragment["pmStart"] == shape["pmStart"]
            })
            .unwrap();
        let measure = output["measured"]
            .as_array()
            .unwrap()
            .iter()
            .find(|measured| measured["block"]["id"] == text["blockId"])
            .unwrap();
        let line = &measure["measure"]["lines"][0];
        let start = text["x"].as_f64().unwrap() + line["leftOffset"].as_f64().unwrap_or(0.0);
        let end = start + line["width"].as_f64().unwrap();
        let shape_x = shape["x"].as_f64().unwrap();
        if expected_page == 1 {
            assert_eq!(shape_x, 96.0);
            assert!(start >= shape_x + 192.0 + 13.0);
        } else {
            assert_eq!(shape_x, 528.0);
            assert_eq!(start, 96.0);
            assert!(
                end <= shape_x - 7.0,
                "text ends at {end}, shape starts at {shape_x}"
            );
        }
    }
}

#[test]
fn dark_cells_preserve_colors_inherited_from_styles_and_defaults() {
    let styles = r#"<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="FF0000"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Green"><w:rPr><w:color w:val="00FF00"/></w:rPr></w:style><w:style w:type="character" w:styleId="Blue"><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>"#;
    for (ppr, rpr, expected) in [
        ("", "", "#FF0000"),
        (r#"<w:pStyle w:val="Green"/>"#, "", "#00FF00"),
        ("", r#"<w:rStyle w:val="Blue"/>"#, "#0000FF"),
    ] {
        let body = format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="222222"/></w:tcPr><w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>Inherited color</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#
        );
        let engine = EngineSession::new(74202);
        seed_from_docx(engine.doc(), &document(&body, styles)).unwrap();
        let blocks = engine
            .with_lowered_story("body", &docx_edit::bridge::RenderEnv::default(), |blocks| {
                serde_json::to_value(blocks).unwrap()
            })
            .unwrap();
        assert_eq!(
            blocks[0]["rows"][0]["cells"][0]["blocks"][0]["runs"][0]["color"],
            expected
        );
    }
}

#[test]
fn undeclared_font_size_matches_word_without_overriding_the_style_hierarchy() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for (defaults, normal, expected) in [("", "", 10.0), ("22", "", 11.0), ("22", "24", 12.0)] {
        let size = |value: &str| {
            if value.is_empty() {
                String::new()
            } else {
                format!(r#"<w:sz w:val="{value}"/>"#)
            }
        };
        let styles = format!(
            r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{}</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>{}</w:rPr></w:style>
<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:sz w:val="26"/></w:rPr></w:style>"#,
            size(defaults),
            size(normal),
        );
        let text = "Inherited text must wrap at the same position as an explicit size. ".repeat(6);
        let body = format!(
            r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:sz w:val="{}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>
<w:p/>
<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>Character style</w:t></w:r>
<w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:sz w:val="28"/></w:rPr><w:t>Direct size</w:t></w:r></w:p>"#,
            expected * 2.0,
        );
        let engine = EngineSession::new(74208);
        seed_from_docx(engine.doc(), &document(&body, &styles)).unwrap();
        let before = engine.doc().encode_state_as_update_v1();
        let output: Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(
                    &json!({
                        "bodyStory": "body", "options": {}, "renderEnv": {},
                        "measurement": {
                            "fontChains": {"arial|0|0": [font], "calibri|0|0": [font]},
                            "defaults": {"fontSize": 11, "fontFamily": "Arial"},
                            "authoritativeShaping": true
                        }
                    })
                    .to_string(),
                )
                .unwrap(),
        )
        .unwrap();
        let measured = output["measured"].as_array().unwrap();
        assert_eq!(measured[0]["block"]["runs"][0]["fontSize"], expected);
        assert_eq!(measured[0]["measure"], measured[1]["measure"]);
        assert!(measured[0]["measure"]["lines"].as_array().unwrap().len() > 1);
        assert_eq!(measured[2]["block"]["attrs"]["defaultFontSize"], expected);
        assert_eq!(measured[3]["block"]["runs"][0]["fontSize"], 13.0);
        assert_eq!(measured[3]["block"]["runs"][1]["fontSize"], 14.0);
        engine.build_display_list_json(&output.to_string()).unwrap();
        assert_eq!(engine.doc().encode_state_as_update_v1(), before);
    }
}

#[test]
fn table_paragraph_spacing_overrides_defaults_but_preserves_paragraph_formatting() {
    let styles = r#"<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="279" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
<w:style w:type="paragraph" w:styleId="Spaced"><w:pPr><w:spacing w:after="100" w:line="360" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="table" w:styleId="Base"><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="table" w:styleId="Grid"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="Inner"><w:pPr><w:spacing w:after="60" w:line="300" w:lineRule="auto"/></w:pPr></w:style>"#;
    let body = r#"<w:p><w:r><w:t>Before</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc>
<w:p><w:r><w:t>Table defaults</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Spaced"/></w:pPr><w:r><w:t>Paragraph style</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Spaced"/><w:spacing w:after="200" w:line="480" w:lineRule="auto"/></w:pPr><w:r><w:t>Direct formatting</w:t></w:r></w:p>
<w:sdt><w:sdtContent><w:p><w:r><w:t>Content control</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:tbl><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Nested table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>Outer restored</w:t></w:r></w:p>
</w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74213);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &docx_edit::bridge::RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    let table_cell = &blocks[1]["rows"][0]["cells"][0]["blocks"];
    for (paragraph, after_twips, line) in [
        (&blocks[0], 160.0, 279.0 / 240.0),
        (&table_cell[0], 0.0, 1.0),
        (&table_cell[1], 100.0, 1.5),
        (&table_cell[2], 200.0, 2.0),
        (&table_cell[3], 0.0, 1.0),
        (
            &table_cell[4]["rows"][0]["cells"][0]["blocks"][0],
            60.0,
            1.25,
        ),
        (&table_cell[5], 0.0, 1.0),
        (&blocks[2], 160.0, 279.0 / 240.0),
    ] {
        let spacing = &paragraph["attrs"]["spacing"];
        assert_eq!(spacing["after"], after_twips / 15.0, "{paragraph}");
        assert_eq!(spacing["line"], line, "{paragraph}");
    }
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn table_spacing_also_overrides_application_defaults_when_document_defaults_are_absent() {
    let styles = r#"<w:style w:type="table" w:styleId="Grid"><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>"#;
    let body = r#"<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Single spaced</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let engine = EngineSession::new(74214);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &docx_edit::bridge::RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    let spacing = &blocks[0]["rows"][0]["cells"][0]["blocks"][0]["attrs"]["spacing"];
    assert_eq!(spacing["after"], 0.0);
    assert_eq!(spacing["line"], 1.0);
}

#[test]
fn conditional_table_paragraph_spacing_follows_regions_and_cascade() {
    let body = include_str!(
        "../../../packages/docx/src/yrs/__fixtures__/table-conditional-spacing/body.xml"
    );
    let styles = include_str!(
        "../../../packages/docx/src/yrs/__fixtures__/table-conditional-spacing/styles.xml"
    );
    let engine = EngineSession::new(74219);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &docx_edit::bridge::RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    let expected = [
        vec![
            vec![190, 110, 110, 110, 200],
            vec![130, 170, 180, 170, 140],
            vec![130, 170, 180, 170, 140],
            vec![130, 170, 180, 170, 140],
            vec![210, 120, 120, 120, 220],
        ],
        vec![
            vec![110; 5],
            vec![150; 5],
            vec![150; 5],
            vec![160; 5],
            vec![160; 5],
        ],
        vec![vec![10; 5]; 5],
        vec![
            vec![190, 110, 200],
            vec![170, 180, 170],
            vec![130, 180, 140],
            vec![210, 120, 220],
        ],
        vec![vec![400, 420, 120, 120, 220]],
        vec![vec![130, 170, 170, 180, 180]; 5],
        vec![
            vec![190, 110, 110, 110, 110],
            vec![130, 150, 150, 150, 150],
            vec![130, 160, 160, 160, 160],
            vec![130, 150, 150, 150, 150],
            vec![130, 160, 160, 160, 160],
        ],
        vec![vec![170, 180, 170, 180, 170]; 5],
        vec![vec![170, 180, 170, 180, 170]; 5],
    ];
    let tables = blocks
        .as_array()
        .unwrap()
        .iter()
        .filter(|block| block["kind"] == "table")
        .collect::<Vec<_>>();
    assert_eq!(tables.len(), expected.len());
    for (table, rows) in tables.iter().zip(expected) {
        for (row, cells) in table["rows"].as_array().unwrap().iter().zip(rows) {
            assert_eq!(row["cells"].as_array().unwrap().len(), cells.len());
            for (cell, after) in row["cells"].as_array().unwrap().iter().zip(cells) {
                assert!(
                    (cell["blocks"][0]["attrs"]["spacing"]["after"]
                        .as_f64()
                        .unwrap()
                        - f64::from(after) / 15.0)
                        .abs()
                        < 1e-8,
                    "{cell}"
                );
            }
        }
    }
    assert_eq!(
        tables[0]["rows"][0]["cells"][0]["blocks"][0]["attrs"]["spacing"]["line"],
        1.5
    );
    assert_eq!(
        tables[4]["rows"][0]["cells"][0]["blocks"][0]["attrs"]["spacing"]["line"],
        2.0
    );
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn hidden_runs_preserve_edit_positions_and_can_be_revealed() {
    let body = r#"<w:p><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>😀secret</w:t><w:tab/><w:br/></w:r><w:hyperlink w:anchor="hidden"><w:r><w:rPr><w:vanish/></w:rPr><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>Z</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74209);
    seed_from_docx(engine.doc(), &document(body, "")).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let lower = |show_hidden_text| {
        engine
            .with_lowered_story(
                "body",
                &docx_edit::bridge::RenderEnv {
                    show_hidden_text,
                    ..Default::default()
                },
                |blocks| serde_json::to_value(blocks).unwrap(),
            )
            .unwrap()
    };
    let hidden = lower(false);
    let shown = lower(true);
    assert_eq!(hidden[0]["runs"].as_array().unwrap().len(), 2, "{hidden}");
    assert_eq!(hidden[0]["runs"][0]["text"], "A");
    assert_eq!(hidden[0]["runs"][1]["text"], "Z");
    assert_eq!(hidden[0]["runs"][1]["pmStart"], 16.0);
    assert_eq!(hidden[0]["pmEnd"], shown[0]["pmEnd"]);
    assert_eq!(hidden[1], shown[1]);
    let shown_runs = shown[0]["runs"].as_array().unwrap();
    assert!(shown_runs.iter().any(|run| run["kind"] == "tab"));
    assert!(shown_runs.iter().any(|run| run["kind"] == "lineBreak"));
    assert!(shown_runs.iter().any(|run| run["text"] == "link"));
    assert!(shown_runs.iter().all(|run| run["hidden"] != true));
    assert_eq!(hidden, lower(false));
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);

    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![
                docx_edit::RawOp::Delete { index: 15, len: 1 },
                docx_edit::RawOp::Insert {
                    index: 15,
                    text: "Y".to_owned(),
                    attrs: Default::default(),
                },
            ],
            &docx_edit::EditCtx::local("", ""),
        )
        .unwrap();
    assert_eq!(lower(false)[0]["runs"][1]["text"], "Y");
    assert_eq!(lower(true)[0]["pmEnd"], shown[0]["pmEnd"]);
}

#[test]
fn only_hidden_paragraph_marks_remove_hidden_paragraph_spacing() {
    let styles = r#"<w:style w:type="paragraph" w:styleId="Instructions"><w:rPr><w:vanish/></w:rPr></w:style>"#;
    let body = r#"<w:p><w:pPr><w:pStyle w:val="Instructions"/><w:spacing w:before="480" w:after="480"/><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>Hidden instructions</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr></w:p><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden content, visible mark</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74210);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &docx_edit::bridge::RenderEnv::default())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(blocks.as_array().unwrap().len(), 3);
    assert_eq!(blocks[0]["kind"], "sectionBreak");
    assert_eq!(blocks[1]["kind"], "paragraph");
    assert_eq!(blocks[1]["runs"], json!([]));
    assert_eq!(blocks[2]["runs"][0]["text"], "After");
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn revealed_hidden_text_is_measured_instead_of_painted_at_zero_width() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(74211);
    let text = "Instructions wrap as ordinary text when explicitly revealed. ".repeat(12);
    let body = format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>{text}</w:t></w:r><w:r><w:t>Z</w:t></w:r></w:p>"#,
    );
    seed_from_docx(engine.doc(), &document(&body, "")).unwrap();
    let measure = |show_hidden_text| -> Value {
        serde_json::from_str(
            &engine
                .layout_document_with_regions_json(
                    &json!({
                        "bodyStory": "body", "options": {},
                        "renderEnv": {"showHiddenText": show_hidden_text},
                        "measurement": {
                            "fontChains": {"calibri|0|0": [font]},
                            "defaults": {"fontFamily": "Calibri", "fontSize": 10}
                        }
                    })
                    .to_string(),
                )
                .unwrap(),
        )
        .unwrap()
    };
    let hidden = measure(false);
    let shown = measure(true);
    let hidden_lines = hidden["measured"][0]["measure"]["lines"]
        .as_array()
        .unwrap();
    let shown_lines = shown["measured"][0]["measure"]["lines"].as_array().unwrap();
    assert_eq!(hidden_lines.len(), 1);
    assert!(shown_lines.len() > 1);
    assert!(hidden_lines[0]["width"].as_f64().unwrap() < 30.0);
    assert!(shown_lines[0]["width"].as_f64().unwrap() > 100.0);
    let display = engine.build_display_list_json(&hidden.to_string()).unwrap();
    assert!(!display.contains("Instructions"));
    let target = engine
        .with_display_list(|list| {
            list.pages[0].primitives.iter().find_map(|primitive| {
                let docx_layout::display_list::Primitive::Text(run) = primitive else {
                    return None;
                };
                (run.text == "Z").then(|| {
                    (
                        run.x.as_f64().unwrap() + run.width.as_f64().unwrap() / 2.0,
                        run.baseline_y.as_f64().unwrap(),
                        run.attrs.doc_start.unwrap(),
                        run.attrs.doc_end.unwrap(),
                    )
                })
            })
        })
        .flatten()
        .unwrap();
    assert_eq!(target.2, 2 + text.encode_utf16().count() as i64);
    let hit: Value = serde_json::from_str(
        &engine
            .display_hit_test_regions_json(0, target.0, target.1)
            .unwrap(),
    )
    .unwrap();
    assert_eq!(hit["region"], "body");
    assert!((target.2..=target.3).contains(&hit["pos"].as_i64().unwrap()));
}

#[test]
fn hidden_drawings_and_fields_follow_run_visibility() {
    let shape = SHAPE.replacen("<w:r>", "<w:r><w:rPr><w:vanish/></w:rPr>", 1);
    let body = format!(
        r#"<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>{shape}</w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"#,
    );
    let engine = EngineSession::new(74212);
    seed_from_docx(engine.doc(), &document(&body, "")).unwrap();
    let lower = |show_hidden_text| -> Value {
        serde_json::from_str(
            &engine
                .lower_story_json(
                    "body",
                    &docx_edit::bridge::RenderEnv {
                        show_hidden_text,
                        ..Default::default()
                    },
                )
                .unwrap(),
        )
        .unwrap()
    };
    let hidden = lower(false);
    let shown = lower(true);
    assert_eq!(hidden.as_array().unwrap().len(), 1);
    assert_eq!(hidden[0]["runs"][0]["text"], "After");
    assert_eq!(shown[0]["kind"], "shape");
    assert_eq!(shown[1]["runs"][0]["kind"], "field");
    assert_eq!(hidden[0], shown[2]);
}

#[test]
fn hidden_inline_content_control_breaks_follow_run_visibility() {
    let body = r#"<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t><w:br/></w:r></w:sdtContent></w:sdt></w:p><w:p><w:r><w:t>Visible</w:t><w:br/><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74213);
    seed_from_docx(engine.doc(), &document(body, "")).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let lower = |show_hidden_text| -> Value {
        serde_json::from_str(
            &engine
                .lower_story_json(
                    "body",
                    &docx_edit::bridge::RenderEnv {
                        show_hidden_text,
                        ..Default::default()
                    },
                )
                .unwrap(),
        )
        .unwrap()
    };
    let hidden = lower(false);
    let shown = lower(true);
    assert_eq!(hidden.as_array().unwrap().len(), 1);
    assert_eq!(shown.as_array().unwrap().len(), 2);
    assert_eq!(shown[0]["runs"][0]["text"], "Hidden");
    assert_eq!(shown[0]["runs"][1]["kind"], "lineBreak");
    assert_eq!(hidden[0]["runs"][1]["kind"], "lineBreak");
    assert_eq!(hidden[0], shown[1]);
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn line_unit_paragraph_spacing_honors_style_precedence_and_auto_spacing() {
    let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:before="50" w:after="50" w:beforeLines="50" w:afterLines="100"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Automatic"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr></w:style>"#;
    let body = r#"<w:p><w:r><w:t>Inherited</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="80" w:after="80" w:line="480" w:lineRule="auto"/></w:pPr><w:r><w:t>Direct twips</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="80" w:after="80" w:beforeLines="0" w:afterLines="0"/></w:pPr><w:r><w:t>Reset lines</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Automatic"/></w:pPr><w:r><w:t>Automatic</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:beforeLines="25" w:afterLines="25" w:line="600" w:lineRule="exact"/></w:pPr></w:p>
<w:p><w:pPr><w:pStyle w:val="Automatic"/><w:spacing w:beforeAutospacing="0" w:afterAutospacing="0"/></w:pPr><w:r><w:t>Automatic disabled</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74230);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &Default::default())
            .unwrap(),
    )
    .unwrap();
    for (index, before, after) in [
        (0, 8.0, 16.0),
        (1, 8.0, 16.0),
        (2, 80.0 / 15.0, 80.0 / 15.0),
        (3, 14.0, 14.0),
        (4, 4.0, 4.0),
        (5, 8.0, 16.0),
    ] {
        assert_eq!(blocks[index]["attrs"]["spacing"]["before"], before);
        assert_eq!(blocks[index]["attrs"]["spacing"]["after"], after);
    }
    assert_eq!(
        blocks[4]["attrs"]["spacingExplicit"],
        json!({"before":true,"after":true})
    );
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn line_unit_paragraph_spacing_uses_each_sections_grid_pitch() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let p = r#"<w:p><w:pPr><w:spacing w:beforeLines="100" w:afterLines="50" w:line="600" w:lineRule="exact"/></w:pPr><w:r><w:t>Grid spacing</w:t></w:r></w:p>"#;
    let boundary = r#"<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>"#;
    let table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc>{p}</w:tc></w:tr></w:tbl>"#
    );
    let body = format!("{p}{table}{boundary}{p}{boundary}{p}");
    let engine = EngineSession::new(74231);
    seed_from_docx(engine.doc(), &document(&body, "")).unwrap();
    let before = engine.doc().encode_state_as_update_v1();
    let output: Value = serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory":"body", "renderEnv":{}, "options":{},
        "regions":{"sections":[
            {"properties":{"docGrid":{"type":"linesAndChars","linePitch":326}}},
            {"properties":{"docGrid":{"type":"lines","linePitch":480}}},
            {"properties":{"docGrid":{"linePitch":720}}}
        ]},
        "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}
    }).to_string()).unwrap()).unwrap();
    let measured = output["measured"].as_array().unwrap();
    let paragraphs = measured
        .iter()
        .filter(|m| m["block"]["runs"][0]["text"] == "Grid spacing")
        .collect::<Vec<_>>();
    assert_eq!(paragraphs.len(), 3);
    for (paragraph, line) in paragraphs.into_iter().zip([326.0 / 15.0, 32.0, 16.0]) {
        assert_eq!(paragraph["block"]["attrs"]["spacing"]["before"], line);
        assert_eq!(paragraph["block"]["attrs"]["spacing"]["after"], line / 2.0);
        assert!(
            (paragraph["measure"]["totalHeight"].as_f64().unwrap() - (40.0 + line * 1.5)).abs()
                < 0.001
        );
    }
    let table = measured
        .iter()
        .find(|m| m["block"]["kind"] == "table")
        .unwrap();
    assert_eq!(
        table["block"]["rows"][0]["cells"][0]["blocks"][0]["attrs"]["spacing"]["before"],
        326.0 / 15.0
    );
    assert_eq!(engine.doc().encode_state_as_update_v1(), before);
}

#[test]
fn line_unit_paragraph_spacing_reuses_clean_blocks_after_editing() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:beforeLines="50" w:afterLines="50"/></w:pPr></w:style>"#;
    let body = r#"<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
<w:p><w:r><w:t>Charlie</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74232);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let request = json!({"bodyStory":"body", "renderEnv":{}, "options":{},
        "regions":{"sections":[{"properties":{"docGrid":{"type":"lines","linePitch":326}}}]},
        "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}}).to_string();
    let extras = json!({"fontChains":{"calibri|0|0":[font]}}).to_string();
    engine.layout_document_with_regions_json(&request).unwrap();
    engine.build_display_list_frame(&extras, 0).unwrap();
    engine
        .doc()
        .insert_text(
            &docx_edit::EditCtx::local("", ""),
            docx_edit::Position::new("body", 9),
            "xx",
            docx_edit::FormatPolicy::Inherit,
        )
        .unwrap();
    let before = engine.stats();
    engine.apply_and_layout("body", 1).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.resident_measure_calls - before.resident_measure_calls,
        1
    );
    assert_eq!(
        after.resident_reused_blocks - before.resident_reused_blocks,
        3
    );
    let fast = engine.retained_kernel_inputs_json().unwrap();
    let fast_display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    engine.layout_document_with_regions_json(&request).unwrap();
    engine.build_display_list_frame(&extras, 2).unwrap();
    assert_eq!(engine.retained_kernel_inputs_json().unwrap(), fast);
    assert_eq!(
        engine
            .with_display_list(|display| serde_json::to_value(display).unwrap())
            .unwrap(),
        fast_display
    );
}

#[test]
fn line_unit_paragraph_spacing_yields_to_twip_edits_and_round_trips() {
    use docx_edit::{EditCtx, ParaAttrDelta, ParaSelector, Patch};
    let body = r#"<w:p><w:pPr><w:spacing w:before="50" w:after="50" w:beforeLines="100" w:afterLines="50" w:beforeAutospacing="1"/></w:pPr><w:r><w:t>Edited spacing</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(74233);
    seed_from_docx(engine.doc(), &document(body, "")).unwrap();
    let id = engine.doc().paragraphs("body").unwrap()[0].para_id.clone();
    engine
        .doc()
        .set_paragraph_attrs(
            &EditCtx::local("", ""),
            &ParaSelector::One(id),
            &ParaAttrDelta {
                space_before: Patch::Set(0.0),
                space_after: Patch::Set(120.0),
                ..Default::default()
            },
        )
        .unwrap();
    let blocks: Value = serde_json::from_str(
        &engine
            .lower_story_json("body", &Default::default())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(blocks[0]["attrs"]["spacing"]["before"], 0.0);
    assert_eq!(blocks[0]["attrs"]["spacing"]["after"], 8.0);
    let snapshot = engine.doc().paragraphs("body").unwrap().remove(0);
    let original =
        serde_json::to_value(snapshot.properties.get("_originalFormatting").unwrap()).unwrap();
    let formatting = serde_json::from_value(original).unwrap();
    let xml = docx_parse::serializer::serialize_paragraph_formatting(
        Some(&formatting),
        None,
        None,
        None,
        false,
        None,
    )
    .unwrap();
    let reopened = EngineSession::new(74234);
    seed_from_docx(
        reopened.doc(),
        &document(
            &format!("<w:p>{xml}<w:r><w:t>Edited spacing</w:t></w:r></w:p>"),
            "",
        ),
    )
    .unwrap();
    let blocks: Value = serde_json::from_str(
        &reopened
            .lower_story_json("body", &Default::default())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(blocks[0]["attrs"]["spacing"]["before"], 0.0);
    assert_eq!(blocks[0]["attrs"]["spacing"]["after"], 8.0);
}

#[test]
fn line_unit_paragraph_spacing_style_changes_round_trip() {
    use docx_edit::{EditCtx, ParaSelector, ResolvedStyleProjection};
    let body = r#"<w:p><w:pPr><w:spacing w:before="80" w:after="80" w:beforeLines="100" w:afterLines="50" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr><w:r><w:t>Styled spacing</w:t></w:r></w:p>"#;
    let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:style>"#;
    let engine = EngineSession::new(74235);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let id = engine.doc().paragraphs("body").unwrap()[0].para_id.clone();
    for (attrs, expected_before, expected_after) in [
        (
            json!({"spaceBefore":0,"spaceAfter":120,"spaceAfterLines":75,"afterAutospacing":false}),
            0.0,
            12.0,
        ),
        (json!({}), 0.0, 0.0),
    ] {
        engine
            .doc()
            .apply_paragraph_style(
                &EditCtx::local("", ""),
                &ParaSelector::One(id.clone()),
                &ResolvedStyleProjection {
                    style_id: "Body".into(),
                    known: true,
                    paragraph_attrs: serde_json::from_value(attrs.clone()).unwrap(),
                    ..Default::default()
                },
            )
            .unwrap();
        let blocks: Value = serde_json::from_str(
            &engine
                .lower_story_json("body", &Default::default())
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            blocks[0]["attrs"]["spacing"]["before"]
                .as_f64()
                .unwrap_or(0.0),
            expected_before
        );
        assert_eq!(
            blocks[0]["attrs"]["spacing"]["after"]
                .as_f64()
                .unwrap_or(0.0),
            expected_after
        );
        let snapshot = engine.doc().paragraphs("body").unwrap().remove(0);
        let original =
            serde_json::to_value(snapshot.properties.get("_originalFormatting").unwrap()).unwrap();
        for key in [
            "spaceBefore",
            "spaceAfter",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
        ] {
            assert_eq!(original.get(key), attrs.get(key));
        }
        let formatting = serde_json::from_value(original).unwrap();
        let xml = docx_parse::serializer::serialize_paragraph_formatting(
            Some(&formatting),
            None,
            None,
            None,
            false,
            None,
        )
        .unwrap();
        let reopened = EngineSession::new(74236);
        seed_from_docx(
            reopened.doc(),
            &document(
                &format!("<w:p>{xml}<w:r><w:t>Styled spacing</w:t></w:r></w:p>"),
                styles,
            ),
        )
        .unwrap();
        let blocks: Value = serde_json::from_str(
            &reopened
                .lower_story_json("body", &Default::default())
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            blocks[0]["attrs"]["spacing"]["before"]
                .as_f64()
                .unwrap_or(0.0),
            expected_before
        );
        assert_eq!(
            blocks[0]["attrs"]["spacing"]["after"]
                .as_f64()
                .unwrap_or(0.0),
            expected_after
        );
    }
}

#[test]
fn line_unit_paragraph_spacing_clears_and_round_trips() {
    use docx_edit::{EditCtx, ParaAttrDelta, ParaSelector, Patch};
    let body = r#"<w:p><w:pPr><w:spacing w:before="80" w:after="80" w:beforeLines="100" w:afterLines="50" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr><w:r><w:t>Cleared spacing</w:t></w:r></w:p>"#;
    let styles = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:style>"#;
    let engine = EngineSession::new(74237);
    seed_from_docx(engine.doc(), &document(body, styles)).unwrap();
    let id = engine.doc().paragraphs("body").unwrap()[0].para_id.clone();
    engine
        .doc()
        .set_paragraph_attrs(
            &EditCtx::local("", ""),
            &ParaSelector::One(id),
            &ParaAttrDelta {
                space_before: Patch::Clear,
                space_after: Patch::Clear,
                ..Default::default()
            },
        )
        .unwrap();
    let snapshot = engine.doc().paragraphs("body").unwrap().remove(0);
    let original =
        serde_json::to_value(snapshot.properties.get("_originalFormatting").unwrap()).unwrap();
    for key in [
        "spaceBefore",
        "spaceAfter",
        "spaceBeforeLines",
        "spaceAfterLines",
        "beforeAutospacing",
        "afterAutospacing",
    ] {
        assert!(!snapshot.properties.contains_key(key));
        assert!(original.get(key).is_none());
    }
    let formatting = serde_json::from_value(original).unwrap();
    let xml = docx_parse::serializer::serialize_paragraph_formatting(
        Some(&formatting),
        None,
        None,
        None,
        false,
        None,
    )
    .unwrap();
    let reopened = EngineSession::new(74238);
    seed_from_docx(
        reopened.doc(),
        &document(
            &format!("<w:p>{xml}<w:r><w:t>Cleared spacing</w:t></w:r></w:p>"),
            styles,
        ),
    )
    .unwrap();
    for session in [&engine, &reopened] {
        let blocks: Value = serde_json::from_str(
            &session
                .lower_story_json("body", &Default::default())
                .unwrap(),
        )
        .unwrap();
        for side in ["before", "after"] {
            assert_eq!(
                blocks[0]["attrs"]["spacing"][side].as_f64().unwrap_or(0.0),
                0.0
            );
        }
    }
}
