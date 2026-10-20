#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position};
use serde_json::{Value, json};

fn section(kind: &str) -> String {
    format!(
        r#"<w:sectPr>{kind}<w:pgSz w:w="7200" w:h="4320"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#
    )
}

fn render_body(body: &str) -> Value {
    let bytes = fixture::with_body(body);
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
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
    let engine = EngineSession::new(9410);
    docx_edit::seed_from_docx(engine.doc(), &bytes).unwrap();
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(
                &json!({
                    "bodyStory": "body", "renderEnv": {},
                    "regions": {"sections": sections, "settings": package.settings},
                    "measurement": {
                        "fontChains": {"arial|0|0": [font]},
                        "defaults": {"fontFamily": "Arial", "fontSize": 12}
                    }
                })
                .to_string(),
            )
            .unwrap(),
    )
    .unwrap()
}

fn text_page(output: &Value, text: &str) -> u64 {
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
    fragment_page(output, &block["block"]["id"]).0
}

fn fragment_page<'a>(output: &'a Value, block_id: &Value) -> (u64, &'a Value) {
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .find_map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .find(|fragment| fragment["blockId"] == *block_id)
                .map(|fragment| (page["number"].as_u64().unwrap(), fragment))
        })
        .unwrap()
}

fn assert_section_mark(output: &Value, page: u64) {
    let mark = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["kind"] == "paragraph"
                && measured["measure"]["lines"]
                    .as_array()
                    .is_some_and(Vec::is_empty)
        })
        .unwrap();
    let (number, fragment) = fragment_page(output, &mark["block"]["id"]);
    assert_eq!(number, page);
    assert_eq!(fragment["height"], 0.0);
    assert_eq!(fragment["pmStart"], mark["block"]["pmStart"]);
    assert_eq!(fragment["pmEnd"], mark["block"]["pmEnd"]);
    assert!(fragment["pmStart"].is_number());
    assert_eq!(fragment["fromLine"], 0);
    assert_eq!(fragment["toLine"], 0);
}

fn assert_two_pages(output: &Value) {
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(output, "A"), 1);
    assert_eq!(text_page(output, "B"), 2);
    assert_eq!(output["layout"]["pages"][1]["sectionIndex"], 1);
}

#[test]
fn bare_section_mark_after_break_only_paragraph_reuses_the_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 2);
}

#[test]
fn bare_section_mark_after_trailing_page_break_reuses_the_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 2);
}

#[test]
fn page_break_before_section_mark_keeps_a_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_eq!(output["layout"]["pages"][2]["sectionIndex"], 1);
    assert_section_mark(&output, 2);
}

#[test]
fn bare_section_mark_does_not_flush_spacing_past_the_body_bottom() {
    let geometry = section("");
    for (line, after) in [(2880, 0), (2640, 1080)] {
        let output = render_body(&format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="exact"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="240" w:after="240"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
        ));
        assert_two_pages(&output);
        assert_section_mark(&output, 1);
    }
}

#[test]
fn bare_section_mark_after_a_full_page_table_does_not_add_a_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2880" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:before="240"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| measured["block"]["kind"] == "table")
        .unwrap();
    assert_eq!(fragment_page(&output, &table["block"]["id"]).0, 2);
    assert_section_mark(&output, 2);
}

#[test]
fn keep_next_does_not_reserve_spacing_for_a_bare_section_mark() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>PREFIX</w:t></w:r></w:p><w:p><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="2400" w:lineRule="exact"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="480"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_two_pages(&output);
    assert_section_mark(&output, 1);
}

#[test]
fn authored_empty_paragraph_after_manual_break_keeps_its_line() {
    for separate_break in [false, true] {
        let first = if separate_break {
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>"#
        } else {
            r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p>"#
        };
        let output = render_body(&format!(
            r#"{first}<w:p/><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section("")
        ));
        assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), 2);
        let measured = output["measured"].as_array().unwrap();
        let empty = measured
            .iter()
            .rev()
            .find(|block| {
                block["block"]["kind"] == "paragraph"
                    && block["block"]["runs"].as_array().is_some_and(Vec::is_empty)
            })
            .unwrap();
        let (page, fragment) = fragment_page(&output, &empty["block"]["id"]);
        assert_eq!(page, 2);
        assert!(fragment["height"].as_f64().unwrap() > 0.0);
    }
}

#[test]
fn authored_empty_paragraphs_before_section_mark_keep_a_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p/><w:p/><w:p/><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    let empties: Vec<_> = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|measured| {
            measured["block"]["kind"] == "paragraph"
                && measured["block"]["runs"]
                    .as_array()
                    .is_some_and(Vec::is_empty)
                && measured["measure"]["lines"]
                    .as_array()
                    .is_some_and(|lines| !lines.is_empty())
        })
        .collect();
    assert_eq!(empties.len(), 3);
    for empty in empties {
        let (page, fragment) = fragment_page(&output, &empty["block"]["id"]);
        assert_eq!(page, 2);
        assert!(fragment["height"].as_f64().unwrap() > 0.0);
    }
    assert_section_mark(&output, 2);
}

#[test]
fn odd_page_section_after_bare_mark_keeps_the_parity_filler() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(""),
        section(r#"<w:type w:val="oddPage"/>"#)
    ));
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_eq!(pages[1]["parityFiller"], true);
    assert_section_mark(&output, 1);
}

#[test]
fn manual_break_before_bare_mark_leaves_parity_fillers_blank() {
    for (kind, preceding, expected_page) in [
        ("oddPage", "", 3),
        (
            "evenPage",
            r#"<w:p><w:r><w:t>C</w:t><w:br w:type="page"/></w:r></w:p>"#,
            4,
        ),
    ] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p>{preceding}<w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section(""),
            section(&format!(r#"<w:type w:val="{kind}"/>"#))
        ));
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), expected_page as usize);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), expected_page);
        let filler = &pages[expected_page as usize - 2];
        assert_eq!(filler["parityFiller"], true);
        assert!(
            filler["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .all(|fragment| fragment["height"] == 0.0)
        );
        assert_section_mark(&output, expected_page - 1);
    }
}

#[test]
fn authored_empty_paragraph_between_manual_breaks_keeps_a_blank_page() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p/><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
}

#[test]
fn section_mark_before_a_continuous_section_keeps_spacing() {
    for resized in [false, true] {
        let mut next = section(r#"<w:type w:val="continuous"/>"#);
        if resized {
            next = next.replace(r#"w:w="7200""#, r#"w:w="8640""#);
        }
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="4320" w:after="4320"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{next}"#,
            section("")
        ));
        assert!(output["layout"]["pages"].as_array().unwrap().len() >= 2);
        assert_eq!(text_page(&output, "A"), 1);
        assert!(text_page(&output, "B") >= 2);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn section_mark_spacing_participates_in_column_balancing() {
    let columns = r#"<w:cols w:num="2" w:space="360"/>"#;
    let mut positions = Vec::new();
    for spacing in [0, 4320] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="{spacing}" w:after="{spacing}"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
            section(columns),
            section(&format!(r#"<w:type w:val="continuous"/>{columns}"#))
        ));
        if spacing == 0 {
            assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
            assert_eq!(text_page(&output, "B"), 1);
            assert_section_mark(&output, 1);
        }
        positions.push(
            output["layout"]["pages"][0]["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .map(|fragment| (fragment["x"].clone(), fragment["y"].clone()))
                .collect::<Vec<_>>(),
        );
    }
    assert_ne!(positions[0], positions[1]);
}

#[test]
fn anchored_cover_section_mark_keeps_before_spacing() {
    let geometry = r#"<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="500" w:right="500" w:bottom="500" w:left="500"/></w:sectPr>"#;
    let output = render_body(&format!(
        r#"<w:p/><w:p/><w:p><w:pPr><w:spacing w:before="13000" w:after="200"/>{geometry}</w:pPr>{}</w:p><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1800" w:lineRule="exact"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#,
        anchored_text_box()
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(&output, "B"), 2);
    let anchor = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|mb| {
            mb["block"]["attrs"]["hasSourceContent"] == true
                && mb["block"]["runs"].as_array().is_some_and(Vec::is_empty)
        })
        .unwrap();
    assert!(
        (anchor["block"]["attrs"]["spacing"]["before"]
            .as_f64()
            .unwrap()
            - 13000.0 / 15.0)
            .abs()
            < 1e-6
    );
    assert_eq!(fragment_page(&output, &anchor["block"]["id"]).0, 1);
    assert!(output["measured"].as_array().unwrap().iter().any(|mb| {
        mb["block"]["kind"] == "shape"
            && mb["block"]["innerText"]
                .as_array()
                .is_some_and(|text| !text.is_empty())
    }));
}

fn anchored_text_box() -> &'static str {
    concat!(
        r#"<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">"#,
        r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>"#,
        r#"<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="Cover"/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>"#,
        r#"<wps:txbx><w:txbxContent><w:p><w:r><w:t>COVER</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
    )
}

#[test]
fn anchored_section_mark_before_next_page_keeps_spacing() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="13000"/>{geometry}</w:pPr>{}</w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#,
        anchored_text_box()
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
}

#[test]
fn two_column_continuous_section_mark_after_page_break_keeps_blank_page() {
    let ending = section(r#"<w:type w:val="continuous"/><w:cols w:num="2" w:space="360"/>"#);
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{ending}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "A"), 1);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn section_mark_after_whitespace_and_page_break_keeps_blank_page() {
    let geometry = section("");
    for whitespace in [
        "<w:tab/>".repeat(7),
        r#"<w:t xml:space="preserve">   </w:t>"#.to_owned(),
    ] {
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r>{whitespace}<w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
        ));
        assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
        assert_eq!(text_page(&output, "A"), 1);
        assert_eq!(text_page(&output, "B"), 3);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn section_mark_after_page_break_with_trailing_space_keeps_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn continuous_section_mark_after_full_page_table_keeps_blank_page() {
    let output = render_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2880" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:before="240"/>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(r#"<w:type w:val="continuous"/>"#),
        section("")
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn one_column_continuous_section_mark_after_terminal_break_reuses_page() {
    for resized in [false, true] {
        let incoming = if resized {
            section("").replace(r#"w:w="7200""#, r#"w:w="8640""#)
        } else {
            section("")
        };
        let output = render_body(&format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/>{}</w:pPr><w:bookmarkStart w:id="0" w:name="boundary"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{incoming}"#,
            section(r#"<w:type w:val="continuous"/>"#),
        ));
        assert_two_pages(&output);
        assert_section_mark(&output, 2);
    }
}

#[test]
fn framed_section_mark_after_page_break_keeps_blank_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:framePr w:w="1440" w:hAnchor="margin" w:vAnchor="text"/>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn section_mark_entering_two_columns_after_page_break_keeps_blank_page() {
    let output = render_body(&format!(
        r#"<w:p><w:r><w:t>A</w:t><w:br w:type="page"/></w:r></w:p><w:p><w:pPr>{}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{}"#,
        section(""),
        section(r#"<w:cols w:num="2" w:space="360"/>"#)
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 3);
    assert_eq!(text_page(&output, "B"), 3);
    assert_section_mark(&output, 2);
}

#[test]
fn genuinely_empty_section_keeps_its_line_and_page() {
    let geometry = section("");
    let output = render_body(&format!(
        r#"<w:p><w:pPr>{geometry}</w:pPr></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>{geometry}"#
    ));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 2);
    assert_eq!(text_page(&output, "B"), 2);
    assert!(
        output["layout"]["pages"][0]["fragments"][0]["height"]
            .as_f64()
            .unwrap()
            > 0.0
    );
}

const UNEQUAL_COLUMNS: &str = r#"<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3256" w:space="2870"/><w:col w:w="3954"/></w:cols>"#;

fn layout_page_body(width: u32, height: u32, columns: &str, body: &str) -> Value {
    let body = format!(
        r#"{body}<w:sectPr><w:pgSz w:w="{width}" w:h="{height}"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>{columns}</w:sectPr>"#,
    );
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9403);
    serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap()
}

fn column_layout(columns: &str, height: u32, first: &str, second: &str) -> Value {
    let body = format!(
        "{}{}",
        fixture::p(
            "71000001",
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0"/><w:widowControl w:val="0"/></w:pPr>{}"#,
                fixture::r(first)
            ),
        ),
        fixture::p(
            "71000002",
            &format!(
                r#"<w:pPr><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="column"/></w:r>{}"#,
                fixture::r(second)
            ),
        ),
    );
    layout_page_body(12240, height, columns, &body)
}

fn assert_column_paragraphs(
    output: &Value,
    widths: [f64; 2],
    second_x: f64,
    fragment_widths: [f64; 2],
) {
    let paragraphs: Vec<_> = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|block| {
            block["block"]["kind"] == "paragraph"
                && block["block"]["runs"].as_array().is_some_and(|runs| {
                    runs.iter()
                        .any(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
                })
        })
        .collect();
    let first_id = &paragraphs[0]["block"]["id"];
    let second_id = &paragraphs[1]["block"]["id"];
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "paragraph")
        .collect();
    for (id, x, width) in [
        (first_id, 72.0, fragment_widths[0]),
        (second_id, second_x, fragment_widths[1]),
    ] {
        let fragments: Vec<_> = fragments
            .iter()
            .filter(|fragment| &fragment["blockId"] == id)
            .collect();
        assert!(!fragments.is_empty());
        for fragment in fragments {
            assert!((fragment["x"].as_f64().unwrap() - x).abs() < 0.01);
            assert!((fragment["width"].as_f64().unwrap() - width).abs() < 0.01);
        }
    }
    let first_lines = paragraphs[0]["measure"]["lines"].as_array().unwrap();
    assert!(first_lines.len() > 1);
    assert!(
        first_lines
            .iter()
            .all(|line| visible_line_width(paragraphs[0], line) <= widths[0] + 0.01)
    );

    let right_edges = painted_line_right_edges(output, "71000002");
    assert!(!right_edges.is_empty());
    let second_lines = paragraphs[1]["measure"]["lines"].as_array().unwrap();
    assert_eq!(right_edges.len(), second_lines.len());
    for (index, right) in &right_edges {
        let line = &second_lines[*index as usize];
        assert!(visible_line_width(paragraphs[1], line) <= widths[1] + 0.01);
        let expected_right = second_x + fragment_widths[1].max(line["width"].as_f64().unwrap());
        assert!((right - expected_right).abs() < 0.1, "line ends at {right}");
    }
}

fn visible_line_width(measured: &Value, line: &Value) -> f64 {
    let trailing_width: f64 = line["clusterAdvances"]
        .as_array()
        .unwrap()
        .iter()
        .rev()
        .take_while(|cluster| {
            let run = cluster["runIndex"].as_u64().unwrap() as usize;
            let text = measured["block"]["runs"][run]["text"].as_str().unwrap();
            let start = cluster["startChar"].as_u64().unwrap() as usize;
            let end = cluster["endChar"].as_u64().unwrap() as usize;
            start < end
                && text
                    .encode_utf16()
                    .skip(start)
                    .take(end - start)
                    .all(|unit| unit == u16::from(b' '))
        })
        .map(|cluster| cluster["advance"].as_f64().unwrap())
        .sum();
    line["width"].as_f64().unwrap() - trailing_width
}

fn painted_line_right_edges(output: &Value, para_id: &str) -> std::collections::BTreeMap<u64, f64> {
    let display = docx_layout::build_display_list_value(&output.to_string()).unwrap();
    let display = serde_json::to_value(display).unwrap();
    let mut right_edges = std::collections::BTreeMap::<u64, f64>::new();
    for primitive in display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().unwrap())
    {
        if primitive["paraId"] != para_id {
            continue;
        }
        let right = match primitive["kind"].as_str() {
            Some("glyphRun") => primitive["glyphs"]
                .as_array()
                .unwrap()
                .iter()
                .map(|glyph| glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap())
                .fold(f64::NEG_INFINITY, f64::max),
            Some("text") => primitive["x"].as_f64().unwrap() + primitive["width"].as_f64().unwrap(),
            _ => continue,
        };
        let line = primitive["lineIndex"].as_u64().unwrap();
        right_edges
            .entry(line)
            .and_modify(|edge| *edge = edge.max(right))
            .or_insert(right);
    }
    right_edges
}

#[test]
fn unequal_columns_measure_and_paint_inside_the_active_column() {
    let output = column_layout(
        UNEQUAL_COLUMNS,
        15840,
        &"Column text wraps within its own width. ".repeat(6),
        &"Upper right title with several words. ".repeat(5),
    );
    assert_column_paragraphs(
        &output,
        [3256.0 / 15.0, 3954.0 / 15.0],
        72.0 + 6126.0 / 15.0,
        [3256.0 / 15.0, 3954.0 / 15.0],
    );

    let control = layout_page_body(
        6114,
        15840,
        "",
        &fixture::p(
            "71000002",
            &fixture::r(&"Upper right title with several words. ".repeat(5)),
        ),
    );
    let title = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["paraId"] == "71000002"
                && measured["block"]["runs"].as_array().is_some_and(|runs| {
                    runs.iter()
                        .any(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
                })
        })
        .unwrap();
    assert_eq!(line_ranges(title), line_ranges(&control["measured"][0]));
}

fn line_ranges(measured: &Value) -> Vec<(u64, u64, u64, u64)> {
    measured["measure"]["lines"]
        .as_array()
        .unwrap()
        .iter()
        .map(|line| {
            (
                line["headRun"].as_u64().unwrap(),
                line["headChar"].as_u64().unwrap(),
                line["tailRun"].as_u64().unwrap(),
                line["tailChar"].as_u64().unwrap(),
            )
        })
        .collect()
}

#[test]
fn equal_columns_keep_their_widths_and_origins() {
    let output = column_layout(
        r#"<w:cols w:num="2" w:space="720"/>"#,
        15840,
        &"Column text wraps within its own width. ".repeat(6),
        "Upper right title",
    );
    assert_column_paragraphs(&output, [312.0, 312.0], 432.0, [672.0, 672.0]);
}

#[test]
fn a_paragraph_rewraps_when_it_flows_between_unequal_columns() {
    let body = fixture::p(
        "71000003",
        &format!(
            r#"<w:pPr><w:jc w:val="right"/><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"Words continue through columns with different widths. ".repeat(20))
        ),
    );
    let output = layout_page_body(12240, 3360, UNEQUAL_COLUMNS, &body);
    assert_spanning_columns(&output);
}

fn assert_spanning_columns(output: &Value) {
    let lines = output["measured"][0]["measure"]["lines"]
        .as_array()
        .unwrap();
    let right_edges = painted_line_right_edges(output, "71000003");
    assert_eq!(right_edges.len(), lines.len());
    let mut next_line = 0;
    let mut columns_seen = [false; 2];
    for fragment in output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "paragraph")
    {
        let x = fragment["x"].as_f64().unwrap();
        let column = usize::from(x > 73.0);
        columns_seen[column] = true;
        let width = [3256.0 / 15.0, 3954.0 / 15.0][column];
        assert!((fragment["width"].as_f64().unwrap() - width).abs() < 0.01);
        let from = fragment["fromLine"].as_u64().unwrap() as usize;
        let to = fragment["toLine"].as_u64().unwrap() as usize;
        assert_eq!(from, next_line);
        next_line = to;
        for (index, line) in lines.iter().enumerate().take(to).skip(from) {
            assert!(visible_line_width(&output["measured"][0], line) <= width + 0.01);
            assert!(line["rightOffset"].as_f64().unwrap_or(0.0).abs() < 0.01);
            let expected_right = x + width.max(line["width"].as_f64().unwrap());
            assert!((right_edges[&(index as u64)] - expected_right).abs() < 0.1);
        }
    }
    assert_eq!(next_line, lines.len());
    assert_eq!(columns_seen, [true, true]);

    let second_column_lines: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["x"].as_f64().unwrap() > 73.0)
        .flat_map(|fragment| {
            let from = fragment["fromLine"].as_u64().unwrap() as usize;
            let to = fragment["toLine"].as_u64().unwrap() as usize;
            &lines[from..to]
        })
        .collect();
    assert!(
        second_column_lines
            .iter()
            .any(|line| line["width"].as_f64().unwrap() > 3256.0 / 15.0 + 1.0)
    );
}

#[test]
fn a_paragraph_rewraps_when_it_flows_into_a_narrower_final_column() {
    let body = fixture::p(
        "71000003",
        &format!(
            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"Words continue through columns with different widths. ".repeat(12))
        ),
    );
    let output = layout_page_body(
        12240,
        4560,
        r#"<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3954" w:space="2870"/><w:col w:w="3256"/></w:cols>"#,
        &body,
    );
    let measured = &output["measured"][0];
    let lines = measured["measure"]["lines"].as_array().unwrap();
    assert!(lines.iter().all(|line| line["syntheticFallback"] != true));
    let mut columns_seen = [false; 2];
    for fragment in output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "paragraph")
    {
        let column = usize::from(fragment["x"].as_f64().unwrap() > 73.0);
        columns_seen[column] = true;
        let width = [3954.0 / 15.0, 3256.0 / 15.0][column];
        let from = fragment["fromLine"].as_u64().unwrap() as usize;
        let to = fragment["toLine"].as_u64().unwrap() as usize;
        for line in &lines[from..to] {
            assert!(visible_line_width(measured, line) <= width + 0.01);
        }
    }
    assert_eq!(columns_seen, [true, true]);
    let last = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .rfind(|fragment| fragment["kind"] == "paragraph")
        .unwrap();
    assert!(last["x"].as_f64().unwrap() > 73.0);
}

#[test]
fn balancing_rewraps_a_paragraph_in_each_unequal_column() {
    let body = fixture::p(
        "71000003",
        &format!(
            r#"<w:pPr><w:jc w:val="right"/><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"Words continue through columns with different widths. ".repeat(6))
        ),
    );
    let output = layout_page_body(12240, 15840, UNEQUAL_COLUMNS, &body);
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    assert_spanning_columns(&output);
}

#[test]
fn keep_lines_remeasures_a_paragraph_after_moving_it_to_column_two() {
    let title = fixture::p(
        "71000008",
        &format!(
            r#"<w:pPr><w:keepLines/><w:spacing w:before="0" w:after="0"/></w:pPr>{}"#,
            fixture::r(&"A paragraph kept together inside its column. ".repeat(3))
        ),
    );
    let body = format!(
        "{}{}",
        fixture::p(
            "71000009",
            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="720" w:lineRule="exact"/></w:pPr><w:r><w:t>One</w:t><w:br/><w:t>Two</w:t><w:br/><w:t>Three</w:t></w:r>"#,
        ),
        title,
    );
    let output = layout_page_body(12240, 4560, UNEQUAL_COLUMNS, &body);
    let control = layout_page_body(6114, 15840, "", &title);
    let measured = &output["measured"][1];
    assert_eq!(line_ranges(measured), line_ranges(&control["measured"][0]));
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["blockId"] == measured["block"]["id"])
        .collect();
    assert_eq!(fragments.len(), 1);
    assert!((fragments[0]["x"].as_f64().unwrap() - (72.0 + 6126.0 / 15.0)).abs() < 0.01);
    assert!((fragments[0]["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
}

#[test]
fn keep_lines_column_width_oscillation_uses_one_stable_measure() {
    let title = fixture::p(
        "7100000A",
        &format!(
            r#"<w:pPr><w:keepLines/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"A paragraph kept together inside its column. ".repeat(4)),
        ),
    );
    let narrow = layout_page_body(3960, 15840, "", &title);
    let wide = layout_page_body(10080, 15840, "", &title);
    let narrow_height = narrow["measured"][0]["measure"]["totalHeight"]
        .as_f64()
        .unwrap();
    let wide_height = wide["measured"][0]["measure"]["totalHeight"]
        .as_f64()
        .unwrap();
    let remainder = (narrow_height + wide_height) / 2.0;
    assert!(wide_height < remainder && remainder < narrow_height);
    let filler_height = narrow_height + 32.0;
    let body = format!(
        "{}{}{}",
        fixture::p(
            "7100000B",
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{}" w:lineRule="exact"/></w:pPr>{}"#,
                (filler_height * 15.0).round() as u32,
                fixture::r("Filler"),
            ),
        ),
        title,
        fixture::p("7100000C", r#"<w:r><w:br w:type="page"/></w:r>"#),
    );
    let columns = r#"<w:cols w:num="2" w:space="360" w:equalWidth="0"><w:col w:w="1800" w:space="360"/><w:col w:w="7920"/></w:cols>"#;
    let height = 2160 + ((filler_height + remainder) * 15.0).round() as u32;
    let output = layout_page_body(12240, height, columns, &body);
    let measured = &output["measured"][1];
    assert_eq!(line_ranges(measured), line_ranges(&narrow["measured"][0]));
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["blockId"] == measured["block"]["id"])
        .collect();
    assert_eq!(fragments.len(), 1);
    assert_eq!(fragments[0]["x"], 216.0);
    assert_eq!(fragments[0]["width"], 528.0);
    assert_eq!(fragments[0]["fromLine"], 0);
    assert_eq!(
        fragments[0]["toLine"].as_u64().unwrap() as usize,
        measured["measure"]["lines"].as_array().unwrap().len(),
    );
}

#[test]
fn single_column_ignores_an_unequal_column_definition() {
    let body = fixture::p(
        "71000004",
        &fixture::r(&"Ordinary paragraph text. ".repeat(8)),
    );
    let plain = layout_page_body(12240, 15840, "", &body);
    let authored = layout_page_body(
        12240,
        15840,
        r#"<w:cols w:num="1" w:equalWidth="0"><w:col w:w="3256" w:space="2870"/></w:cols>"#,
        &body,
    );
    assert_eq!(plain["measured"], authored["measured"]);
    assert_eq!(
        plain["layout"]["pages"][0]["fragments"],
        authored["layout"]["pages"][0]["fragments"]
    );
    assert_eq!(
        authored["layout"]["pages"][0]["fragments"][0]["width"],
        672.0
    );
}

#[test]
fn a_table_after_a_column_break_uses_the_second_column_width() {
    let body = format!(
        r#"{}{}<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3954"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        fixture::p("71000005", &fixture::r("First column")),
        fixture::p("71000006", r#"<w:r><w:br w:type="column"/></w:r>"#),
        fixture::p("71000007", &fixture::r("Table text")),
    );
    let output = layout_page_body(12240, 15840, UNEQUAL_COLUMNS, &body);
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|block| block["block"]["kind"] == "table")
        .unwrap();
    assert!((table["measure"]["totalWidth"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
    let fragment = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .find(|fragment| fragment["kind"] == "table")
        .unwrap();
    assert!((fragment["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
}

#[test]
fn table_fragments_stay_inside_unequal_columns() {
    let rows: String = (0..8)
        .map(|index| {
            format!(
                r#"<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:fill="FFFF00"/></w:tcPr>{}</w:tc></w:tr>"#,
                fixture::p(&format!("720000{index:02X}"), &fixture::r("Table row")),
            )
        })
        .collect();
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="7920"/></w:tblGrid>{rows}</w:tbl>{}"#,
        fixture::p("72000010", r#"<w:r><w:br w:type="page"/></w:r>"#),
    );
    let columns = r#"<w:cols w:num="2" w:space="360" w:equalWidth="0"><w:col w:w="7920" w:space="360"/><w:col w:w="1800"/></w:cols>"#;
    let output = layout_page_body(12240, 3360, columns, &body);
    let mut columns_seen = [false; 2];
    for fragment in output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "table")
    {
        let x = fragment["x"].as_f64().unwrap();
        let column = usize::from(x > 73.0);
        columns_seen[column] = true;
        let right = [600.0, 744.0][column];
        assert!(x + fragment["width"].as_f64().unwrap() <= right + 0.01);
    }
    assert_eq!(columns_seen, [true, true]);

    let display = docx_layout::build_display_list_value(&output.to_string()).unwrap();
    let display = serde_json::to_value(display).unwrap();
    let mut clipped_in_narrow_column = false;
    for primitive in display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().unwrap())
        .filter(|primitive| primitive["table"].is_object())
    {
        let clip = &primitive["clipGroup"]["clip"];
        let x = clip["x"].as_f64().unwrap();
        let right = if x > 601.0 { 744.0 } else { 600.0 };
        assert!(x + clip["w"].as_f64().unwrap() <= right + 0.01);
        clipped_in_narrow_column |= x > 601.0;
    }
    assert!(clipped_in_narrow_column);
}

#[test]
fn a_continuous_section_applies_changed_authored_column_widths() {
    let body = format!(
        "{}{}",
        fixture::p(
            "71000010",
            &format!(
                r#"<w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>{UNEQUAL_COLUMNS}</w:sectPr></w:pPr>{}"#,
                fixture::r("First region")
            ),
        ),
        fixture::p("71000011", &fixture::r("Second region")),
    );
    let output = layout_page_body(
        12240,
        15840,
        r#"<w:type w:val="continuous"/><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3954" w:space="2870"/><w:col w:w="3256"/></w:cols>"#,
        &body,
    );
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    let fragments: Vec<_> = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "paragraph")
        .collect();
    assert_eq!(fragments.len(), 2);
    assert!((fragments[0]["width"].as_f64().unwrap() - 3256.0 / 15.0).abs() < 0.01);
    assert!((fragments[1]["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
    assert_eq!(fragments[1]["x"], 72.0);
    assert!(fragments[1]["y"].as_f64().unwrap() > fragments[0]["y"].as_f64().unwrap());
}

fn staggered_row(exact: bool) -> Vec<u8> {
    let cells: String = [("left", 90, 240), ("right", 85, 260)]
        .into_iter()
        .enumerate()
        .map(|(cell, (prefix, count, spacing))| {
            let paragraphs: String = (0..count)
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7100_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("{prefix}{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let height = if exact {
        r#"<w:trPr><w:trHeight w:val="7200" w:hRule="exact"/></w:trPr>"#
    } else {
        ""
    };
    fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>{height}{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ))
}

fn staggered_row_layout(exact: bool) -> (Value, Value) {
    let (engine, request) = fixture::laid_out(&staggered_row(exact), 9403);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    (
        serde_json::from_str(&output).unwrap(),
        serde_json::from_str(&display).unwrap(),
    )
}

#[test]
fn centred_cells_with_staggered_lines_split_within_the_page_body() {
    let (output, display) = staggered_row_layout(false);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let mut fragment_count = 0;
    let mut texts = [String::new(), String::new()];
    for (page, painted) in pages.iter().zip(display["pages"].as_array().unwrap()) {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        if fragments.is_empty() {
            continue;
        }
        assert_eq!(fragments.len(), 1);
        fragment_count += 1;
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        let clips = fragment["cellClips"].as_array().unwrap();
        assert_eq!(clips.len(), 2);
        assert_ne!(clips[0]["bottom"], clips[1]["bottom"]);
        for row in fragment["rowStart"].as_u64().unwrap()..fragment["rowEnd"].as_u64().unwrap() {
            assert!(
                painted["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|primitive| {
                        primitive["cell"]["row"].as_u64() == Some(row)
                            && primitive["text"]
                                .as_str()
                                .is_some_and(|text| !text.is_empty())
                    }),
                "row {row} has no line text in {fragment}"
            );
        }
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            let measured_cell = &output["measured"][0]["measure"]["rows"][0]["cells"][cell];
            let line = &measured_cell["blocks"][0]["lines"][0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_bottom = line_top + line["lineHeight"].as_f64().unwrap();
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
    }
    assert!(fragment_count > 1);
    for (cell, (prefix, count)) in [("left", 90), ("right", 85)].into_iter().enumerate() {
        assert_eq!(
            texts[cell],
            (0..count)
                .map(|index| format!("{prefix}{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn a_short_bottom_aligned_cell_stays_on_the_last_row_fragment() {
    assert_short_aligned_cells(&["top", "top", "bottom"]);
}

#[test]
fn short_centred_and_bottom_cells_keep_their_whole_row_alignment() {
    assert_short_aligned_cells(&["top", "top", "center", "bottom"]);
}

fn assert_short_aligned_cells(alignments: &[&str]) {
    let width = 8640 / alignments.len();
    let counts = [225, 216, 1, 1];
    let cells: String = alignments
        .iter()
        .enumerate()
        .map(|(cell, alignment)| {
            let spacing = if cell == 1 { 250 } else { 240 };
            let paragraphs: String = (0..counts[cell])
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7200_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("cell{cell}line{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{alignment}"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let grid = format!(r#"<w:gridCol w:w="{width}"/>"#).repeat(alignments.len());
    let bytes = fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ));
    let (engine, request) = fixture::laid_out(&bytes, 9404);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    let display: Value = serde_json::from_str(&display).unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 5);
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let row_height = output["measured"][0]["measure"]["rows"][0]["height"]
        .as_f64()
        .unwrap();
    let mut consumed = 0.0;
    let mut texts = vec![String::new(); alignments.len()];
    for (page_index, (page, painted)) in pages
        .iter()
        .zip(display["pages"].as_array().unwrap())
        .enumerate()
    {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        assert_eq!(fragments.len(), 1);
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            if cell > 1 {
                let expected_page = if alignments[cell] == "center" { 2 } else { 4 };
                assert_eq!(page_index, expected_page, "{text}");
            }
            let line = &output["measured"][0]["measure"]["rows"][0]["cells"][cell]["blocks"][0]["lines"]
                [0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_height = line["lineHeight"].as_f64().unwrap();
            let line_bottom = line_top + line_height;
            if cell > 1 {
                let slack = row_height - line_height;
                let offset = if alignments[cell] == "center" {
                    slack / 2.0
                } else {
                    slack
                };
                let expected_top = top + offset - consumed;
                assert!((line_top - expected_top).abs() < 0.02, "{text}: {line_top}");
            }
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
        consumed += fragment["height"].as_f64().unwrap();
    }
    for (cell, text) in texts.iter().enumerate() {
        assert_eq!(
            *text,
            (0..counts[cell])
                .map(|index| format!("cell{cell}line{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn an_exact_height_row_with_staggered_lines_stays_whole_and_clips() {
    let (output, display) = staggered_row_layout(true);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments: Vec<_> = pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(fragments.len(), 1);
    let fragment = fragments[0];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["height"], 480.0);
    assert_ne!(fragment["carriedToNext"], true);
    assert!(fragment["cellClips"].is_null());
    let top = fragment["y"].as_f64().unwrap();
    let bottom = top + fragment["height"].as_f64().unwrap();
    let mut text = String::new();
    for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
        let Some(run) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
            continue;
        };
        let clip = &primitive["clipGroup"]["clip"];
        assert!((clip["y"].as_f64().unwrap() - top).abs() < 0.02);
        assert!(clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap() <= bottom + 0.02);
        text.push_str(run);
    }
    assert!(text.contains("left000"));
    assert!(text.contains("right000"));
    assert!(!text.contains("left089"));
    assert!(!text.contains("right084"));
}

#[test]
fn an_edit_before_page_breaks_repaginates_only_its_own_page() {
    let paragraph = |id: u32| fixture::p(&format!("{id:08X}"), &fixture::r("text on its own page"));
    let page_break =
        |id: u32| fixture::p(&format!("{id:08X}"), r#"<w:r><w:br w:type="page"/></w:r>"#);
    let body: String = (0..6_u32)
        .map(|index| {
            format!(
                "{}{}",
                paragraph(0x7000_0000 + index * 2),
                page_break(0x7000_0001 + index * 2)
            )
        })
        .collect();
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9401);
    let before = engine.stats();
    assert!(before.retained_pages >= 6);

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 0),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );
    assert_eq!(after.rebuilt_pages, 1, "later pages converge");

    let fresh = EngineSession::new(9402);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        layout
    );
}
