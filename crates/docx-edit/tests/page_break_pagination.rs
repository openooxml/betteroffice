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
