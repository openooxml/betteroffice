#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use docx_layout::cell_layout::layout_cell_content;
use docx_layout::types::{BlockExtent, LayoutBlock, MeasuredBlock};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const AUTO: &str = r#"w:beforeAutospacing="1" w:afterAutospacing="1""#;

fn paragraph(spacing: &str, properties: &str) -> String {
    format!(
        r#"<w:p><w:pPr>{properties}<w:spacing w:line="240" w:lineRule="exact" {spacing}/></w:pPr><w:r><w:t>Sample</w:t></w:r></w:p>"#
    )
}

fn numbered(id: u32, level: u32) -> String {
    format!(r#"<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{id}"/></w:numPr>"#)
}

fn document(body: &str, optout: bool) -> Vec<u8> {
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        let xml = String::from_utf8(bytes.clone()).unwrap();
        *bytes = match name.as_str() {
            "word/document.xml" => format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            ),
            "word/_rels/document.xml.rels" => xml.replace(
                "</Relationships>",
                r#"<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#,
            ),
            "[Content_Types].xml" => xml.replace(
                "</Types>",
                r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#,
            ),
            _ => xml,
        }
        .into_bytes();
    }
    let levels: String = (0..2)
        .map(|level| format!(
            r#"<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>"#
        ))
        .collect();
    parts.push((
        "word/numbering.xml".to_owned(),
        format!(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0">{levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        )
        .into_bytes(),
    ));
    let compat = if optout {
        "<w:doNotUseHTMLParagraphAutoSpacing/>"
    } else {
        ""
    };
    parts.push((
        "word/settings.xml".to_owned(),
        format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>{compat}</w:compat></w:settings>"#
        )
        .into_bytes(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn render(body: &str, optout: bool) -> Value {
    render_with_page_height(body, optout, None)
}

fn render_with_page_height(body: &str, optout: bool, height: Option<f64>) -> Value {
    docx_layout::with_private_measure_fonts(|| {
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let bytes = document(body, optout);
        let engine = EngineSession::new(76403);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let mut request = fixture::region_request(&engine, &bytes, font);
        request["measurement"]["defaults"] = json!({"fontFamily": "Arial", "fontSize": 10});
        if let Some(height) = height {
            request["regions"]["sections"] = json!([{
                "pageSize": {"w": 816, "h": height},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
            }]);
        }
        serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap()
    })
}

fn close(actual: f64, expected: f64) {
    assert!((actual - expected).abs() < 0.01, "{actual} != {expected}");
}

fn body_gap(output: &Value, index: usize, twips: f64) {
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments = pages[0]["fragments"].as_array().unwrap();
    close(
        fragments[index + 1]["y"].as_f64().unwrap()
            - fragments[index]["y"].as_f64().unwrap()
            - output["measured"][index]["measure"]["totalHeight"]
                .as_f64()
                .unwrap(),
        twips / 15.0,
    );
}

#[test]
fn auto_sides_collapse_to_280_twips() {
    body_gap(&render(&paragraph(AUTO, "").repeat(2), false), 0, 280.0);
}

#[test]
fn mixed_explicit_and_auto_sides_take_the_maximum() {
    let body = format!(
        "{}{}{}{}",
        paragraph(r#"w:after="200""#, ""),
        paragraph(r#"w:beforeAutospacing="1""#, ""),
        paragraph(r#"w:afterAutospacing="1""#, ""),
        paragraph(r#"w:before="570""#, ""),
    );
    let output = render(&body, false);
    body_gap(&output, 0, 280.0);
    body_gap(&output, 2, 570.0);
}

#[test]
fn first_body_auto_before_is_zero_and_explicit_before_survives() {
    let output = render(&paragraph(AUTO, ""), false);
    close(
        output["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap(),
        96.0,
    );
    let explicit = render(&paragraph(r#"w:before="570""#, ""), false);
    close(
        explicit["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap(),
        96.0 + 570.0 / 15.0,
    );
}

#[test]
fn same_numbering_suppresses_auto_across_levels_and_non_list_restores_it() {
    let body = format!(
        "{}{}{}{}",
        paragraph(AUTO, &numbered(1, 0)),
        paragraph(AUTO, &numbered(1, 1)),
        paragraph(AUTO, ""),
        paragraph(AUTO, &numbered(1, 0)),
    );
    let output = render(&body, false);
    body_gap(&output, 0, 0.0);
    body_gap(&output, 1, 280.0);
    body_gap(&output, 2, 280.0);
}

#[test]
fn same_numbering_preserves_explicit_sides_and_distinct_ids_keep_auto() {
    for (after, before, expected) in [
        (r#"w:after="200""#, r#"w:beforeAutospacing="1""#, 200.0),
        (r#"w:afterAutospacing="1""#, r#"w:before="570""#, 570.0),
    ] {
        let body = format!(
            "{}{}{}",
            paragraph(after, &numbered(1, 0)),
            paragraph(before, &numbered(1, 1)),
            paragraph(AUTO, &numbered(2, 0)),
        );
        let output = render(&body, false);
        body_gap(&output, 0, expected);
        body_gap(&output, 1, 280.0);
    }
}

#[test]
fn optout_cell_edges_and_contextual_spacing_keep_their_contracts() {
    body_gap(&render(&paragraph(AUTO, "").repeat(2), false), 0, 280.0);
    let body = format!(
        "{}{}",
        paragraph(r#"w:afterAutospacing="1""#, ""),
        paragraph("", ""),
    );
    body_gap(&render(&body, true), 0, 200.0);
    let body = paragraph(AUTO, "<w:contextualSpacing/>").repeat(2);
    body_gap(&render(&body, false), 0, 0.0);
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        paragraph(AUTO, ""),
    );
    let output = render(&body, false);
    let measured: MeasuredBlock = serde_json::from_value(output["measured"][0].clone()).unwrap();
    let (LayoutBlock::Table(table), BlockExtent::Table(extent)) =
        (&measured.block, &measured.measure)
    else {
        panic!("expected table");
    };
    let cell = &table.rows[0].cells[0];
    let measured_cell = &extent.rows[0].cells[0];
    let placement = layout_cell_content(Some(&cell.blocks), Some(&measured_cell.blocks), 0.0);
    close(placement.line_tops[0][0], 0.0);
    let BlockExtent::Paragraph(paragraph) = &measured_cell.blocks[0] else {
        panic!("expected paragraph");
    };
    close(extent.rows[0].height, paragraph.total_height);
}

#[test]
fn same_list_cell_measurement_placement_and_row_breaks_use_resolved_spacing() {
    let content = format!(
        "{}{}{}",
        paragraph(AUTO, &numbered(1, 0)),
        paragraph(AUTO, &numbered(1, 1)),
        paragraph(AUTO, ""),
    );
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>"#,
    );
    let output = render(&body, false);
    let measured: MeasuredBlock = serde_json::from_value(output["measured"][0].clone()).unwrap();
    let (LayoutBlock::Table(table), BlockExtent::Table(extent)) =
        (&measured.block, &measured.measure)
    else {
        panic!("expected table");
    };
    let cell = &table.rows[0].cells[0];
    let measured_cell = &extent.rows[0].cells[0];
    let placement = layout_cell_content(Some(&cell.blocks), Some(&measured_cell.blocks), 0.0);
    close(placement.line_tops[0][0], 0.0);
    close(placement.line_tops[1][0], 16.0);
    close(placement.line_tops[2][0], 32.0 + 280.0 / 15.0);
    close(extent.rows[0].height, 48.0 + 280.0 / 15.0);
    let breaks = docx_layout::table_row_break::build_table_row_break_info(table, extent);
    assert!(
        breaks.break_offsets[0]
            .iter()
            .any(|offset| (*offset - 32.0).abs() < 0.01)
    );
    close(*breaks.row_tops.last().unwrap(), 48.0 + 280.0 / 15.0);
}

#[test]
fn changing_list_identity_recomputes_both_boundary_sides_on_relayout() {
    docx_layout::with_private_measure_fonts(|| {
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let body = paragraph(AUTO, &numbered(1, 0)).repeat(3);
        let bytes = document(&body, false);
        let engine = EngineSession::new(76404);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let request = fixture::region_request(&engine, &bytes, font).to_string();
        let initial: Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        body_gap(&initial, 0, 0.0);
        body_gap(&initial, 1, 0.0);
        let id = engine.doc().paragraphs("body").unwrap()[1].para_id.clone();
        for (num_id, expected) in [(2, 280.0), (0, 280.0), (1, 0.0)] {
            let num_pr = yrs::Any::Map(std::sync::Arc::new(std::collections::HashMap::from([
                ("numId".to_owned(), yrs::Any::Number(f64::from(num_id))),
                ("ilvl".to_owned(), yrs::Any::Number(0.0)),
            ])));
            engine
                .doc()
                .set_paragraph_attr(&id, "numPr", num_pr)
                .unwrap();
            let actual: Value =
                serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                    .unwrap();
            body_gap(&actual, 0, expected);
            body_gap(&actual, 1, expected);
            let fresh = EngineSession::new(76405);
            fresh
                .doc()
                .apply_update_v1(&engine.doc().encode_state_as_update_v1())
                .unwrap();
            let cold: Value =
                serde_json::from_str(&fresh.layout_document_with_regions_json(&request).unwrap())
                    .unwrap();
            assert_eq!(actual["layout"], cold["layout"]);
        }
    });
}

#[test]
fn keep_group_budget_uses_suppressed_list_spacing() {
    let body = format!(
        "{}{}{}",
        paragraph("", ""),
        paragraph(
            r#"w:afterAutospacing="1""#,
            &format!("{}<w:keepNext/>", numbered(1, 0)),
        ),
        paragraph(r#"w:beforeAutospacing="1""#, &numbered(1, 1)),
    );
    let output = render_with_page_height(&body, false, Some(240.0));
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    body_gap(&output, 0, 0.0);
    body_gap(&output, 1, 0.0);
}
