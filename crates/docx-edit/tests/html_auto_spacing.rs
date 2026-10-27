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
    document_with_bands(body, optout, None, None)
}

fn document_with_bands(
    body: &str,
    optout: bool,
    header: Option<&str>,
    footer: Option<&str>,
) -> Vec<u8> {
    let mut parts = fixture::page_parts(false);
    let mut references = String::new();
    if header.is_some() {
        references.push_str(r#"<w:headerReference w:type="default" r:id="rIdHeader2"/>"#);
    }
    if footer.is_some() {
        references.push_str(r#"<w:footerReference w:type="default" r:id="rIdFooter1"/>"#);
    }
    for (name, bytes) in &mut parts {
        let xml = String::from_utf8(bytes.clone()).unwrap();
        *bytes = match name.as_str() {
            "word/document.xml" => format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr>{references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#
            ),
            "word/header2.xml" if header.is_some() => format!(
                r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}</w:hdr>"#,
                header.unwrap(),
            ),
            "word/footer1.xml" if footer.is_some() => format!(
                r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}</w:ftr>"#,
                footer.unwrap(),
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
    render_package(&document(body, optout), height)
}

fn render_package(bytes: &[u8], height: Option<f64>) -> Value {
    docx_layout::with_private_measure_fonts(|| {
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let engine = EngineSession::new(76403);
        seed_from_docx(engine.doc(), bytes).unwrap();
        let mut request = fixture::region_request(&engine, bytes, font);
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
    let height = fragments[index]["height"].as_f64().unwrap();
    close(height, 16.0);
    close(
        fragments[index + 1]["y"].as_f64().unwrap()
            - fragments[index]["y"].as_f64().unwrap()
            - height,
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

#[test]
fn first_header_auto_before_does_not_push_the_body_below_its_top_margin() {
    let header = paragraph(AUTO, "").replace("Sample", "");
    let bytes = document_with_bands(&paragraph("", ""), false, Some(&header), None);
    let output = render_package(&bytes, None);
    let variants = output["headersFooters"]["variants"].as_array().unwrap();
    assert!(
        variants
            .iter()
            .all(|variant| variant["rId"] == "rIdHeader2")
    );
    let header = &variants[0];
    assert_eq!(header["kind"], "header");
    let spacing = &header["measured"][0]["block"]["attrs"]["spacing"];
    close(spacing["before"].as_f64().unwrap(), 0.0);
    close(spacing["after"].as_f64().unwrap(), 280.0 / 15.0);
    close(header["flowHeight"].as_f64().unwrap(), 16.0 + 280.0 / 15.0);
    close(
        output["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap(),
        96.0,
    );
}

#[test]
fn first_header_auto_before_is_dropped_when_a_watermark_shape_leads_the_story() {
    let watermark = r#"<w:r xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="6096000" cy="2286000"/><wp:wrapNone/><wp:docPr id="1" name="Watermark"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6096000" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;
    let header = paragraph(AUTO, "").replace("<w:r><w:t>Sample</w:t></w:r>", watermark);
    let bytes = document_with_bands(&paragraph("", ""), false, Some(&header), None);
    let output = render_package(&bytes, None);
    let variants = output["headersFooters"]["variants"].as_array().unwrap();
    assert!(!variants.is_empty());
    for variant in variants {
        assert_eq!(variant["kind"], "header");
        close(variant["flowHeight"].as_f64().unwrap(), 16.0 + 280.0 / 15.0);
        let measured = variant["measured"].as_array().unwrap();
        let paragraph = measured
            .iter()
            .find(|entry| entry["block"]["kind"] == "paragraph")
            .unwrap();
        close(
            paragraph["block"]["attrs"]["spacing"]["before"]
                .as_f64()
                .unwrap(),
            0.0,
        );
        close(
            paragraph["block"]["attrs"]["spacing"]["after"]
                .as_f64()
                .unwrap(),
            280.0 / 15.0,
        );
    }
    close(
        output["layout"]["pages"][0]["fragments"][0]["y"]
            .as_f64()
            .unwrap(),
        96.0,
    );
}

#[test]
fn header_interior_and_footer_outer_auto_spacing_remain_280_twips() {
    let content = format!(
        "{}{}",
        paragraph(AUTO, "").replace("Sample", ""),
        paragraph(AUTO, ""),
    );
    let bytes = document_with_bands(&paragraph("", ""), false, Some(&content), Some(&content));
    let output = render_package(&bytes, None);
    let variants = output["headersFooters"]["variants"].as_array().unwrap();
    for (kind, before, gaps) in [("header", 0.0, 2.0), ("footer", 280.0 / 15.0, 3.0)] {
        let variant = variants
            .iter()
            .find(|variant| variant["kind"] == kind)
            .unwrap();
        close(
            variant["measured"][0]["block"]["attrs"]["spacing"]["before"]
                .as_f64()
                .unwrap(),
            before,
        );
        close(
            variant["measured"][1]["block"]["attrs"]["spacing"]["before"]
                .as_f64()
                .unwrap(),
            280.0 / 15.0,
        );
        close(
            variant["flowHeight"].as_f64().unwrap(),
            32.0 + gaps * 280.0 / 15.0,
        );
    }
}

#[test]
fn auto_before_is_zero_after_an_authored_page_break_and_explicit_before_survives() {
    for properties in ["<w:pageBreakBefore/>", ""] {
        let target = if properties.is_empty() {
            format!(
                r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact" {AUTO}/></w:pPr><w:r><w:br w:type="page"/></w:r><w:r><w:t>Sample</w:t></w:r></w:p>"#
            )
        } else {
            paragraph(AUTO, properties)
        };
        let body = format!("{}{}", paragraph(r#"w:after="0""#, ""), target);
        let output = render(&body, false);
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 2);
        close(pages[1]["fragments"][0]["y"].as_f64().unwrap(), 96.0);
    }
    let body = format!(
        "{}{}",
        paragraph(r#"w:after="0""#, ""),
        paragraph(r#"w:before="570""#, "<w:pageBreakBefore/>"),
    );
    let output = render(&body, false);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 2);
    close(
        pages[1]["fragments"][0]["y"].as_f64().unwrap(),
        96.0 + 570.0 / 15.0,
    );
}

#[test]
fn first_header_explicit_before_survives_auto_after() {
    let header = paragraph(r#"w:before="570" w:afterAutospacing="1""#, "");
    let bytes = document_with_bands(&paragraph("", ""), false, Some(&header), None);
    let output = render_package(&bytes, None);
    let header = &output["headersFooters"]["variants"][0];
    close(
        header["measured"][0]["block"]["attrs"]["spacing"]["before"]
            .as_f64()
            .unwrap(),
        570.0 / 15.0,
    );
    close(
        header["flowHeight"].as_f64().unwrap(),
        16.0 + (570.0 + 280.0) / 15.0,
    );
}
