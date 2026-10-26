#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_layout::types::{BlockExtent, TableExtent};
use serde_json::Value;

fn measured_table(layout: &str, no_wrap: &str, text: &str, page_width: u32) -> TableExtent {
    let rows: String = (0..3)
        .map(|row| {
            let cells: String = (0..2)
                .map(|column| {
                    format!(
                        r#"<w:tc><w:tcPr><w:tcW w:w="641" w:type="dxa"/>{no_wrap}<w:tcBorders><w:top w:val="single" w:sz="6"/><w:bottom w:val="single" w:sz="6"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:p w14:paraId="{:08X}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:tc>"#,
                        0x7500_0000 + row * 2 + column,
                    )
                })
                .collect();
            format!(r#"<w:tr><w:trPr><w:trHeight w:val="240"/></w:trPr>{cells}</w:tr>"#)
        })
        .collect();
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="1282" w:type="dxa"/>{layout}<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="641"/><w:gridCol w:w="641"/></w:tblGrid>{rows}</w:tbl><w:sectPr><w:pgSz w:w="{page_width}" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#,
    );
    measured_body(&body).0
}

fn measured_body(body: &str) -> (TableExtent, Value) {
    let bytes = fixture::with_body(body);
    let (engine, request) = fixture::laid_out(&bytes, 9501);
    let mut output: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    let request: Value = serde_json::from_str(&request).unwrap();
    output["fontChains"] = request["measurement"]["fontChains"].clone();
    let table = serde_json::from_value(output["measured"][0]["measure"].clone()).unwrap();
    (table, output)
}

fn assert_lines(table: &TableExtent, count: usize) {
    assert_eq!(table.rows.len(), 3);
    for row in &table.rows {
        for cell in &row.cells {
            let BlockExtent::Paragraph(paragraph) = &cell.blocks[0] else {
                panic!("expected cell paragraph");
            };
            assert_eq!(paragraph.lines.len(), count);
            assert!(
                paragraph
                    .lines
                    .iter()
                    .all(|line| line.synthetic_fallback != Some(true))
            );
            let line_height: f64 = paragraph.lines.iter().map(|line| line.line_height).sum();
            assert!((paragraph.total_height - line_height).abs() < 0.01);
            assert!(cell.height + 0.01 >= paragraph.total_height);
            assert!(row.height + 0.01 >= cell.height);
        }
    }
}

#[test]
fn autofit_nowrap_numeric_cells_wrap_when_constrained() {
    docx_layout::with_private_measure_fonts(|| {
        for layout in ["", r#"<w:tblLayout w:type="autofit"/>"#] {
            let table = measured_table(layout, "<w:noWrap/>", "3,39", 2722);
            assert_lines(&table, 2);
            let control = measured_table(layout, "", "3,39", 2722);
            assert_lines(&control, 2);
        }
    });
}

#[test]
fn autofit_nowrap_cells_wrap_at_spaces_when_constrained_and_grow_with_room() {
    docx_layout::with_private_measure_fonts(|| {
        let table = measured_table("", "<w:noWrap/>", "3 39", 2722);
        assert_lines(&table, 2);
        let growing = measured_table("", "<w:noWrap/>", "3 39", 12240);
        assert_lines(&growing, 1);
        assert!(growing.rows[0].cells[0].width > 641.0 / 15.0);
        let control = measured_table("", "", "3 39", 12240);
        assert_lines(&control, 2);
    });
}

#[test]
fn nowrap_preserves_explicit_line_breaks() {
    docx_layout::with_private_measure_fonts(|| {
        let table = measured_table("", "<w:noWrap/>", r#"3,39</w:t><w:br/><w:t>3,39"#, 12240);
        assert_lines(&table, 2);
    });
}

#[test]
fn fixed_layout_and_explicitly_disabled_nowrap_keep_wrapping() {
    docx_layout::with_private_measure_fonts(|| {
        let fixed = measured_table(
            r#"<w:tblLayout w:type="fixed"/>"#,
            "<w:noWrap/>",
            "3,39",
            12240,
        );
        assert_lines(&fixed, 2);
        let disabled = measured_table("", r#"<w:noWrap w:val="0"/>"#, "3,39", 2722);
        assert_lines(&disabled, 2);
    });
}

fn full_width_body(no_wrap: &str, spacing: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="3600" w:type="dxa"/><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/>{no_wrap}</w:tcPr><w:p w14:paraId="75000100"><w:pPr>{spacing}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Aberdeen City Council</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p w14:paraId="75000101"/></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="5040" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#,
    )
}

#[test]
fn full_width_autofit_nowrap_wraps_and_paints_inside_the_cell() {
    docx_layout::with_private_measure_fonts(|| {
        let (table, output) = measured_body(&full_width_body(
            "<w:noWrap/>",
            r#"<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>"#,
        ));
        assert!((table.total_width - 240.0).abs() < 0.01);
        let cell = &table.rows[0].cells[0];
        assert!((cell.width - 120.0).abs() < 0.01);
        let BlockExtent::Paragraph(paragraph) = &cell.blocks[0] else {
            panic!("expected cell paragraph");
        };
        assert_eq!(paragraph.lines.len(), 2);
        assert!(
            paragraph
                .lines
                .iter()
                .all(|line| line.synthetic_fallback != Some(true))
        );
        let paragraph_id = output["measured"][0]["block"]["rows"][0]["cells"][0]["blocks"][0]["id"]
            .as_str()
            .unwrap();
        let display = docx_layout::build_display_list_value(&output.to_string()).unwrap();
        let display = serde_json::to_value(display).unwrap();
        let cell_right = 48.0 - 108.0 / 15.0 + cell.width;
        let mut glyphs = 0;
        for primitive in display["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["primitives"].as_array().unwrap())
            .filter(|primitive| primitive["blockKey"] == paragraph_id)
        {
            if primitive["kind"] == "glyphRun" {
                for glyph in primitive["glyphs"].as_array().unwrap() {
                    let right = glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap();
                    assert!(right <= cell_right + 0.01, "glyph right {right}");
                    glyphs += 1;
                }
            }
        }
        assert!(glyphs > 0);
    });
}

#[test]
fn nowrap_row_height_includes_wrapped_content_spacing_and_margins() {
    docx_layout::with_private_measure_fonts(|| {
        let spacing = r#"<w:spacing w:before="120" w:after="180" w:line="360" w:lineRule="auto"/>"#;
        let (table, _) = measured_body(&full_width_body("<w:noWrap/>", spacing));
        let (control, _) = measured_body(&full_width_body("", spacing));
        let row = &table.rows[0];
        let cell = &row.cells[0];
        let BlockExtent::Paragraph(paragraph) = &cell.blocks[0] else {
            panic!("expected cell paragraph");
        };
        assert_eq!(paragraph.lines.len(), 2);
        let lines: f64 = paragraph.lines.iter().map(|line| line.line_height).sum();
        assert!((paragraph.total_height - lines - 300.0 / 15.0).abs() < 0.01);
        assert!((cell.height - paragraph.total_height - 150.0 / 15.0).abs() < 0.01);
        assert!((row.height - cell.height).abs() < 0.01);
        assert!((row.height - control.rows[0].height).abs() < 0.01);
        assert!(row.height > paragraph.lines[0].line_height + 30.0);
    });
}

#[test]
fn nowrap_fractional_floor_reflows_auto_but_preserves_absolute_neighbor() {
    docx_layout::with_private_measure_fonts(|| {
        for (preference, expected_lines) in [
            (r#"<w:tcW w:w="0" w:type="auto"/>"#, 1),
            (r#"<w:tcW w:w="8400" w:type="dxa"/>"#, 2),
        ] {
            let paragraph = |id: &str, text: &str| {
                format!(
                    r#"<w:p w14:paraId="{id}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#,
                )
            };
            let first = paragraph("75000200", "ID 001");
            let second = paragraph(
                "75000201",
                "Row 001 alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu",
            );
            let body = |no_wrap: &str| {
                format!(
                    r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="720"/><w:gridCol w:w="720"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>{no_wrap}</w:tcPr>{first}</w:tc><w:tc><w:tcPr>{preference}</w:tcPr>{second}</w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#,
                )
            };
            let (ordinary, _) = measured_body(&body(""));
            let (table, _) = measured_body(&body("<w:noWrap/>"));
            let BlockExtent::Paragraph(paragraph) = &table.rows[0].cells[0].blocks[0] else {
                panic!("expected cell paragraph");
            };
            assert_eq!(paragraph.lines.len(), expected_lines);
            assert!(
                paragraph
                    .lines
                    .iter()
                    .all(|line| line.synthetic_fallback != Some(true))
            );
            assert!((table.total_width - 9026.0 / 15.0).abs() < 0.01);
            if expected_lines == 1 {
                let growth = table.rows[0].cells[0].width - ordinary.rows[0].cells[0].width;
                assert!(growth > 0.1 && growth < 0.2, "growth {growth}");
                assert!((table.rows[0].cells[0].width - 47.140625).abs() < 0.01);
                assert!(table.rows[0].cells[1].width < ordinary.rows[0].cells[1].width);
            } else {
                for (cell, control) in table.rows[0].cells.iter().zip(&ordinary.rows[0].cells) {
                    assert!((cell.width - control.width).abs() < 0.001);
                }
            }
        }
    });
}
