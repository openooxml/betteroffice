#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_layout::types::{BlockExtent, TableExtent};
use serde_json::Value;

const CELL_LABELS: [&str; 3] = [
    "Alphabet Data Module",
    "Modules and Abstract Cluster",
    "Number Parameters Module",
];

fn label_table_body(borders: &str) -> String {
    let rows: String = CELL_LABELS
        .iter()
        .enumerate()
        .map(|(row, &label)| {
            let cells: String = (0..5)
                .map(|column| {
                    let preference = if column == 0 {
                        r#"<w:tcW w:w="0" w:type="auto"/>"#
                    } else {
                        r#"<w:tcW w:w="1800" w:type="dxa"/>"#
                    };
                    let text = if column == 0 { label } else { "2,000,000" };
                    format!(
                        r#"<w:tc><w:tcPr>{preference}<w:noWrap/>{borders}</w:tcPr><w:p w14:paraId="{:08X}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:tc>"#,
                        0x7500_0300 + row * 5 + column,
                    )
                })
                .collect();
            format!("<w:tr>{cells}</w:tr>")
        })
        .collect();
    let grid = r#"<w:gridCol w:w="1800"/>"#.repeat(5);
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl><w:sectPr><w:pgSz w:w="10440" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#,
    )
}

#[test]
fn retained_editor_paths_wrap_constrained_auto_nowrap_council_cells() {
    docx_layout::with_private_measure_fonts(|| {
        for borders in [
            "",
            r#"<w:tcBorders><w:left w:val="single" w:sz="4"/></w:tcBorders>"#,
        ] {
            for path in ["retained", "prefix", "exact-prefix", "sliced"] {
                docx_layout::clear_measure_fonts();
                let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
                let bytes = fixture::with_body(&label_table_body(borders));
                let engine = docx_edit::EngineSession::new(9502);
                docx_edit::seed_from_docx(engine.doc(), &bytes).unwrap();
                let request = fixture::region_request(&engine, &bytes, font);
                let input = request.to_string();
                match path {
                    "retained" => engine
                        .layout_document_with_regions_retained(&input)
                        .unwrap(),
                    "prefix" => {
                        engine
                            .layout_document_with_regions_prefix_retained_json(&input, 1)
                            .unwrap();
                    }
                    "exact-prefix" => {
                        engine
                            .layout_document_with_regions_exact_prefix_retained(&input, 1)
                            .unwrap();
                    }
                    "sliced" => {
                        let mut progress = engine.begin_region_layout(&input).unwrap();
                        while progress.layout_json.is_none() {
                            progress = engine.resume_region_layout(1).unwrap();
                        }
                    }
                    _ => unreachable!(),
                }
                let kernel: Value =
                    serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
                let table: TableExtent =
                    serde_json::from_value(kernel["measured"][0]["measure"].clone()).unwrap();
                assert!(
                    (table.total_width - 600.0).abs() < 0.01,
                    "{path}, {borders:?}"
                );
                assert_eq!(table.rows.len(), 3, "{path}, {borders:?}");
                engine.set_display_window(Some(0..0));
                let extras = serde_json::json!({
                    "fontChains": request["measurement"]["fontChains"],
                });
                engine
                    .build_display_list_frame(&extras.to_string(), 0)
                    .unwrap();
                engine.build_display_pages_frame(&[0], 0).unwrap();
                let display = engine
                    .with_display_list(|list| serde_json::to_value(list).unwrap())
                    .unwrap();
                let layout: Value =
                    serde_json::from_str(&engine.retained_layout_json().unwrap()).unwrap();
                let fragment = layout["layout"]["pages"][0]["fragments"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .find(|fragment| fragment["kind"] == "table")
                    .unwrap();
                let table_x = fragment["x"].as_f64().unwrap();
                for (row_index, row) in table.rows.iter().enumerate() {
                    let cell = &row.cells[0];
                    let BlockExtent::Paragraph(paragraph) = &cell.blocks[0] else {
                        panic!("{path}, {borders:?}, row {row_index}: expected label paragraph");
                    };
                    assert!(
                        (cell.width - 120.0).abs() < 0.01,
                        "{path}, {borders:?}, row {row_index}: cell width {}",
                        cell.width
                    );
                    assert!(
                        paragraph.lines.len() >= 2,
                        "{path}, {borders:?}, row {row_index}: expected wrapped lines"
                    );
                    for (line_index, line) in paragraph.lines.iter().enumerate() {
                        assert!(
                            line.synthetic_fallback != Some(true),
                            "{path}, {borders:?}, row {row_index}, line {line_index}: fallback"
                        );
                        let text = &CELL_LABELS[row_index][line.head_char..line.tail_char];
                        let trailing_start = line.head_char + text.trim_end().len();
                        let trailing_width: f64 = line
                            .cluster_advances
                            .as_ref()
                            .unwrap()
                            .iter()
                            .filter(|cluster| {
                                cluster.start_char.unwrap() as usize >= trailing_start
                            })
                            .map(|cluster| cluster.advance.unwrap())
                            .sum();
                        let visible_width = line.width - trailing_width;
                        assert!(
                            visible_width <= cell.width - 216.0 / 15.0 + 0.01,
                            "{path}, {borders:?}, row {row_index}, line {line_index}: \
                             width {}, visible {visible_width}, cell {}, fallback {:?}",
                            line.width,
                            cell.width,
                            line.synthetic_fallback
                        );
                        if row_index == 0 && line_index == 0 {
                            assert!(
                                line.width > cell.width - 216.0 / 15.0 + 0.01,
                                "{path}, {borders:?}, row {row_index}: trailing-space fixture"
                            );
                        }
                    }
                    let height: f64 = paragraph.lines.iter().map(|line| line.line_height).sum();
                    assert!(
                        cell.height + 0.01 >= height,
                        "{path}, {borders:?}, row {row_index}: cell height"
                    );
                    assert!(
                        row.height + 0.01 >= cell.height,
                        "{path}, {borders:?}, row {row_index}: row height"
                    );
                    assert!(
                        row.height > paragraph.lines[0].line_height + 1.0,
                        "{path}, {borders:?}, row {row_index}: wrapped row height"
                    );
                    let id = &kernel["measured"][0]["block"]["rows"][row_index]["cells"][0]["blocks"]
                        [0]["id"];
                    let mut glyphs = 0;
                    for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
                        if primitive["kind"] != "glyphRun" || primitive["blockKey"] != *id {
                            continue;
                        }
                        let visible_end = primitive["text"].as_str().unwrap().trim_end().len();
                        for glyph in primitive["glyphs"].as_array().unwrap() {
                            if glyph["cluster"].as_u64().unwrap() as usize >= visible_end {
                                continue;
                            }
                            let right =
                                glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap();
                            assert!(
                                right <= table_x + cell.width - 108.0 / 15.0 + 0.01,
                                "{path}, {borders:?}, row {row_index}: glyph right {right}"
                            );
                            glyphs += 1;
                        }
                    }
                    assert!(glyphs > 0, "{path}, {borders:?}, row {row_index}");
                }
            }
        }
    });
}

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
fn percentage_autofit_all_nowrap_keeps_authored_grid_when_content_cannot_fit() {
    docx_layout::with_private_measure_fonts(|| {
        let grid = [3672.0, 1792.0, 1960.0, 2096.0];
        let preferences = [2223, 824, 869, 1082];
        let labels = [
            "Module label",
            "Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu",
            "Beta gamma delta epsilon zeta eta theta iota kappa lambda mu alpha",
            "Gamma delta epsilon zeta eta theta iota kappa lambda mu alpha beta",
        ];
        let headers = [
            "Module label",
            "Module group indicator",
            "Module balance indicator",
            "Module group description",
        ];
        let rows: String = labels
            .iter()
            .enumerate()
            .map(|(row, label)| {
                let cells: String = preferences
                    .iter()
                    .enumerate()
                    .map(|(column, preference)| {
                        let text = if row == 0 {
                            headers[column]
                        } else if column == 0 {
                            label
                        } else if column == 3 {
                            "8888888888888"
                        } else {
                            "2.000"
                        };
                        format!(
                            r#"<w:tc><w:tcPr><w:tcW w:w="{preference}" w:type="pct"/><w:noWrap/></w:tcPr><w:p w14:paraId="{:08X}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:tc>"#,
                            0x7500_0400 + row * 4 + column,
                        )
                    })
                    .collect();
                format!("<w:tr>{cells}</w:tr>")
            })
            .collect();
        let columns: String = grid
            .iter()
            .map(|width| format!(r#"<w:gridCol w:w="{width}"/>"#))
            .collect();
        let body = format!(
            r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{columns}</w:tblGrid>{rows}</w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#,
        );
        let (table, _) = measured_body(&body);
        let grid_total: f64 = grid.iter().sum();
        assert!((table.total_width - grid_total / 15.0).abs() < 0.1);
        for row in &table.rows {
            for (cell, authored) in row.cells.iter().zip(grid) {
                let expected = table.total_width * authored / grid_total;
                assert!(
                    (cell.width - expected).abs() < 1.0,
                    "column width {}, authored proportion {expected}",
                    cell.width
                );
            }
        }
        let BlockExtent::Paragraph(header) = &table.rows[0].cells[1].blocks[0] else {
            panic!("expected header paragraph");
        };
        assert_eq!(header.lines.len(), 2);
        assert!(
            header
                .lines
                .iter()
                .all(|line| line.synthetic_fallback != Some(true))
        );
    });
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
        r#"<w:tbl><w:tblPr><w:tblW w:w="3600" w:type="dxa"/><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/>{no_wrap}</w:tcPr><w:p w14:paraId="75000100"><w:pPr>{spacing}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Alphabet Data Module</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p w14:paraId="75000101"/></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="5040" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#,
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
