use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{BlockExtent, LayoutBlock, TableBlock, TableExtent};
use serde_json::json;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn lower(properties: &str, span: u32) -> Vec<LayoutBlock> {
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="300"/><w:gridCol w:w="1200"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:trPr>{properties}</w:trPr><w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>ECM Description Title</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#
    );
    lower_xml(xml)
}

fn lower_xml(xml: String) -> Vec<LayoutBlock> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", xml),
    ];
    let bytes = ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap();
    let doc = EditingDoc::new(75140);
    seed_from_docx(&doc, &bytes).unwrap();
    let before = doc.encode_state_as_update_v1();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    assert_eq!(doc.encode_state_as_update_v1(), before);
    blocks
}

fn cell(properties: &str, text: &str) -> String {
    format!(
        r#"<w:tc><w:tcPr>{properties}</w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:tc>"#
    )
}

fn row(properties: &str, cells: &str) -> String {
    format!("<w:tr><w:trPr>{properties}</w:trPr>{cells}</w:tr>")
}

fn measured_fixed_table(grid: &[u32], width: u32, rows: &str) -> (TableBlock, TableExtent) {
    let grid = grid
        .iter()
        .map(|width| format!(r#"<w:gridCol w:w="{width}"/>"#))
        .collect::<String>();
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl></w:body></w:document>"#
    );
    let mut blocks = lower_xml(xml);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let config = MeasurementConfig {
        font_chains: [("arial|0|0".to_owned(), vec![font])].into(),
        defaults: json!({"fontFamily": "Arial", "fontSize": 12}),
        authoritative_shaping: true,
        ..Default::default()
    };
    let measures = measure_blocks(&mut blocks, f64::from(width) / 15.0, &config).unwrap();
    let LayoutBlock::Table(table) = &blocks[0] else {
        panic!("table");
    };
    let BlockExtent::Table(measure) = &measures[0] else {
        panic!("table extent");
    };
    (table.clone(), measure.clone())
}

fn assert_widths(actual: &[f64], expected: &[f64]) {
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(expected) {
        assert!((actual - expected).abs() < 1e-9, "{actual} != {expected}");
    }
}

#[test]
fn bridge_preserves_leading_grid_columns_and_drops_legacy_synthetic_cell_widths() {
    docx_layout::with_private_measure_fonts(|| {
        let mut blocks = lower(r#"<w:gridBefore w:val="1"/>"#, 2);
        let LayoutBlock::Table(table) = &blocks[0] else {
            panic!("table");
        };
        assert_eq!(table.rows[0].grid_before, Some(1));
        assert_eq!(table.rows[0].cells[0].width_value, None);
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let config = MeasurementConfig {
            font_chains: [("arial|0|0".to_owned(), vec![font])].into(),
            defaults: json!({"fontFamily": "Arial", "fontSize": 12}),
            authoritative_shaping: true,
            ..Default::default()
        };
        let measures = measure_blocks(&mut blocks, 220.0, &config).unwrap();
        let BlockExtent::Table(table) = &measures[0] else {
            panic!("table extent");
        };
        assert_eq!(table.column_widths, [20.0, 80.0, 120.0]);
        assert_eq!(table.rows[0].cells[0].width, 200.0);
        let BlockExtent::Paragraph(paragraph) = &table.rows[0].cells[0].blocks[0] else {
            panic!("paragraph extent");
        };
        assert_eq!(paragraph.lines.len(), 1);
        assert!(
            paragraph
                .lines
                .iter()
                .all(|line| line.synthetic_fallback != Some(true))
        );
    });
}

#[test]
fn bridge_preserves_trailing_grid_columns() {
    let blocks = lower(r#"<w:gridAfter w:val="1"/>"#, 2);
    let LayoutBlock::Table(table) = &blocks[0] else {
        panic!("table");
    };
    assert_eq!(table.rows[0].grid_before, None);
    assert_eq!(table.rows[0].grid_after, Some(1));
    assert_eq!(docx_layout::table_grid::count_table_columns(table), 3);
}

#[test]
fn bridge_preserves_both_skips_and_their_preferred_width_units() {
    let blocks = lower(
        concat!(
            r#"<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>"#,
            r#"<w:wBefore w:w="300" w:type="dxa"/><w:wAfter w:w="500" w:type="pct"/>"#,
        ),
        1,
    );
    let LayoutBlock::Table(table) = &blocks[0] else {
        panic!("table");
    };
    let row = &table.rows[0];
    assert_eq!((row.grid_before, row.grid_after), (Some(1), Some(1)));
    let before = row.width_before.as_ref().unwrap();
    let after = row.width_after.as_ref().unwrap();
    assert_eq!(
        (before.value, before.r#type.as_deref()),
        (Some(300.0), Some("dxa"))
    );
    assert_eq!(
        (after.value, after.r#type.as_deref()),
        (Some(500.0), Some("pct"))
    );
}

#[test]
fn bridge_leaves_unskipped_rows_without_new_layout_properties() {
    for properties in [
        "",
        r#"<w:wBefore w:w="300" w:type="dxa"/><w:wAfter w:w="500" w:type="pct"/>"#,
    ] {
        let blocks = lower(properties, 3);
        let LayoutBlock::Table(table) = &blocks[0] else {
            panic!("table");
        };
        let row = &table.rows[0];
        assert_eq!((row.grid_before, row.grid_after), (None, None));
        assert_eq!((&row.width_before, &row.width_after), (&None, &None));
        assert_eq!(
            table.column_widths.as_deref(),
            Some([20.0, 80.0, 120.0].as_slice())
        );
    }
}

#[test]
fn imported_fixed_layout_keeps_cell_preferences_with_grid_before() {
    docx_layout::with_private_measure_fonts(|| {
        let cells = ["First", "Second", "Third"]
            .map(|text| cell(r#"<w:tcW w:w="3000" w:type="dxa"/>"#, text));
        let full = row("", &cells.concat());
        let skipped = row(r#"<w:gridBefore w:val="1"/>"#, &cells[1..].concat());
        let (table, measure) =
            measured_fixed_table(&[15, 4485, 4500], 9000, &format!("{full}{skipped}"));
        assert_eq!(table.table_layout.as_deref(), Some("fixed"));
        assert_eq!(docx_layout::table_grid::count_table_columns(&table), 3);
        assert_widths(&measure.column_widths, &[200.0, 200.0, 200.0]);
        assert!((measure.total_width - 600.0).abs() < 1e-9);
        for cell in &measure.rows[0].cells {
            assert!((cell.width - 200.0).abs() < 1e-9);
            let BlockExtent::Paragraph(paragraph) = &cell.blocks[0] else {
                panic!("paragraph extent");
            };
            assert_eq!(paragraph.lines.len(), 1);
            assert!(
                paragraph
                    .lines
                    .iter()
                    .all(|line| line.synthetic_fallback != Some(true))
            );
        }
    });
}

#[test]
fn imported_single_cell_grid_after_does_not_add_a_grid_column() {
    docx_layout::with_private_measure_fonts(|| {
        let cells = ["First", "Second"]
            .map(|text| cell(r#"<w:tcW w:w="1500" w:type="dxa"/>"#, text));
        let full = row("", &cells.concat());
        let skipped = row(r#"<w:gridAfter w:val="1"/>"#, &cell("", "Last"));
        let (table, measure) =
            measured_fixed_table(&[1500, 1500], 3000, &format!("{full}{skipped}"));
        assert_eq!(table.rows[1].cells[0].col_span, Some(1.0));
        assert_eq!(table.rows[1].cells[0].width_value, None);
        assert_eq!(docx_layout::table_grid::count_table_columns(&table), 2);
        assert_eq!(measure.column_widths, [100.0, 100.0]);
        assert_eq!(measure.total_width, 200.0);
        assert_eq!(measure.rows[0].cells[0].width, 100.0);
        assert_eq!(measure.rows[0].cells[1].width, 100.0);
        assert_eq!(measure.rows[1].cells[0].width, 100.0);
    });
}

#[test]
fn mixed_skipped_rows_keep_main_grid_and_unskipped_widths() {
    docx_layout::with_private_measure_fonts(|| {
        let cells = ["First", "Second", "Third"]
            .map(|text| cell(r#"<w:tcW w:w="3000" w:type="dxa"/>"#, text));
        let full = row("", &cells.concat());
        let leading = row(r#"<w:gridBefore w:val="1"/>"#, &cells[1..].concat());
        let trailing = row(r#"<w:gridAfter w:val="1"/>"#, &cells[..2].concat());
        let both = row(
            r#"<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>"#,
            &cell("", "Middle"),
        );
        let rows = format!("{full}{leading}{trailing}{both}{full}");
        let main_rows = rows
            .replace(r#"<w:gridBefore w:val="1"/>"#, "")
            .replace(r#"<w:gridAfter w:val="1"/>"#, "");
        let (main_table, main) = measured_fixed_table(&[15, 4485, 4500], 9000, &main_rows);
        let (table, measure) = measured_fixed_table(&[15, 4485, 4500], 9000, &rows);
        assert_eq!(docx_layout::table_grid::count_table_columns(&main_table), 3);
        assert_eq!(docx_layout::table_grid::count_table_columns(&table), 3);
        assert_widths(&main.column_widths, &[200.0, 200.0, 200.0]);
        assert_eq!(measure.column_widths, main.column_widths);
        assert_eq!(measure.total_width, main.total_width);
        assert_eq!(table.rows[3].cells[0].col_span, Some(1.0));
        assert_eq!(table.rows[3].cells[0].width_value, None);
        for row in [0, 4] {
            let widths: Vec<_> = measure.rows[row].cells.iter().map(|cell| cell.width).collect();
            let main_widths: Vec<_> = main.rows[row].cells.iter().map(|cell| cell.width).collect();
            assert_eq!(widths, main_widths);
        }
    });
}
